VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBurnerDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -------------------------------------------------------------------------------
'
'  DVD Flick - A DVD authoring program
'  Copyright (C) 2006-2009  Dennis Meuwissen
'
'  This program is free software; you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation; either version 2 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software
'  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'
' -------------------------------------------------------------------------------
'
'   File purpose: Use SPTI and MMC commands to query a device for CD\DVD information.
'                 See http://www.t10.org/drafts.htm
'
Option Explicit
Option Compare Binary
Option Base 0


' SPTI
Private Const OPLEN_6 As Byte = 6
Private Const OPLEN_10 As Byte = 10
Private Const OPLEN_12 As Byte = 12

' SCSI opcodes
Private Const SCSIOP_INQUIRY As Byte = &H12
Private Const SCSIOP_MODE_SENSE As Byte = &H5A
Private Const SCSIOP_GETCONFIGURATION As Byte = &H46
Private Const SCSIOP_GETPERFORMANCE As Byte = &HAC

' Mode sense codepage
Private Const MODE_PAGE_CAPABILITIES As Byte = &H2A


' Properties
Public deviceName As String
Public deviceDriveChar As String
Public deviceBufferSize As Long
Public deviceReadSpeed As Long
Public deviceWriteSpeed As Long
Public deviceLoadMechanism As String
Public deviceVolumeLevels As Long

' Gathered information
Public deviceSpeeds As Dictionary
Public deviceCaps As Dictionary
Public deviceFeatures As Dictionary
Public deviceProfiles As Dictionary


' Get the details of a drive behind a drive character
Public Function getProperties(ByVal driveChar As String) As Boolean

    Dim Device As clsSPTIDevice
    
    
    ' Reset drive info
    deviceName = "Unknown"
    deviceDriveChar = driveChar
    Set deviceSpeeds = New Dictionary
    Set deviceCaps = New Dictionary
    Set deviceFeatures = New Dictionary
    Set deviceProfiles = New Dictionary

    ' Connect and retrieve basic info
    Set Device = New clsSPTIDevice
    If Not Device.connectTo(driveChar) Then Exit Function
    If Not getDeviceInquiry(Device) Then Exit Function
    
    ' Retrieve further info
    getDeviceCaps Device
    getDeviceFeatures Device
    getDeviceWriteSpeeds Device

    getProperties = True

End Function


' Read the write speeds according to GET PERFORMANCE
Private Function getDeviceWriteSpeeds(ByRef Device As clsSPTIDevice) As Boolean

    Dim dataLen As Long
    Dim Offset As Long
    Dim writeSpeed As Long
    Dim Index As Long
    
    
    ' SPTI command
    Device.clearCDB
    Device.setCDBByte 0, SCSIOP_GETPERFORMANCE
    Device.setCDBInt 8, 32
    Device.setCDBByte 10, &H3
    If Not Device.sendSCSIOp(OPLEN_12) Then Exit Function
   
    ' Parse write speeds
    dataLen = Device.getDataLong(0)
    If dataLen > Device.getIOBufferSize Then Exit Function
    Offset = 8
    Do
        Index = Index + 1
        
        writeSpeed = Device.getDataLong(Offset + 12)
        If Not deviceSpeeds.Exists("writeSpeed" & Index) Then deviceSpeeds.Add "writeSpeed" & Index, writeSpeed
        
        Offset = Offset + 16
    Loop Until Offset >= dataLen Or Offset > Device.getIOBufferSize

End Function


' Read all features supported by the device
Private Function getDeviceFeatures(ByRef Device As clsSPTIDevice) As Boolean

    Dim A As Long
    Dim Offset As Long
    Dim featNum As Integer, featSize As Byte, profileNum As Integer
    
    
    ' SPTI call
    Device.clearCDB
    Device.setCDBByte 0, SCSIOP_GETCONFIGURATION
    Device.setCDBInt 7, Device.getIOBufferSize
    If Not Device.sendSCSIOp(OPLEN_10) Then Exit Function
    
    ' Add list of features
    Offset = 8
    Do
        featNum = Device.getDataInt(Offset)
        featSize = Device.getDataByte(Offset + 3)

        ' Bail out if this feature was the last one
        If Offset > 8 And featNum = 0 And featSize = 0 Then Exit Do
        
        ' Add feature
        If mmcFeatures.Exists(featNum) Then
            If deviceFeatures.Exists(mmcFeatures(featNum)) Then Exit Function
            deviceFeatures.Add mmcFeatures(featNum), -1
        Else
            appLog.Add "Unknown device MMC feature " & Hex$(featNum), 1
            If deviceFeatures.Exists(Hex$(featNum)) Then Exit Function
            deviceFeatures.Add Hex$(featNum), -1
        End If
        
        ' Profile list feature
        If featNum = 0 Then
            For A = 0 To featSize - 4 Step 4
                If Offset + 4 + A > Device.getIOBufferSize Then Exit For
                profileNum = Device.getDataInt(Offset + 4 + A)
                
                ' Add profile
                If mmcProfiles.Exists(profileNum) Then
                    If deviceProfiles.Exists(mmcProfiles(profileNum)) Then Exit Function
                    deviceProfiles.Add mmcProfiles(profileNum), -1
                Else
                    appLog.Add "Unknown device MMC profile " & Hex$(profileNum), 1
                    If deviceProfiles.Exists(Hex$(profileNum)) Then Exit Function
                    deviceProfiles.Add Hex$(profileNum), -1
                End If
                
            Next A
        End If
        
        Offset = Offset + 4 + featSize
    Loop Until Offset > Device.getIOBufferSize

    getDeviceFeatures = True

End Function


' Read device's mode sense capabilities
Private Function getDeviceCaps(ByRef Device As clsSPTIDevice) As Boolean

    Dim A As Long
    Dim Data() As Byte
    Dim writeSpeeds As Integer
    Dim Offset As Long, writeSpeed As Integer
    Dim bitSet As String
    
    
    ' SPTI call
    Device.clearCDB
    Device.setCDBByte 0, SCSIOP_MODE_SENSE
    Device.setCDBByte 1, &H8
    Device.setCDBByte 2, MODE_PAGE_CAPABILITIES
    Device.setCDBInt 7, Device.getIOBufferSize

    If Device.sendSCSIOp(OPLEN_10) Then
    
        ' Parse caps if mode page was returned
        If Device.getDataByte(8) = MODE_PAGE_CAPABILITIES And Device.getBytesReturned >= 27 Then
            Device.getDataArray Data()
            
            With deviceCaps
                
                ' CD
                If (Data(10) And &H1) Then .Add "readCDR", -1
                If (Data(10) And &H2) Then .Add "readCDRW", -1
                If (Data(11) And &H1) Then .Add "writeCDR", -1
                If (Data(11) And &H2) Then .Add "writeCDRW", -1
                
                ' DVD
                If (Data(10) And &H8) Then .Add "readDVDROM", -1
                If (Data(10) And &H10) Then .Add "readDVDR", -1
                If (Data(10) And &H20) Then .Add "readDVDRAM", -1
                If (Data(11) And &H10) Then .Add "writeDVDR", -1
                If (Data(11) And &H20) Then .Add "writeDVDRAM", -1
                
                ' Misc.
                If (Data(10) And &H4) Then .Add "allowMethod2", -1
                
                If (Data(11) And &H4) Then .Add "writeTesting", -1
                
                If (Data(12) And &H1) Then .Add "audioPlay", -1
                If (Data(12) And &H2) Then .Add "compositeSignal", -1
                If (Data(12) And &H4) Then .Add "digitalPort1", -1
                If (Data(12) And &H8) Then .Add "digitalPort2", -1
                If (Data(12) And &H10) Then .Add "mode2Form1", -1
                If (Data(12) And &H20) Then .Add "mode2Form2", -1
                If (Data(12) And &H40) Then .Add "Multisession", -1
                If (Data(12) And &H80) Then .Add "bufferUnderrunFree", -1
                
                If (Data(13) And &H1) Then .Add "cddaCommands", -1
                If (Data(13) And &H2) Then .Add "cddaAccurate", -1
                If (Data(13) And &H4) Then .Add "subChanRW", -1
                If (Data(13) And &H8) Then .Add "subChanDIEC", -1
                If (Data(13) And &H10) Then .Add "c2ErrorPtrs", -1
                If (Data(13) And &H20) Then .Add "readISRC", -1
                If (Data(13) And &H40) Then .Add "readUPC", -1
                If (Data(13) And &H80) Then .Add "readBarCode", -1
                
                If (Data(14) And &H1) Then .Add "canLockMedia", -1
                If (Data(14) And &H2) Then .Add "allowLock", -1
                If (Data(14) And &H4) Then .Add "preventJumper", -1
                If (Data(14) And &H8) Then .Add "canEject", -1
                
                If (Data(15) And &H1) Then .Add "separateVolume", -1
                If (Data(15) And &H2) Then .Add "separateMute", -1
                If (Data(15) And &H4) Then .Add "discPresent", -1
                If (Data(15) And &H8) Then .Add "swSlotSelect", -1
                If (Data(15) And &H10) Then .Add "canChangeSide", -1
                If (Data(15) And &H20) Then .Add "readRawLeadIn", -1
                
                ' Loading tray mechanism
                bitSet = Device.getDataBit(14, 7) & Device.getDataBit(14, 6) & Device.getDataBit(14, 5)
                If bitSet = "000" Then deviceLoadMechanism = "Caddy"
                If bitSet = "001" Then deviceLoadMechanism = "Tray"
                If bitSet = "010" Then deviceLoadMechanism = "Popup"
                If bitSet = "100" Then deviceLoadMechanism = "Changer (individual)"
                If bitSet = "000" Then deviceLoadMechanism = "Changer (magazine)"
                
                ' Supported volume levels
                deviceVolumeLevels = Device.getDataInt(18)
                
                ' Buffer size
                deviceBufferSize = Device.getDataInt(20)
                
                ' Max speeds
                deviceReadSpeed = Device.getDataInt(16)
                deviceWriteSpeed = Device.getDataInt(26)
                
                ' Write speeds
                writeSpeeds = Device.getDataInt(38)
                Offset = 40
                For A = 0 To writeSpeeds
                    writeSpeed = Device.getDataInt(Offset + 2)
                    If writeSpeed > 0 Then
                        If Not deviceSpeeds.Exists("writeSpeed" & A + 1) Then deviceSpeeds.Add "writeSpeed" & A + 1, writeSpeed
                    End If
                    Offset = Offset + 4
                    If Offset + 4 > Device.getIOBufferSize Then Exit For
                Next A
            
            End With
            
        End If
    
    End If

    getDeviceCaps = True
    
End Function


' Read device inquiry data
Private Function getDeviceInquiry(ByRef Device As clsSPTIDevice) As Boolean

    Dim sData As String
    Dim devType As Byte
    
    
    ' SPTI call
    Device.clearCDB
    Device.setCDBByte 0, SCSIOP_INQUIRY
    Device.setCDBByte 4, 255
    If Not Device.sendSCSIOp(OPLEN_6) Then Exit Function
    
    ' Abort if device type is not CD\DVD
    ' Clears last 3 bits of inspection byte
    If (Device.getDataByte(0) And Not &HE0) <> 5 Then Exit Function
    
    ' Parse vendor, ID and revision
    deviceName = RTrim$(Device.getDataString(9, 8)) & " " & RTrim$(Device.getDataString(17, 16)) & " " & RTrim$(Device.getDataString(33, 4))

    getDeviceInquiry = True

End Function
