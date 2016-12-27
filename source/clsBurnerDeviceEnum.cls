VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBurnerDeviceEnum"
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
'   File purpose: Enumerate installed CD\DVD burner drive objects.
'
Option Explicit
Option Compare Binary
Option Base 0


' Enumerated drives
Private Devices() As clsBurnerDevice
Private nDevices As Long


' API
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function getBit Lib "dvdflick" (ByVal bitMask As Long, ByVal bitNum As Long) As Long


' Return number of devices available
Public Property Get deviceCount() As Long

    deviceCount = nDevices

End Property


' Return a device's object
Public Property Get getDevice(ByVal Index As Long) As clsBurnerDevice

    If Index < 0 Or Index > nDevices Then Exit Property
    Set getDevice = Devices(Index)

End Property
    

' Scan for devices
Public Function scanDevices() As Boolean
    
    Dim A As Long
    Dim cDevice As Long
    Dim Device As clsBurnerDevice
    Dim deviceMask As Long, driveL As String
    
        
    ' Erase possible previous
    Erase Devices
    nDevices = 0
    
    ' Get bitmask of available drives
    deviceMask = GetLogicalDrives
    If deviceMask = 0 Then Exit Function
    
    
    ' Scan attached devices
    For A = 0 To 25

        If getBit(deviceMask, A) Then
            driveL = Chr$(65 + A)

            ' Add new object
            ReDim Preserve Devices(nDevices)
            nDevices = nDevices + 1
            cDevice = nDevices - 1
    
            ' Get device properties, if not valid remove it
            Set Devices(cDevice) = New clsBurnerDevice
            If Not Devices(cDevice).getProperties(driveL) Then
                Set Devices(cDevice) = Nothing
                nDevices = nDevices - 1
            End If
            
        End If

    Next A
    
    ' Ensure last one is removed
    If nDevices > 0 Then
        ReDim Preserve Devices(nDevices - 1)
    Else
        Erase Devices
    End If
    
    scanDevices = True

End Function
