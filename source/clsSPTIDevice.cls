VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSPTIDevice"
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
'   File purpose: SPTI device wrapper. Allows one to send SCSI commands and
'                 read returned buffer data.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Const ioBufferSize As Long = 952
Private Const ioSenseSize As Long = 32
Private Const ioCDBSize As Long = 16


' SCSI passthrough
Private Type SCSI_PASS_THROUGH
    Length As Integer
    scsiStatus As Byte
    pathId As Byte
    targetId As Byte
    Lun As Byte
    cdbLength As Byte
    senseInfoLength As Byte
    dataIn As Byte
    dataTransferLength As Long
    timeOutValue As Long
    dataBufferOffset As Long
    senseInfoOffset As Long
    Cdb(ioCDBSize - 1) As Byte
End Type

' SCSI buffered passthrough
Private Type SCSI_PASS_THROUGH_BUFFER
    SPTI As SCSI_PASS_THROUGH
    Filler As Long
    senseData(ioSenseSize - 1) As Byte
    Data(ioBufferSize - 1) As Byte
End Type


' CreateFile constants
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3

' DeviceIoControl SCSI passthrough command
Private Const IOCTL_SCSI_PASS_THROUGH  As Long = &H4D004

' SCSI passthrough
Private Const SCSI_IOCTL_DATA_OUT As Byte = 0
Private Const SCSI_IOCTL_DATA_IN As Byte = 1
Private Const SCSI_IOCTL_DATA_UNSPECIFIED As Byte = 2
Private Const SPTI_TIMEOUT As Long = 1


' Device handle
Private hDevice As Long

' SPTI passthrough buffer
Private sptiBuf As SCSI_PASS_THROUGH_BUFFER


' API
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal bytes As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

' dvdflicklib
Private Declare Function reverseLong Lib "dvdflick" (ByVal Value As Long) As Long
Private Declare Function reverseInteger Lib "dvdflick" (ByVal Value As Integer) As Integer


' Return IO buffer size
Public Property Get getIOBufferSize() As Long

    getIOBufferSize = ioBufferSize

End Property


' CDB data set functions
Public Sub setCDBByte(ByVal Index As Long, ByVal Data As Byte)

    If Index < 0 Or Index > ioCDBSize - 1 Then Exit Sub
    
    CopyMemory sptiBuf.SPTI.Cdb(Index), Data, 1

End Sub

Public Sub setCDBInt(ByVal Index As Long, ByVal Data As Integer)

    If Index < 0 Or Index + 1 > ioCDBSize - 1 Then Exit Sub
    
    Data = reverseInteger(Data)
    CopyMemory sptiBuf.SPTI.Cdb(Index), Data, 2

End Sub

Public Sub setCDBLong(ByVal Index As Long, ByVal Data As Long)

    If Index < 0 Or Index + 3 > ioCDBSize - 1 Then Exit Sub
    
    Data = reverseLong(Data)
    CopyMemory sptiBuf.SPTI.Cdb(Index), Data, 4

End Sub


' Data buffer get functions
Public Function getDataBit(ByVal byteIndex As Long, ByVal bitIndex As Long) As Byte

    If byteIndex < 0 Or byteIndex > ioBufferSize - 1 Then Exit Function
    If bitIndex > 7 Or bitIndex < 0 Then Exit Function
    
    If (sptiBuf.Data(byteIndex) And (2 ^ bitIndex)) Then getDataBit = 1

End Function

Public Function getDataByte(ByVal Index As Long) As Byte

    If Index < 0 Or Index > ioBufferSize - 1 Then Exit Function
    
    CopyMemory getDataByte, sptiBuf.Data(Index), 1

End Function

Public Function getDataInt(ByVal Index As Long) As Integer

    If Index < 0 Or Index + 1 > ioBufferSize - 1 Then Exit Function

    CopyMemory getDataInt, sptiBuf.Data(Index), 2
    getDataInt = reverseInteger(getDataInt)

End Function

Public Function getDataLong(ByVal Index As Long) As Long

    If Index < 0 Or Index + 3 > ioBufferSize - 1 Then Exit Function
    
    CopyMemory getDataLong, sptiBuf.Data(Index), 4
    getDataLong = reverseLong(getDataLong)

End Function

Public Function getDataString(ByVal Start As Long, ByVal Length As Long) As String

    Dim A As Long
    
    
    If Start < 0 Or Start + Length > ioBufferSize - 1 Then Exit Function
    
    getDataString = Space$(Length)
    For A = 0 To Length
        CopyMemory ByVal StrPtr(getDataString) + A * 2, ByVal VarPtr(sptiBuf.Data(A)) + Start - 1, 1
    Next A

End Function


' Copy entire SPTI data buffer
Public Sub getDataArray(ByRef byteArray() As Byte)

    ReDim byteArray(ioBufferSize - 1)
    CopyMemory byteArray(0), sptiBuf.Data(0), ioBufferSize

End Sub


' Clear CDB data
Public Sub clearCDB()

    FillMemory sptiBuf.SPTI.Cdb(0), ioCDBSize, 0

End Sub


' Send a SCSI operation
Public Function sendSCSIOp(ByVal opLength As Byte) As Boolean

    Dim bytesReturned As Long
    
    
    If hDevice = -1 Then Exit Function
    
    ' Reset SPTI structure
    With sptiBuf.SPTI
        .scsiStatus = 0
        .targetId = 1
        .Lun = 0
        .pathId = 0
        .cdbLength = opLength
        .Length = LenB(sptiBuf.SPTI)
        .senseInfoLength = ioSenseSize
        .dataIn = SCSI_IOCTL_DATA_IN
        .dataTransferLength = ioBufferSize
        .timeOutValue = SPTI_TIMEOUT
        .dataBufferOffset = LenB(sptiBuf.SPTI) + 4 + ioSenseSize
        .senseInfoOffset = LenB(sptiBuf.SPTI) + 4
    End With
    
    ' Send
    If DeviceIoControl(hDevice, IOCTL_SCSI_PASS_THROUGH, sptiBuf, LenB(sptiBuf), sptiBuf, LenB(sptiBuf), bytesReturned, 0) = 0 Then Exit Function
    If sptiBuf.SPTI.scsiStatus <> 0 Then Exit Function
    
    sendSCSIOp = True

End Function


' Return status after last command
Public Property Get getSCSIStatus() As Byte

    getSCSIStatus = sptiBuf.SPTI.scsiStatus

End Property


' Return number of returned data bytes from last command
Public Property Get getBytesReturned() As Long

    getBytesReturned = sptiBuf.SPTI.dataTransferLength

End Property


' Connect to a device by supplying a drive character (A, C, D, etc.)
Public Function connectTo(ByVal driveChar As String) As Boolean

    hDevice = CreateFile("\\.\" & driveChar & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If hDevice = -1 Then Exit Function

    connectTo = True

End Function


' Class
Private Sub Class_Initialize()

    hDevice = -1

End Sub

Private Sub Class_Terminate()

    If hDevice <> -1 Then CloseHandle hDevice

End Sub
