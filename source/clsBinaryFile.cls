VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBinaryFile"
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
'   File purpose: Binary file I/O using Windows API.
'
Option Explicit
Option Compare Binary
Option Base 0


' Security attributes
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


' Private data
Private mFileName As String
Private mHandle As Long
Private mAllowWrite As Boolean

' Bytes actually processed
Private bytesRead As Long
Private bytesWritten As Long


' Desired access
Private Const FILE_SHARE_READ As Long = 1
Private Const FILE_SHARE_WRITE As Long = 2
Private Const FILE_SHARE_DELETE As Long = 4

' Share mode
Private Const FILE_READ_DATA As Long = &H1
Private Const FILE_LIST_DIRECTORY As Long = &H1
Private Const FILE_WRITE_DATA As Long = &H2
Private Const FILE_ADD_FILE As Long = &H2
Private Const FILE_APPEND_DATA As Long = &H4
Private Const FILE_ADD_SUBDIRECTORY As Long = &H4
Private Const FILE_CREATE_PIPE_INSTANCE  As Long = &H4
Private Const FILE_READ_EA As Long = &H8
Private Const FILE_WRITE_EA As Long = &H10
Private Const FILE_EXECUTE As Long = &H20
Private Const FILE_TRAVERSE As Long = &H20
Private Const FILE_DELETE_CHILD As Long = &H40
Private Const FILE_READ_ATTRIBUTES As Long = &H80
Private Const FILE_WRITE_ATTRIBUTES As Long = &H100
Private Const FILE_GENERIC_READ As Long = FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA
Private Const FILE_GENERIC_WRITE As Long = FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA
Private Const FILE_GENERIC_EXECUTE As Long = FILE_READ_ATTRIBUTES Or FILE_EXECUTE

' Creation disposition
Private Const CREATE_NEW As Long = 1
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const TRUNCATE_EXISTING As Long = 5

' File postition move method
Private Const FILE_BEGIN As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END As Long = 2


' API
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplate As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" Alias "SetFilePointerEx" (ByVal hFile As Long, ByVal distanceToMove As Currency, ByRef newFilePointer As Currency, ByVal moveMethod As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, ptrFileSize As Currency) As Long


' Filename property
Public Property Get fileName() As String

    fileName = mFileName

End Property


' Return whether writing can be done
Public Property Get allowWrite() As Boolean

    allowWrite = mAllowWrite

End Property


' Writing functions
Public Function writeByte(ByVal Data As Byte) As Boolean

    If mHandle = -1 Then Exit Function
    If mAllowWrite = False Then Exit Function
    
    If WriteFile(mHandle, VarPtr(Data), 1, bytesWritten, ByVal 0&) = 0 Then Exit Function
    writeByte = True

End Function

Public Function writeInteger(ByVal Data As Integer) As Boolean

    If mHandle = -1 Then Exit Function
    If mAllowWrite = False Then Exit Function
    
    If WriteFile(mHandle, VarPtr(Data), 2, bytesWritten, ByVal 0&) = 0 Then Exit Function
    writeInteger = True

End Function

Public Function writeLong(ByVal Data As Long) As Boolean

    If mHandle = -1 Then Exit Function
    If mAllowWrite = False Then Exit Function
    
    If WriteFile(mHandle, VarPtr(Data), 4, bytesWritten, ByVal 0&) = 0 Then Exit Function
    writeLong = True

End Function

Public Function writeString(ByVal Data As String) As Boolean

    Dim dataLen As Long
    

    If mHandle = -1 Then Exit Function
    If mAllowWrite = False Then Exit Function
    
    dataLen = Len(Data)
    If WriteFile(mHandle, VarPtr(dataLen), 4, bytesWritten, ByVal 0&) = 0 Then Exit Function
    If WriteFile(mHandle, StrPtr(Data), LenB(Data), bytesWritten, ByVal 0&) = 0 Then Exit Function
    
    writeString = True

End Function

Public Function writeData(ByVal dataPtr As Long, ByVal dataLength As Long) As Boolean

    If mHandle = -1 Then Exit Function
    If mAllowWrite = False Then Exit Function
    
    If WriteFile(mHandle, dataPtr, dataLength, bytesWritten, ByVal 0&) = 0 Then Exit Function
    writeData = True

End Function


' Return whether the end of this file has been reached
Public Function fileEndReached() As Boolean

    If bytesRead = 0 Then fileEndReached = True

End Function


' Return the length of the opened file
Public Function fileLength() As Currency
    
    If mHandle = -1 Then Exit Function
    If GetFileSizeEx(mHandle, fileLength) = 0 Then
        fileLength = -1
    Else
        fileLength = fileLength * 10000
    End If

End Function


' Return the current file offset
Public Function fileOffset() As Currency
    
    If SetFilePointer(mHandle, 0, fileOffset, FILE_CURRENT) = 0 Then
        fileOffset = -1
    Else
        fileOffset = fileOffset * 10000
    End If

End Function


' Seek to an offset starting from the beginning of the file
Public Function fileSeek(ByVal Offset As Currency) As Boolean

    Dim newP As Currency
    
    
    If mHandle = -1 Then Exit Function
    
    Offset = Offset / 10000
    If SetFilePointer(mHandle, Offset, newP, FILE_BEGIN) = 0 Then Exit Function
    
    fileSeek = True

End Function


' Open a file
Public Function fileOpen(ByVal fileName As String, ByVal writeMode As Boolean) As Boolean

    Dim accessFlags As Long
    Dim shareFlags As Long
    
    
    ' Create opening flags for read\write modes
    accessFlags = FILE_GENERIC_READ
    If writeMode Then accessFlags = accessFlags Or FILE_GENERIC_WRITE
    shareFlags = FILE_SHARE_READ
    If writeMode Then shareFlags = shareFlags Or FILE_SHARE_WRITE
    
    CloseHandle mHandle
    mHandle = CreateFile(StrPtr(fileName), accessFlags, shareFlags, ByVal 0&, OPEN_ALWAYS, &H80, 0)
    If mHandle = -1 Then Exit Function
    
    mFileName = fileName
    mAllowWrite = writeMode
    fileOpen = True

End Function


' Close this file
Public Sub fileClose()

    CloseHandle mHandle
    mHandle = -1
    mAllowWrite = False

End Sub


' Reading functions
Public Function readByte() As Byte

    If mHandle = -1 Then Exit Function
    ReadFile mHandle, VarPtr(readByte), 1, bytesRead, ByVal 0&

End Function

Public Function readInteger() As Integer

    If mHandle = -1 Then Exit Function
    ReadFile mHandle, VarPtr(readInteger), 2, bytesRead, ByVal 0&

End Function

Public Function readLong() As Long

    If mHandle = -1 Then Exit Function
    ReadFile mHandle, VarPtr(readLong), 4, bytesRead, ByVal 0&

End Function

Public Function readString() As String

    Dim stringLen As Long
    
    
    If mHandle = -1 Then Exit Function
    
    ReadFile mHandle, VarPtr(stringLen), 4, bytesRead, ByVal 0&
    readString = Space$(stringLen)
    ReadFile mHandle, StrPtr(readString), stringLen * 2, bytesRead, ByVal 0&

End Function

Public Function readData(ByVal dataPtr As Long, ByVal dataLength As Long) As Boolean

    If mHandle = -1 Then Exit Function
    ReadFile mHandle, dataPtr, dataLength, bytesRead, ByVal 0&

End Function


' Special function, read data and return it as a converted unicode string
Public Function readStringData(ByVal Length As Long) As String

    Dim Data() As Byte
    
    
    ReDim Data(Length - 1)
    
    If mHandle = -1 Then Exit Function
    ReadFile mHandle, VarPtr(Data(0)), Length, bytesRead, ByVal 0&
    readStringData = StrConv(Data, vbUnicode)

End Function


' Class
Private Sub Class_Initialize()

    mHandle = -1

End Sub

Private Sub Class_Terminate()

    CloseHandle mHandle

End Sub
