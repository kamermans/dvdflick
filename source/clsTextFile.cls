VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextFile"
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
'   File purpose: Custom text file reader\writer.
'
Option Explicit
Option Compare Binary
Option Base 0


' Text encoding type
Public Enum enumTextEncodings
    encodeASCIIorUTF8 = 0
    encodeUTF16
End Enum


' Security attributes
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


' Private data
Private mFileName As String
Private mHandle As Long
Private mEncoding As enumTextEncodings
Private mCodepage As Long
Private mWriteMode As Boolean

Private Buffer As String
Private bufferOffset As Long

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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Long, ByVal Source As Long, ByVal bytes As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplate As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal distanceToMove As Long, ByRef newFilePointer As Long, ByVal moveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal Handle As Long, ByVal ptrHigh As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal ptrIn As Long, ByVal sizeIn As Long, ByVal ptrOut As Long, ByVal sizeOut As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal ptrIn As Long, ByVal sizeIn As Long, ByVal ptrOut As Long, ByVal sizeOut As Long, ByVal ptrDefaultChar As Long, ByVal ptrUsedDefaultChar As Long) As Long
Private Declare Function GetACP Lib "kernel32" () As Long


' Codepage property
Public Property Get codePage() As Long

    codePage = mCodepage

End Property


' Filename property
Public Property Get fileName() As String

    fileName = mFileName

End Property


' Write mode property
Public Property Get writeMode() As Boolean

    writeMode = mWriteMode

End Property


' File encoding property
Public Property Get Encoding() As enumTextEncodings

    Encoding = mEncoding

End Property


' Writing functions
Public Function writeText(ByVal Data As String, Optional ByVal NewLine As Boolean = False) As Boolean

    Dim byteData() As Byte
    
    
    If mHandle = -1 Or Not mWriteMode Then Exit Function
    
    If NewLine Then Data = Data & vbCrLf
    
    ' UTF8\ASCII
    If mEncoding = encodeASCIIorUTF8 Then
        If convertWideToByte(Data, mCodepage, byteData) = False Then Exit Function
        If WriteFile(mHandle, VarPtr(byteData(0)), UBound(byteData), bytesWritten, ByVal 0&) = 0 Then Exit Function
    
    ' UTF16
    ElseIf mEncoding = encodeUTF16 Then
        ReDim byteData(Len(Data) * 2 - 1)
        CopyMemory VarPtr(byteData(0)), StrPtr(Data), Len(Data) * 2
        If WriteFile(mHandle, VarPtr(byteData(0)), Len(Data) * 2, bytesWritten, ByVal 0&) = 0 Then Exit Function
        
    End If
    
    writeText = True

End Function


' Return whether the end of this file has been reached
Public Function fileEndReached() As Boolean

    If mWriteMode Or bufferOffset >= Len(Buffer) Then fileEndReached = True

End Function


' Return the length of the opened file
Public Function fileLength() As Long
    
    ' Writing
    If mWriteMode Then
        If mHandle = -1 Then Exit Function
        fileLength = GetFileSize(mHandle, 0)
        
    ' Reading
    Else
        fileLength = Len(Buffer)
    
    End If

End Function


' Return the current file offset
Public Function fileOffset() As Long
    
    ' Writing
    If mWriteMode Then
        If mHandle = -1 Then Exit Function
        If SetFilePointer(mHandle, 0, 0, FILE_CURRENT) = 0 Then fileOffset = -1
        
    ' Reading
    Else
        fileOffset = bufferOffset
    
    End If

End Function


' Seek to an offset starting from the beginning of the file
Public Function fileSeek(ByVal Offset As Long) As Boolean

    Dim newP As Currency
    
    
    ' Writing
    If mWriteMode Then
        If mHandle = -1 Then Exit Function
        If SetFilePointer(mHandle, Offset, 0, FILE_BEGIN) = 0 Then Exit Function
        
    ' Reading
    Else
        bufferOffset = Offset
    
    End If
    
    fileSeek = True

End Function


' Create a new file
Public Function fileCreate(ByVal fileName As String, ByVal textEncoding As enumTextEncodings, Optional ByVal codePage As Long = -1) As Boolean

    Dim byteData() As Byte
    
    
    mHandle = CreateFile(StrPtr(fileName), FILE_GENERIC_READ Or FILE_GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, &H80, 0)
    If mHandle = -1 Then Exit Function
    
    ' Write UTF16 BOM
    ' BOM is not written for UTF8 for backward compatibilty with apps that do not support UTF8
    If textEncoding = encodeUTF16 Then
        ReDim byteData(1)
        byteData(0) = &HFF
        byteData(1) = &HFE
        
        If WriteFile(mHandle, VarPtr(byteData(0)), 2, bytesWritten, ByVal 0&) = 0 Then Exit Function
        
    End If
    
    If codePage = -1 Then codePage = GetACP
    
    mEncoding = textEncoding
    mCodepage = codePage
    mFileName = fileName
    mWriteMode = True
    
    fileCreate = True

End Function


' Open a file
Public Function fileOpen(ByVal fileName As String, Optional ByVal codePage As Long = -1) As Boolean

    Dim tmpBuffer() As Byte
    Dim dataLen As Long
    Dim Handle As Long


    ' Refuse to open large files (> 10 Megabytes)
    ' clsTextFile was not made for large files since all data is read into memory
    If FS.GetFile(fileName).Size > 10 * 1024& * 1024& Then Exit Function

    ' Open file
    Handle = CreateFile(StrPtr(fileName), FILE_GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, &H80, 0)
    If Handle = -1 Then Exit Function
    
    ' Read file
    dataLen = GetFileSize(Handle, 0)
    ReDim tmpBuffer(dataLen)
    ReadFile Handle, VarPtr(tmpBuffer(0)), dataLen, bytesRead, 0&
    
    ' Detect encoding
    If codePage = -1 Then
        mCodepage = GetACP
    Else
        mCodepage = codePage
    End If
    mEncoding = detectFileEncoding(tmpBuffer(), mCodepage, bufferOffset)

    ' Convert file contents
    ' UTF-16
    If mEncoding = encodeUTF16 Then
        Buffer = Space$(dataLen / 2)
        CopyMemory StrPtr(Buffer), VarPtr(tmpBuffer(0)), dataLen

    ' UTF-8
    ElseIf mEncoding = encodeASCIIorUTF8 Then
        Buffer = convertByteToWide(tmpBuffer, dataLen, mCodepage)
    
    End If

    ' Detect line ending style
    If hasUnixLineEndings Then Buffer = Replace(Buffer, vbLf, vbCrLf)

    ' Clean up
    Erase tmpBuffer
    CloseHandle Handle

    mFileName = fileName
    mWriteMode = False
    mHandle = -1
    
    fileOpen = True

End Function


' Detect a buffer's type of text encoding
Private Function detectFileEncoding(ByRef Data() As Byte, ByRef codePage As Long, ByRef Offset As Long) As enumTextEncodings

    Dim Encoding As enumTextEncodings
    
    
    ' Prevent out of bounds lookups for 1-byte sized files
    If UBound(Data) <= 1 Then
        detectFileEncoding = encodeASCIIorUTF8
        codePage = 0
        Offset = 1
        Exit Function
    End If
    
    ' Note: Offset is offset in encoded characters, not bytes, starting at 1
    ' UTF16 LE
    If Data(0) = &HFF And Data(1) = &HFE Then
        Encoding = encodeUTF16
        codePage = -1
        Offset = 2
        
    ' UTF16 BE
    ElseIf Data(0) = &HFE And Data(1) = &HFF Then
        Encoding = encodeUTF16
        codePage = -1
        Offset = 2
        
    ' UTF8
    ElseIf (Data(0) = &HEF And Data(1) = &HBB And Data(2) = &HBF) Then
        Encoding = encodeASCIIorUTF8
        codePage = 65001
        Offset = 2
        
    ' Detect whether it's UTF8 without a BOM, or ASCII
    Else
        Encoding = encodeASCIIorUTF8
        Offset = 1
        
        If isFileUTF8(Data()) Then
            codePage = 65001
            If (Data(0) = &HEF And Data(1) = &HBB And Data(2) = &HBF) Then Offset = 2
        End If

    End If

    detectFileEncoding = Encoding

End Function


' Attempt to detect whether a file uses Unix-style line endings
Private Function hasUnixLineEndings() As Boolean

    Dim A As Long
    Dim Char As Long
    Dim prevChar As Long
    Dim maxLen As Long
    
    
    ' Limit search
    maxLen = 20480
    If Len(Buffer) < maxLen Then maxLen = Len(Buffer)
    
    ' If a CrLf is found, then style is DOS, if only a lone Lf then Unix
    For A = 1 To maxLen
        
        Char = AscW(Mid$(Buffer, A, 1))
        
        If Char = 10 Then
            If prevChar = 13 Then
                hasUnixLineEndings = False
            Else
                hasUnixLineEndings = True
            End If
            
            Exit Function
        End If
        
        prevChar = Char
        
    Next A

End Function


' Attempt to detect whether a file is encoded using UTF8
' Adapted from http://mail.nl.linux.org/linux-utf8/1999-09/msg00110.html
Private Function isFileUTF8(ByRef Data() As Byte) As Boolean

    Dim A As Long
    Dim goodUTF As Long, badUTF As Long
    Dim prevChar As Byte
    

    ' Count number of good UTF8 sequences vs bad ones
    For A = 0 To UBound(Data)
    
        If (Data(A) And &HC0) = &H80 Then
            If (prevChar And &HC0) = &HC0 Then
                goodUTF = goodUTF + 1

            ElseIf (prevChar And &H80) = 0 Then
                badUTF = badUTF + 1
                
            End If

        ElseIf (prevChar And &HC0) = &HC0 Then
            badUTF = badUTF + 1
            
        End If

        prevChar = Data(A)
        
    Next A
    
    ' Determine outcome
    If goodUTF > badUTF Then isFileUTF8 = True

End Function


' Close this file
Public Sub fileClose()

    If mWriteMode And mHandle <> -1 Then
        FlushFileBuffers mHandle
        CloseHandle mHandle
    End If
    
    mHandle = -1
    mWriteMode = False
    mEncoding = encodeASCIIorUTF8
    mCodepage = 0
    mFileName = vbNullString
    
    Buffer = vbNullString
    bufferOffset = 1

End Sub


' Read a single line
Public Function readLine() As String

    Dim A As Long
    Dim lastWasCr As Boolean
    Dim Char As Long
    

    For A = bufferOffset To Len(Buffer)
    
        Char = AscW(Mid$(Buffer, A, 1))

        If (Char = 10 And lastWasCr) Then
            readLine = Mid$(Buffer, bufferOffset, A - bufferOffset - 1)
            Exit For
            
        ElseIf A = Len(Buffer) Then
            readLine = Mid$(Buffer, bufferOffset)
            Exit For
            
        ElseIf Char = 13 Then
            lastWasCr = True
            
        Else
            lastWasCr = False
            
        End If
    
    Next A

    bufferOffset = A + 1

End Function


' Return the entire buffer
Public Function readAll() As String

    readAll = Buffer

End Function


' Convert byte data into a UTF16 string
Private Function convertByteToWide(ByRef Data() As Byte, ByVal byteCount As Long, ByVal codePage As Long) As String
    
    Dim bufferLen As Long
    Dim tempBuffer() As Byte
    
    
    If byteCount = 0 Then Exit Function
    
    ' Get required space
    bufferLen = MultiByteToWideChar(codePage, 0, VarPtr(Data(0)), byteCount, 0, 0)
    If bufferLen = 0 Then Err.Raise vbObjectError, , "Could not convert byte data... " & bufferLen & "... " & Err.LastDllError
    
    ' Allocate and actually convert
    ReDim tempBuffer(bufferLen * 2)
    MultiByteToWideChar codePage, 0, VarPtr(Data(0)), byteCount, VarPtr(tempBuffer(0)), bufferLen
    
    convertByteToWide = CStr(tempBuffer)

End Function


' Convert a UTF16 string to ASCII
Private Function convertWideToByte(ByVal Data As String, ByVal codePage As Long, ByRef outData() As Byte) As Boolean

    Dim bufferLen As Long


    ' Get required space
    bufferLen = WideCharToMultiByte(codePage, 0, StrPtr(Data), Len(Data), 0, 0, 0, 0)
    If bufferLen = 0 Then Err.Raise vbObjectError, , "Could not convert wide character data... " & bufferLen & "... " & Err.LastDllError
    
    ' Allocate and actually convert
    ReDim outData(bufferLen)
    WideCharToMultiByte codePage, 0, StrPtr(Data), Len(Data), VarPtr(outData(0)), bufferLen, 0, 0
    
    convertWideToByte = True

End Function


' Class
Private Sub Class_Initialize()
    
    mHandle = -1
    mWriteMode = False
    mEncoding = encodeASCIIorUTF8
    mCodepage = 0
    mFileName = vbNullString

End Sub

Private Sub Class_Terminate()

    If mHandle <> -1 Then
        FlushFileBuffers mHandle
        CloseHandle mHandle
    End If

End Sub
