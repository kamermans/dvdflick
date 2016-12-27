VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRIFFParser"
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
'   File purpose: General RIFF file type parser
'                 Only ever tested with AVI files
'
Option Explicit
Option Compare Binary
Option Base 0


' File info
Public fileFourCC As String
Public fileSize As Long
Public fileObj As clsBinaryFile
Public fileName As String

' Chunk info to read
Public chunkType As String
Public chunkSize As Long
Public chunkFourCC As String
Public currentChunk As Long

Private nextChunkOffset As Long


' Open a file and parse it's possible RIFF structure
Public Function openFile(ByVal myFileName As String) As Boolean

    Dim Header As String


    Set fileObj = New clsBinaryFile
    If Not fileObj.fileOpen(myFileName, False) Then Exit Function

    ' Check header
    Header = fileObj.readStringData(4)
    If Header <> "RIFF" Then
        openFile = False
        Exit Function
    End If
    
    fileSize = fileObj.readLong
    fileFourCC = fileObj.readStringData(4)
    fileName = myFileName
    
    openFile = True

End Function


Public Sub getChunk()

    Dim fourCC As String
    Dim Size As Long
    

    ' Seek to next chunk if we already read one
    If nextChunkOffset > 0 Then fileObj.fileSeek nextChunkOffset
    
    fourCC = fileObj.readStringData(4)
    Size = fileObj.readLong

    
    ' List type
    If fourCC = "LIST" Then
        chunkType = "list"
        chunkSize = Size - 4
        chunkFourCC = fourCC
        chunkFourCC = fileObj.readStringData(4)
        nextChunkOffset = 0
    
    ' Plain chunk
    Else
        chunkType = "chunk"
        chunkSize = Size
        
        ' Sizes are padded on 2 byte boundaries no matter what size
        ' the chunk indicates itself
        chunkSize = chunkSize + (chunkSize Mod 2)
        
        chunkFourCC = fourCC
        nextChunkOffset = fileObj.fileOffset + chunkSize
    
    End If
    
    currentChunk = currentChunk + 1
    
    
    ' Detect EOF (we have just read an imaginary chunk\list)
    If fileObj.fileEndReached Then
        chunkType = "eof"
        Exit Sub
    End If

End Sub


' Skip over the last read chunk
Public Sub skipChunk()

    fileObj.fileSeek fileObj.fileOffset + chunkSize

End Sub


Private Sub Class_Terminate()

    Set fileObj = Nothing

End Sub
