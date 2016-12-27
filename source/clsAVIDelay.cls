VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAVIDelay"
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
'   File purpose: AVI AC3\MP3 audio delay detector
'                 Only detects delay in CBR audio, VBR todo.
'
Option Explicit
Option Compare Binary
Option Base 0


' AVI Header
Private Type aviHeader
    uSPerFrame As Long
    bytePerSec As Long
    Reserved1 As Long
    flags As Long
    frameCount As Long
    intialFrames As Long
    streamCount As Long
    bufferSize As Long
    Width As Long
    Height As Long
    Reserved2 As Long
End Type

' Stream header
Private Type streamHeader
    lFourCC As Long
    lCodecFourCC As Long
    flags As Long
    Priority As Integer
    Language As Integer
    initialFrames As Long
    timeScale As Long
    timeRate As Long
    startTime As Long
    Length As Long
    bufferSize As Long
    Quality As Long
    sampleSize As Long
    frameLeft As Integer
    frameTop As Integer
    frameRight As Integer
    frameBottom As Integer
End Type

' WAVE sound information
Private Type WAVEFORMATEX
    formatTag As Integer
    Channels As Integer
    samplesPerSec As Long
    avgBytesPerSec As Long
    blockAlign As Integer
    bitsPerSample As Integer
    Size As Integer
End Type

' AVI stream information
Private Type aviStream
    Header As streamHeader
    audioFormat As WAVEFORMATEX
    
    audioData As Long
    Delay As Long
    VBR As Boolean
End Type


' Syncable audio format tags
Private Const AUDIO_MPEG1L3 As Long = 85
Private Const AUDIO_AC3 As Long = 8192

' Maximum amount of data to try to sync to
Private Const MAX_SYNC_DATA As Long = 1000000
Private Const MAX_FILE_SYNC_DATA As Long = 1500000

' Invalid delay value
Private Const INVALID_DELAY As Long = &HFFFFFF

' "AUDS" fourCC value
Private Const FCC_AUDS As Long = 1935963489

' Filename of source
Private fileName As String

' Streams
Private Streams() As aviStream
Private nStreams As Long


' Number of streams
Public Property Get streamCount() As Long

    streamCount = nStreams

End Property


' Delay of stream
Public Property Get streamDelay(ByVal Index As Long) As Long

    streamDelay = Streams(Index).Delay

End Property


' Return whether a stream has VBR audio
Public Property Get streamIsVBR(ByVal Index As Long) As Boolean

    streamIsVBR = Streams(Index).VBR

End Property


' Return whether a stream is an audio stream or not
Public Property Get streamIsAudio(ByVal Index As Long) As Boolean

    If Streams(Index).Header.lFourCC = FCC_AUDS Then
        streamIsAudio = True
    Else
        streamIsAudio = False
    End If

End Property


' Open and parse an AVI file
Public Function openFile(ByVal myFile As String) As Boolean

    Dim Parser As clsRIFFParser
    Dim chunkName As String
    Dim Header As aviHeader
    
    Dim Data As String
    Dim streamIndex As Long
    Dim audioData As Long
    Dim syncOffset As Long
    Dim formatTag As Long
    Dim chunkDuration As Long
    
    On Error GoTo OpenFile_Error
    

    Set Parser = New clsRIFFParser

    If Not Parser.openFile(myFile) Then
        openFile = False
        Exit Function
    End If
    
    ' Check filetype
    If Parser.fileFourCC <> "AVI " Then
        openFile = False
        Exit Function
    End If
    
    
    Do
        Parser.getChunk
        
        ' Lists only
        If Parser.chunkType = "list" Then
            
            ' Skip index
            If Parser.chunkFourCC = "idx1" Then Parser.skipChunk
        
        ' Chunks only
        ElseIf Parser.chunkType = "chunk" Then
            
            ' AVI header
            If Parser.chunkFourCC = "avih" Then
                Parser.fileObj.readData VarPtr(Header), Len(Header)
            
            
            ' Stream header
            ElseIf Parser.chunkFourCC = "strh" Then
                ReDim Preserve Streams(nStreams)
                nStreams = nStreams + 1
                
                Streams(nStreams - 1).Delay = INVALID_DELAY
                Parser.fileObj.readData VarPtr(Streams(nStreams - 1).Header), Len(Streams(nStreams - 1).Header)
                
                ' Divide by 0 preventions
                With Streams(nStreams - 1)
                    If .Header.timeRate = 0 Then .Header.timeRate = 1
                    If .audioFormat.avgBytesPerSec = 0 Then .audioFormat.avgBytesPerSec = 1
                End With


            ' Stream format
            ElseIf Parser.chunkFourCC = "strf" Then
            
                ' Audio stream
                With Streams(nStreams - 1)
                    If .Header.lFourCC = FCC_AUDS Then Parser.fileObj.readData VarPtr(.audioFormat), Len(.audioFormat)
                End With
                
                
            ' Audio stream data
            ElseIf Right$(Parser.chunkFourCC, 2) = "wb" Then
                streamIndex = CLng(Left$(Parser.chunkFourCC, 2))

                If Streams(streamIndex).Delay = INVALID_DELAY Then
                    formatTag = Streams(streamIndex).audioFormat.formatTag
                    If (formatTag = AUDIO_AC3 Or formatTag = AUDIO_MPEG1L3) And Parser.chunkSize > 0 Then
                        
                        ' Read chunk data
                        Data = Parser.fileObj.readStringData(Parser.chunkSize)
                        
                        ' Get sync offset from chunk
                        If Streams(streamIndex).audioFormat.formatTag = AUDIO_AC3 Then
                            syncOffset = getAC3Sync(Data)
                        ElseIf formatTag = AUDIO_MPEG1L3 Then
                            syncOffset = getMP3Sync(Data)
                        End If
                        
                        If syncOffset <> -1 Then
                            
                            With Streams(streamIndex)
                            
                                .audioData = .audioData + syncOffset
                                
                                ' Determine if VBR MP3
                                If .audioFormat.formatTag = AUDIO_MPEG1L3 And .audioFormat.blockAlign = 1152 Then .VBR = True
                                
                                ' Calculate delay based on type of stream
                                ' MP3 VBR
                                ' Duration of a single chunk of data * number of missed chunks
                                If .VBR Then
                                    chunkDuration = ((Parser.chunkSize / .audioFormat.blockAlign) * (.Header.timeRate / .Header.timeScale))
                                    .Delay = (.audioData / Parser.chunkSize) * chunkDuration
                                
                                ' CBR audio
                                Else
                                    .Delay = (.audioData / .audioFormat.avgBytesPerSec * 1000)
                                
                                End If
                                
                                ' Add dwStartTime too
                                .Delay = .Delay + (.Header.startTime * (.Header.timeRate / .Header.timeScale) * 1000)
                                
                            End With
                            
                        End If

                        Streams(streamIndex).audioData = Streams(streamIndex).audioData + Parser.chunkSize
                        If Streams(streamIndex).audioData > MAX_SYNC_DATA Then Streams(streamIndex).Delay = INVALID_DELAY
                        
                    End If
                        
                End If
            
            End If
            
        End If

        
        If Parser.fileObj.fileOffset > MAX_FILE_SYNC_DATA Then Exit Do
    
    Loop Until Parser.chunkType = "eof"

    
    Set Parser = Nothing
    fileName = myFile
    openFile = True
    
    On Error GoTo 0
    
    Exit Function
    
    
OpenFile_Error:
    openFile = False
    Set Parser = Nothing
    
    On Error GoTo 0

End Function


Private Function getAC3Sync(ByRef Data As String) As Long

    Dim Char As String
    Dim Rover As Long
    Dim syncWord As String
    
    
    getAC3Sync = -1
    
    syncWord = ChrW$(&HB) & ChrW$(&H77)
    Rover = 1

    Do
        Char = Mid$(Data, Rover, 2)
        If Char = syncWord Then
            getAC3Sync = Rover - 1
            Exit Function
        End If
        
        Rover = Rover + 1
    Loop Until Rover > Len(Data) - 1

End Function


Private Function getMP3Sync(ByRef Data As String) As Long

    Dim Char As String
    Dim Rover As Long
    
    
    getMP3Sync = -1
    Rover = 1

    Do
        Char = AscW(Mid$(Data, Rover, 1))
        If Char = 255 Then
            Char = AscW(Mid$(Data, Rover + 1, 1))
            If (Char And 3) = 3 Then
                getMP3Sync = Rover - 1
                Exit Function
            End If
        End If
        
        Rover = Rover + 1
    Loop Until Rover > Len(Data) - 1
    
End Function
