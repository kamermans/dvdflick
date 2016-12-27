VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubFile"
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
'   File purpose: Detect, read and parse subtitle files.
'                 Supports SubRip, SSA\ASS, SubViewer (2.0) and MicroDVD formats.
'
'                 Limitations: HTML tags are stripped and unused.
'                              In SSA subtitles only the dialogue text is used.
'
Option Explicit
Option Compare Binary
Option Base 0


' Parser states
Private Enum parseStates
    NOSTATE = 0
    
    SRT_NONE
    SRT_TIME
    SRT_TEXT
    
    SSA_NONE
    SSA_DIALOGUE
    SSA_DIALOGUEEXT
    
    MDVD_NONE
    
    SUBV_NONE
    SUBV_TIME
    SUBV_TEXT
End Enum


' Subtitle format strings
Private Const FORMAT_SUBRIP As String = "SubRip"
Private Const FORMAT_SUBSTATION As String = "SubStation Alpha"
Private Const FORMAT_MICRODVD As String = "MicroDVD"
Private Const FORMAT_SUBVIEWER As String = "SubViewer"

' SubViewer timecode regexp pattern
Private Const PAT_SUBVIEWERTIME As String = "^[0-9]{2}:[0-9]{2}:[0-9]{2}.[0-9]{2},[0-9]{2}:[0-9]{2}:[0-9]{2}.[0-9]{2}"
Private Const PAT_MDVDLINE As String = "^{[0-9]*}{[0-9]*}[\s]*"


' Private properties
Private mFileName As String
Private mFormat As String
Private mFPS As Single
Private mFrameBased As Boolean
Private mCodepage As Long

' Current state of subtitle parser
Private parseState As parseStates
Private cBlock As Long

' Subtitles
Private Blocks() As clsSubBlock
Private mBlockCount As Long


' Codepage property
Public Property Get codePage() As Long

    codePage = mCodepage

End Property


' File name property
Public Property Get fileName() As String

    fileName = mFileName

End Property


' File format property
Public Property Get fileFormat() As String

    fileFormat = mFormat

End Property


' Frame based property
Public Property Get frameBased() As Boolean

    frameBased = mFrameBased

End Property


' Block count property
Public Property Get blockCount() As Long

    blockCount = mBlockCount

End Property


' Framerate property
Public Property Get FPS() As Single

    FPS = mFPS

End Property

Public Property Let FPS(ByVal Value As Single)

    mFPS = Value

End Property


' Convert this subtitle to a new framerate by adjusting timings
Public Function convertToFPS(ByVal newFPS As Single)

    Dim A As Long
    Dim Modif As Single
    
    
    Dbg.addLog DM_Subtitles, "Converting subtitle FPS from " & mFPS & " to " & newFPS
    
    Modif = newFPS / mFPS
    For A = 0 To mBlockCount - 1
        Blocks(A).startFrame = Blocks(A).startFrame * Modif
        Blocks(A).endFrame = Blocks(A).endFrame * Modif
    Next A
    
    mFPS = newFPS

End Function


' Convert frame numbers to display times using current FPS
Public Sub convertFramesToTime()

    Dim A As Long
    
    
    If Not mFrameBased Then Exit Sub
    
    Dbg.addLog DM_Subtitles, "Converting frame numbers to timestamps"
    
    For A = 0 To mBlockCount - 1
        Blocks(A).startTime = Blocks(A).startFrame / mFPS
        Blocks(A).endTime = Blocks(A).endFrame / mFPS
    Next A

End Sub


' Try to correct overlapping subtitles, returning number of fixed ones
Public Function fixOverlaps(Optional ByVal fixTimes As Boolean = True) As Long
    
    Dim A As Long
    Dim sTime As Single
    
    
    sTime = -1
    
    For A = 0 To mBlockCount - 1
        If Blocks(A).startTime < sTime Then
            If fixTimes Then Blocks(A).startTime = sTime + 0.01
            fixOverlaps = fixOverlaps + 1
        End If
        
        sTime = Blocks(A).endTime
    Next A
    
    Dbg.addLog DM_Subtitles, "Found " & fixOverlaps & " timestamp overlaps. Fixing enabled: " & fixTimes

End Function


' Try to correct too short subtitles, returning number of fixed ones
Public Function fixShorts(Optional ByVal fixTimes As Boolean = True) As Long
    
    Dim A As Long
    
    
    For A = 0 To mBlockCount - 1
        If Blocks(A).endTime - Blocks(A).startTime < 0.01 Then
            If fixTimes Then Blocks(A).endTime = Blocks(A).startTime + 0.01
            fixShorts = fixShorts + 1
        End If
    Next A
    
    Dbg.addLog DM_Subtitles, "Found " & fixShorts & " short timestamps. Fixing enabled: " & fixTimes

End Function


' Split lines that are too long
Private Function fixLengths(Optional ByVal maxLineLength As Long = 40) As Boolean

    Dim A As Long, B As Long
    Dim newBlock As String
    Dim nextSpace As Long
    Dim Lines() As String
    
    
    Dbg.addLog DM_Subtitles, "Attempting to fix long subtitle line lengths"
    
    For A = 0 To mBlockCount - 1
        
        ' Split the block into lines and remove newlines
        newBlock = Replace(Blocks(A).Text, vbLf, "")
        Lines = Split(newBlock, vbCr)
        newBlock = vbNullString
        
        For B = 0 To UBound(Lines)
            
            Do
            
                If Len(Lines(B)) > maxLineLength Then
                
                    ' Search for the next space after maxLineLength and cut & paste it into the new block
                    nextSpace = InStr(maxLineLength, Lines(B), " ")
                    If nextSpace Then
                        newBlock = newBlock & Mid$(Lines(B), 1, nextSpace - 1) & vbNewLine
                        Lines(B) = Mid$(Lines(B), nextSpace + 1)
                    Else
                        newBlock = newBlock & Lines(B)
                        Exit Do
                    End If
                    
                Else
                    newBlock = newBlock & Lines(B)
                    Exit Do
                    
                End If
            Loop
            
            ' Do not attach newlines to last lines but to all other
            If B < UBound(Lines) Then newBlock = newBlock & vbNewLine
            
        Next B
        
        Blocks(A).Text = newBlock
        
    Next A

End Function


' Return a block
Public Function getBlock(ByVal Index As Long) As clsSubBlock

    Set getBlock = Blocks(Index)

End Function


' Load a subtitle file
Public Function openFrom(ByVal fileName As String, Optional ByVal codePage As Long = -1) As Boolean

    Dim openFile As Long
    Dim Data As String
    Dim Result As Boolean
    
    Dim Stream As clsTextFile


    Dbg.addLog DM_Subtitles, "Loading subtitle from " & fileName
    
    closeFile
    
    ' Probe for type of file
    If probeSRT(fileName) Then
        mFormat = FORMAT_SUBRIP
        mFrameBased = False
    ElseIf probeSSA(fileName) Then
        mFormat = FORMAT_SUBSTATION
        mFrameBased = False
    ElseIf probeMDVD(fileName) Then
        mFormat = FORMAT_MICRODVD
        mFrameBased = True
    ElseIf probeSUBV(fileName) Then
        mFormat = FORMAT_SUBVIEWER
        mFrameBased = False
    End If

    If LenB(mFormat) = 0 Then Exit Function

    Dbg.addLog DM_Subtitles, "Format is " & mFormat


    ' Read each line and pass it to the appropriate parser
    Set Stream = New clsTextFile
    If Not Stream.fileOpen(fileName, codePage) Then Exit Function
    Do
        Data = Trim$(Stream.readLine)

        If mFormat = FORMAT_SUBRIP Then Result = parseSRT(Data)
        If mFormat = FORMAT_SUBSTATION Then Result = parseSSA(Data)
        If mFormat = FORMAT_MICRODVD Then Result = parseMDVD(Data)
        If mFormat = FORMAT_SUBVIEWER Then Result = parseSUBV(Data)
        
        If Result = False Then Exit Function
    Loop Until Stream.fileEndReached

    stripTags
    fixLengths

    mFileName = fileName
    If codePage = -1 Then mCodepage = Stream.codePage
    Set Stream = Nothing
    openFrom = True

End Function


' Strip all tags from this subtitle
Private Sub stripTags()

    Dim A As Long
    
    
    Dbg.addLog DM_Subtitles, "Stripping tags from subtitle"
    
    For A = 0 To mBlockCount - 1
        Blocks(A).stripHTML
        Blocks(A).stripRTF
    Next A

End Sub


' Reset subtitle
Public Sub closeFile()

    mFormat = ""
    parseState = NOSTATE
    mFileName = ""
    
    Erase Blocks()
    mBlockCount = 0

End Sub


' Add a new subtitle block
Private Function addBlock() As clsSubBlock

    ReDim Preserve Blocks(mBlockCount)
    mBlockCount = mBlockCount + 1
    Set Blocks(mBlockCount - 1) = New clsSubBlock

    cBlock = mBlockCount - 1
    Set addBlock = Blocks(cBlock)

End Function


' Parse SubViewer subtitle
Private Function parseSUBV(ByVal Data As String) As Boolean

    Dim RE As New RegExp
    
    
    RE.Pattern = PAT_SUBVIEWERTIME
    If RE.Test(Data) = True Then parseState = SUBV_TIME

    ' Time start and end
    If parseState = SUBV_TIME Then
        addBlock
        
        Blocks(cBlock).startTime = parseTime(Mid$(Data, 1, 11))
        Blocks(cBlock).endTime = parseTime(Mid$(Data, 13, 11))
        If Blocks(cBlock).startTime = -1 Or Blocks(cBlock).endTime = -1 Then Exit Function
        
        parseState = SUBV_TEXT
        
    ' Text or blank
    ElseIf parseState = SUBV_TEXT Then
        If LenB(Data) = 0 Then
            parseState = SUBV_NONE
            parseSUBV = True
            Exit Function
        End If
        
        Data = Replace(Data, "[br]", vbNewLine)
        If LenB(Blocks(cBlock).Text) = 0 Then
            Blocks(cBlock).Text = Blocks(cBlock).Text & Data
        Else
            Blocks(cBlock).Text = Blocks(cBlock).Text & vbNewLine & Data
        End If
        
    End If
    
    parseSUBV = True

End Function


' Parse MicroDVD file
Private Function parseMDVD(ByVal Data As String) As Boolean

    Dim Temp() As String
    
    
    If LenB(Data) = 0 Then
        parseMDVD = True
        Exit Function
    End If
    
    Temp = Split(Data, "}")
    If Not UBound(Temp) = 2 Then Exit Function
    
    ' Framerate is indicated
    If Temp(0) = "{1" And Temp(1) = "{1" Then
        mFPS = Val(Temp(2))
    
    ' Add as new block
    Else
        addBlock
        Blocks(cBlock).startFrame = CLng(Mid$(Temp(0), 2))
        Blocks(cBlock).endFrame = CLng(Mid$(Temp(1), 2))
        Blocks(cBlock).Text = Replace(Temp(2), "|", vbNewLine)
        
    End If
    
    parseMDVD = True

End Function


' Parse SubStation file
Private Function parseSSA(ByVal Data As String) As Boolean

    Dim A As Long
    Dim Temp() As String
    

    If Left$(LCase$(Data), 10) = "dialogue: " Then parseState = SSA_DIALOGUE

    ' Dialogue line
    If parseState = SSA_DIALOGUE Then
        addBlock

        Temp = Split(Data, ",")
        If UBound(Temp) < 9 Then Exit Function
        
        Blocks(cBlock).startTime = parseTime(Temp(1))
        Blocks(cBlock).endTime = parseTime(Temp(2))
        If Blocks(cBlock).startTime = -1 Or Blocks(cBlock).endTime = -1 Then Exit Function
        
        ' Add all parts of the line
        For A = 9 To UBound(Temp)
            Blocks(cBlock).Text = Blocks(cBlock).Text & Temp(A)
        Next A
        
        ' Replace \N with newlines
        Blocks(cBlock).Text = Replace(Blocks(cBlock).Text, "\n", vbNewLine, , , vbTextCompare)
        
        parseState = SSA_DIALOGUEEXT
        
    ' 2nd dialogue line
    ElseIf parseState = SSA_DIALOGUEEXT Then
        Blocks(cBlock).Text = Blocks(cBlock).Text & vbNewLine & Data
        parseState = SSA_NONE

    End If
    
    parseSSA = True

End Function


' Parse SubRip file
Private Function parseSRT(ByVal Data As String) As Boolean

    ' Time start and end
    If parseState = SRT_TIME Then
        
        ' Skip empty lines while waiting for timing
        If LenB(Data) = 0 Then
            parseSRT = True
            Exit Function
        End If
    
        addBlock
        Blocks(cBlock).startTime = parseTime(Mid$(Data, 1, 12))
        Blocks(cBlock).endTime = parseTime(Mid$(Data, 18, 12))
        If Blocks(cBlock).startTime = -1 Or Blocks(cBlock).endTime = -1 Then Exit Function

        parseState = SRT_TEXT

    ' Text or blank
    ElseIf parseState = SRT_TEXT Then
        If LenB(Data) = 0 Then
            parseState = SRT_NONE
            parseSRT = True
            Exit Function
        End If
        
        If LenB(Blocks(cBlock).Text) = 0 Then
            Blocks(cBlock).Text = Blocks(cBlock).Text & Data
        Else
            Blocks(cBlock).Text = Blocks(cBlock).Text & vbNewLine & Data
        End If

    ' Numerical lines mark the start of a new subtitle
    ElseIf IsNumeric(Data) Then
        parseState = SRT_TIME
        
    End If
    
    parseSRT = True

End Function


' Parse a time in *0:00:00.0* format to seconds
Private Function parseTime(ByVal Data As String) As Single

    Dim Temp() As String
    Dim Frac As String
    Dim parseData
    

    parseTime = -1

    Temp = Split(Data, ":")
    If Not UBound(Temp) = 2 Then Exit Function
    
    parseData = parseData + (CLng(Temp(0)) * 60 * 60)
    parseData = parseData + (CLng(Temp(1)) * 60)
    parseData = parseData + CLng(Left$(Temp(2), 2))
    
    ' Calculate milliseconds
    Frac = Mid$(Temp(2), 4)
    If Not IsNumeric(Frac) Then Exit Function
    parseTime = parseData + (CLng(Frac) * 0.001)

End Function


' Probe for SuBView subtitles
' Find timecode regexp match
Private Function probeSUBV(ByVal fileName As String) As Boolean

    Dim Data As String
    Dim RE As New RegExp
    Dim Stream As clsTextFile
    
    
    Set Stream = New clsTextFile
    If Not Stream.fileOpen(fileName, False) Then Exit Function
    
    RE.Pattern = PAT_SUBVIEWERTIME
    Do
        Data = Trim$(Stream.readLine)
    Loop Until Stream.fileEndReached Or RE.Test(Data) = True
    
    If Not Stream.fileEndReached Then probeSUBV = True

End Function


' Probe for MicroDVD subtitles
' {number}{number}string
Private Function probeMDVD(ByVal fileName As String) As Boolean
    
    Dim Data As String
    Dim RE As New RegExp
    Dim Stream As clsTextFile
    
    
    Set Stream = New clsTextFile
    If Not Stream.fileOpen(fileName, False) Then Exit Function
    
    RE.Pattern = PAT_MDVDLINE
    Do
        Data = Trim$(Stream.readLine)
    Loop Until Stream.fileEndReached Or RE.Test(Data) = True
    
    If Not Stream.fileEndReached Then probeMDVD = True
    
End Function


' Probe for SubRip subtitles
' First non-empty line is a numerical value
Private Function probeSRT(ByVal fileName As String) As Boolean

    Dim Data As String
    Dim Stream As clsTextFile
    
    
    Set Stream = New clsTextFile
    If Not Stream.fileOpen(fileName, False) Then Exit Function
        
    Do
        Data = Trim$(Stream.readLine)
    Loop Until Stream.fileEndReached Or Data <> ""
    
    If Not Stream.fileEndReached Then
        If IsNumeric(Data) Then probeSRT = True
    End If

End Function


' Probe for SubStation Alpha\ASS subtitles
' [Script info] in first non-empty, non-commented (;) line
Private Function probeSSA(ByVal fileName As String) As Boolean

    Dim Data As String
    Dim Stream As clsTextFile
    
    
    Set Stream = New clsTextFile
    If Not Stream.fileOpen(fileName, False) Then Exit Function
    
    Do
        Data = Trim$(Stream.readLine)
    Loop Until Stream.fileEndReached Or (LenB(Data) <> 0 And Left$(Data, 1) <> ";")

    If Not Stream.fileEndReached Then
        If LCase$(Data) = "[script info]" Then
            probeSSA = True
        ElseIf LCase$(Data) = Chr$(239) & Chr$(187) & Chr$(191) & "[script info]" Then
            probeSSA = True
        End If
    End If

End Function
