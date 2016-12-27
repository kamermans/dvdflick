VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSource"
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
'   File purpose: Media file source detail gatherer.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Type streamInfoType
    Type As enumStreamTypes
    Index As Long
    
    bitRate As Long
    Compression As String
    
    FPS As Single
    Width As Long
    Height As Long
    Image As clsGDIImage
    sourceAR As Single
    pixelAR As Single
    Interlaced As Boolean
    
    Channels As Long
    sampleRate As Long
    Delay As Long
    VBR As Boolean
End Type


Private mDuration As Single
Private mStartTime As Single

Private mFileName As String
Private mTimeStamp As String
Private mTimeIndex As Long

' Filetype
' TODO: Generalize in a string and parse all types
Private mIsAVI As Boolean
Private mIsMatroska As Boolean

Private mStreamCount As Long
Private Streams() As streamInfoType

' Number of times this source is referenced in a project
Public refCount As Long


' Parse FFmpeg's output to retrieve media information from it
Public Function openFile(ByVal myFile As String) As Boolean

    Dim A As Long, B As Long
    Dim Report As clsLog
    Dim cmdLine As String
    Dim AVI As clsAVIDelay
    Dim Stream As clsTextFile
    Dim Ext As String
    Dim Size As Variant
    Dim Rate As Long
    Dim Failed As Boolean
    Dim Reason As String
    Dim realWidth As Long
    
    Dim readLine As String
    Dim Offset As Long
    Dim token As String, Temp As String
    Dim Pieces() As String, tempPieces() As String, nPieces As Long
    
    Dim cStream As Long
    Dim videoOffset As Long
    Dim audioOffset As Long
    Dim fOffset As Long
    Dim indexOffset As Long
    Dim currentIndex As Long
    
    
    Dbg.addLog DM_SourceParser, "Analyzing " & myFile
    
    cmdLine = "-analyzeduration 3M -threads " & cpuInfo.getLogicalCPUCount & " -i " & vbQuote & myFile & vbQuote
    executeToFile APP_PATH & "bin\ffmpeg.exe", cmdLine, TEMP_PATH & "analyse.txt", SM_Nothing, EP_Normal, ""
    If FileLen(TEMP_PATH & "analyse.txt") = 0 Then Exit Function
    
    mStreamCount = 0
    
    ' Parse output from FFmpeg
    Set Stream = New clsTextFile
    If Not Stream.fileOpen(TEMP_PATH & "analyse.txt", False) Then Exit Function
    Do
    
        readLine = Stream.readLine

        ' FFmpeg did not know what to do
        If InStr(readLine, "could not find codec parameters") <> 0 Or InStr(readLine, "Unknown format") <> 0 Then
            Failed = True
            Reason = "Unsupported file format or the file is corrupted."
            Exit Do
        End If
        
        ' Ignore codec parameters error
        If Left$(readLine, 1) = "[" Then readLine = ""
        
        
        ' Detect filetypes
        If InStr(readLine, ", matroska,") Then mIsMatroska = True
        If InStr(readLine, ", avi,") Then mIsAVI = True
        
        ' Stream type\index
        audioOffset = InStr(readLine, "Audio: ")
        videoOffset = InStr(readLine, "Video: ")
        
        indexOffset = InStr(readLine, "Stream #")
        If indexOffset Then currentIndex = Val(Mid$(readLine, indexOffset + 10, 1))
                    
        If audioOffset Or videoOffset Then
            ReDim Preserve Streams(mStreamCount)
            mStreamCount = mStreamCount + 1
            cStream = mStreamCount - 1
            
            Streams(cStream).Index = currentIndex
        End If
        
        
        ' Audio
        If audioOffset Then
            Pieces = Split(Mid$(readLine, audioOffset + 7), ", ")
            nPieces = UBound(Pieces)
            
            Dbg.addLog DM_SourceParser, "Encountered audio stream"
            
            If nPieces >= 2 Then
                Streams(cStream).Type = ST_Audio
                Streams(cStream).Compression = remapAudio(Pieces(0))
                Streams(cStream).sampleRate = CLng(Left$(Pieces(1), Len(Pieces(1)) - 3))
            
                Streams(cStream).Delay = -160
            
                ' Channel count
                If InStr(Pieces(2), "channels") Then
                    Streams(cStream).Channels = CLng(Left$(Pieces(2), Len(Pieces(2)) - 9))
                
                Else
                
                    ' channel_layout_map
                    ' See ffmpeg\libavcodec\audioconvert.c
                    If Pieces(2) = "mono" Then
                        Streams(cStream).Channels = 1
                    ElseIf Pieces(2) = "stereo" Then
                        Streams(cStream).Channels = 2
                    ElseIf Pieces(2) = "surround" Then
                        Streams(cStream).Channels = 3
                    ElseIf Pieces(2) = "quad" Then
                        Streams(cStream).Channels = 4
                    ElseIf Pieces(2) = "5.0" Then
                        Streams(cStream).Channels = 5
                    ElseIf Pieces(2) = "5.1" Then
                        Streams(cStream).Channels = 6
                    ElseIf Pieces(2) = "5.1+downmix" Then
                        Streams(cStream).Channels = 8
                    ElseIf Pieces(2) = "7.1" Then
                        Streams(cStream).Channels = 8
                    ElseIf Pieces(2) = "7.1(wide)" Then
                        Streams(cStream).Channels = 8
                    ElseIf Pieces(2) = "7.1+downmix" Then
                        Streams(cStream).Channels = 10
                    End If
                    
                End If
                
            
                ' BitRate
                If InStr(Pieces(nPieces), "kb/s") <> 0 Then
                    Offset = InStr(Pieces(nPieces), " ")
                    If Offset Then Streams(cStream).bitRate = CLng(Left$(Pieces(nPieces), Offset - 1))
                End If
                
                If unsupportedAudio(Streams(cStream).Compression) Then
                    frmDialog.Display "The audio format " & Streams(cStream).Compression & " in file " & myFile & " is currently unsupported. It will not be added to the project.", Exclamation Or OkOnly
                    
                    mStreamCount = mStreamCount - 1
                    If mStreamCount = 0 Then
                        Erase Streams
                    Else
                        ReDim Preserve Streams(mStreamCount - 1)
                    End If
                    cStream = -1
                End If
            End If
            
        
        ' Video
        ElseIf videoOffset Then
            Dbg.addLog DM_SourceParser, "Encountered video stream"
        
            fOffset = 0
            Pieces = Split(Mid$(readLine, videoOffset + 7), ", ")
            nPieces = UBound(Pieces)
            
            Streams(cStream).Type = ST_Video
            Streams(cStream).Compression = modUtil.remapVideo(Pieces(fOffset))
            fOffset = fOffset + 1
        
            ' Colorspace?
            If colorSpaces.Exists(Pieces(fOffset)) Then fOffset = fOffset + 1
            
            ' Dimensions \ aspect ratio
            Offset = InStr(Pieces(fOffset), "x")
            If Offset Then
                
                ' Check if there is a space, if so there is a PAR after the dimensions
                If InStr(Pieces(fOffset), " ") Then
                    Temp = Left$(Pieces(fOffset), InStr(Pieces(fOffset), " ") - 1)
                Else
                    Temp = Pieces(fOffset)
                End If
                
                ' Get width and height
                Streams(cStream).Width = CLng(Left$(Temp, Offset - 1))
                Streams(cStream).Height = CLng(Mid$(Temp, Offset + 1))
                
                fOffset = fOffset + 1
            End If
            
            ' Pixel aspect ratio
            If InStr(Pieces(fOffset), "PAR") <> 0 Then
                
                Temp = Mid$(Pieces(fOffset), InStr(Pieces(fOffset), " ") + 1)
                tempPieces = Split(Temp, " ")
                Streams(cStream).pixelAR = getARFromString(tempPieces(1))
                If Streams(cStream).pixelAR = 0 Then Streams(cStream).pixelAR = 1
                
                ' FFmpeg does Width = Width * AR, DVD Flick does Width = Height * AR
                If Streams(cStream).pixelAR <> 1 Then
                    realWidth = Streams(cStream).Width * Streams(cStream).pixelAR
                    Streams(cStream).pixelAR = realWidth / Streams(cStream).Height
                End If
                
                fOffset = fOffset + 1
                
                Dbg.addLog DM_SourceParser, "Found aspect ratio " & Streams(cStream).pixelAR & ", " & Streams(cStream).sourceAR
            End If
        
            ' BitRate
            If InStr(Pieces(fOffset), " kb/s") <> 0 Then
                Offset = InStr(Pieces(fOffset), " ")
                If Offset Then Streams(cStream).bitRate = CLng(Left$(Pieces(fOffset), Offset - 1))
                fOffset = fOffset + 1
                
                Dbg.addLog DM_SourceParser, "Found bitrate " & Streams(cStream).bitRate
            End If
                
            ' Framerate
            Offset = InStr(Pieces(fOffset), " tb")
            If Offset Then
                Offset = InStr(Pieces(fOffset), " ")
                If Offset Then Streams(cStream).FPS = Val(Left$(Pieces(fOffset), Offset - 1))
                fOffset = fOffset + 1
                
                Dbg.addLog DM_SourceParser, "Found framerate " & Streams(cStream).FPS
            End If
        
        End If
        
        
        ' Duration
        If Left$(readLine, 12) = "  Duration: " Then
            Offset = InStr(readLine, ",")
            If Offset Then mDuration = parseTime(Mid$(readLine, 13, Offset - 13))
            
            ' Consider files without duration broken or incomplete
            If mDuration = 0 Then
                Failed = True
                Reason = "File may be incomplete or corrupted."
                Exit Do
            End If
            
            Dbg.addLog DM_SourceParser, "Found duration " & mDuration
        End If
        
        ' Starting time
        Offset = InStr(readLine, "start: ")
        If Offset Then
            mStartTime = Val(Mid$(readLine, Offset + 7, InStr(Offset + 7, readLine, ",") - Offset - 7))
        End If
                    
    
    Loop Until Stream.fileEndReached
    Set Stream = Nothing
    
    
    ' Clean up
    killIfExists TEMP_PATH & "analyse.txt"
    If Failed Then
        openFile = False
        Exit Function
    End If
    

    ' Check for invalid streams and remove them
    If mStreamCount > 0 Then
        A = 0
        Do
            If Streams(A).Type = ST_Video And (Streams(A).Width = 0 Or Streams(A).Height = 0 Or Streams(A).FPS > 250) Then
                Dbg.addLog DM_SourceParser, "Removing invalid stream " & A
                For B = A To mStreamCount - 2
                    Streams(B) = Streams(B + 1)
                Next B
                
                mStreamCount = mStreamCount - 1
                If mStreamCount > 0 Then
                    ReDim Preserve Streams(mStreamCount)
                Else
                    Erase Streams
                End If
            End If
            
            A = A + 1
        Loop Until A >= mStreamCount
    End If
    
    ' File is valid at this point
    mFileName = myFile
    

    ' If file is AVI, parse audio delays
'    Set AVI = New clsAVIDelay
'    If isAVI = True And mStreamCount > 0 And AVI.openFile(myFile) Then
'        For A = 0 To AVI.streamCount - 1
'            If AVI.streamIsAudio(A) Then
'                Streams(A).VBR = AVI.streamIsVBR(A)
'                Streams(A).Delay = AVI.streamDelay(A)
'                If Streams(A).Delay = &HFFFFFF Then Streams(A).Delay = 0
'
'                Dbg.addLog DM_SourceParser, "Parsed AVI audio delay for stream " & A & ": " & Streams(A).Delay
'            End If
'        Next A
'    End If

    ' Get possible MPEG information
    Ext = LCase$(FS.GetExtensionName(mFileName))
    If Ext = "vob" Or Ext = "mpg" Or Ext = "mpeg" Or Ext = "m2p" Or Ext = "mpv" Or Ext = "m1v" Or Ext = "m2v" Or Ext = "m2ts" Then
        getMPEGInfo
        
        Size = FS.GetFile(mFileName).Size
        Rate = Size / mDuration
        Rate = CDbl(Rate / 1000) * 8
        
        For A = 0 To mStreamCount - 1
            If Streams(A).Type = ST_Audio Then Rate = Rate - Streams(A).bitRate
        Next A
        For A = 0 To mStreamCount - 1
            If Streams(A).Type = ST_Video Then
                Streams(A).bitRate = Rate
                Exit For
            End If
        Next A
    End If
    
    ' Correct aspect ratios
    For A = 0 To mStreamCount - 1
        If Streams(A).Type = ST_Video Then
            Streams(A).sourceAR = Streams(A).Width / Streams(A).Height
            If Streams(A).pixelAR = 0 Then Streams(A).pixelAR = 1
        End If
    Next A
    

    ' Output report
    Set Report = New clsLog
    Report.Start DATA_PATH & "report.txt"
    Report.Add "Stream information of " & mFileName
    Report.Add "Duration  : " & mDuration
    Report.Add "Starttime : " & mStartTime
    
    For A = 0 To mStreamCount - 1
        
        With Streams(A)
            
            If .Type = ST_Audio Then
                Report.Add "Stream " & A & " (Audio)"
                
                Report.Add "Channels    : " & .Channels, 1
                Report.Add "Samplerate  : " & .sampleRate & " Hz", 1
                Report.Add "Compression : " & .Compression, 1
                
            ElseIf .Type = ST_Video Then
                Report.Add "Stream " & A & " (Video)"
                
                Report.Add "Width       : " & .Width & " pixels", 1
                Report.Add "Height      : " & .Height & " pixels", 1
                Report.Add "Framerate   : " & .FPS, 1
                Report.Add "Compression : " & .Compression, 1
                Report.Add "Source AR   : " & visualAspectRatio(.sourceAR), 1
                Report.Add "Pixel AR    : " & visualAspectRatio(.pixelAR), 1
                Report.Add "Interlaced  : " & CBool(.Interlaced), 1
                
            End If
            
            Report.Add "BitRate     : " & .bitRate & " Kbit\s", 1
            
        End With
        
    Next A

    
    mTimeStamp = FS.GetFile(myFile).DateLastModified
    Set Report = Nothing
    
    openFile = True

End Function


Public Function createStreamThumb(ByVal Index As Long, ByVal timeIndex As Long) As Boolean

    Dim Asp As Single
    Dim thumbWidth As Long
    Dim thumbHeight As Long
    
    
    With Streams(Index)
    
        ' Calculate thumbnail size
        Asp = getResizeValue(.Width, .Height, 340, 256)
        thumbWidth = .Width * Asp
        thumbHeight = .Height * Asp
        thumbWidth = thumbWidth + (thumbWidth Mod 2)
        thumbHeight = thumbHeight + (thumbHeight Mod 2)
        
        ' Get time index
        If timeIndex < 0 Then timeIndex = 0
        If timeIndex > mDuration Then timeIndex = mDuration
        
        ' Grab thumbnail
        Set .Image = getFrameBitmap(mFileName, thumbWidth, thumbHeight, timeIndex)
        If .Image Is Nothing Then Exit Function

    End With
    
    createStreamThumb = True

End Function


' Stream image property
Public Property Get streamImage(ByVal Stream As Long) As clsGDIImage

    Set streamImage = Streams(Stream).Image

End Property

Public Property Set streamImage(ByVal Stream As Long, ByRef newImage As clsGDIImage)

    Set Streams(Stream).Image = newImage

End Property


' Return dictionary with stream information in it
Public Function streamInfo(ByVal Stream As Long) As Dictionary

    Set streamInfo = New Dictionary
    With streamInfo
        
        If Streams(Stream).Type = ST_Audio Then
            .Add "Channels", Streams(Stream).Channels
            .Add "sampleRate", Streams(Stream).sampleRate
            .Add "Delay", Streams(Stream).Delay
            .Add "VBR", Streams(Stream).VBR
            
        ElseIf Streams(Stream).Type = ST_Video Then
            .Add "FPS", Streams(Stream).FPS
            .Add "Width", Streams(Stream).Width
            .Add "Height", Streams(Stream).Height
            .Add "sourceAR", Streams(Stream).sourceAR
            .Add "pixelAR", Streams(Stream).pixelAR
            .Add "Interlaced", Streams(Stream).Interlaced

        End If
        
        .Add "Type", Streams(Stream).Type
        .Add "Index", Streams(Stream).Index
        .Add "Duration", mDuration
        .Add "startTime", mStartTime
        .Add "Compression", Streams(Stream).Compression
        .Add "bitRate", Streams(Stream).bitRate
        
    End With

End Function


' Write to dictionary
Public Function Serialize() As Dictionary

    Dim A As Long
    Dim MyStream As Dictionary
    
    
    Set Serialize = New Dictionary
    
    Serialize.Add "fileName", mFileName
    Serialize.Add "Duration", mDuration
    Serialize.Add "startTime", mStartTime
    Serialize.Add "timeStamp", mTimeStamp
    Serialize.Add "isAVI", mIsAVI
    Serialize.Add "isMatroska", mIsMatroska
        
    For A = 0 To mStreamCount - 1
        Set MyStream = New Dictionary
        
        With Streams(A)
            MyStream.Add "Type", .Type
            MyStream.Add "Index", .Index
            MyStream.Add "Compression", .Compression
            MyStream.Add "bitRate", .bitRate
            MyStream.Add "FPS", .FPS
            MyStream.Add "Width", .Width
            MyStream.Add "Height", .Height
            MyStream.Add "Channels", .Channels
            MyStream.Add "sampleRate", .sampleRate
            MyStream.Add "Delay", .Delay
            MyStream.Add "sourceAR", .sourceAR
            MyStream.Add "pixelAR", .pixelAR
            MyStream.Add "Interlaced", .Interlaced
            MyStream.Add "VBR", .VBR
        End With
        
        Serialize.Add "Stream" & A, MyStream
    Next A

End Function


' Read from dictionary
Public Function unSerialize(ByRef Dict As Dictionary)

    Dim A As Long
    Dim MyStream As Dictionary
    
    
    Reset
    
    mFileName = Dict("fileName")
    mDuration = Dict("Duration")
    mTimeStamp = Dict("timeStamp")
    mIsAVI = Dict("isAVI")
    mIsMatroska = Dict("isMatroska")
    
    For A = 0 To Dict.Count - 1
        
        If Left$(Dict.Keys(A), 6) = "Stream" Then
            Set MyStream = Dict.Items(A)
        
            ReDim Preserve Streams(mStreamCount)
            mStreamCount = mStreamCount + 1
        
            With Streams(mStreamCount - 1)
                .Type = MyStream("Type")
                .Index = MyStream("Index")
                .Compression = MyStream("Compression")
                .bitRate = MyStream("bitRate")
                .FPS = MyStream("FPS")
                .Width = MyStream("Width")
                .Height = MyStream("Height")
                .Channels = MyStream("Channels")
                .sampleRate = MyStream("sampleRate")
                .Delay = MyStream("Delay")
                .sourceAR = MyStream("sourceAR")
                .pixelAR = MyStream("pixelAR")
                .Interlaced = MyStream("Interlaced")
                .VBR = MyStream("VBR")
            End With
            
        End If
        
    Next A

End Function


Public Property Get fileName() As String

    fileName = mFileName

End Property


Public Property Get streamCount() As Long

    streamCount = mStreamCount

End Property


Public Property Get timeStamp() As String

    timeStamp = mTimeStamp

End Property


Public Property Get isAVI() As Boolean

    isAVI = mIsAVI

End Property


Public Property Get isMatroska() As Boolean

    isMatroska = mIsMatroska

End Property


' Retrieve MPEG file information using MPGTX
Private Sub getMPEGInfo()

    Dim A As Long
    Dim cmdLine As String
    Dim Stream As clsTextFile
    
    Dim Data As String
    Dim Have As Long
    Dim Spatial As Long
    Dim AR As String
    
    Dim streamAR As Single
    Dim streamFPS As Single
    Dim streamBitrate As Long
    Dim streamInterlaced As Boolean
    
    
    cmdLine = "-i " & vbQuote & mFileName & vbQuote
    executeToFile APP_PATH & "bin\mpgtx.exe", cmdLine, TEMP_PATH & "mpgtx.txt", SM_Nothing, EP_Normal, ""
    
    Set Stream = New clsTextFile
    Stream.fileOpen TEMP_PATH & "mpgtx.txt", False
    
    Do
        Data = Stream.readLine

        ' FPS
        Have = InStr(Data, " fps")
        If Have Then
            Spatial = InStrRev(Data, " ", Have - 1)
            streamFPS = Val(Mid$(Data, Spatial + 1, Have - Spatial - 1))
        End If

        ' Video bitrate
        Have = InStr(Data, " Mbps")
        If Have <> 0 And InStr(Data, "Muxrate") = 0 Then
            Spatial = InStrRev(Data, " ", Have - 1)
            streamBitrate = Val(Mid$(Data, Spatial + 1, Have - Spatial - 1)) * 1000
        End If
        
        ' Interlaced
        Have = InStr(Data, "interlaced")
        If Have <> 0 And InStr(Data, "Not ") = 0 Then streamInterlaced = True
    
    Loop Until Stream.fileEndReached
    Set Stream = Nothing
    
    
    ' Copy stream info
    For A = 0 To mStreamCount - 1
        If Streams(A).Type = ST_Video Then
            If streamBitrate Then Streams(A).bitRate = streamBitrate
            If streamFPS Then Streams(A).FPS = streamFPS
            If streamAR Then Streams(A).pixelAR = streamAR
            If streamInterlaced Then Streams(A).Interlaced = True
            Exit For
        End If
    Next A
    
    
    killIfExists TEMP_PATH & "mpgtx.txt"
    
End Sub


Private Sub Reset()

    Erase Streams
    mStreamCount = 0
    
    mDuration = 0
    mFileName = ""
    mTimeStamp = ""
    mIsAVI = False

End Sub


' Class
Private Sub Class_Initialize()

    Reset

End Sub

Private Sub Class_Terminate()

    Reset

End Sub
