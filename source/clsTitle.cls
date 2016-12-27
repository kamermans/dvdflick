VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTitle"
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
'   File purpose: Project title
'
Option Explicit
Option Compare Binary
Option Base 0


' Subclasses
Public Videos As New clsObjectArray
Public audioTracks As New clsObjectArray
Public Subtitles As New clsObjectArray

' Chapters
Public chapterInterval As Long
Public chapterCount As Long
Public chapterOnSource As Byte

' Misc.
Public Name As String
Public targetAspect As enumVideoAspects
Public copyTS As Byte
Private mThumbTimeIndex As Long

' Menus
Public Menus As New Dictionary


' Generate menu graphics
Public Sub generateMenus(ByRef Templates As Dictionary, ByRef Captions As Dictionary)

    Dim A As Long
    Dim Menu As clsMenu
    Dim Obj As clsMenuObject
    Dim newButton As clsMenuButton
    
    
    Set Menus = New Dictionary
    
    ' Audio menu
    Set Menu = New clsMenu
    Menu.copyFrom Templates("audioMenu")
    Menus.Add "audio", Menu

    For A = 0 To Menu.Objects.Count - 1
        Set Obj = Menu.Objects.Items(A)
        
        ' Hardcoded title caption
        If Obj.Text = "%titleTitle%" Then Obj.Text = Me.Name
        
        ' Audio item
        If Obj.audioID Then
            If Obj.audioID <= audioTracks.Count Then
                Obj.Text = langCodes.Items(audioTracks.Item(Obj.audioID - 1).Language)
                Set newButton = New clsMenuButton
                Menu.Buttons.Add "Audio" & Obj.audioID, newButton
                newButton.Action = "audio = " & Obj.audioID - 1 & "; resume;"
                newButton.addSurface Obj
            
            Else
                Obj.audioID = 0

            End If
        End If
    Next A
    
    
    ' Subtitle menu
    Set Menu = New clsMenu
    Menu.copyFrom Templates("subtitleMenu")
    Menus.Add "subtitle", Menu

    For A = 0 To Menu.Objects.Count - 1
        Set Obj = Menu.Objects.Items(A)
        
        ' Hardcoded title caption
        If Obj.Text = "%titleTitle%" Then Obj.Text = Me.Name
        
            ' Subtitle item
        If Obj.subID Then
            If Obj.subID <= Subtitles.Count Then
                Obj.Text = langCodes.Items(Subtitles.Item(Obj.subID - 1).Language)
                Set newButton = New clsMenuButton
                Menu.Buttons.Add "Sub" & Obj.subID, newButton
                newButton.Action = "subtitle = " & Obj.subID - 1 + 64 & "; resume;"
                newButton.addSurface Obj

            Else
                Obj.subID = 0

            End If
        End If
    Next A
    
    
    ' Create remaining buttons and captions for all menus
    For A = 0 To Menus.Count - 1
        Menus.Items(A).replaceCaptions Captions
        Menus.Items(A).generateButtons
    Next A

End Sub


' Generate title data fro DVD Author
Public Function generateAuthorData(ByVal titleIndex As Long) As String

    Dim A As Long
    Dim Menu As clsMenu
    Dim XML As String
    
    Dim myTrack As clsAudioTrack
    Dim videoString As String
    Dim Encode As Dictionary
    Dim Chapters As clsChapterList
    Dim cTime As Single
    
    Dim Offset As Long
    Dim menuName As String
    
    
    Set Encode = Me.encodeInfo
    XML = "<titleset>"
    
    ' Generate menu pages
    If Project.menuTemplateName <> STR_DISABLED_MENU Then
        XML = XML & "<menus>"
    
        XML = XML & "<video aspect=""4:3"" "
        If Project.targetFormat = VF_PAL Or Project.targetFormat = VF_MIXED Then
            XML = XML & "format=""pal"""
        Else
            XML = XML & "format=""ntsc"""
        End If
        XML = XML & " />"
            
        For A = 0 To Menus.Count - 1
            Set Menu = Menus.Items(A)
    
            Offset = InStr(Menus.Keys(A), "_")
            If Offset Then
                menuName = Left$(Menus.Keys(A), Offset - 1)
            Else
                menuName = Menus.Keys(A)
            End If
            
            XML = XML & "<pgc entry=""" & menuName & """>" & vbNewLine
            XML = XML & Menu.generateAuthorData
            XML = XML & "<vob file=""" & pathEntities(Project.destinationDir) & "\titlemenu" & titleIndex & "_" & A & "d.mpg"" pause=""inf"" />" & vbNewLine
            XML = XML & "</pgc>" & vbNewLine
        Next A
        
        XML = XML & "</menus>"
        
    End If
    
    
    XML = XML & "<titles>"
    
    ' Audio streams
    For A = 0 To audioTracks.Count - 1
        Set myTrack = audioTracks.Item(A)
        XML = XML & "<audio lang=""" & langCodes.Keys(myTrack.Language) & """ format=""ac3"" channels=""" & myTrack.targetChannels & """ />"
    Next A
    
    ' Subtitle streams
    For A = 0 To Subtitles.Count - 1
        XML = XML & "<subpicture lang=""" & langCodes.Keys(Subtitles.Item(A).Language) & """ />"
    Next A
        
        
    ' Video stream
    ' Format
    videoString = "<video format="""
    Select Case Project.targetFormat
        Case VF_MIXED
            If Encode("Format") = VF_NTSC Then videoString = videoString & "ntsc"""
            If Encode("Format") = VF_NTSCFILM Then videoString = videoString & "ntsc-film"""
            If Encode("Format") = VF_PAL Then videoString = videoString & "pal"""
        Case VF_NTSC, VF_NTSCFILM
            videoString = videoString & "ntsc"""
        Case VF_PAL
            videoString = videoString & "pal"""
    End Select
        
    
    ' Aspect ratio
    If targetAspect = VA_169 Then
        videoString = videoString & " aspect=""16:9"" />"
    Else
        videoString = videoString & " aspect=""4:3"" />"
    End If
    XML = XML & videoString
        
    
    XML = XML & "<pgc>"
            
    ' PRE commands
    XML = XML & "<pre>"
    If Project.enableFirstSub = 1 And Subtitles.Count > 0 Then
        XML = XML & "subtitle=" & 64 & ";"
    Else
        For A = 0 To Subtitles.Count - 1
            If Subtitles.Item(A).displayDefault = 1 Then
                XML = XML & "subtitle = " & A + 64 & ";"
                Exit For
            End If
        Next A
    End If
    
    ' Show subtitle page
    If Project.menuTemplateName <> STR_DISABLED_MENU Then
        If Project.menuShowSubtitleFirst = 1 And Subtitles.Count > 0 Then
            XML = XML & "call menu entry subtitle;"
        End If
        If Project.menuShowAudioFirst = 1 And audioTracks.Count > 1 Then
            XML = XML & "call menu entry audio;"
        End If
    End If
    
    XML = XML & "</pre>"
                    
    ' Video VOB file
    XML = XML & "<vob file=""" & pathEntities(Project.destinationDir) & "\title" & titleIndex & ".mpg"" "
        
        
    ' Chapters
    ' Chapters are present by default at the start (done in getAuthorString for string concat reasons) and end of title
    Set Chapters = New clsChapterList
    Chapters.Duration = Duration
    Chapters.addEntry Duration
    If chapterInterval <> -1 Then Chapters.addByInterval chapterInterval * 60
    If chapterCount <> -1 Then Chapters.addByCount chapterCount

    ' Chapters on sources
    If Videos.Count > 1 Then
        For A = 0 To Videos.Count - 2
            cTime = cTime + Videos.Item(A).streamInfo("Duration")
            Chapters.addEntry cTime
        Next A
    End If
    
    Chapters.Sort
    Chapters.removeDuplicates
    XML = XML & "chapters=""" & Chapters.getAuthorString & """"
    
    XML = XML & " />"
        
        
    ' POST commands
    XML = XML & "<post>"
    If Project.whenPlayed = PA_NextTitle Then
        If titleIndex < Project.Titles.Count - 1 Then
            XML = XML & "g0 = " & titleIndex + 2 & "; call vmgm fpc;"
        ElseIf Project.loopPlayback = 1 Then
            XML = XML & "g0 = 1; call vmgm fpc;"
        End If
        
    ElseIf Project.whenPlayed = PA_SameTitle Then
        XML = XML & "g0 = " & titleIndex + 1 & "; call vmgm fpc;"
        
    ElseIf Project.whenPlayed = PA_Menu Then
        XML = XML & "call vmgm menu;"
        
    End If
    XML = XML & "</post>"
    
    XML = XML & "</pgc>"
    XML = XML & "</titles>"
    XML = XML & "</titleset>"
    
    generateAuthorData = XML

End Function


Public Property Get thumbTimeIndex() As Long

    thumbTimeIndex = mThumbTimeIndex

End Property

Public Property Let thumbTimeIndex(ByVal newValue As Long)

    Dim Video As clsVideo
    Static Busy As Boolean
    
    
    If Busy = True Then Exit Property
    
    Busy = True
    mThumbTimeIndex = newValue
    
    Set Video = Videos.Item(0)
    Video.Source.createStreamThumb Video.streamIndex, mThumbTimeIndex
    Busy = False

End Property


' Returns guesstimated BitRate of subtitle (100 Kbit/s)
Public Function subBitRate() As Long

    subBitRate = 100 * (Me.Subtitles.Count + 1)

End Function


' Calculate the BitRate to encode this title's video at
' If AvgBitRate + audioBitRate is bigger than maximum allowed stream BitRate,
'   function returns maximum allowed stream BitRate minus audio BitRate
Public Function videoBitRate(ByVal avgBitRate As Long) As Long

    Dim Available As Long
    Dim bitRate As Long
    Dim audioBitRate As Long
    Dim myInfo As Dictionary
    Dim encodeInfo As Dictionary
    
    
    Set myInfo = Videos.Item(0).streamInfo
    Set encodeInfo = Me.encodeInfo
    
    ' Don't recompress
    If Project.MPEG2Copy = 1 And myInfo("Compression") = "mpeg2video" And encodeInfo("Width") = myInfo("Width") And encodeInfo("Height") = myInfo("Height") And encodeInfo("bitRate") < MAX_VIDEO_BITRATE Then
        videoBitRate = myInfo("bitRate")
    
    ' Set to custom bitrate
    ElseIf Project.targetBitrate <> TB_Auto Then
        videoBitRate = Project.customBitrate
            
    ' Auto-fit
    ' Substract audio and subtitle bitrate from maximum stream bitrate
    ' Bitrate left is the room to compress video in
    Else
        Available = MAX_STREAM_BITRATE - Me.audioBitRate - Me.subBitRate
        If Available > MAX_VIDEO_BITRATE Then Available = MAX_VIDEO_BITRATE
        
        If avgBitRate < Available Then
            videoBitRate = avgBitRate
        Else
            videoBitRate = Available
        End If
    
    End If
    
End Function


' Return total audio BitRate in kilobits per second
Public Function audioBitRate() As Long

    Dim A As Long
    
    
    For A = 0 To audioTracks.Count - 1
        audioBitRate = audioBitRate + audioTracks.Item(A).targetBitrate
    Next A

End Function


' Save title to dictionary
Public Function Serialize() As Dictionary

    Dim A As Long
    Dim myData As Dictionary
    
    
    Set Serialize = New Dictionary
    
    ' Settings
    Serialize.Add "Name", Name
    Serialize.Add "chapterInterval", chapterInterval
    Serialize.Add "chapterCount", chapterCount
    Serialize.Add "chapterOnSource", chapterOnSource
    Serialize.Add "targetAspect", targetAspect
    Serialize.Add "copyTS", copyTS
    Serialize.Add "thumbTimeIndex", thumbTimeIndex
    
    ' Videos
    Set myData = New Dictionary
    For A = 0 To Videos.Count - 1
        myData.Add "Video" & A, Videos.Item(A).Serialize
    Next A
    Serialize.Add "Videos", myData
    
    ' Audio tracks
    Set myData = New Dictionary
    For A = 0 To audioTracks.Count - 1
        myData.Add "audioTrack" & A, audioTracks.Item(A).Serialize
    Next A
    Serialize.Add "audioTracks", myData

    ' Subtitles
    Set myData = New Dictionary
    For A = 0 To Subtitles.Count - 1
        myData.Add "Subtitle" & A, Subtitles.Item(A).Serialize
    Next A
    Serialize.Add "Subtitles", myData

End Function


' Restore title from dictionary
Public Function unSerialize(ByRef myDict As Dictionary) As Boolean

    Dim A As Long
    Dim myData As Dictionary
    Dim myData2 As Dictionary
    Dim myVideo As clsVideo
    Dim myTrack As clsAudioTrack
    Dim mySub As clsSubtitle
    
    
    unSerialize = False
    Reset
    
    Name = myDict("Name")
    chapterInterval = myDict("chapterInterval")
    chapterCount = myDict("chapterCount")
    chapterOnSource = myDict("chapterOnSource")
    targetAspect = myDict("targetAspect")
    copyTS = myDict("copyTS")
    mThumbTimeIndex = myDict("thumbTimeIndex")
    
    ' Read video sources
    Set myData = myDict("Videos")
    For A = 0 To myData.Count - 1
        Set myVideo = New clsVideo
        Set myData2 = myData.Items(A)
        If myVideo.unSerialize(myData2) Then Videos.Add myVideo
    Next A
    
    ' Any video sources at all?
    If Videos.Count = 0 Then Exit Function
    
    ' Read audio tracks
    Set myData = myDict("audioTracks")
    For A = 0 To myData.Count - 1
        Set myTrack = New clsAudioTrack
        Set myData2 = myData.Items(A)
        If myTrack.unSerialize(myData2) Then audioTracks.Add myTrack
    Next A
    
    ' Read Subtitles
    Set myData = myDict("Subtitles")
    For A = 0 To myData.Count - 1
        Set mySub = New clsSubtitle
        Set myData2 = myData.Items(A)
        If mySub.unSerialize(myData2) Then Subtitles.Add mySub
    Next A
    
    unSerialize = True
    
End Function


' Add a new video source
Public Function addVideo(ByRef Source As clsSource, ByVal Stream As Long) As clsVideo

    Dim Video As clsVideo
    

    Set Video = New clsVideo
    If Not Video.openFrom(Source, Stream) Then
        Set Video = Nothing
        Exit Function
    End If
    
    Set addVideo = Video
    Videos.Add addVideo

End Function


' Add a subtitle file object to this title
Public Function addSub(ByRef subFile As clsSubFile) As clsSubtitle

    Dim mySub As clsSubtitle
    
    
    ' Add
    Set mySub = New clsSubtitle
    Me.Subtitles.Add mySub
    
    ' Set properties
    mySub.fileName = subFile.fileName
    If subFile.frameBased = True Then
        mySub.frameBased = 1
    Else
        mySub.frameBased = 0
    End If
    mySub.fileFormat = subFile.fileFormat
    mySub.FPS = Videos.Item(0).streamInfo("FPS")
    If subFile.FPS <> 0 Then mySub.FPS = subFile.FPS
    mySub.codePage = subFile.codePage

End Function


' Scan a source file for usable streams and add them to the title
Public Function scanFile(ByRef Source As clsSource) As Boolean

    Dim A As Long
    Dim Info As Dictionary
    Dim myVideo As clsVideo
    Dim myTrack As clsAudioTrack
    Dim AR As Single
    
    
    scanFile = True
    
    For A = 0 To Source.streamCount - 1
        
        Set Info = Source.streamInfo(A)
        
        ' Add video streams
        If Info("Type") = ST_Video Then
            Set myVideo = Me.addVideo(Source, A)
            
            If myVideo Is Nothing Then
                scanFile = False
                Exit Function
            End If
            
            ' Find best matching target aspect ratio
            If Info("pixelAR") <> 1 Then
                AR = (Info("Height") * Info("pixelAR")) / Info("Height")
            Else
                AR = Info("sourceAR")
            End If
            If AR >= 1.56 Then
                targetAspect = VA_169
            Else
                targetAspect = VA_43
            End If
            
            ' Set thumbnail time index and generate it
            Me.thumbTimeIndex = 0.1 * myVideo.streamInfo("Duration")
            If myVideo.Thumbnail Is Nothing Then frmDialog.Display "No thumbnail could be generated for the file " & Source.fileName & ". This usually means the file is damaged or truncated.", Exclamation Or OkOnly
                        
        ' Add AudioTracks
        ElseIf Info("Type") = ST_Audio Then
            Set myTrack = Me.addTrack(Source, A)
            If Not myTrack Is Nothing Then
                myTrack.Language = Config.ReadSetting("titleAudioLanguage", myTrack.Language)
            End If
        
        End If
        
    Next A
    
    Set myVideo = Nothing
    Set Info = Nothing

End Function


' Return dictionary with information used to encode this title
Public Function encodeInfo() As Dictionary

    Dim FPS As Single
    
    
    Set encodeInfo = New Dictionary
    FPS = Videos.Item(0).streamInfo("FPS")
    
    ' Mixed mode
    If Project.targetFormat = VF_MIXED Then
        
        ' NTSC
        If FPS > 27 Then
            encodeInfo.Add "Format", VF_NTSC
            encodeInfo.Add "FPS", 29.97
            encodeInfo.Add "Height", 480
            encodeInfo.Add "GOP", GOPSIZE_NTSC
            encodeInfo.Add "Pulldown", 0
            
        ' NTSC-Film to NTSC
        ElseIf FPS <= 24 Then
            encodeInfo.Add "Format", VF_NTSC
            encodeInfo.Add "Height", 480
            encodeInfo.Add "GOP", GOPSIZE_NTSC
            
            If Project.Pulldown = 1 Then
                encodeInfo.Add "FPS", 23.976
                encodeInfo.Add "Pulldown", 1
            Else
                encodeInfo.Add "FPS", 29.97
                encodeInfo.Add "Pulldown", 0
            End If
            
        ' PAL
        Else
            encodeInfo.Add "Format", VF_PAL
            encodeInfo.Add "FPS", 25
            encodeInfo.Add "Height", 576
            encodeInfo.Add "GOP", GOPSIZE_PAL
            encodeInfo.Add "Pulldown", 0
        End If
        
    ' NTSC
    ElseIf Project.targetFormat = VF_NTSC Then
        encodeInfo.Add "Format", VF_NTSC
        encodeInfo.Add "Height", 480
        encodeInfo.Add "GOP", GOPSIZE_NTSC
        
        If FPS = 23.976 And Project.Pulldown = 1 Then
            encodeInfo.Add "FPS", 23.976
            encodeInfo.Add "Pulldown", 1
        Else
            encodeInfo.Add "FPS", 29.97
            encodeInfo.Add "Pulldown", 0
        End If
    
    ' NTSC-Film
    ElseIf Project.targetFormat = VF_NTSCFILM Then
        encodeInfo.Add "Format", VF_NTSCFILM
        encodeInfo.Add "FPS", 23.976
        encodeInfo.Add "Height", 480
        encodeInfo.Add "GOP", GOPSIZE_NTSC
        encodeInfo.Add "Pulldown", 0
    
    ' PAL
    ElseIf Project.targetFormat = VF_PAL Then
        encodeInfo.Add "Format", VF_PAL
        encodeInfo.Add "FPS", 25
        encodeInfo.Add "Height", 576
        encodeInfo.Add "GOP", GOPSIZE_PAL
        encodeInfo.Add "Pulldown", 0
    
    End If
    
    ' Half horizontal resolution
    If Project.halfRes Then encodeInfo.Add "Width", 352 Else encodeInfo.Add "Width", 720

End Function


' Add a new audio track
Public Function addTrack(ByRef Source As clsSource, ByVal streamIndex As Long) As clsAudioTrack

    Dim Track As New clsAudioTrack
        
    
    If Track.addSource(Source, streamIndex) Is Nothing Then
        Set Track = Nothing
        Exit Function
    End If

    Set addTrack = Track
    audioTracks.Add Track

End Function


' Return total duration of title
Public Function Duration() As Single

    Dim A As Long
    
    
    For A = 0 To Videos.Count - 1
        Duration = Duration + Videos.Item(A).streamInfo("Duration")
    Next A

End Function


' Class
Private Sub Class_Initialize()

    Reset

End Sub

Private Sub Class_Terminate()

    Reset

End Sub


' Reset class data
Private Sub Reset()

    Dim A As Long
    
    
    Videos.Clear
    audioTracks.Clear
    Subtitles.Clear
    
    Name = "Unnamed title"
    chapterInterval = Default_ChapterInterval
    chapterCount = Default_ChapterCount
    chapterOnSource = Default_ChapterOnSource
    targetAspect = VA_43
    mThumbTimeIndex = 0
    copyTS = 0

End Sub


' Move a video upwards in the list
Public Sub moveVideoUp(ByVal Index As Long)

    Dim A As Long, B As Long
    Dim myAudio As clsAudio
    
    
    ' Also move any audio track sources with the same source file
    For A = 0 To audioTracks.Count - 1
        For B = 0 To audioTracks.Item(A).Sources.Count - 1
            Set myAudio = audioTracks.Item(A).Sources.Item(B)
            
            If myAudio.Source.fileName = Videos.Item(Index).Source.fileName Then
                audioTracks.Item(A).MoveSourceUp B
            End If
        Next B
    Next A
    
    Videos.moveBackward Index

End Sub


' Move a video downwards in the list
Public Sub moveVideoDown(ByVal Index As Long)

    Dim A As Long, B As Long
    Dim myAudio As clsAudio
    
    
    ' Also move any audio track sources with the same source file
    For A = 0 To audioTracks.Count - 1
        For B = audioTracks.Item(A).Sources.Count - 1 To 0 Step -1
            Set myAudio = audioTracks.Item(A).Sources.Item(B)
            
            If myAudio.Source.fileName = Videos.Item(Index).Source.fileName Then
                audioTracks.Item(A).MoveSourceDown B
            End If
        Next B
    Next A
    
    Videos.moveForward Index

End Sub
