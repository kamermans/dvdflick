VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
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
'   File purpose: A DVD Flick main project class
'
'
'   Project class structure
'   () = array
'
'    Project
'       |
'        ---> Title()
'       |       |
'       |        ---> Video source()
'       |       |
'       |        ---> Audio track()
'       |       |           |
'       |       |            ---> Audio source()
'       |       |
'       |        ---> Subtitle()
'       |
'        ---> Source()
'       |
'        ---> Menu data
'
Option Explicit
Option Compare Binary
Option Base 0


' General
Public Title As String
Public targetFormat As enumVideoEncodeFormats
Public encodeProfile As enumVideoEncodeProfiles
Public destinationDir As String
Public keepFiles As Byte
Public threadCount As Long

' Playback
Public loopPlayback As Byte
Public whenPlayed As enumWhenPlayedActions
Public enableFirstSub As Byte

' Video
Public targetSize As enumTargetSizes
Public customSize As Long
Public encodePriority As enumEncodePriorities
Public targetBitrate As enumTargetBitRates
Public customBitrate As Long
Public Pulldown As Byte

' Audio
Public volumeMod As Long
Public channelCount As enumAudioChannels
Public audioBitRate As enumAudioBitRates

' Video advanced
Public PSNR As Byte
Public halfRes As Byte
Public overscanBorders As Byte
Public overscanSize As Single
Public Deinterlace As Byte
Public MPEG2Copy As Byte
Public dcPrecision As Byte

' Burning
Public createISO As Byte
Public discLabel As String
Public burnerName As String
Public deleteISO As Byte
Public eraseRW As Byte
Public enableBurning As Byte
Public burnSpeed As String
Public verifyDisc As Byte
Public ejectTray As Byte

' Menu
Public menuTemplateName As String
Public menuAutoPlay As Byte
Public menuShowSubtitleFirst As Byte
Public menuShowAudioFirst As Byte

' Project State
Public Modified As Boolean
Public fileName As String

' Titles present
Public Titles As New clsObjectArray

' Source file pool
Private Sources As Dictionary

' Menus
Public Menus As New Dictionary

' Version
Private Const PROJECT_VERSION As Long = 1


' Render all menus
Public Sub renderMenus(ByRef Template As clsMenuTemplate)

    Dim A As Long, B As Long
    Dim myTitle As clsTitle


    ' Render menus
    For A = 0 To Menus.Count - 1
        Dbg.addLog DM_Menus, "Starting rendering of Project menu " & Menus.Keys(A)
        Menus.Items(A).Render Template
    Next A
    For A = 0 To Titles.Count - 1
        Set myTitle = Titles.Item(A)
        For B = 0 To myTitle.Menus.Count - 1
            Dbg.addLog DM_Menus, "Starting rendering of Title menu " & Menus.Keys(B)
            myTitle.Menus.Items(B).Render Template
        Next B
    Next A

End Sub


' Generate the menus for this project
Public Sub generateMenus(ByRef Templates As Dictionary)

    Dim A As Long
    Dim Menu As clsMenu
    Dim Title As clsTitle
    Dim Captions As Dictionary
    Dim Obj As clsMenuObject
    Dim myThumb As clsGDIImage
    Dim newButton As clsMenuButton
    Dim buttonAction As String
    Dim pageTitles As Long, curTitle As Long, curMenu As Long, Index As Long
    
    
    ' Generate captions
    Set Captions = New Dictionary
    With Captions
        .Add "%projectTitle%", Me.Title
    End With
    
    
    Set Menus = New Dictionary
    
    ' Main menu
    Set Menu = New clsMenu
    Menu.copyFrom Templates("mainMenu")
    Menus.Add "title", Menu
    
    ' Title selection menus
    Do
        Set Menu = New clsMenu
        Menu.copyFrom Templates("titleMenu")
        
        ' Create contents of thumbnails
        pageTitles = 0
        For A = 0 To Menu.Objects.Count - 1
            Set Obj = Menu.Objects.Items(A)
            
            ' Thumbnail
            If Obj.thumbID Then
                If curTitle + Obj.thumbID <= Titles.Count Then

                    Set Title = Project.Titles.Item((curTitle + Obj.thumbID - 1))

                    ' Set object properties to match that of the current title
                    If Obj.typeName = "text" Then Obj.Text = Title.Name
                    If Obj.typeName = "image" And LenB(Obj.Image) = 0 Then
                        Set myThumb = Title.Videos.Item(0).Thumbnail
                        If myThumb Is Nothing Then
                            Set Obj.imageLD = New clsGDIImage
                            Obj.imageLD.copyFrom noThumb
                            Obj.imageLD.Resize 192, 144, 32
                        Else
                            Set Obj.imageLD = modUtil.resizeToMatch(myThumb, myThumb.Width, myThumb.Height, Title.Videos.Item(0).PAR)
                        End If
                        Obj.imageRef = True
                    End If

                    ' If this is a new thumbnail add a button for it too
                    If Obj.thumbID > pageTitles Then
                        pageTitles = Obj.thumbID
                        Set newButton = New clsMenuButton
                        Menu.Buttons.Add "Thumb" & Obj.thumbID, newButton
                        newButton.Action = "g0 = " & curTitle + Obj.thumbID & "; jump vmgm fpc;"
                    End If

                ' This thumbnail isn't going to be used anymore, so object should "hide"
                Else
                    Obj.thumbID = 0
                    Obj.Text = vbNullString
                    Obj.Image = vbNullString
                    Obj.selectImage = vbNullString
                
                End If
            End If
        Next A
        
        ' Remove previous button from first menu
        If curMenu = 0 Then Menu.Objects.Remove "buttonPrevious"
        
        
        ' Create previous\next button actions
        For A = 0 To Menu.Objects.Count - 1
            Set Obj = Menu.Objects.Items(A)
            
            buttonAction = ""
            If Obj.Name = "buttonNext" Then buttonAction = "jump vmgm menu " & Menus.Count + 2 & ";"
            If Obj.Name = "buttonPrevious" Then buttonAction = "jump vmgm menu " & Menus.Count & ";"
            
            If LenB(buttonAction) <> 0 Then
                Set newButton = New clsMenuButton
                Menu.Buttons.Add Obj.Name, newButton
                newButton.Action = buttonAction
                newButton.addSurface Obj
            End If
        Next A
        
        
        ' Advance to next titles menu
        curTitle = curTitle + pageTitles
        curMenu = curMenu + 1
        Menus.Add "titleselect_" & curMenu, Menu
        
    Loop Until curTitle >= Titles.Count
    
    ' Delete next button from last menu
    Menu.Objects.Remove "buttonNext"
    Menu.Buttons.Remove "buttonNext"
    
    
    ' Create remaining buttons and captions for all menus
    For A = 0 To Menus.Count - 1
        Menus.Items(A).replaceCaptions Captions
        Menus.Items(A).generateButtons
    Next A
    
    
    ' Generate title menus
    For A = 0 To Titles.Count - 1
        Set Title = Titles.Item(A)
        Title.generateMenus Templates, Captions
    Next A

End Sub


' Generate spumux XML files for each menu
Public Function generateMenuFiles() As Boolean

    Dim A As Long, B As Long
    Dim myTitle As clsTitle
    

    For A = 0 To Menus.Count - 1
        If Not Menus.Items(A).generateFiles("root", A) Then Exit Function
    Next A
    For A = 0 To Titles.Count - 1
        Set myTitle = Titles.Item(A)
        For B = 0 To myTitle.Menus.Count - 1
            If Not myTitle.Menus.Items(B).generateFiles("title", B, A) Then Exit Function
        Next B
    Next A
    
    FS.DeleteFile Project.destinationDir & "\*.png"
    generateMenuFiles = True
    
End Function


' Rescale all the menus to DVD dimensions
Public Sub rescaleMenus()

    Dim A As Long, B As Long
    Dim myTitle As clsTitle
    Dim modifX As Single, modifY As Single
    
    
    modifX = 720 / MENU_WIDTH
    modifY = Me.menuHeight / MENU_HEIGHT

    For A = 0 To Menus.Count - 1
        Menus.Items(A).Rescale modifX, modifY
    Next A
    
    For A = 0 To Titles.Count - 1
        Set myTitle = Titles.Item(A)
        For B = 0 To myTitle.Menus.Count - 1
            myTitle.Menus.Items(B).Rescale modifX, modifY
        Next B
    Next A

End Sub


' Generate dvdauthor XML data
Public Function generateAuthorData() As String

    Dim A As Long
    Dim XML As String
    
    Dim Offset As Long
    Dim menuName As String
    
    
    ' VMGM area, create empty menu always, for compatibility purposes
    XML = XML & "<vmgm>"
    
    
    ' Title jumping handler
    XML = XML & "<fpc>"
    For A = 0 To Titles.Count - 1
        XML = XML & "if (g0 == " & A + 1 & ") { jump title " & A + 1 & "; } "
    Next A
    If Project.menuTemplateName <> STR_DISABLED_MENU And menuAutoPlay Then
        XML = XML & "jump vmgm menu entry title;"
    Else
        XML = XML & "jump title 1;"
    End If
    XML = XML & "</fpc>"
    
    
    ' English as hardcoded language
    XML = XML & "<menus lang=""en"">"

    XML = XML & "<video aspect=""4:3"" "
    If targetFormat = VF_PAL Or targetFormat = VF_MIXED Then
        XML = XML & "format=""pal"""
    Else
        XML = XML & "format=""ntsc"""
    End If
    XML = XML & " />"
        
    If Project.menuTemplateName <> STR_DISABLED_MENU Then
        
        For A = 0 To Menus.Count - 1
        
            ' Parse for menu entry types
            ' Only root allowed by dvdauthor for now
            Offset = InStr(Menus.Keys(A), "_")
            If Offset Then
                menuName = Left$(Menus.Keys(A), Offset - 1)
            Else
                menuName = Menus.Keys(A)
            End If
            If menuName = "title" Then
                XML = XML & "<pgc entry=""" & menuName & """>" & vbNewLine
            Else
                XML = XML & "<pgc>" & vbNewLine
            End If
            
            XML = XML & Menus.Items(A).generateAuthorData
            XML = XML & "<vob file=""" & pathEntities(destinationDir) & "\rootmenu_" & A & "d.mpg"" pause=""inf"" />" & vbNewLine
            XML = XML & "</pgc>" & vbNewLine
        Next A
    
    End If
    
    XML = XML & "</menus>"
    
    XML = XML & "</vmgm>"
    
    
    generateAuthorData = XML

End Function


' Calculate sizes of some project statistics
Public Function calculateSizes() As Dictionary

    Dim A As Long
    Dim Size As Long
    Dim discSize As Long
    Dim audioRate As Long
    Dim videoRate As Long
    Dim subRate As Long
    Dim avgBitRate As Long
    

    Set calculateSizes = New Dictionary
    
    ' All sizes are in Kilobytes
    With calculateSizes
    
        ' Full discsize
        discSize = Project.customSize * 1024
        .Add "discSize", discSize
        
        ' Room for muxing overhead is 4% (which is a lot for a 4.3 Gb DVD)
        'discSize = 0.96 * discSize
        
        ' Space for IFO and BUP files (128k for each title, 64 for VIDEO_TS)
        'discSize = discSize - (64 + (Titles.Count * 128))
        discSize = discSize - (32 + (Titles.Count * 64))

        
        ' Auto-fit video BitRate
        If Me.targetBitrate = TB_Auto Then
            If Me.Duration > 0 Then
            
                ' Substract audio, subtitle and menu size from total disc size
                ' Divide remaining space by duration to get kilobytes per second
                ' Multiply by 8 to get kilobits per second
                ' Result is default video BitRate of all titles on disc
                ' clsTitle.TargetBitRate can return different (lower) BitRate if
                '   video BitRate + audio BitRate is bigger than maximum allowed
                '   stream BitRate
                avgBitRate = CLng(((discSize - Me.getAudioSize - Me.getSubSize - Me.getMenuSize) / Me.Duration) * 8)

                If avgBitRate > MAX_VIDEO_BITRATE Then avgBitRate = MAX_VIDEO_BITRATE
                If avgBitRate < MIN_VIDEO_BITRATE Then avgBitRate = MIN_VIDEO_BITRATE

            Else
                avgBitRate = 0
                
            End If
            
        ' User video BitRate, just assume value entered
        ' clsTitle.TargetBitRate can still modify this on a per-title basis
        Else
            avgBitRate = Me.customBitrate

        End If

        ' Calculate total size in use
        Size = Me.getAudioSize + Me.getSubSize + Me.getMenuSize + Me.getVideoSize(avgBitRate)
        
        
        .Add "sizeUsed", Size
        .Add "avgBitRate", avgBitRate
        
    End With

End Function


' Return total duration in seconds of all titles combined
Public Function Duration() As Single

    Dim A As Long
    
    
    For A = 0 To Titles.Count - 1
        Duration = Duration + Titles.Item(A).Duration
    Next A

End Function


' Return total estimated size of menus
Public Function getMenuSize() As Long

    Dim A As Long
    
    Const PAGE_SIZE As Long = 2048
    
    
    If Project.menuTemplateName = STR_DISABLED_MENU Then Exit Function
    
    getMenuSize = getMenuSize + PAGE_SIZE + (PAGE_SIZE * Titles.Count)
    For A = 0 To Titles.Count - 1
        getMenuSize = getMenuSize + (PAGE_SIZE * 2)
    Next A

End Function


' Return total size of all audio streams combined in kilobyte
Public Function getAudioSize() As Long

    Dim A As Long
    Dim audioRate As Long
    
    
    For A = 0 To Titles.Count - 1
        audioRate = ((Titles.Item(A).audioBitRate * 1000) / 8) / 1024
        getAudioSize = getAudioSize + (audioRate * Titles.Item(A).Duration)
    Next A

End Function


' Return total guessed size of all subtitle streams combined in kilobyte
Public Function getSubSize() As Long

    Dim A As Long
    Dim subRate As Long
    
    
    For A = 0 To Titles.Count - 1
        subRate = ((Titles.Item(A).subBitRate * 1000) / 8) / 1024
        getSubSize = getSubSize + (subRate * Titles.Item(A).Duration) * Titles.Item(A).Subtitles.Count
    Next A

End Function


' Return total guessed size of video
Public Function getVideoSize(ByVal avgRate As Long) As Long

    Dim A As Long
    Dim videoRate As Long
    
    
    For A = 0 To Titles.Count - 1
        videoRate = ((Titles.Item(A).videoBitRate(avgRate) * 1000) / 8) / 1024
        getVideoSize = getVideoSize + (videoRate * Titles.Item(A).Duration)
    Next A

End Function


' Count sources' reference count
Private Sub setSourceRefCount()

    Dim A As Long
    Dim B As Long
    Dim C As Long
    
    Dim mySource As clsSource
    Dim myTitle As clsTitle
    Dim myVideo As clsVideo
    Dim myTrack As clsAudioTrack
    Dim myAudio As clsAudio
    
    
    For A = 0 To Sources.Count - 1
        Set mySource = Sources.Items(A)
        mySource.refCount = 0
    Next A
    
    For A = 0 To Titles.Count - 1
        Set myTitle = Titles.Item(A)
        
        For B = 0 To myTitle.Videos.Count - 1
            Set myVideo = myTitle.Videos.Item(B)
            myVideo.Source.refCount = myVideo.Source.refCount + 1
        Next B
        
        For B = 0 To myTitle.audioTracks.Count - 1
            Set myTrack = myTitle.audioTracks.Item(B)
            
            For C = 0 To myTrack.Sources.Count - 1
                Set myAudio = myTrack.Sources.Item(C)
                myAudio.Source.refCount = myAudio.Source.refCount + 1
            Next C
        Next B
    Next A
    
End Sub


' Return a source interface for a file to be loaded
Public Function getSource(ByVal myFileName As String) As clsSource

    ' Only if file exists
    myFileName = myFileName
    If Not FS.FileExists(myFileName) Then
        Set getSource = Nothing
        Exit Function
    End If
    
    ' Return existing from pool
    If Sources.Exists(myFileName) Then
        Set getSource = Sources(myFileName)
        
        ' Reload if file modify time has changed
        If getSource.timeStamp <> FS.GetFile(myFileName).DateLastModified Then
            Sources.Remove myFileName
        Else
            Exit Function
        End If
    End If
        
        
    ' Load new
    frmStatus.setStatus "Loading " & FS.GetFileName(myFileName)
    appLog.Add "Loading source " & myFileName
    Set getSource = New clsSource
    
    ' Attempt to load
    If Not getSource.openFile(myFileName) Then
        Set getSource = Nothing
        
    ' Add to pool for future referencing
    Else
        Sources.Add myFileName, getSource
        
    End If

End Function


' Add a new title
Public Function addTitle() As clsTitle

    If Titles.Count = 50 Then
        frmDialog.Display "Maximum number of titles reached.", OkOnly Or Exclamation
        
        Set addTitle = Nothing
        Exit Function
    End If
    
    Set addTitle = New clsTitle
    Titles.Add addTitle

End Function


' Reset all project info and titles
Public Sub Reset()

    Dim A As Long
    
    
    Titles.Clear
    
    If Not (Sources Is Nothing) Then
        For A = 0 To Sources.Count - 1
            Set Sources.Items(A) = Nothing
        Next A
    End If
    Set Sources = New Dictionary
    Sources.CompareMode = TextCompare
    
    With Config
        
        ' General
        Title = .ReadSetting("projectTitle", Default_ProjectTitle)
        targetSize = .ReadSetting("targetSize", Default_TargetSize)
        customSize = .ReadSetting("customSize", Default_CustomSize)
        encodePriority = .ReadSetting("encodePriority", Default_EncodePriority)
        threadCount = .ReadSetting("threadCount", Default_ThreadCount)
        keepFiles = .ReadSetting("keepFiles", Default_KeepFiles)
        
        ' Video
        targetFormat = .ReadSetting("targetFormat", Default_TargetFormat)
        encodeProfile = .ReadSetting("encodingProfile", Default_EncodeProfile)
        targetBitrate = .ReadSetting("targetBitRate", Default_TargetBitRate)
        customBitrate = .ReadSetting("customBitRate", Default_CustomBitRate)
        Pulldown = .ReadSetting("allowPulldown", Default_Pulldown)
        
        ' Audio
        volumeMod = .ReadSetting("volumeMod", Default_VolumeMod)
        channelCount = .ReadSetting("channelCount", Default_ChannelCount)
        audioBitRate = .ReadSetting("audioBitRate", Default_AudioBitRate)
        
        ' Playback
        loopPlayback = .ReadSetting("loopPlayback", Default_LoopPlayback)
        whenPlayed = .ReadSetting("whenPlayed", Default_WhenPlayed)
        enableFirstSub = .ReadSetting("enableFirstSub", Default_EnableFirstSub)
        
        ' Video advanced
        PSNR = .ReadSetting("PSNR", Default_PSNR)
        halfRes = .ReadSetting("halfRes", Default_HalfRes)
        overscanBorders = .ReadSetting("overscanBorders", Default_OverscanBorders)
        overscanSize = .ReadSetting("overscanSize", Default_OverscanSize)
        Deinterlace = .ReadSetting("Deinterlace", Default_Deinterlace)
        MPEG2Copy = .ReadSetting("MPEG2Copy", Default_MPEG2Copy)
        dcPrecision = .ReadSetting("dcPrecision", Default_DCPrecision)
        
        ' Burning
        createISO = .ReadSetting("createISO", Default_CreateISO)
        discLabel = .ReadSetting("discLabel", Default_DiscLabel)
        burnerName = .ReadSetting("burnerName", Default_BurnerName)
        deleteISO = .ReadSetting("deleteISO", Default_DeleteISO)
        eraseRW = .ReadSetting("eraseRW", Default_EraseRW)
        enableBurning = .ReadSetting("enableBurning", Default_EnableBurning)
        burnSpeed = .ReadSetting("burnSpeed", Default_BurnSpeed)
        verifyDisc = .ReadSetting("verifyDisc", Default_VerifyDisc)
        ejectTray = .ReadSetting("ejectTray", Default_EjectTray)
        
        ' Menu
        menuTemplateName = .ReadSetting("menuTemplateName", Default_menuTemplateName)
        menuAutoPlay = .ReadSetting("menuAutoPlay", Default_menuAutoPlay)
        menuShowSubtitleFirst = .ReadSetting("menuShowSubtitleFirst", Default_menuShowSubtitleFirst)
        menuShowAudioFirst = .ReadSetting("menuShowAudioFirst", Default_menuShowAudioFirst)
        
        destinationDir = .ReadSetting("lastOutputDir", Default_LastOutputDir)
        
    End With
    
    Modified = False
    fileName = ""
    
End Sub


' Read a project from a file
Public Function unSerialize(ByVal myFileName As String) As Boolean

    Dim A As Long, B As Long
    Dim tempPath As String
    Dim picFile As String
    Dim oldSourceFile As String
    
    Dim myConf As clsConfiguration
    Dim myDict As Dictionary
    Dim myData As Dictionary
    Dim myData2 As Dictionary
    Dim newSource As clsSource
    Dim newTitle As clsTitle
    Dim newSources As Dictionary
    
    Dim Ver As clsTextFile
    Dim cVersion As Long
    

    If Not FS.FileExists(myFileName) Then
        frmDialog.Display "The project file " & myFileName & " does not exist.", OkOnly Or Exclamation
        Exit Function
    End If

    myFileName = FS.GetFile(myFileName).Path
    tempPath = TEMP_PATH & "project"

    ' Unarchive
    If FS.FolderExists(tempPath) Then FS.DeleteFolder tempPath, True
    FS.CreateFolder tempPath
    folderFromZip myFileName, tempPath
    
    ' Valid project file
    If Not FS.FileExists(tempPath & "\project.cfg") Then Exit Function

    ' Check version
    If FS.FileExists(tempPath & "\version.txt") Then
        Set Ver = New clsTextFile
        Ver.fileOpen tempPath & "\version.txt"
        cVersion = CLng(Ver.readLine)
        Ver.fileClose
        
    Else
        cVersion = 0
        
    End If

    If cVersion < PROJECT_VERSION And Not unattendMode Then
        If frmDialog.Display("The project file you are trying to load is from an older version of DVD Flick. If you are sure you want to load it, please check all of it's settings. It is not guaranteed that it will produce a working DVD." & vbNewLine & vbNewLine & "Are you sure you want to load it?", Exclamation Or YesNo) = buttonNo Then
            FS.DeleteFolder tempPath, True
            unSerialize = True
            Exit Function
        End If
    End If

    Reset

    ' Read configuration from file
    Set myConf = New clsConfiguration
    myConf.LoadConfiguration tempPath & "\project.cfg"
    Set myDict = myConf.Root(True)
    
    ' General
    Title = getDictValue(myDict, "Title", Default_ProjectTitle)
    customSize = getDictValue(myDict, "customSize", Default_CustomSize)
    encodePriority = getDictValue(myDict, "encodePriority", Default_EncodePriority)
    threadCount = getDictValue(myDict, "threadCount", Default_ThreadCount)
    keepFiles = getDictValue(myDict, "keepFiles", Default_KeepFiles)
    
    ' Playback
    loopPlayback = getDictValue(myDict, "loopPlayback", Default_LoopPlayback)
    whenPlayed = getDictValue(myDict, "whenPlayed", Default_WhenPlayed)
    enableFirstSub = getDictValue(myDict, "enableFirstSub", Default_EnableFirstSub)
        
    ' Video
    targetFormat = getDictValue(myDict, "targetFormat", Default_TargetFormat)
    encodeProfile = getDictValue(myDict, "encodeProfile", Default_EncodeProfile)
    targetBitrate = getDictValue(myDict, "targetBitrate", Default_TargetBitRate)
    customBitrate = getDictValue(myDict, "customBitrate", Default_CustomBitRate)
    Pulldown = getDictValue(myDict, "allowPulldown", Default_Pulldown)
    
    ' Audio
    volumeMod = getDictValue(myDict, "volumeMod", Default_VolumeMod)
    channelCount = getDictValue(myDict, "channelCount", Default_ChannelCount)
    audioBitRate = getDictValue(myDict, "audioBitrate", Default_AudioBitRate)
    
    ' Video advanced
    PSNR = getDictValue(myDict, "PSNR", Default_PSNR)
    halfRes = getDictValue(myDict, "halfRes", Default_HalfRes)
    overscanBorders = getDictValue(myDict, "overscanBorders", Default_OverscanBorders)
    overscanSize = getDictValue(myDict, "overscanSize", Default_OverscanSize)
    Deinterlace = getDictValue(myDict, "Deinterlace", Default_Deinterlace)
    MPEG2Copy = getDictValue(myDict, "MPEG2Copy", Default_MPEG2Copy)
    dcPrecision = getDictValue(myDict, "dcPrecision", Default_DCPrecision)
    
    ' Burning
    createISO = getDictValue(myDict, "createISO", Default_CreateISO)
    discLabel = getDictValue(myDict, "discLabel", Default_DiscLabel)
    burnerName = getDictValue(myDict, "burnerName", Default_BurnerName)
    deleteISO = getDictValue(myDict, "deleteISO", Default_DeleteISO)
    eraseRW = getDictValue(myDict, "eraseRW", Default_EraseRW)
    enableBurning = getDictValue(myDict, "enableBurning", Default_EnableBurning)
    burnSpeed = getDictValue(myDict, "burnSpeed", Default_BurnSpeed)
    verifyDisc = getDictValue(myDict, "verifyDisc", Default_VerifyDisc)
    ejectTray = getDictValue(myDict, "ejectTray", Default_EjectTray)
    
    ' Menu
    menuTemplateName = getDictValue(myDict, "menuTemplateName", Default_menuTemplateName)
    menuAutoPlay = getDictValue(myDict, "menuAutoPlay", Default_menuAutoPlay)
    menuShowSubtitleFirst = getDictValue(myDict, "menuShowSubtitleFirst", Default_menuShowSubtitleFirst)
    menuShowAudioFirst = getDictValue(myDict, "menuShowAudioFirst", Default_menuShowAudioFirst)
    
    destinationDir = getDictValue(myDict, "destinationDir", Default_LastOutputDir)
    Debug.Print 6

    ' Read sources
    Set myData = myDict("Sources")
    For A = 0 To myData.Count - 1
        Set newSource = New clsSource
        Set myData2 = myData.Items(A)
        newSource.unSerialize myData2
        oldSourceFile = newSource.fileName

        ' Source exists and has been modified?
        If FS.FileExists(oldSourceFile) Then
            If newSource.timeStamp <> FS.GetFile(oldSourceFile).DateLastModified Then
                Set newSource = Nothing
                Set newSource = New clsSource
                newSource.openFile oldSourceFile
            End If
        End If

        ' Add it whether the file exists or not, it will be checked later
        Sources.Add Replace(myData.Keys(A), vbNullChar, " "), newSource
    Next A

    ' Read thumbnails
    For A = 0 To Sources.Count - 1
        Set newSource = Sources.Items(A)
        
        For B = 0 To newSource.streamCount - 1
            If newSource.streamInfo(B)("Type") = ST_Video Then
                picFile = tempPath & "\thumb" & A & "-" & B & ".bmp"
                If FS.FileExists(picFile) Then
                    Set newSource.streamImage(B) = New clsGDIImage
                    newSource.streamImage(B).openFrom picFile
                End If
            End If
        Next B
    Next A
    
    ' Remove sources that no longer exist
    Set newSources = New Dictionary
    For A = 0 To Sources.Count - 1
        Set newSource = Sources.Items(A)
        If Not FS.FileExists(newSource.fileName) Then
            frmDialog.Display "The file " & newSource.fileName & " could not be found. Any related video or audio items will not be added to the project.", Exclamation Or OkOnly
        Else
            newSources.Add Sources.Keys(A), Sources.Items(A)
        End If
    Next A
    Set Sources = newSources
    
    ' Read titles
    Set myData = myDict("Titles")
    For A = 0 To myData.Count - 1
        Set newTitle = New clsTitle
        Set myData2 = myData.Items(A)
        
        If newTitle.unSerialize(myData2) Then
            Titles.Add newTitle
        Else
            frmDialog.Display "Title " & A + 1 & " could not be added.", Exclamation Or OkOnly
        End If
    Next A
    
    
    ' Opening was succesful
    FS.DeleteFolder tempPath, True
    fileName = myFileName
    Modified = False
    unSerialize = True

End Function


' Save the project to a file
Public Sub Serialize(ByVal myFileName As String)

    Dim A As Long, B As Long
    Dim myConf As clsConfiguration
    Dim myDict As Dictionary
    Dim myData As Dictionary
    Dim mySource As clsSource
    Dim tempPath As String
    Dim sourceCount As Long
    Dim Pic As clsGDIImage
    Dim Ver As New clsTextFile
    
    
    Set myConf = New clsConfiguration

    ' Create temporary destination folder
    tempPath = TEMP_PATH & "project"
    If Not FS.FolderExists(tempPath) Then FS.CreateFolder tempPath

    ' Remove if already exists
    myFileName = FS.GetAbsolutePathName(myFileName)
    killIfExists myFileName

    ' Version file
    Set Ver = New clsTextFile
    Ver.fileCreate tempPath & "\version.txt", encodeASCIIorUTF8, CODEPAGE_LATIN1
    Ver.writeText PROJECT_VERSION
    Ver.fileClose

    ' Get dictionary of configuration
    Set myDict = myConf.Root(True)
    
    
    With myDict
    
        ' General
        .Add "Title", Title
        .Add "targetSize", targetSize
        .Add "customSize", customSize
        .Add "encodePriority", encodePriority
        .Add "threadCount", threadCount
        .Add "keepFiles", keepFiles
        
        ' Video
        .Add "targetFormat", targetFormat
        .Add "encodeProfile", encodeProfile
        .Add "targetBitrate", targetBitrate
        .Add "customBitrate", customBitrate
        .Add "allowPulldown", Pulldown
        
        ' Audio
        .Add "volumeMod", volumeMod
        .Add "channelCount", channelCount
        .Add "audioBitrate", audioBitRate
        
        ' Playback
        .Add "loopPlayback", loopPlayback
        .Add "whenPlayed", whenPlayed
        .Add "enableFirstSub", enableFirstSub
        
        ' Video advanced
        .Add "PSNR", PSNR
        .Add "halfRes", halfRes
        .Add "overscanBorders", overscanBorders
        .Add "overscanSize", overscanSize
        .Add "Deinterlace", Deinterlace
        .Add "MPEG2Copy", MPEG2Copy
        .Add "dcPrecision", dcPrecision
        
        ' Burning
        .Add "createISO", createISO
        .Add "discLabel", discLabel
        .Add "burnerName", burnerName
        .Add "deleteISO", deleteISO
        .Add "eraseRW", eraseRW
        .Add "enableBurning", enableBurning
        .Add "burnSpeed", burnSpeed
        .Add "verifyDisc", verifyDisc
        .Add "ejectTray", ejectTray
        
        ' Menu
        .Add "menuTemplateName", menuTemplateName
        .Add "menuAutoPlay", menuAutoPlay
        .Add "menuShowSubtitleFirst", menuShowSubtitleFirst
        .Add "menuShowAudioFirst", menuShowAudioFirst
        
        .Add "destinationDir", destinationDir
        
    End With
    

    ' Write sources
    setSourceRefCount
    Set myData = New Dictionary
    For A = 0 To Sources.Count - 1
        Set mySource = Sources.Items(A)

        If mySource.refCount > 0 Then
            myData.Add Replace(mySource.fileName, " ", vbNullChar), mySource.Serialize

            ' Write thumbnails of video streams
            For B = 0 To mySource.streamCount - 1
                If mySource.streamInfo(B)("Type") = ST_Video And Not (mySource.streamImage(B) Is Nothing) Then
                    Set Pic = mySource.streamImage(B)
                    If Not (Pic Is Nothing) Then SavePicture Pic.getPicture, tempPath & "\thumb" & sourceCount & "-" & B & ".bmp"
                End If
            Next B
            sourceCount = sourceCount + 1
        End If
        
    Next A
    myDict.Add "Sources", myData

    ' Write titles
    Set myData = New Dictionary
    For A = 0 To Titles.Count - 1
        myData.Add "Title" & A, Titles.Item(A).Serialize
    Next A
    myDict.Add "Titles", myData

    ' Write to file
    myConf.SaveConfiguration tempPath & "\project.cfg"
    Set myConf = Nothing

    ' Archive
    zipFromFolder myFileName, tempPath
    FS.DeleteFolder tempPath, True

    fileName = myFileName

End Sub


' Class
Private Sub Class_Initialize()

    Reset

End Sub

Private Sub Class_Terminate()

    Reset

End Sub


' Return the ideal height of a menu
Public Property Get menuHeight() As Long

    If targetFormat = VF_MIXED Or targetFormat = VF_PAL Then
        menuHeight = 576
    Else
        menuHeight = 480
    End If

End Property


' Return whether a video conforms to the project's target format
Public Function canCopyVideo(ByRef encInfo As Dictionary, ByRef vidInfo As Dictionary)

    If vidInfo("Width") = encInfo("Width") And vidInfo("Height") = encInfo("Height") And vidInfo("FPS") = encInfo("FPS") And vidInfo("Compression") = "mpeg2video" Then
        canCopyVideo = True
    Else
        canCopyVideo = False
    End If

End Function


' Return number of sources
Public Property Get nSources() As Long

    nSources = Sources.Count
    
End Property


' Return a source from an index number
Public Property Get getSourceIndex(ByVal Index As Long) As clsSource

    Set getSourceIndex = Sources.Items(Index)

End Property
