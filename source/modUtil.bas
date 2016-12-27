Attribute VB_Name = "modUtil"
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
'   File purpose: General utility functions.
'
Option Explicit
Option Compare Binary
Option Base 0


' Enum values for special paths for GetSpecialpath
Public Enum specialPathEnum
    myDocuments = 1
    applicationData = 2
    Temp = 3
    userDesktop = 4
End Enum
   
   
' OS version struct
Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

' Menu info for a form
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


' Local user ID
Private Type LUID
    dwLowPart As Long
    dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    udtLUID As LUID
    dwAttributes As Long
End Type

' Process priveledge tokens
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    laa As LUID_AND_ATTRIBUTES
End Type

' Win32 data for describing a file's time\date
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

' Win32 data for finding a file....?
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

' Font structure
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(31) As Byte
End Type

' Common controls init data
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type


' Windows NT 5 version
Private Const VER_PLATFORM_WIN32_NT = 2

' Eh?
Private Const MD As Long = &H5

' What?
Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

' Maximum length of some path variables
Private Const MAX_PATH = 260

' Windows NT system control
Public Const EWX_LOGOFF As Long = 0
Public Const EWX_SHUTDOWN As Long = &H1
Public Const EWX_REBOOT = &H2
Public Const EWX_FORCE As Long = &H4
Public Const EWX_POWEROFF As Long = &H8
Public Const EWX_FORCEIFHUNG As Long = &H10

' Special folder constants
Private Const CSIDL_APPDATA As Long = &H1A
Private Const CSIDL_MYDOCUMENTS As Long = &H5
Private Const CSIDL_DESKTOP As Long = &H0
Private Const CSIDL_FLAG_CREATE As Long = &H8000

' SetThreadExecutionState
Public Const ES_SYSTEM_REQUIRED As Long = &H1
Public Const ES_DISPLAY_REQUIRED As Long = &H2
Public Const ES_USER_PRESENT  As Long = &H4
Public Const ES_AWAYMODE_REQUIRED As Long = &H40
Public Const ES_CONTINUOUS As Long = &H80000000

' To get default system font with SystemParameterInfo
Private Const SPI_GETICONTITLELOGFONT As Long = 31

' Process priveledge token adjustment
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2


' File filter lists
Public titleFiles As String
Public audioFiles As String
Public subFiles As String

' Handle to app's mutex
Public hMutex As Long


' API calls
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal bytes As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExW" (ByVal lpRootPathName As Long, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetACP Lib "kernel32" () As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, ByVal ptrParam As Long, ByVal fWinIni As Long) As Long
Private Declare Function getDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal ptrDriver As String, ByVal ptrDevice As Long, ByVal ptrOutput As Long, ByVal printerData As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32" (iCCEx As tagInitCommonControlsEx) As Boolean
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long) As Long
Public Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Integer, ByVal fForce As Integer) As Integer
Public Declare Function SetSuspendState Lib "powrprof" (ByVal Hibernate As Boolean, ByVal forceCritical As Boolean, ByVal disableWakeEvent As Boolean) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


' Return a value from a dictionary if it exists, else return a default value
Public Function getDictValue(ByRef Dict As Dictionary, ByVal Value As String, ByVal Default As Variant) As Variant

    If Dict.Exists(Value) Then
        getDictValue = Dict.Item(Value)
    Else
        getDictValue = Default
    End If

End Function


' Return the height of a titlebar in pixels
Public Function getTitleBarHeight() As Long

    getTitleBarHeight = GetSystemMetrics(55)

End Function


' Initialise common controls for themed look
Public Function initCommonControls() As Boolean
   
   Dim iCCEx As tagInitCommonControlsEx
   
   On Error Resume Next
      
   
   With iCCEx
       .lngSize = LenB(iCCEx)
       .lngICC = &H200
   End With
   
   InitCommonControlsEx iCCEx
   initCommonControls = (Err.Number = 0)
   
   On Error GoTo 0
   
End Function


' Validate the current configuration
Public Sub validateConfig()

    If Config.Root.Exists("lastBrowseDir") Then
        If Not FS.FolderExists(Config.ReadSetting("lastBrowseDir")) Then Config.WriteSetting "lastBrowseDir", getSpecialPath(myDocuments)
    End If

End Sub


' Get address of a function through AddressOf
Public Function getAddressOfFunction(Address As Long) As Long

    getAddressOfFunction = Address
  
End Function


' Check whether a MPEG-2 video file is a valid video stream by checking the magic bytes
Public Function isValidMPEG2Video(ByVal fileName As String) As Boolean

    Dim fileObj As clsBinaryFile
    Dim magicBytes As Long
    
    
    Set fileObj = New clsBinaryFile
    If Not fileObj.fileOpen(fileName, False) Then Exit Function
    
    magicBytes = fileObj.readLong
    If magicBytes = &H1B3 Then isValidMPEG2Video = True

End Function


' Append an MPEG-2 end of sequence code to a file
Public Function appendEndSequence(ByVal fileName As String) As Boolean

    Dim Obj As clsBinaryFile
    
    
    Set Obj = New clsBinaryFile
    If Not Obj.fileOpen(fileName, True) Then Exit Function
    
    Obj.fileSeek Obj.fileLength
    Obj.writeByte &H0
    Obj.writeByte &H0
    Obj.writeByte &H1
    Obj.writeByte &HB7
    Obj.fileClose
    
    appendEndSequence = True

End Function


' Resize an image to fit into a smaller version, taking care of possible DAR
Public Function resizeToMatch(ByRef Pic As clsGDIImage, ByVal Width As Long, ByVal Height As Long, ByVal DAR As Single) As clsGDIImage

    Dim Modif As Double
    Dim mTop As Long, mLeft As Long
    Dim mWidth As Long, mHeight As Long
    Dim newPic As clsGDIImage
    
    
    If Pic Is Nothing Then Exit Function
    
    Set newPic = New clsGDIImage
    newPic.copyFrom Pic
    If DAR <> 1 Then newPic.Resize newPic.Height * DAR, newPic.Height, STRETCH_HALFTONE
    
    ' Resize keeping aspect ratio
    Modif = getResizeValue(newPic.Width, newPic.Height, Width, Height)
    mWidth = newPic.Width * Modif
    mHeight = newPic.Height * Modif
        
    ' Keep inside image
    If mWidth < Width Then mLeft = (Width / 2) - (mWidth / 2)
    If mHeight < Height Then mTop = (Height / 2) - (mHeight / 2)
    
    ' Resize to fit onto a new image
    Set resizeToMatch = New clsGDIImage
    resizeToMatch.createNew Width, Height, 32
    resizeToMatch.renderImage newPic, mLeft, mTop, mWidth, mHeight, Render_Copy, STRETCH_HALFTONE
    
    Set newPic = Nothing

End Function


' Get the default system font and apply it to all controls in this program
Public Function applyDefaultSystemFont() As Boolean

    Dim useFont As LOGFONT
    Dim myForm As Form
    Dim myControl As Control
    Dim desktopDC As Long
    Dim oldFont As StdFont

    
    ' Get default system font
    If SystemParametersInfo(SPI_GETICONTITLELOGFONT, LenB(useFont), VarPtr(useFont), 0) = 0 Then Exit Function
    
    ' Get desktop device context
    desktopDC = CreateDC("DISPLAY", 0, 0, 0)
    If desktopDC = 0 Then Exit Function
    
    
    ' Apply font object to all controls on all loaded forms
    For Each myForm In Forms
        For Each myControl In myForm.Controls
            
            ' Only apply font to objects that need it
            Select Case typeName(myControl)
                Case "CommandButton", "Label", "TextBox", "ComboBox", "CheckBox", "ctlFancyList", "OptionButton", "PictureBox", "ListBox", "DTPicker"
                    
                    Set oldFont = myControl.Font
                
                    ' Create new font object for each control to keep specific size and styles
                    Set myControl.Font = New StdFont
                    With myControl.Font
                        .Name = StrConv(useFont.lfFaceName, vbUnicode)
                        .Charset = useFont.lfCharSet
                        .Bold = oldFont.Bold
                        .Italic = oldFont.Italic
                        .Underline = oldFont.Underline
                        .Size = oldFont.Size
                    End With

            End Select
            
        Next myControl
    Next myForm
    
    
    applyDefaultSystemFont = True

End Function


' Unused in code but sometimes useful to generate Long FourCC values
Public Function createLongFourCC(ByVal fourCC As String) As Long

    Dim Data(3) As Byte
    
    
    Data(0) = AscW(Mid$(fourCC, 1, 1))
    Data(1) = AscW(Mid$(fourCC, 2, 1))
    Data(2) = AscW(Mid$(fourCC, 3, 1))
    Data(3) = AscW(Mid$(fourCC, 4, 1))
    CopyMemory ByVal VarPtr(createLongFourCC), ByVal VarPtr(Data(0)), 4

End Function


' Set\get DateTimePicker values from time indices
Public Sub setDTValue(ByRef DT As DTPicker, ByVal timeIndex As Long)

    DT.Hour = timeIndex \ 3600
    timeIndex = timeIndex - (DT.Hour * 3600)
    
    DT.Minute = timeIndex \ 60
    timeIndex = timeIndex - (DT.Minute * 60)
    
    DT.Second = timeIndex

End Sub

Public Function getDTValue(ByRef DT As DTPicker) As Long

    getDTValue = getDTValue + (DT.Hour * 60 * 60)
    getDTValue = getDTValue + (DT.Minute * 60)
    getDTValue = getDTValue + DT.Second

End Function


' Create and destroy app's mutex
Public Sub createAppMutex()

    hMutex = CreateMutex(ByVal 0&, 1, App.Title)
    
End Sub

Public Sub destroyAppMutex()

    ReleaseMutex hMutex
    CloseHandle hMutex

End Sub


' Return a default aspect ratio as value
Public Function getAspect(ByVal Aspect As enumVideoAspects) As Single

    If Aspect = VA_169 Then getAspect = 16 / 9
    If Aspect = VA_43 Then getAspect = 4 / 3

End Function


' Convert a BMP to PNG using an external utility
Public Function convertToPNG(ByVal inFileName As String, ByVal outFileName As String, Optional ByVal transColor As Long = -1, Optional ByVal logToProject As Boolean = False) As Boolean

    Dim cmdLine As String
    Dim binPath As String
    
    
    binPath = APP_PATH & "bin\bmp2png.exe"
    cmdLine = cmdLine & " -X"
    cmdLine = cmdLine & " -0"
    If transColor <> -1 Then cmdLine = cmdLine & " -P " & transColor
    cmdLine = cmdLine & " -E " & vbQuote & inFileName & vbQuote
    cmdLine = cmdLine & " -O " & vbQuote & outFileName & vbQuote
    
    If logToProject Then
        executeToFile binPath, cmdLine, Project.destinationDir & "\bmp2png.txt", SM_Nothing, EP_Normal, ""
    Else
        Execute vbQuote & binPath & vbQuote, cmdLine, WS_HIDE, True, True
    End If
    
    If Not FS.FileExists(outFileName) Then Exit Function
    convertToPNG = True

End Function


' Delete only files generated by the encoding process and DVD subfolder
Public Function emptyDestinationFolder(ByVal folderName As String) As Boolean

    Dim A As Long
    Dim MyFiles As Files
    Dim MyFolders As Folders
    Dim myFile As File
    Dim myFolder As Folder
    
    On Error GoTo NoDelete
    
    
    If FS.FolderExists(folderName) = False Then
        emptyDestinationFolder = True
        Exit Function
    End If
    
    Set MyFiles = FS.GetFolder(folderName).Files
    Set MyFolders = FS.GetFolder(folderName).SubFolders
    
    For Each myFile In MyFiles
        Select Case LCase$(FS.GetExtensionName(myFile.Name))
            Case "m2v", "txt", "log", "xml", "vob", "ifo", "bup", "iso", "ac3", "mds", "bat"
                myFile.Delete
        End Select
    Next myFile
    
    For Each myFolder In MyFolders
        If myFolder.Name = "dvd" Or myFolder.Name = "VIDEO_TS" Then emptyDestinationFolder myFolder.Path
    Next myFolder
    
    Set MyFiles = Nothing
    Set MyFolders = Nothing
    
    On Error GoTo 0
    
    emptyDestinationFolder = True
    Exit Function
    
    
NoDelete:
    On Error GoTo 0

End Function


' Return a special path
' No trailing slash is ensured
Public Function getSpecialPath(ByVal Path As specialPathEnum) As String

    Dim sPath As String
    Dim flags As Long
    

    If Path = applicationData Then
        flags = CSIDL_APPDATA
    ElseIf Path = myDocuments Then
        flags = CSIDL_MYDOCUMENTS
    ElseIf Path = userDesktop Then
        flags = CSIDL_DESKTOP
    End If
    
    sPath = String(MAX_PATH, vbNullChar)
    If Path = Temp Then
        If GetTempPath(Len(sPath), sPath) <> 0 Then getSpecialPath = stripNull(sPath)
    Else
        If SHGetFolderPath(0, flags, 0, 0, sPath) = 0 Then getSpecialPath = stripNull(sPath)
    End If

    If Right$(getSpecialPath, 1) = "\" Then getSpecialPath = Left$(getSpecialPath, Len(getSpecialPath) - 1)

End Function


Public Function remapAudio(ByVal Audio As String) As String

    If audioMappings.Exists(Audio) Then
        remapAudio = audioMappings(Audio)
    Else
        remapAudio = Audio
    End If

End Function


Public Function unsupportedAudio(ByVal Audio As String) As Boolean
    
    unsupportedAudio = False
    
    If Audio = "Windows Media Audio Professional v9" Then unsupportedAudio = True
    If Audio = "Windows Media Audio v9" Then unsupportedAudio = True
    If InStr(Audio, " / ") <> 0 Then unsupportedAudio = True
    If InStr(Audio, "0x") <> 0 Then unsupportedAudio = True

End Function


Public Function remapVideo(ByVal Video As String) As String

    If videoMappings.Exists(Video) Then
        remapVideo = videoMappings(Video)
    Else
        remapVideo = Video
    End If

End Function


Public Function unsupportedVideo(ByVal Video As String) As Boolean
    
    unsupportedVideo = False
    
    If Video = "Windows Media Video 9" Then unsupportedVideo = True
    If Video = "iTunes DRM" Then unsupportedVideo = True
    If Video = "Microsoft Photostory 2" Then unsupportedVideo = True
    If Video = "RealVideo 4" Then unsupportedVideo = True
    If Video = "Microsoft Windows Screen Video" Then unsupportedVideo = True
    If Video = "Picture To Exe Video" Then unsupportedVideo = True

End Function


' Return whether a burner device can actually write
Public Function burnerCanWrite(ByRef Device As clsBurnerDevice) As Boolean

    If Device.deviceCaps("writeCDR") Or Device.deviceCaps("writeCDRW") Then burnerCanWrite = True: Exit Function
    If Device.deviceCaps("writeDVDR") Or Device.deviceCaps("writeDVDRAM") Then burnerCanWrite = True: Exit Function
    If Device.deviceFeatures("randomWrite") Or Device.deviceFeatures("cdRWWrite") Then burnerCanWrite = True: Exit Function
    If Device.deviceFeatures("dvdMinusRW") Then burnerCanWrite = True: Exit Function

End Function


Public Sub createFileTypes()

    Dim myList As clsFilterList
    
    
    Set myList = New clsFilterList
    
    ' Titles (video files)
    myList.addType "3GPP files", "*.3gp;*.3g2"
    myList.addType "AVI files", "*.avi"
    myList.addType "AviSynth script files", "*.avs"
    myList.addType "DivX video files", "*.divx"
    myList.addType "Flash Video files", "*.flv"
    myList.addType "HD QuickTime files", "*.hdmov"
    myList.addType "Matroska files", "*.mkv"
    myList.addType "Motion-JPEG video files", "*.mjpg"
    myList.addType "MPEG video files", "*.mpg;*.m2v;*.mpeg;*.mpv"
    myList.addType "MPEG-4 files", "*.mp4;*.m4v"
    myList.addType "Nullsoft Video files", "*.nsv"
    myList.addType "NUT files", "*.nut"
    myList.addType "QuickTime files", "*.qt;*.mov"
    myList.addType "RealMedia files", "*.rm"
    myList.addType "Smacker files", "*.smk"
    myList.addType "MPEG-2 Transport Stream files", "*.ts"
    myList.addType "Vorbis files", "*.ogm"
    myList.addType "Windows Media Video files", "*.wmv;*.asf"
    titleFiles = myList.fullString
    
    ' Audio
    ' Do not clear list, video files can have audio streams too
    myList.addType "Advanced Audio Coding files", "*.aac"
    myList.addType "Dolby AC3 files", "*.ac3"
    myList.addType "Digital Theatre System files", "*.dts"
    myList.addType "FLAC files", "*.flac"
    myList.addType "Matroska Audio files", "*.mka"
    myList.addType "MPEG Layer 3 files", "*.mp3"
    myList.addType "MPEG Layer 2 files", "*.mp2"
    myList.addType "MPEG Audio files", "*.mpa"
    myList.addType "OGG Vorbis files", "*.ogg"
    myList.addType "PCM WAV files", "*.wav"
    myList.addType "Windows Media Audio files", "*.wma"
    audioFiles = myList.fullString
    
    ' Subtitles
    myList.Clear
    myList.addType "MicroDVD files", "*.sub;*.txt"
    myList.addType "SSA\ASS files", "*.ass;*.ssa"
    myList.addType "SubRip files", "*.srt"
    myList.addType "SubView files", "*.sub"
    myList.addType "Other text-based files", "*.txt"
    subFiles = myList.fullString
    
    Set myList = Nothing

End Sub


' Get a framegrab from a video file using FFmpeg
Public Function getFrameBitmap(ByVal fileName As String, ByVal Width As Long, ByVal Height As Long, ByVal timeIndex As Single) As clsGDIImage

    Dim cmdLine As String
    Dim inFile As String
    

    inFile = TEMP_PATH & "img.jpg"
    killIfExists inFile
    
    ' Play 1 frame, and capture it to jpeg format
    cmdLine = "-vframes 1 -sws_flags bilinear -ss" & Str$(timeIndex) & " -i " & vbQuote & fileName & vbQuote & " -s " & Width & "x" & Height & " -f image2 " & vbQuote & inFile & vbQuote
    Execute APP_PATH & "bin\ffmpeg.exe", cmdLine, WS_HIDE, True, False
    
    If FS.FileExists(inFile) Then
        Set getFrameBitmap = New clsGDIImage
        getFrameBitmap.openFrom inFile
        FS.DeleteFile inFile
    
    Else
        Set getFrameBitmap = Nothing
        
    End If
    
End Function


' Enable or disable a form's close button
Public Function setCloseButton(ByVal hWnd As Long, Enable As Boolean) As Long
    
    Dim hMenu As Long
    Dim MII As MENUITEMINFO
    Dim lngMenuID As Long
    
    Const xSC_CLOSE As Long = -10


    ' Check that the window handle passed is valid
    setCloseButton = -1
    If IsWindow(hWnd) = 0 Then Exit Function
    
    ' Retrieve a handle to the window's system menu
    hMenu = GetSystemMenu(hWnd, 0)
    
    ' Retrieve the menu item information for the close menu item/button
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If
    
    setCloseButton = -1
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    
    ' Switch the ID of the menu item so that VB can not undo the action itself
    lngMenuID = MII.wID
    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If
    
    MII.fMask = MIIM_ID
    setCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    
    
    ' Set the enabled / disabled state of the menu item
    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If
    
    MII.fMask = MIIM_STATE
    setCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    
    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the close button in its new state.
    SendMessage hWnd, WM_NCACTIVATE, True, 0
    
    setCloseButton = 0
    
End Function


' Determine if the current OS is at least NT version 5 (Windows 2000 and higher)
Public Function isNT5() As Boolean

    Dim myOS As OSVERSIONINFOEX
    
    
    isNT5 = False
    myOS.dwOSVersionInfoSize = Len(myOS)
    
    GetVersionEx myOS
    If myOS.dwPlatformId = VER_PLATFORM_WIN32_NT And myOS.dwMajorVersion >= 5 Then isNT5 = True
    
End Function


' Return current OS version + build
Public Function windowsVersion() As String

    Dim myOS As OSVERSIONINFOEX
    Dim osCSD As String
    
    
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    
    windowsVersion = "Windows " & myOS.dwMajorVersion & "." & myOS.dwMinorVersion & " build " & myOS.dwBuildNumber
    
    ' Service packs
    osCSD = StrConv(myOS.szCSDVersion, 0)
    If LenB(osCSD) <> 0 Then windowsVersion = windowsVersion & " " & osCSD
    
End Function


' Fill a combobox with language names
Public Sub fillLangCombo(ByRef Combo As ComboBox)

    Dim A As Long
    
    
    Combo.Clear
    For A = 0 To langCodes.Count - 1
        Combo.addItem langCodes.Items(A) & " (" & langCodes.Keys(A) & ")"
    Next A

End Sub


' Kill a file, but only if it exists
Public Sub killIfExists(ByVal fileName As String)

    appLog.Add "Attempting to delete " & fileName
    
    ' Wildcard, delete anyway
    If InStr(fileName, "*") Then
        FS.DeleteFile fileName, True
    
    ' Delete single file
    Else
        If FS.FileExists(fileName) Then
            FS.DeleteFile fileName, True
        Else
            appLog.Add "File does not exist.", 1
        End If
        
    End If

End Sub


' Set our process to have the shutdown security priveledge
Public Function adjustShutdownToken() As Boolean
    
    Dim hProcessHandle As Long
    Dim hTokenHandle As Long
    Dim lpv_la As LUID
    Dim token As TOKEN_PRIVILEGES
    
    
    hProcessHandle = GetCurrentProcess()
    If hProcessHandle <> 0 Then
    
        ' Open access token associated with current process
        If OpenProcessToken(hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hTokenHandle) <> 0 Then
        
            ' Obtain LUID
            If LookupPrivilegeValue(vbNullString, "SeShutdownPrivilege", lpv_la) <> 0 Then
        
                ' Prepare the TOKEN_PRIVILEGES structure
                With token
                    .PrivilegeCount = 1
                    .laa.udtLUID = lpv_la
                    .laa.dwAttributes = SE_PRIVILEGE_ENABLED
                End With
        
                ' Enable the shutdown privilege
                If AdjustTokenPrivileges(hTokenHandle, False, token, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then adjustShutdownToken = True
        
            End If
        
        End If
    
    End If

End Function


' Populate a Dictionary from an external file
Public Function loadDict(ByVal fileName As String) As Dictionary

    Dim readLine As String
    Dim Stream As clsTextFile
    Dim Separator As Long
    Dim Key As String
    Dim Item As String
    

    Set Stream = New clsTextFile
    If Not Stream.fileOpen(fileName, False) Then Exit Function
    Set loadDict = New Dictionary
    
    Do
        readLine = Stream.readLine

        ' Key and item are separated by the first comma
        Separator = InStr(readLine, ",")
        If Separator Then
            Key = Trim$(Left$(readLine, Separator - 1))
            Item = Trim$(Mid$(readLine, Separator + 1))

            If IsNumeric(Key) Then
                loadDict.Add Val(Key), Item
            Else
                loadDict.Add Key, Item
            End If
        End If
    Loop Until Stream.fileEndReached
    Set Stream = Nothing
    
    appLog.Add "Loaded " & fileName, 1

End Function


' Stretch to fit
' It will keep aspect ratio for letterboxing, or for pan&scan
Public Function getResizeValue(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal destWidth As Long, ByVal destHeight As Long) As Double

    ' Divide by zero check
    If srcWidth = 0 Or srcHeight = 0 Or destWidth = 0 Or destHeight = 0 Then
        getResizeValue = 1
        Exit Function
    End If
    
    If srcWidth / destWidth > srcHeight / destHeight Then
        getResizeValue = destWidth / srcWidth
    Else
        getResizeValue = destHeight / srcHeight
    End If

End Function


' Return information about current locale
Public Function readLocaleInfo(ByVal lInfo As Long) As String
    
    Dim Buffer As String
    Dim Ret As String
    
    
    Buffer = String(256, 0)

    Ret = GetLocaleInfo(&H400, lInfo, StrPtr(Buffer), LenB(Buffer) - 16)
    If Ret > 0 Then
        readLocaleInfo = Left$(Buffer, Ret - 1)
    Else
        readLocaleInfo = ""
    End If
    
End Function


Public Sub initDefaults()

    Dim A As Long
    Dim Locale As String
    Dim Burner As clsBurnerDevice
    
    
    ' Try to get NTSC or PAL standard, but default to NTSC if not found
    Config.WriteSetting "targetFormat", VF_NTSC
    Locale = LCase$(readLocaleInfo(&H1002))
    For A = 0 To Countries.Count - 1
        If InStr(Countries.Keys(A), Locale) Then
            If Countries.Items(A) = 1 Then
                Config.WriteSetting "targetFormat", VF_PAL
            Else
                Config.WriteSetting "targetFormat", VF_NTSC
            End If
        End If
    Next A
    
    ' Default window position based on screen resolution
    Default_WindowTop = Screen.Height * 0.1
    Default_WindowLeft = Screen.Width * 0.1
    Default_WindowWidth = Screen.Width * 0.8
    Default_WindowHeight = Screen.Height * 0.8
    
    Default_ThreadCount = cpuInfo.getLogicalCPUCount
    Default_LastOutputDir = getSpecialPath(myDocuments) & "\dvd"
    Default_LastBrowseDir = getSpecialPath(myDocuments)
    
    ' Default burner device
    For A = 0 To Burners.deviceCount - 1
        Set Burner = Burners.getDevice(A)
        If burnerCanWrite(Burner) Then
            Default_BurnerName = Burner.deviceName & " (" & Burner.deviceDriveChar & ")"
            Exit For
        End If
    Next A
    
End Sub


' If a path does not exist, construct every folder
Public Function constructPath(ByVal Path As String) As Boolean

    Dim A As Long
    Dim Pts() As String
    Dim cTry As String
    
    On Error GoTo noConstruct
    
    
    Pts = Split(Path, "\")
    For A = 0 To UBound(Pts())
        cTry = cTry & Pts(A) & "\"
        If Not FS.FolderExists(cTry) Then FS.CreateFolder cTry
    Next A
    
    On Error GoTo 0
    
    constructPath = True
    Exit Function
    
    
noConstruct:
    On Error GoTo 0

End Function


' Create a zip file from a folder
Public Sub zipFromFolder(ByVal zipFile As String, ByVal Folder As String)

    Dim cmdLine As String
    

    killIfExists zipFile
    
    cmdLine = "a -y -r -tzip -mx=5 " & vbQuote & zipFile & vbQuote & " " & vbQuote & Folder & "\*" & vbQuote
    Execute APP_PATH & "bin\7za.exe", cmdLine, WS_HIDE, True

End Sub


' Extract a folder from a zip file
Public Sub folderFromZip(ByVal zipFile As String, ByVal Folder As String)

    Dim cmdLine As String
    
    
    cmdLine = "x " & vbQuote & zipFile & vbQuote & " -y -o" & vbQuote & Folder & vbQuote
    Execute APP_PATH & "bin\7za.exe", cmdLine, WS_HIDE, True

End Sub


' Return whether a file is binary or not
Public Function isBinaryFile(ByVal fileName As String) As Boolean

    Dim fileObj As clsBinaryFile
    Dim Buffer As String
    
    Const BUFFER_SIZE As Long = 10240
    
    
    isBinaryFile = False
    
    Set fileObj = New clsBinaryFile
    If Not fileObj.fileOpen(fileName, False) Then Err.Raise vbObjectError, , "Unable to open binary file"

    ' Limit checking to one BUFFER_SIZE
    If fileObj.fileLength >= BUFFER_SIZE Then
        Buffer = fileObj.readStringData(BUFFER_SIZE)
    Else
        Buffer = fileObj.readStringData(fileObj.fileLength)
    End If
    
    Set fileObj = Nothing
    
    ' UTF-16 BOM
    If Left$(Buffer, 2) = Chr$(&HFF) & Chr$(&HFE) Then
        isBinaryFile = False
    
    ' Crude but effective for other cases
    ElseIf InStr(Buffer, vbNullChar) <> 0 Or InStr(Buffer, ChrW$(1)) <> 0 Or InStr(Buffer, ChrW$(2)) <> 0 Then
        isBinaryFile = True
        
    End If

End Function
