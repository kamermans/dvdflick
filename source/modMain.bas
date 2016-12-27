Attribute VB_Name = "modMain"
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
'   File purpose: Main, startup
'
Option Explicit
Option Compare Binary
Option Base 0


' Our friendly neighbourhood classes
Public Project As clsProject
Public langCodes As Dictionary
Public Countries As Dictionary
Public colorSpaces As Dictionary
Public codePages As Dictionary
Public audioMappings As Dictionary
Public videoMappings As Dictionary
Public mmcProfiles As Dictionary
Public mmcFeatures As Dictionary
Public appLog As clsLog
Public Config As clsConfiguration
Public FS As FileSystemObject
Public Burners As clsBurnerDeviceEnum
Public CMD As clsParameters
Public cpuInfo As clsCPUDetect
Public fileDialog As clsCommonDialog
Public Dbg As clsDebug

' Warnings given before encoding process started
Public Warnings As String

' Used to halt pipe\shellex processes
Public haltShellEx As Boolean

' Image to indicate no thumbnail is available
Public noThumb As clsGDIImage

' Character values
Public vbQuote As String * 1

' Run in unattended mode, meaning encoding will start after program start and will show no messages
Public unattendMode As Boolean

' Paths
Public APP_PATH As String
Public DATA_PATH As String
Public TEMP_PATH As String

' Program beta
Public Const APP_BETA As Boolean = False
Public Const APP_RC As Long = 0


Public Sub Main()
    
    Dim A As Long
    Dim showWelcome As Boolean
    
    On Error GoTo startupError
    
    
    ' Character "constants"
    vbQuote = ChrW$(34)
    
    ' Create some objects
    Set FS = New FileSystemObject
    Set Burners = New clsBurnerDeviceEnum
    Set CMD = New clsParameters
    Set fileDialog = New clsCommonDialog
    Set Dbg = New clsDebug


    ' Pre-init
    createAppMutex
    modUtil.initCommonControls
    Load frmDialog
    
    ' Create application path
    APP_PATH = App.Path
    If Right$(APP_PATH, 1) <> "\" Then APP_PATH = APP_PATH & "\"

    ' Create user data path
    If CMD.Switches.Exists("datapath") Then
        DATA_PATH = CMD.Switches("datapath")
        If Right$(DATA_PATH, 1) <> "\" Then DATA_PATH = DATA_PATH & "\"
        If FS.FolderExists(DATA_PATH) = False Then
            frmDialog.Display "Data path " & DATA_PATH & " does not exist.", OkOnly Or Exclamation
            End
        End If
    Else
        DATA_PATH = getSpecialPath(applicationData) & "\DVD Flick\"
        If Not FS.FolderExists(DATA_PATH) And Not CMD.Switches.Exists("portable") Then FS.CreateFolder DATA_PATH
    End If
    
    ' Create temp data path
    If CMD.Switches.Exists("temppath") Then
        TEMP_PATH = CMD.Switches("temppath")
        If Right$(TEMP_PATH, 1) <> "\" Then TEMP_PATH = TEMP_PATH & "\"
        If FS.FolderExists(TEMP_PATH) = False Then
            frmDialog.Display "Temp path " & TEMP_PATH & " does not exist.", OkOnly Or Exclamation
            End
        End If
    Else
        TEMP_PATH = getSpecialPath(Temp) & "\"
    End If
    
    ' Paths for portable version
    If CMD.Switches.Exists("portable") Then
        TEMP_PATH = APP_PATH & "temp\"
        DATA_PATH = APP_PATH & "data\"
        If Not FS.FolderExists(TEMP_PATH) Then FS.CreateFolder TEMP_PATH
        If Not FS.FolderExists(DATA_PATH) Then FS.CreateFolder DATA_PATH
    End If
    
    ' Batch mode run
    If CMD.Switches.Exists("startunattended") And CMD.Switches.Exists("load") Then unattendMode = True


    ' Detect Windows version, NT5 minimum
    If Not isNT5 Then
        frmDialog.Display "Windows 2000 Professional or higher is required in order to use DVD Flick.", Exclamation Or OkOnly
        End
    End If


    Set appLog = New clsLog
    appLog.Start DATA_PATH & "dvdflick.log"
    
    ' Enable debug modes
    Dbg.setMode "pipes"
    If CMD.Switches.Exists("debug") Then Dbg.setMode CMD.Switches("debug")
   
    
    ' Add info for debugging purposes
    appLog.Add App.Title & " " & versionString
    appLog.Add windowsVersion
    appLog.Add "Application path: " & APP_PATH
    appLog.Add "Data path: " & DATA_PATH
    appLog.Add "Temp path: " & TEMP_PATH
    
    appLog.Add "CPU info..."
    Set cpuInfo = New clsCPUDetect
    appLog.Add cpuInfo.getProcessorBrandString, 1
    appLog.Add cpuInfo.getLogicalCPUCount & " Logical CPUs", 1
    If cpuInfo.has3DNOW Then appLog.Add "Supports 3DNow! extensions", 1
    If cpuInfo.hasMMX Then appLog.Add "Supports MMX extensions", 1
    
    
    ' Load and display the splash screen
    Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.fadeIn
    

    ' Load dictionaries
    appLog.Add "Initializing dictionaries..."
    Set langCodes = loadDict(APP_PATH & "data\langcodes.txt")
    Set Countries = loadDict(APP_PATH & "data\countries.txt")
    Set codePages = loadDict(APP_PATH & "data\codepages.txt")
    Set audioMappings = loadDict(APP_PATH & "data\audiomaps.txt")
    Set videoMappings = loadDict(APP_PATH & "data\videomaps.txt")
    Set mmcProfiles = loadDict(APP_PATH & "data\mmcprofiles.txt")
    Set mmcFeatures = loadDict(APP_PATH & "data\mmcfeatures.txt")
    Set colorSpaces = loadDict(APP_PATH & "data\colorspaces.txt")

    
    ' Load\create configuration
    appLog.Add "Setting up configuration..."
    
    Set Config = New clsConfiguration
    appLog.Add "Initializing defaults"
    initDefaults
    
    appLog.Add "Loading configuration file", 1
    If FS.FileExists(DATA_PATH & "dvdflick.cfg") Then
        Config.LoadConfiguration DATA_PATH & "dvdflick.cfg"
    Else
        showWelcome = True
    End If
    
    appLog.Add "Validating configuration", 1
    validateConfig
    
    
    ' Init misc. stuff
    appLog.Add "Initialising file types..."
    modUtil.createFileTypes
    
    appLog.Add "Initialising shutdown priveledge token..."
    modUtil.adjustShutdownToken
    
    appLog.Add "Loading menu template fonts..."
    modRenderMenu.registerTemplateFonts
    
    appLog.Add "Loading noThumbnail image..."
    Set noThumb = New clsGDIImage
    noThumb.openFrom APP_PATH & "data\nothumbnail.bmp"
    
    appLog.Add "Creating blank project..."
    Set Project = New clsProject

    ' Enumerate installed burners
    appLog.Add "Enumerating burners..."
    Burners.scanDevices
    appLog.Add "Found " & Burners.deviceCount & " devices.", 1
    For A = 0 To Burners.deviceCount - 1
        appLog.Add Burners.getDevice(A).deviceName, 1
    Next A
    
    ' Load file parameter
    If CMD.Switches.Exists("load") Then
        appLog.Add "Loading project from command line..."
        If Not Project.unSerialize(CMD.Switches("load")) Then
            frmDialog.Display "Could not load project file.", Exclamation Or OkOnly
            modMain.Quit
        End If
    End If

    
    ' Load all forms
    Load frmStatus
    Load frmWelcome
    Load frmTetris
    Load frmSubtitle
    Load frmProgress
    Load frmEncodeError
    Load frmAbout
    Load frmProjectSettings
    Load frmSelectTrack
    Load frmTitle
    Load frmEditTrack
    Load frmAdvOptsVideo
    Load frmSubPreview
    Load frmMenuSettings
    Load frmMenuPreview
    Load frmMain
    
    modUtil.applyDefaultSystemFont

    ' Copy some images
    Set frmProgress.imgTop.Picture = frmDialog.imgTop.Picture
    Set frmWelcome.imgTop.Picture = frmDialog.imgTop.Picture
       
    ' Set main window position and size
    frmMain.Top = Config.ReadSetting("windowTop", Default_WindowTop)
    frmMain.Left = Config.ReadSetting("windowLeft", Default_WindowLeft)
    frmMain.Width = Config.ReadSetting("windowWidth", Default_WindowWidth)
    frmMain.Height = Config.ReadSetting("windowHeight", Default_WindowHeight)
    frmMain.WindowState = Config.ReadSetting("windowState", Default_WindowState)
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
    
    fileDialog.lastFolder = Config.ReadSetting("lastBrowseDir", Default_LastBrowseDir)
    If Not FS.FolderExists(fileDialog.lastFolder) Then fileDialog.lastFolder = APP_PATH
    frmAbout.lblDebugFlags.Caption = Dbg.debugMode
    
    frmSplash.Hide
    On Error GoTo 0
    
    If unattendMode Then
        appLog.Add "Running in unattended mode, starting encoding."
        frmMain.startEncode
        modMain.Quit
    Else
        appLog.Add "Displaying main form."
        frmMain.Show
        If showWelcome Then frmWelcome.Show 1
    End If

    Exit Sub
    
    
startupError:
    frmDialog.Display "An error occured during startup. Number " & Err.Number & " from " & Err.Source & ":" & vbNewLine & Err.Description & vbNewLine & "Last DLL error: " & Err.LastDllError, Critical Or OkOnly
    destroyAppMutex
    End

End Sub


' Quit, clean up
Public Sub Quit(Optional ByVal endProgram As Boolean = True)

    ' Write config
    If Not (Config Is Nothing) And Not (Project Is Nothing) Then
        appLog.Add "Saving configuration..."
    
        If frmMain.WindowState = vbNormal Then
            Config.WriteSetting "windowTop", frmMain.Top
            Config.WriteSetting "windowLeft", frmMain.Left
            Config.WriteSetting "windowWidth", frmMain.Width
            Config.WriteSetting "windowHeight", frmMain.Height
            Config.WriteSetting "windowState", vbNormal
        Else
            Config.WriteSetting "windowState", vbMaximized
        End If
        Config.WriteSetting "lastOutputDir", Project.destinationDir
        Config.WriteSetting "lastBrowseDir", fileDialog.lastFolder
    End If
    
    ' Destroy global objects
    appLog.Add "Destroying classes..."
    Set Project = Nothing
    Set langCodes = Nothing
    Set Countries = Nothing
    Set colorSpaces = Nothing
    Set codePages = Nothing
    Set audioMappings = Nothing
    Set videoMappings = Nothing
    Set Burners = Nothing
    Set CMD = Nothing
    Set FS = Nothing
    Set mmcProfiles = Nothing
    Set mmcFeatures = Nothing
    Set cpuInfo = Nothing
    Set fileDialog = Nothing
    Set Dbg = Nothing
    
    ' Unload forms
    appLog.Add "Unloading forms..."
    Unload frmSplash
    Unload frmDialog
    Unload frmStatus
    Unload frmWelcome
    Unload frmTetris
    Unload frmSubtitle
    Unload frmProgress
    Unload frmEncodeError
    Unload frmAbout
    Unload frmProjectSettings
    Unload frmSelectTrack
    Unload frmTitle
    Unload frmEditTrack
    Unload frmAdvOptsVideo
    Unload frmSubPreview
    Unload frmMenuSettings
    Unload frmMenuPreview
    Unload frmMain
    
    Config.SaveConfiguration DATA_PATH & "dvdflick.cfg"
    Set Config = Nothing
    
    appLog.Add "Destroying application mutex"
    destroyAppMutex
    
    appLog.Add "Bye-bye."
    Set appLog = Nothing

    If endProgram Then End

End Sub
