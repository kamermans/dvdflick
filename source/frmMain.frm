VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "DVD Flick"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   539
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ilEdit 
      Left            =   10200
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3672
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4164
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5748
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":623A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToobar 
      Left            =   9480
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":781E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":84F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB86
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B860
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C53A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D214
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DEEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbEditFunctions 
      Height          =   450
      Left            =   9540
      TabIndex        =   2
      Top             =   840
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   794
      ButtonWidth     =   3016
      ButtonHeight    =   794
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilEdit"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Add title...  "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Edit title...  "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Remove title  "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Move up  "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Move down  "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Compact list  "
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1429
      ButtonWidth     =   2328
      ButtonHeight    =   1429
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ilToobar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save project"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Project settings"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Menu settings"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Create DVD"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guide"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Update"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin dvdflick.ctlFancyList flTitles 
      Height          =   6375
      Left            =   780
      TabIndex        =   1
      Top             =   840
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   11245
   End
   Begin VB.PictureBox picMeter 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   180
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   495
      Begin VB.Label lblPercentage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdBrowseDestFolder 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   9420
      TabIndex        =   4
      ToolTipText     =   "Browse for destination folder"
      Top             =   7620
      Width           =   1695
   End
   Begin VB.TextBox txtDestFolder 
      Height          =   315
      Left            =   780
      TabIndex        =   3
      Top             =   7620
      Width           =   8535
   End
   Begin VB.Label lblProjectInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   9540
      TabIndex        =   8
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblDestFolder 
      Caption         =   "Project destination folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   7380
      Width           =   2835
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'   File purpose: Main form, display titles and actions
'
Option Explicit
Option Compare Binary
Option Base 0


Private compactList As Boolean

Private keyShift As Long

' List item
Private Const itemHeight As Long = 88
Private Const itemPadding As Long = 8

' Disc meter
Private discPercentage As Single
Private discFull As Boolean


Private Sub changeListMode()

    If compactList Then
        compactList = False
        tlbEditFunctions.Buttons(8).Caption = "Compact list  "
        
        flTitles.itemHeight = itemHeight
        flTitles.Padding = itemPadding
        flTitles.Font.Size = 8
        flTitles.Thumbnails = True
        
    Else
        compactList = True
        tlbEditFunctions.Buttons(8).Caption = "Expand list  "
        
        flTitles.itemHeight = itemHeight - 52
        flTitles.Padding = itemPadding - 2
        flTitles.Font.Size = 7
        flTitles.Thumbnails = False
        
    End If
    
    refreshData

End Sub


Private Sub flTitles_DblClick()

    titleEdit

End Sub


Private Sub flTitles_KeyUp(ByVal Key As Long, ByVal Shift As Boolean)

    If Key = 46 And Shift = False Then titleRemove

End Sub


Private Sub Form_Activate()

    Me.Caption = App.Title & " - " & Project.Title
    
End Sub


Private Sub titleAdd()

    Dim A As Long
    Dim fileList As Dictionary
    

    Set fileList = fileDialog.openFile(Me.hWnd, "Select video file", titleFiles, "", "", cdlOFNFileMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly)
    If fileList.Count Then

        Me.Enabled = False
        For A = 0 To fileList.Count - 1
            If Not addTitleFile(fileList.Items(A)) Then Exit For
        Next A
    
        refreshData
        
        ' Focus on last added item
        flTitles.selectedItem = flTitles.Count - 1
        flTitles.focusOn flTitles.selectedItem
        
        Me.Enabled = True
        
    End If
    
    Set fileList = Nothing

End Sub


' Adds a new title to the current project
Private Function addTitleFile(ByVal fileName As String) As Boolean

    Dim Source As clsSource
    Dim myTitle As clsTitle
    Dim baseName As String
    Dim myFolder As Folder
    Dim myFile As File
    Dim mySub As clsSubFile
    
    On Error GoTo addTitleError
    
    
    ' Get source class
    Set Source = Project.getSource(fileName)
    
    ' Load
    If Not (Source Is Nothing) Then
    
        Set myTitle = Project.addTitle
        If myTitle Is Nothing Then
            Set Source = Nothing
            
            On Error GoTo 0
            Exit Function
        End If
        
        If Not myTitle.scanFile(Source) Then
            Project.Titles.Remove Project.Titles.Count - 1
                
        ElseIf myTitle.Videos.Count = 0 Then
            frmDialog.Display fileName & " has no usable video track. It will not be added to the project.", Exclamation Or OkOnly
            Project.Titles.Remove Project.Titles.Count - 1
        
        ' Title scanned succesfully
        Else
            Project.Modified = True
            
            ' Set default settings
            myTitle.chapterCount = Config.ReadSetting("titleChapterCount", myTitle.chapterCount)
            myTitle.chapterInterval = Config.ReadSetting("titleChapterInterval", myTitle.chapterInterval)
            myTitle.chapterOnSource = Config.ReadSetting("titleChapterOnSource", myTitle.chapterOnSource)
            myTitle.Name = FS.GetBaseName(fileName)
            
            ' Add subtitles with similar filename
            baseName = LCase$(FS.GetBaseName(fileName))
            Set myFolder = FS.GetFolder(FS.GetParentFolderName(fileName))
            For Each myFile In myFolder.Files
                If myFile.Name <> FS.GetFileName(fileName) And LCase$(FS.GetBaseName(myFile.Name)) = baseName Then
                    Set mySub = New clsSubFile
                    If mySub.openFrom(myFolder.Path & "\" & myFile.Name) Then myTitle.addSub mySub
                End If
            Next myFile
        End If
        
    ' Invalid
    Else
        frmDialog.Display fileName & " could not be opened. The file may be corrupted or it's format may be unsupported.", Exclamation Or OkOnly
    
    End If
    
    
    frmStatus.Hide
    
    addTitleFile = True
    Set Source = Nothing
    Set myTitle = Nothing
    On Error GoTo 0
    Exit Function
    
    
addTitleError:
    frmDialog.Display "Error " & Err.Number & " from " & Err.Source & ": " & Err.Description, Critical Or OkOnly

End Function


' Refreshes list of titles
Private Sub refreshData()

    Dim A As Long
    Dim B As Long
    Dim myTitle As clsTitle
    Dim myVideo As clsVideo
    Dim oldMod As Boolean
    Dim oldSel As Long
    Dim itemText As String
    Dim requiredSpace  As Long
    
    Dim useThumb As clsGDIImage
    Dim Sizes As Dictionary
    Dim Info As String
    
    
    oldSel = flTitles.selectedItem
    
    ' Populate fancy list
    flTitles.Refresh = False
    flTitles.Clear
    For A = 0 To Project.Titles.Count - 1
        Set myTitle = Project.Titles.Item(A)
        Set myVideo = myTitle.Videos.Item(0)
        
        ' Create multi-line item text
        If InStr(myTitle.Name, vbNewLine) > 0 Then
            itemText = Left$(myTitle.Name, InStr(myTitle.Name, vbNewLine) - 1) & vbNewLine
        Else
            itemText = myTitle.Name & vbNewLine
        End If
        
        If compactList = False Then
            
            ' Video name concatenation
            If myTitle.Videos.Count > 1 Then
                itemText = itemText & myVideo.Source.fileName
                For B = 1 To myTitle.Videos.Count - 1
                    itemText = itemText & ", " & FS.GetFileName(myTitle.Videos.Item(B).Source.fileName)
                Next B
                itemText = itemText & vbNewLine
            Else
                itemText = itemText & myVideo.Source.fileName & vbNewLine
            End If
            
            ' Items
            itemText = itemText & "Duration: " & visualTime(myTitle.Duration) & vbNewLine
            itemText = itemText & myTitle.audioTracks.Count & " audio track(s)" & vbNewLine
            itemText = itemText & myTitle.Subtitles.Count & " subtitle(s)"
            
            ' Thumbnail
            Set useThumb = myVideo.Thumbnail
            If useThumb Is Nothing Then Set useThumb = noThumb
            
        Else
            itemText = itemText & visualTime(myTitle.Duration) & ", " & myTitle.Videos.Count & " sources, " & myTitle.audioTracks.Count & " audio track(s), " & myTitle.Subtitles.Count & " subtitle(s)"
            
        End If
            
        flTitles.addItem itemText, resizeToMatch(useThumb, flTitles.imageWidth, flTitles.imageHeight, myVideo.PAR)
    Next A
    
    If oldSel >= 0 And oldSel < flTitles.Count Then
        flTitles.selectedItem = oldSel
    ElseIf flTitles.Count > 0 Then
        flTitles.selectedItem = 0
    End If
    flTitles.Refresh = True
    
    ' Remember destination directory
    oldMod = Project.Modified
    txtDestFolder.Text = Project.destinationDir
    Project.Modified = oldMod
    
    ' Update disc meter
    updateMeter
    renderMeter
    
    ' Update project info
    Set Sizes = Project.calculateSizes
    requiredSpace = modEncode.requiredSpace
    
    Info = "Total duration" & vbNewLine & visualTime(Project.Duration) & vbNewLine
    Info = Info & Project.Titles.Count & " titles" & vbNewLine & vbNewLine
    Info = Info & "Average bitrate" & vbNewLine & Sizes("avgBitRate") & " Kbit\s" & vbNewLine & vbNewLine
    Info = Info & "Harddisk space required" & vbNewLine & requiredSpace \ 1024 & " Mb" & vbNewLine & requiredSpace & " Kb"
    lblProjectInfo.Caption = Info
    
    ' Disc meter tooltip
    picMeter.ToolTipText = Sizes("sizeUsed") \ 1024 & " Mb of " & Sizes("discSize") \ 1024 & " Mb used"
    
    Set myTitle = Nothing
    
End Sub


Private Sub cmdBrowseDestFolder_Click()

    Dim Folder As String
    
    
    Folder = fileDialog.selectFolder(Me.hWnd, "Please select the destination folder of the DVD.")
    If LenB(Folder) <> 0 Then
        
        ' Disallow drive root
        If Right$(Folder, 2) = ":\" Or Right$(Folder, 1) = ":" Then
            frmDialog.Display "You cannot use a drive's root as project destination folder.", Information Or OkOnly
            Exit Sub
        End If
        
        ' Disallow My Documents
        If LCase$(FS.GetAbsolutePathName(Folder)) = LCase$(getSpecialPath(myDocuments)) Then
            frmDialog.Display "You cannot use My Documents' root as project destination folder. Please use a folder inside My Documents instead.", Information Or OkOnly
            Exit Sub
        End If
        
        ' Disallow Desktop
        If LCase$(FS.GetAbsolutePathName(Folder)) = LCase$(getSpecialPath(userDesktop)) Then
            frmDialog.Display "You cannot use your desktop as project destination folder. Please use a folder on your desktop instead.", Information Or OkOnly
            Exit Sub
        End If
        
        
        txtDestFolder.Text = Folder
    End If

End Sub


Private Sub titleEdit()

    If flTitles.selectedItem = -1 Then Exit Sub
    
    ' Edit title
    frmTitle.Setup flTitles.selectedItem
    frmTitle.Show 1
    
    refreshData
    
End Sub


Private Sub titleMoveUp()

    Dim A As Long
    Dim moveCount As Long
    Dim Selected As Long
    
    
    ' Valid selection
    If flTitles.selectedItem = -1 Then Exit Sub
    Selected = flTitles.selectedItem
    If Selected <= 0 Then Exit Sub
    
    ' Move by amount
    If keyShift Then moveCount = 10 Else moveCount = 1
    For A = 1 To moveCount
        
        Project.Titles.moveBackward Selected
        Selected = Selected - 1
        flTitles.moveUp Selected
        If Selected <= 0 Then Exit For
        
    Next A
    
    flTitles.selectedItem = Selected

End Sub


Private Sub titleMoveDown()

    Dim A As Long
    Dim moveCount As Long
    Dim Selected As Long
    
    
    ' Valid selection
    If flTitles.selectedItem = -1 Then Exit Sub
    Selected = flTitles.selectedItem
    If Selected = Project.Titles.Count - 1 Then Exit Sub
    
    ' Move by amount
    If keyShift Then moveCount = 10 Else moveCount = 1
    For A = 1 To moveCount
        
        Project.Titles.moveForward Selected
        Selected = Selected + 1
        flTitles.moveDown Selected
        If Selected = Project.Titles.Count - 1 Then Exit For
        
    Next A
    
    flTitles.selectedItem = Selected

End Sub


Private Sub titleRemove()

    If flTitles.selectedItem = -1 Then Exit Sub
    If frmDialog.Display("Are you sure you want to remove this title?", Question Or YesNo) = buttonNo Then Exit Sub
    
    ' Remove and mark modified
    Project.Titles.Remove flTitles.selectedItem
    Project.Modified = True
    
    refreshData

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    keyShift = Shift

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."
    
    ' Setup fancylist
    flTitles.itemHeight = itemHeight
    flTitles.Padding = itemPadding
    flTitles.imageWidth = 90
    flTitles.imageHeight = 72
    compactList = False
    
    refreshData

End Sub


Private Sub flTitles_OLEDragDrop(ByRef Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    Dim A As Long
    
    
    For A = 1 To Data.Files.Count
        If Not addTitleFile(Data.Files.Item(A)) Then Exit For
    Next A
    
    refreshData

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Project Is Nothing Then Exit Sub
    
    ' Project modified, ask if user really wants to exit
    If Project.Modified = True And unattendMode = False Then
        If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
        If frmDialog.Display("Your project has been modified. Are you sure you want to exit?", Question Or YesNo) = buttonNo Then
            Cancel = True
            Exit Sub
        End If
    End If

End Sub


Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    
    flTitles.Top = tlbMain.Height
    tlbEditFunctions.Top = tlbMain.Height
    picMeter.Top = tlbMain.Height
    
    picMeter.Height = Me.ScaleHeight - picMeter.Top - 12
    
    tlbEditFunctions.Width = tlbEditFunctions.ButtonWidth
    
    flTitles.Height = Me.ScaleHeight - flTitles.Top - 64
    flTitles.Width = Me.ScaleWidth - flTitles.Left - tlbEditFunctions.ButtonWidth - 24
    
    tlbEditFunctions.Left = flTitles.Left + flTitles.Width + 10
    
    txtDestFolder.Top = flTitles.Top + flTitles.Height + 28
    cmdBrowseDestFolder.Top = txtDestFolder.Top
    lblDestFolder.Top = txtDestFolder.Top - lblDestFolder.Height - 4
    
    txtDestFolder.Width = flTitles.Width
    cmdBrowseDestFolder.Width = tlbEditFunctions.ButtonWidth
    cmdBrowseDestFolder.Left = txtDestFolder.Left + txtDestFolder.Width + 8
    
    lblProjectInfo.Width = tlbEditFunctions.Width
    lblProjectInfo.Top = (flTitles.Top + flTitles.Height) - lblProjectInfo.Height - 8
    lblProjectInfo.Left = flTitles.Left + flTitles.Width + 12
    
    renderMeter
    
    On Error GoTo 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

    modMain.Quit

End Sub


Private Sub tlbEditFunctions_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            titleAdd
        Case 2
            titleEdit
        Case 3
            titleRemove
            
        Case 5
            titleMoveUp
        Case 6
            titleMoveDown
            
        Case 8
            changeListMode
    End Select

End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            newProject
        Case 2
            openProject
        Case 3
            saveProject
            
        Case 6
            frmMenuSettings.Setup
            frmMenuSettings.Show 1
        Case 5
            frmProjectSettings.manualActivate
            frmProjectSettings.Show 1
            refreshData
            
        Case 8
            startEncode
            
        Case 10
            Execute APP_PATH & "\guide\index.html", "", EP_Normal, False
        Case 11
            frmAbout.Show 1
        Case 12
            Update
    End Select

End Sub


Private Sub Update()

    Dim Ver As String
    Dim Build As String
    
    
    Ver = App.Major & App.Minor
    Build = App.Revision
    
    Execute "http://www.dvdflick.net/update.php?v=" & Ver & "&b=" & Build, "", EP_Normal, False

End Sub


Public Sub startEncode()

    Dim A As Long
    Dim driveName As String
    Dim spaceFree As Variant
    Dim needSpace As Long
    Dim driveChar As String
    Dim Result As dialogResultConstants
    
    
    Warnings = ""

    ' No titles?
    If flTitles.Count = 0 Then
        frmDialog.Display "There are no titles present in your project.", Information Or OkOnly
        Exit Sub
    End If

    ' Project destination folder is on a DVD\Recordable drive?
    If InStr(Project.destinationDir, ":\") <> 0 Then
        driveChar = UCase$(Left$(Project.destinationDir, 1))
        For A = 0 To Burners.deviceCount - 1
            If Burners.getDevice(A).deviceDriveChar = driveChar Then
                frmDialog.Display "The current project destination folder points to a removable disc drive (" & Burners.getDevice(A).deviceDriveChar & "). You must choose a folder on a harddrive as the project destination. It will be used as a temporary place to store and process files before burning them onto a disc.", Information Or OkOnly
                Exit Sub
            End If
        Next A
    End If

    ' Odd paths? This needs to be fixed properly sometime, but something is amiss.
    If oddPath(Project.destinationDir) Then
        frmDialog.Display "Your project's destination folder path contains non-standard characters, like ë or æ. Please use a different path without any of such characters.", Information Or OkOnly
        Exit Sub
    End If

    ' Disc full?
    If discFull And Not unattendMode Then
        If frmDialog.Display("Your project's DVD disc is full. You must remove titles or audio tracks, or set a different target bitrate in order to be able to fit the contents of your project onto the target disc size specified. Are you sure you want to continue?", Exclamation Or YesNo) = buttonNo Then Exit Sub
        Warnings = Warnings & "DVDDiscFull | "
    End If
    
    
    ' Delete and reconstruct destination folder if needed
    If FS.FolderExists(Project.destinationDir) = True Then
        If Not unattendMode Then If frmDialog.Display("WARNING!" & vbNewLine & vbNewLine & "The destination folder (" & Project.destinationDir & ") already exists. If you continue the destination folder's contents will be DELETED. Are you sure you want to continue?", YesNo Or Critical) = buttonNo Then Exit Sub
        
        While emptyDestinationFolder(Project.destinationDir) = False
            If frmDialog.Display("The folder " & Project.destinationDir & " is in use by another program. Please close this program first before continuing.", retryCancel Or Exclamation) = buttonCancel Then Exit Sub
        Wend
    End If
    
    If Not constructPath(Project.destinationDir) Then
        frmDialog.Display "Unable to create the destination folder.", Exclamation Or OkOnly
        Exit Sub
    End If


    ' Disk space? In Megabytes
    driveName = FS.GetDriveName(Project.destinationDir)
    spaceFree = FS.GetDrive(driveName).AvailableSpace / 1024 / 1024
    needSpace = modEncode.requiredSpace / 1024
    
    If spaceFree < needSpace Then
        If frmDialog.Display("There is not enough free hard drive space available on drive " & driveName & " to create your DVD. You need at least " & needSpace & " Megabytes of free space available to be able to complete the encoding. Are you sure you want to continue?", Exclamation Or YesNo) = buttonNo Then Exit Sub
        Warnings = Warnings & "NoFreeSpace | "
    End If

    ' NTFS recommended
    If FS.GetDrive(driveName).FileSystem <> "NTFS" And Not unattendMode Then
        frmDialog.Display "It is recommended to put your project's destination folder on a drive with the NTFS file system. Older file systems like FAT32 cannot handle files larger than 2 Gigabytes in size, which can be created during the encoding process. If this happens, encoding will fail.", Exclamation Or OkOnly
        Warnings = Warnings & "NTFSRequired | "
    End If

    ' Blank disc inserted?
    If Project.enableBurning And Config.ReadSetting("dialogEnsureBlankDisc", 1) = 1 And Not unattendMode Then
        Result = frmDialog.Display("The project will be burnt to disc after it has finished encoding. Please make sure there is a blank disc in the recorder drive you selected.", okCancel Or Information, True)
        If (Result And checkNotAgain) Then Config.WriteSetting "dialogEnsureBlankDisc", 0
        If (Result And buttonCancel) Then Exit Sub
        Warnings = Warnings & "EnsureEmptyDiscInserted | "
    End If

    ' Proceed yes or no?
'    If Config.ReadSetting("dialogProceedEncoding", 1) = 1 And Not unattendMode Then
'        Result = frmDialog.Display("Are you sure you want to proceed?", Question Or YesNo, True)
'        If (Result And checkNotAgain) Then Config.WriteSetting "dialogProceedEncoding", 0
'        If (Result And buttonNo) Then Exit Sub
'    End If

    modEncode.Start

End Sub


Private Sub updateMeter()

    Dim Sizes As Dictionary
    
    
    ' Calculate percentage of space in use
    Set Sizes = Project.calculateSizes
    discPercentage = (Sizes("sizeUsed") / Sizes("discSize")) * 100
    If discPercentage > 100 Then discPercentage = 100
    
    ' Disc full
    If discPercentage = 100 Then
        discFull = True
    Else
        discFull = False
    End If
    
    lblPercentage = CInt(discPercentage) & "%"
    
    Set Sizes = Nothing
    
End Sub


Private Sub renderMeter()
    
    Dim Bit As Single
    Dim drawColor As Long
    Dim drawColorShade As Long
    Dim drawColorShine As Long
    
    
    Bit = picMeter.Height / 100
    
    ' Disc full color
    If discPercentage = 100 Then
        drawColor = RGB(255, 63, 31)
        drawColorShine = RGB(255, 127, 63)
        drawColorShade = RGB(192, 31, 0)
    
    ' Normal color
    Else
        drawColor = RGB(244, 191, 43)
        drawColorShine = RGB(244, 227, 53)
        drawColorShade = RGB(191, 139, 38)
    End If
    
    ' Render
    picMeter.Cls
    picMeter.Line (0, (100 - discPercentage) * Bit)-(picMeter.Width, picMeter.Height), drawColor, BF
    picMeter.Line (0, (100 - discPercentage) * Bit)-(2, picMeter.Height), drawColorShine, BF
    picMeter.Line (picMeter.Width - 7, (100 - discPercentage) * Bit)-(picMeter.Width, picMeter.Height), drawColorShade, BF

End Sub


Private Sub newProject()

    If Project.Modified = True Then
        If frmDialog.Display("Your project has not been saved. Are you sure you want to continue?", Question Or YesNo) = buttonNo Then Exit Sub
    End If
    
    Project.Reset
    
    refreshData
    Form_Activate

End Sub


Private Sub openProject()

    Dim fileList As Dictionary
    
    
    If Project.Modified = True Then
        If frmDialog.Display("Your project has not been saved. Are you sure you want to continue?", Question Or YesNo) = buttonNo Then Exit Sub
    End If
    
    Set fileList = fileDialog.openFile(Me.hWnd, "Open project", "DVD Flick project (*.dfproj)|*.dfproj|All files (*.*)|*.*", "", "", cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly)
    If fileList.Count Then
        If Not Project.unSerialize(fileList.Items(0)) Then
            frmDialog.Display fileList.Items(0) & " is not a valid DVD Flick project file.", Exclamation Or OkOnly
        End If
        
        refreshData
        Form_Activate
    End If

End Sub


Private Sub saveProject()

    Dim fileName As String
    Dim lastFolder As String
    
    
    If LenB(Project.fileName) <> 0 Then lastFolder = FS.GetParentFolderName(Project.fileName)
    
    fileName = fileDialog.saveFile(Me.hWnd, "Save project", "DVD Flick project (*.dfproj)|*.dfproj|All files (*.*)|*.*", lastFolder, Project.fileName, cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt)
    If LenB(fileName) <> 0 Then
        If LCase$(Right$(fileName, 7)) <> ".dfproj" Then fileName = fileName & ".dfproj"

        Project.Serialize fileName
        Project.Modified = False
        
        Form_Activate
    End If

End Sub


Private Sub txtDestFolder_Change()

    Project.Modified = True
    Project.destinationDir = txtDestFolder.Text

End Sub
