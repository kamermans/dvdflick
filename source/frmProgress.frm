VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6AA9104C-4921-11D4-BD2E-0800460222F0}#2.0#0"; "trayicon_handler.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DVD Flick - Encoding unnamed"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9255
   Icon            =   "frmProgress.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar prgBar 
      Height          =   315
      Left            =   360
      TabIndex        =   17
      Top             =   4620
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ilActions 
      Left            =   1560
      Top             =   5580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgress.frx":3672
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgress.frx":4164
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgress.frx":4C56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbWhenDone 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmProgress.frx":5748
      Left            =   6660
      List            =   "frmProgress.frx":5758
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
   End
   Begin dvdflick.ctlSeparator ctlSeparator1 
      Height          =   5115
      Left            =   6420
      Top             =   480
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   9022
   End
   Begin vbRad.TrayIcon Tray 
      Left            =   960
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.ComboBox cmbPriority 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmProgress.frx":578B
      Left            =   6660
      List            =   "frmProgress.frx":579B
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3180
      Width           =   2415
   End
   Begin VB.Timer tElapsed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   420
      Top             =   5640
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   5700
      Width           =   2475
   End
   Begin MSComctlLib.Toolbar tlbActions 
      Height          =   1800
      Left            =   6720
      TabIndex        =   0
      Top             =   540
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   3175
      ButtonWidth     =   3995
      ButtonHeight    =   794
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilActions"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Abort  "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Minimize to tray  "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "        Entertain me  "
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label lblWhenDone 
      Caption         =   "When done..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6660
      TabIndex        =   16
      Top             =   3780
      Width           =   2415
   End
   Begin VB.Label lblPrct 
      Alignment       =   1  'Right Justify
      Caption         =   "FPS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5400
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblSubStatus 
      Caption         =   "Index"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   5280
      Width           =   4935
   End
   Begin VB.Label lblStatus 
      Caption         =   "Index"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label lblProcPriority 
      Caption         =   "Process priority"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6660
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Image imgTop 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label lblFinalize 
      Caption         =   "Finalize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   11
      Top             =   4005
      Width           =   1695
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   6
      Left            =   540
      Top             =   3960
      Width           =   240
   End
   Begin VB.Label lblElapsed 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6540
      TabIndex        =   10
      Top             =   5040
      Width           =   2475
   End
   Begin VB.Label lblPrepare 
      Caption         =   "Prepare files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   9
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label lblEncVideo 
      Caption         =   "Encode video"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   8
      Top             =   1305
      Width           =   1635
   End
   Begin VB.Label lblEncAudio 
      Caption         =   "Encode audio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   7
      Top             =   1845
      Width           =   1395
   End
   Begin VB.Label lblCombine 
      Caption         =   "Combine sources"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   6
      Top             =   2385
      Width           =   1755
   End
   Begin VB.Label lblAddSubs 
      Caption         =   "Add subtitles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   5
      Top             =   2925
      Width           =   2055
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   0
      Left            =   540
      Picture         =   "frmProgress.frx":57C9
      Top             =   720
      Width           =   240
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   1
      Left            =   540
      Picture         =   "frmProgress.frx":5B53
      Top             =   1260
      Width           =   240
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   2
      Left            =   540
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   3
      Left            =   540
      Top             =   2340
      Width           =   240
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   4
      Left            =   540
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image imgProgress 
      Height          =   240
      Index           =   5
      Left            =   540
      Top             =   3420
      Width           =   240
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author DVD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   4
      Top             =   3465
      Width           =   1695
   End
End
Attribute VB_Name = "frmProgress"
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
'   File purpose: DVD creation progress monitoring.
'
Option Explicit
Option Compare Binary
Option Base 0


' Images
Private workingImage As Picture
Private doneImage As Picture

' Time elapsed
Private msPassed As Long
Private lastTick As Long

' Progress count
Private Progress As Long


Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Sub Setup()

    Dim A As Long
    
    
    Progress = 0
    
    For A = 0 To imgProgress.UBound
        Set imgProgress(A).Picture = Nothing
    Next A
    Set imgProgress(0).Picture = workingImage
        
    ' Enable buttons
    tlbActions.Buttons(1).Enabled = True
    cmdClose.Enabled = False
    
    ' Shutdown is disabled by default
    cmbWhenDone.ListIndex = 0
    
    ' Encoding priority class
    cmbPriority.ListIndex = Project.encodePriority
    
    updateCaption
    msPassed = 0
    lastTick = GetTickCount
    tElapsed_Timer
    tElapsed.Enabled = True

End Sub


Private Sub updateCaption()

    Me.Caption = "Encoding " & Project.Title & " - " & App.Title

End Sub


Public Sub resetStatus()

    lblStatus.Caption = ""
    lblSubStatus.Caption = ""
    lblPrct.Caption = ""
    prgBar.Value = 0
    prgBar.Visible = False

End Sub


Public Sub setStatus(ByVal Message As String)

    lblStatus.Caption = Message

End Sub


Public Sub setSubStatus(ByVal Message As String)

    lblSubStatus.Caption = Message

End Sub


' Change encoding priority class
Private Sub cmbPriority_Click()
    
    If cmbPriority.ListIndex = -1 Then Exit Sub
    
    ' Set encoder process priority default
    If cmbPriority.ListIndex = EP_AboveNormal Then
        newPriorityClass = PRIORITY_CLASS_ABOVE_NORMAL
    ElseIf cmbPriority.ListIndex = EP_Normal Then
        newPriorityClass = PRIORITY_CLASS_NORMAL
    ElseIf cmbPriority.ListIndex = EP_BelowNormal Then
        newPriorityClass = PRIORITY_CLASS_BELOW_NORMAL
    ElseIf cmbPriority.ListIndex = EP_Idle Then
        newPriorityClass = PRIORITY_CLASS_IDLE
    End If
    
    Project.encodePriority = cmbPriority.ListIndex

End Sub


' Cancel encoding process
Private Sub progressAbort()

    If frmDialog.Display("Are you sure you want to abort the encoding process?", YesNo Or Question) = buttonNo Then Exit Sub
    
    tlbActions.Buttons(1).Enabled = False
    modEncode.cancelError = True
    haltShellEx = True

End Sub


Private Sub cmdClose_Click()

    Me.Hide
    frmMain.Show

End Sub


Public Sub Advance()

    ' Set current image to done
    Set imgProgress(Progress).Picture = doneImage
    
    ' Advance
    Progress = Progress + 1
    If Progress > imgProgress.UBound Then Exit Sub
    
    ' Set new image to working
    Set imgProgress(Progress).Picture = workingImage

End Sub


' Minimize to tray
Private Sub progressMinimize()

    Set Tray.Icon = Me.Icon
    Tray.ToolTip = "DVD Flick Progress"
    Tray.Start
    
    Me.Hide

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

    ctlSeparator1.Orientation = orientateVertical
    
    ' Disable form close button
    setCloseButton Me.hWnd, False
    
    ' Get images
    Set doneImage = imgProgress(0).Picture
    Set workingImage = imgProgress(1).Picture
    
    ' Caption matches app title
    Me.Caption = App.Title & " Progress"
    
    
    ' Resize to fit larger fonts
    tlbActions.Width = tlbActions.ButtonWidth
    tlbActions.Left = (Me.Width \ 15) - tlbActions.Width - 24
    ctlSeparator1.Left = tlbActions.Left - 24
    
    prgBar.Width = ctlSeparator1.Left - prgBar.Left - 24
    lblPrct.Left = prgBar.Left + prgBar.Width - lblPrct.Width
    lblStatus.Width = prgBar.Width - lblPrct.Width - 24
    lblSubStatus.Width = lblStatus.Width
    
    lblWhenDone.Left = tlbActions.Left
    cmbWhenDone.Left = tlbActions.Left
    cmbWhenDone.Width = tlbActions.Width
    
    lblProcPriority.Left = tlbActions.Left
    cmbPriority.Left = tlbActions.Left
    cmbPriority.Width = tlbActions.Width
    
    lblElapsed.Width = tlbActions.Width
    lblElapsed.Left = tlbActions.Left
    
    cmdClose.Width = tlbActions.Width + 16
    
End Sub


Public Sub showBar(ByVal Visible As Boolean)

    prgBar.Visible = Visible

End Sub


Public Sub Finish()

    If Me.Visible = False Then Tray_Click 0, 0, 0, 0
    
    Me.Caption = App.Title & " - Done"
    
    tElapsed.Enabled = False
    cmdClose.Enabled = True
    tlbActions.Buttons(1).Enabled = False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then Cancel = True

End Sub


Private Sub tBar_Timer()

    prgBar.Value = prgBar.Value

End Sub


Private Sub tElapsed_Timer()

    msPassed = msPassed + (GetTickCount - lastTick)
    lblElapsed.Caption = "Time elapsed:" & vbCrLf & visualTime(msPassed \ 1000)
    
    lastTick = GetTickCount

End Sub


Private Sub Tray_MouseUp(ByVal Button As Integer)

    If Me.Visible = False Then Tray_Click 0, 0, 0, 0

End Sub


Private Sub tlbActions_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            progressAbort
        Case 2
            progressMinimize
        Case 4
            frmTetris.Show
    End Select

End Sub


Private Sub Tray_Click(Button As Integer, Shift As Integer, x As Single, y As Single)

    Me.Show
    Tray.Remove

End Sub
