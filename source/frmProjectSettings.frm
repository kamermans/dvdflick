VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmProjectSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project settings"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjectSettings.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin dvdflick.ctlSeparator ctlSeparator1 
      Height          =   30
      Left            =   180
      Top             =   4860
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   53
   End
   Begin dvdflick.ctlFancyList flMenu 
      Height          =   4515
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2295
      _extentx        =   4048
      _extenty        =   7964
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2580
      TabIndex        =   29
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdUseDefault 
      Caption         =   "Use as defaults"
      Height          =   375
      Left            =   4740
      TabIndex        =   30
      Top             =   5040
      Width           =   2310
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   7140
      TabIndex        =   31
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox framPlayback 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   2640
      ScaleHeight     =   4515
      ScaleWidth      =   6555
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   180
      Width           =   6555
      Begin VB.CheckBox chkLoopPlayback 
         Caption         =   "Loop to first title when done playing last"
         Height          =   315
         Left            =   300
         TabIndex        =   27
         Top             =   840
         Width           =   4275
      End
      Begin VB.ComboBox cmbWhenPlayed 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":000C
         Left            =   3480
         List            =   "frmProjectSettings.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   180
         Width           =   2955
      End
      Begin VB.CheckBox chkEnableFirstSub 
         Caption         =   "Always enable first subtitle"
         Height          =   315
         Left            =   300
         TabIndex        =   28
         Top             =   1320
         Width           =   3195
      End
      Begin VB.Label Label5 
         Caption         =   "After a title has finished playing"
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.PictureBox framGeneral 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   2640
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   180
      Width           =   6555
      Begin VB.ComboBox cmbTargetSize 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":0062
         Left            =   2520
         List            =   "frmProjectSettings.frx":0081
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   2475
      End
      Begin VB.TextBox txtCustomSize 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CheckBox chkKeepFiles 
         Caption         =   "Keep intermediate encoded audio and video files"
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   3000
         Width           =   4635
      End
      Begin VB.ComboBox cmbEncodePriority 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":011C
         Left            =   2520
         List            =   "frmProjectSettings.frx":012C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1860
         Width           =   2475
      End
      Begin VB.TextBox txtThreadCount 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   2340
         Width           =   375
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   2520
         TabIndex        =   1
         Top             =   180
         Width           =   3855
      End
      Begin ComCtl2.UpDown udThreadCount 
         Height          =   315
         Left            =   2940
         TabIndex        =   6
         Top             =   2340
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtThreadCount"
         BuddyDispid     =   196617
         OrigLeft        =   68
         OrigTop         =   136
         OrigRight       =   84
         OrigBottom      =   157
         Max             =   8
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTargetSize 
         Caption         =   "Target size"
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblCustomSize 
         Caption         =   "Custom size"
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "MB"
         Height          =   195
         Left            =   3720
         TabIndex        =   47
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Encoder priority"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Image imgIEncodePriority 
         Height          =   360
         Left            =   5040
         MouseIcon       =   "frmProjectSettings.frx":015A
         MousePointer    =   99  'Custom
         Picture         =   "frmProjectSettings.frx":02AC
         Top             =   1860
         Width           =   360
      End
      Begin VB.Image imgIThreadCount 
         Height          =   360
         Left            =   3240
         MouseIcon       =   "frmProjectSettings.frx":0D8E
         MousePointer    =   99  'Custom
         Picture         =   "frmProjectSettings.frx":0EE0
         Top             =   2340
         Width           =   360
      End
      Begin VB.Label lblThreadCount 
         Caption         =   "Thread count"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Title"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox framVideo 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   2640
      ScaleHeight     =   4515
      ScaleWidth      =   6555
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   180
      Width           =   6555
      Begin VB.ComboBox cmbFormat 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":19C2
         Left            =   2520
         List            =   "frmProjectSettings.frx":19D2
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   2475
      End
      Begin VB.ComboBox cmbProfile 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":19F3
         Left            =   2520
         List            =   "frmProjectSettings.frx":1A03
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   660
         Width           =   2475
      End
      Begin VB.ComboBox cmbVideoBitrate 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":1A24
         Left            =   2520
         List            =   "frmProjectSettings.frx":1A43
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1260
         Width           =   2475
      End
      Begin VB.TextBox txtCustomBitrate 
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   1740
         Width           =   975
      End
      Begin VB.CommandButton cmdAdvancedVideo 
         Caption         =   "Advanced..."
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   2340
         Width           =   2055
      End
      Begin VB.Label lblTargetFormat 
         Caption         =   "Target format"
         Height          =   195
         Left            =   180
         TabIndex        =   55
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblEncodeProfile 
         Caption         =   "Encoding profile"
         Height          =   195
         Left            =   180
         TabIndex        =   54
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblTargetBitrate 
         Caption         =   "Target bitrate"
         Height          =   195
         Left            =   180
         TabIndex        =   53
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Image imgITargetFormat 
         Height          =   360
         Left            =   5040
         MouseIcon       =   "frmProjectSettings.frx":1A9F
         MousePointer    =   99  'Custom
         Picture         =   "frmProjectSettings.frx":1BF1
         Top             =   180
         Width           =   360
      End
      Begin VB.Image imgIEncodingProfile 
         Height          =   360
         Left            =   5040
         MouseIcon       =   "frmProjectSettings.frx":26D3
         MousePointer    =   99  'Custom
         Picture         =   "frmProjectSettings.frx":2825
         Top             =   660
         Width           =   360
      End
      Begin VB.Image imgITargetBitrate 
         Height          =   360
         Left            =   5040
         MouseIcon       =   "frmProjectSettings.frx":3307
         MousePointer    =   99  'Custom
         Picture         =   "frmProjectSettings.frx":3459
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Custom bitrate"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Kbit/s"
         Height          =   195
         Left            =   3600
         TabIndex        =   51
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.PictureBox framBurning 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   2640
      ScaleHeight     =   4515
      ScaleWidth      =   6555
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   180
      Width           =   6555
      Begin VB.CheckBox chkEjectTray 
         Caption         =   "Eject tray when done"
         Height          =   315
         Left            =   300
         TabIndex        =   21
         Top             =   3960
         Width           =   3915
      End
      Begin VB.CheckBox chkVerifyDisc 
         Caption         =   "Verify disc after burning"
         Height          =   315
         Left            =   300
         TabIndex        =   20
         Top             =   3540
         Width           =   3915
      End
      Begin VB.ComboBox cmbDrives 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1680
         Width           =   4635
      End
      Begin VB.CheckBox chkKeepISO 
         Caption         =   "Delete ISO image after burning"
         Height          =   315
         Left            =   300
         TabIndex        =   19
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtDiscLabel 
         Height          =   315
         Left            =   1740
         TabIndex        =   15
         Top             =   1200
         Width           =   4635
      End
      Begin VB.CheckBox chkEraseRewritable 
         Caption         =   "Automatically erase disc if it is rewritable"
         Height          =   315
         Left            =   300
         TabIndex        =   18
         Top             =   2700
         Width           =   3915
      End
      Begin VB.CheckBox chkBurnToDisc 
         Caption         =   "Burn project to disc"
         Height          =   315
         Left            =   300
         TabIndex        =   14
         Top             =   600
         Width           =   1755
      End
      Begin VB.CheckBox chkCreateISO 
         Caption         =   "Create ISO image"
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Top             =   180
         Width           =   3015
      End
      Begin VB.ComboBox cmbBurnSpeed 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":3F3B
         Left            =   1740
         List            =   "frmProjectSettings.frx":3F57
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblDrive 
         Caption         =   "Drive"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label lblDiscLabel 
         Caption         =   "Disc label"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label lblBurnSpeed 
         Caption         =   "Speed"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   2220
         Width           =   1395
      End
   End
   Begin VB.PictureBox framAudio 
      BorderStyle     =   0  'None
      Height          =   4515
      Left            =   2640
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   180
      Width           =   6555
      Begin VB.TextBox txtVolumeMod 
         Height          =   315
         Left            =   2520
         TabIndex        =   22
         Top             =   180
         Width           =   495
      End
      Begin VB.ComboBox cmbChannelCount 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":3F7E
         Left            =   2520
         List            =   "frmProjectSettings.frx":3F8E
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   660
         Width           =   1995
      End
      Begin VB.ComboBox cmbAudioBitrate 
         Height          =   315
         ItemData        =   "frmProjectSettings.frx":3FB4
         Left            =   2520
         List            =   "frmProjectSettings.frx":3FCA
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1140
         Width           =   1395
      End
      Begin ComCtl2.UpDown udVolumeMod 
         Height          =   315
         Left            =   3060
         TabIndex        =   23
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "txtVolumeMod"
         BuddyDispid     =   196655
         OrigLeft        =   3360
         OrigTop         =   360
         OrigRight       =   3615
         OrigBottom      =   675
         Increment       =   10
         Max             =   500
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblVolumeMod 
         Caption         =   "Volume modification"
         Height          =   255
         Left            =   180
         TabIndex        =   42
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label lblVolumeModPerc 
         Caption         =   "%"
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "Channel count"
         Height          =   255
         Left            =   180
         TabIndex        =   40
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "Bitrate"
         Height          =   255
         Left            =   180
         TabIndex        =   39
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "kbit\s"
         Height          =   255
         Left            =   4080
         TabIndex        =   38
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label lblChannelNote 
         Caption         =   $"frmProjectSettings.frx":3FEC
         Height          =   795
         Left            =   300
         TabIndex        =   37
         Top             =   1800
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmProjectSettings"
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
'   File purpose: Project settings dialog
'
Option Explicit
Option Compare Binary
Option Base 0


Private disableUpdates As Boolean


Private Sub chkBurnToDisc_Click()

    updateBurningState
    
End Sub


Private Sub chkCreateISO_Click()

    updateBurningState

End Sub


Private Sub cmbBurnSpeed_Click()

    If (cmbBurnSpeed.ListIndex > 5 Or cmbBurnSpeed.ListIndex < 2) And Not disableUpdates Then
        frmDialog.Display "Warning: burning at too high or too low speeds increases the risk of playback errors.", Exclamation Or OkOnly
    End If

End Sub


Private Sub cmbChannelCount_Click()

    If cmbChannelCount.ListIndex > 0 Then lblChannelNote.Visible = True Else lblChannelNote.Visible = False

End Sub


Private Sub cmbTargetSize_Click()

    ' Target size can be custom, enable editbox
    If cmbTargetSize.ListIndex = TS_Custom Then
        txtCustomSize.Enabled = True
        txtCustomSize.backColor = &H80000005
    Else
        txtCustomSize.Enabled = False
        txtCustomSize.backColor = &H80000014
    End If
    
    ' Target sizes are really stored in the custom box
    Select Case cmbTargetSize.ListIndex
        Case TS_650
            txtCustomSize.Text = 650
        Case TS_700
            txtCustomSize.Text = 700
        Case TS_800
            txtCustomSize.Text = 800
        Case TS_DVD
            txtCustomSize.Text = 4470
        Case TS_DVDDL
            txtCustomSize.Text = 8095
        Case TS_DVDRAM
            txtCustomSize.Text = 4360
        Case TS_MINIDVD
            txtCustomSize.Text = 1380
        Case TS_MINIDVDDL
            txtCustomSize.Text = 2530
    End Select

End Sub


Private Sub cmbVideoBitrate_Change()

    If cmbVideoBitrate.ListIndex = TB_Custom Then
        txtCustomBitrate.Enabled = True
    Else
        txtCustomBitrate.Text = (cmbVideoBitrate.ListIndex + 1) * 1000
        txtCustomBitrate.Enabled = False
    End If

End Sub


Private Sub cmbVideoBitrate_Click()

    cmbVideoBitrate_Change

End Sub


Private Sub cmdAccept_Click()

    ' Custom bitrate threshold warnings
    If cmbVideoBitrate.ListIndex = TB_Custom Then
        If CLng(txtCustomBitrate.Text) < MIN_VIDEO_BITRATE Then
            frmDialog.Display "Your custom bitrate is very low, visible quality loss can occur.", OkOnly Or Exclamation
        ElseIf CLng(txtCustomBitrate.Text) > MAX_VIDEO_BITRATE Then
            frmDialog.Display "Your custom bitrate is very high, your DVD player might not be able to play the resulting DVD.", OkOnly Or Exclamation
        End If
    End If
    
    
    ' General
    Project.Title = txtTitle.Text
    Project.targetSize = cmbTargetSize.ListIndex
    Project.customSize = CLng(txtCustomSize.Text)
    Project.encodePriority = cmbEncodePriority.ListIndex
    Project.keepFiles = chkKeepFiles.Value
    Project.threadCount = CLng(txtThreadCount.Text)
    
    ' Playback
    Project.loopPlayback = chkLoopPlayback.Value
    Project.whenPlayed = cmbWhenPlayed.ListIndex
    Project.enableFirstSub = chkEnableFirstSub.Value
    
    ' Burning
    Project.enableBurning = chkBurnToDisc.Value
    Project.createISO = chkCreateISO.Value
    Project.eraseRW = chkEraseRewritable.Value
    Project.deleteISO = chkKeepISO.Value
    Project.burnerName = cmbDrives.List(cmbDrives.ListIndex)
    Project.discLabel = txtDiscLabel.Text
    Project.burnSpeed = cmbBurnSpeed.List(cmbBurnSpeed.ListIndex)
    Project.verifyDisc = chkVerifyDisc.Value
    Project.ejectTray = chkEjectTray.Value
    
    ' Video
    Project.targetFormat = cmbFormat.ListIndex
    Project.targetBitrate = cmbVideoBitrate.ListIndex
    Project.customBitrate = CLng(txtCustomBitrate.Text)
    Project.encodeProfile = cmbProfile.ListIndex
    
    ' Audio
    Project.volumeMod = CLng(txtVolumeMod.Text)
    Project.channelCount = cmbChannelCount.ListIndex
    Project.audioBitRate = cmbAudioBitrate.ListIndex

    ' Video advanced
    With frmAdvOptsVideo
        Project.overscanBorders = .chkOverscanBorders.Value
        Project.overscanSize = CSng(.txtOverscanSize.Text)
        Project.halfRes = .chkHalfRes.Value
        Project.PSNR = .chkPSNR.Value
        Project.Deinterlace = .chkDeinterlace.Value
        Project.MPEG2Copy = .chkMPEG2Copy.Value
        Project.dcPrecision = CLng(.txtDC.Text)
        Project.Pulldown = .chkPulldown.Value
    End With
    
    Project.Modified = True
    Me.Hide

End Sub


Private Sub cmdAdvancedVideo_Click()

    Me.Hide
    frmAdvOptsVideo.Show 1
    Me.Show 1

End Sub


Private Sub cmdCancel_Click()

    Me.Hide

End Sub


Private Sub cmdUseDefault_Click()

    With Config
    
        ' General
        .WriteSetting "projectTitle", txtTitle.Text
        .WriteSetting "targetSize", cmbTargetSize.ListIndex
        .WriteSetting "customSize", CLng(txtCustomSize.Text)
        .WriteSetting "encodePriority", cmbEncodePriority.ListIndex
        .WriteSetting "keepFiles", chkKeepFiles.Value
        .WriteSetting "threadCount", CLng(txtThreadCount.Text)
        
        ' Playback
        .WriteSetting "loopPlayback", chkLoopPlayback.Value
        .WriteSetting "whenPlayed", cmbWhenPlayed.ListIndex
        .WriteSetting "enableFirstSub", chkEnableFirstSub.Value
        
        ' Video
        .WriteSetting "targetFormat", cmbFormat.ListIndex
        .WriteSetting "targetBitRate", cmbVideoBitrate.ListIndex
        .WriteSetting "customBitRate", CLng(txtCustomBitrate.Text)
        .WriteSetting "encodingProfile", cmbProfile.ListIndex
        
        ' Audio
        .WriteSetting "volumeMod", CLng(txtVolumeMod.Text)
        .WriteSetting "channelCount", cmbChannelCount.ListIndex
        .WriteSetting "audioBitRate", cmbAudioBitrate.ListIndex
         
        ' Burning
        .WriteSetting "enableBurning", chkBurnToDisc.Value
        .WriteSetting "createISO", chkKeepISO.Value
        .WriteSetting "discLabel", txtDiscLabel.Text
        .WriteSetting "burnerName", cmbDrives.List(cmbDrives.ListIndex)
        .WriteSetting "deleteISO", chkKeepISO.Value
        .WriteSetting "eraseRW", chkEraseRewritable.Value
        .WriteSetting "burnSpeed", cmbBurnSpeed.List(cmbBurnSpeed.ListIndex)
        .WriteSetting "verifyDisc", chkVerifyDisc.Value
        .WriteSetting "ejectTray", chkEjectTray.Value
        
    End With
    
    ' Video advanced
    With frmAdvOptsVideo
        Config.WriteSetting "PSNR", .chkPSNR.Value
        Config.WriteSetting "halfRes", .chkHalfRes.Value
        Config.WriteSetting "overscanBorders", .chkOverscanBorders.Value
        Config.WriteSetting "overscanSize", CSng(.txtOverscanSize.Text)
        Config.WriteSetting "Deinterlace", .chkDeinterlace.Value
        Config.WriteSetting "MPEG2Copy", .chkMPEG2Copy
        Config.WriteSetting "dcPrecision", CLng(.txtDC.Text)
        Config.WriteSetting "allowPulldown", .chkPulldown.Value
    End With
    
End Sub


Private Sub flMenu_Click()

    If flMenu.selectedItem = -1 Then Exit Sub

    framGeneral.Visible = False
    framVideo.Visible = False
    framAudio.Visible = False
    framPlayback.Visible = False
    framBurning.Visible = False
    
    Select Case flMenu.selectedText
    Case "General"
        framGeneral.Visible = True
    Case "Video"
        framVideo.Visible = True
    Case "Audio"
        framAudio.Visible = True
    Case "Playback"
        framPlayback.Visible = True
    Case "Burning"
        framBurning.Visible = True
    End Select

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."
    
    ' Setup menu list
    flMenu.Thumbnails = False
    flMenu.Padding = 8
    flMenu.itemHeight = 32
    
    flMenu.addItem "General"
    flMenu.addItem "Video"
    flMenu.addItem "Audio"
    flMenu.addItem "Playback"
    flMenu.addItem "Burning"
    
    
    ' If recorders present enable UI for them
    If Burners.deviceCount > 0 Then
        fillDriveList
        chkBurnToDisc.Enabled = True
    Else
        chkBurnToDisc.Enabled = False
    End If

End Sub


Private Sub fillDriveList()

    Dim A As Long
    Dim Burner As clsBurnerDevice
    

    ' Fill combobox
    cmbDrives.Clear
    For A = 0 To Burners.deviceCount - 1
        Set Burner = Burners.getDevice(A)
        If burnerCanWrite(Burner) Then
            cmbDrives.addItem Burner.deviceName & " (" & Burner.deviceDriveChar & ")"
            cmbDrives.ItemData(cmbDrives.ListCount - 1) = A
        End If
    Next A
    If cmbDrives.ListCount > 0 Then cmbDrives.ListIndex = 0

End Sub


Private Sub imgIEncodePriority_Click()

    frmDialog.Display "Sets the default process priority of the encoding processes created by DVD Flick. Setting this to anything above Normal will give other processes running less time to execute but more to DVD Flick. The opposite is true for priorities below Normal." & vbNewLine & vbNewLine & "Anything above Normal is not recommended if you plan to use your computer while creating your DVD.", OkOnly Or Information

End Sub

Private Sub imgIEncodingProfile_Click()

    frmDialog.Display "Specifies how the video should be encoded; fast but inaccurate or slow and accurate. The 'Best' profile produces only marginally better quality than Normal but is much slower.", OkOnly Or Information

End Sub

Private Sub imgITargetBitRate_Click()

    frmDialog.Display "The target bitrate of the DVD determines how much material you can store on it in a certain quality. Lower bitrates offer more space at lower quality, while higher bitrates offer less space at better quality." & vbNewLine & vbNewLine & "Auto-fit is recommended because it will attempt to fit all the titles in the project on the disc using the highest bitrate possible, while not going below 2 Mbit/s.", OkOnly Or Information

End Sub

Private Sub imgITargetFormat_Click()

    frmDialog.Display "For most European, Asian, African and Oceanic countries this should be set to PAL, for most American countries to NTSC. Some DVD Players support both formats, in which case you should select the Mixed format. When using the Mixed format each separate title will be determined to be either PAL or NTSC, and encoded as such for the best quality.", OkOnly Or Information

End Sub

Private Sub imgIThreadCount_Click()

    frmDialog.Display "The number of threads used to encode audio and video. Setting this to any amount higher than the available number of physical CPU cores in the system will result in a decrease in performance." & vbNewLine & vbNewLine & "This value is autodetected when DVD Flick is first started.", OkOnly Or Information

End Sub


Private Sub txtThreadCount_Change()

    If IsNumeric(txtThreadCount.Text) Then
        If CLng(txtThreadCount.Text) < udThreadCount.Min Then txtThreadCount.Text = udThreadCount.Min
        If CLng(txtThreadCount.Text) > udThreadCount.Max Then txtThreadCount.Text = udThreadCount.Max
    Else
        txtThreadCount.Text = 1
    End If

End Sub


Private Sub txtCustomBitrate_Change()

    If Not IsNumeric(txtCustomBitrate.Text) Then
        txtCustomBitrate.Text = 5500
    End If

End Sub


Private Sub txtCustomSize_Change()

    If Not IsNumeric(txtCustomSize.Text) Then
        txtCustomSize.Text = 4489
    End If

End Sub


Private Sub updateBurningState()
    
    ' Disable burning chekbox if no drives presenet
    If cmbDrives.ListCount Then
        chkBurnToDisc.Enabled = True
    Else
        chkBurnToDisc.Enabled = False
    End If

    ' Burn OR make ISO, enable disc label
    If chkCreateISO.Value = 1 Or chkBurnToDisc.Value = 1 Then
        txtDiscLabel.Enabled = True
    Else
        txtDiscLabel.Enabled = False
    End If
    
    ' Burn AND make ISO, enable keep ISO option
    If chkCreateISO.Value = 1 And chkBurnToDisc.Value = 1 Then
        chkKeepISO.Enabled = True
    Else
        chkKeepISO.Enabled = False
    End If

    ' Burn to disc
    If chkBurnToDisc.Value = 1 Then
        cmbDrives.Enabled = True
        chkEraseRewritable.Enabled = True
        cmbBurnSpeed.Enabled = True
        chkVerifyDisc.Enabled = True
        chkEjectTray.Enabled = True
    Else
        cmbDrives.Enabled = False
        chkEraseRewritable.Enabled = False
        cmbBurnSpeed.Enabled = False
        chkVerifyDisc.Enabled = False
        chkEjectTray.Enabled = False
    End If

End Sub


Public Sub manualActivate()

    Dim A As Long
    
    
    disableUpdates = True
    
    ' General
    txtTitle.Text = Project.Title
    cmbTargetSize.ListIndex = Project.targetSize
    txtCustomSize.Text = Project.customSize
    cmbEncodePriority.ListIndex = Project.encodePriority
    chkKeepFiles.Value = Project.keepFiles
    txtThreadCount = Project.threadCount
    
    ' Playback
    cmbWhenPlayed.ListIndex = Project.whenPlayed
    chkLoopPlayback.Value = Project.loopPlayback
    chkEnableFirstSub.Value = Project.enableFirstSub
    
    ' Video
    cmbFormat.ListIndex = Project.targetFormat
    cmbProfile.ListIndex = Project.encodeProfile
    cmbVideoBitrate.ListIndex = Project.targetBitrate
    txtCustomBitrate.Text = Project.customBitrate
    
    ' Audio
    txtVolumeMod.Text = Project.volumeMod
    cmbChannelCount.ListIndex = Project.channelCount
    cmbAudioBitrate.ListIndex = Project.audioBitRate

    ' Video advanced
    With frmAdvOptsVideo
        .chkHalfRes.Value = Project.halfRes
        .chkOverscanBorders.Value = Project.overscanBorders
        .txtOverscanSize.Text = Project.overscanSize
        .chkPSNR = Project.PSNR
        .chkDeinterlace.Value = Project.Deinterlace
        .chkMPEG2Copy.Value = Project.MPEG2Copy
        .txtDC.Text = Project.dcPrecision
        .chkPulldown.Value = Project.Pulldown
    End With
    
    ' Burning
    chkCreateISO.Value = Project.createISO
    txtDiscLabel = Project.discLabel
    chkBurnToDisc.Value = Project.enableBurning
    chkEraseRewritable.Value = Project.eraseRW
    chkKeepISO.Value = Project.deleteISO
    chkVerifyDisc.Value = Project.verifyDisc
    chkEjectTray.Value = Project.ejectTray
    
    ' Burning speed
    For A = 0 To cmbBurnSpeed.ListCount - 1
        If cmbBurnSpeed.List(A) = Project.burnSpeed Then
            cmbBurnSpeed.ListIndex = A
            Exit For
        End If
    Next A

    ' Burning device
    If LenB(Project.burnerName) <> 0 Then
        For A = 0 To cmbDrives.ListCount - 1
            If cmbDrives.List(A) = Project.burnerName Then
                cmbDrives.ListIndex = A
                Exit For
            End If
        Next A
    End If
    
    disableUpdates = False
    
    
    updateBurningState
    flMenu.selectedItem = 0
    flMenu_Click

End Sub


Private Sub txtVolumeMod_Change()

    If IsNumeric(txtVolumeMod.Text) Then
        If CLng(txtVolumeMod.Text) < 1 Then txtVolumeMod.Text = CLng(1)
        If CLng(txtVolumeMod.Text) > 500 Then txtVolumeMod.Text = CLng(500)
    Else
        txtVolumeMod.Text = CLng(1)
    End If

End Sub

