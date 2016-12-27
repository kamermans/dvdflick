VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmAdvOptsVideo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced video options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
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
   Icon            =   "frmAdvOptsVideo.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPulldown 
      Caption         =   "Apply 2:3 pulldown"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin ComCtl2.UpDown udOverScanSize 
      Height          =   315
      Left            =   4020
      TabIndex        =   7
      Top             =   660
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      Value           =   3
      BuddyControl    =   "txtOverScanSize"
      BuddyDispid     =   196612
      OrigLeft        =   332
      OrigTop         =   48
      OrigRight       =   349
      OrigBottom      =   73
      Max             =   30
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
   End
   Begin VB.TextBox txtDC 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.CheckBox chkMPEG2Copy 
      Caption         =   "Copy MPEG-2 streams"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CheckBox chkDeinterlace 
      Caption         =   "Deinterlace source"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtOverscanSize 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   660
      Width           =   495
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   3540
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox chkPSNR 
      Caption         =   "Log PSNR values"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1515
   End
   Begin VB.CheckBox chkOverscanBorders 
      Caption         =   "Add overscan borders"
      Height          =   315
      Left            =   3180
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.CheckBox chkHalfRes 
      Caption         =   "Half horizontal resolution"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2475
   End
   Begin ComCtl2.UpDown udDC 
      Height          =   315
      Left            =   3720
      TabIndex        =   9
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      Value           =   8
      BuddyControl    =   "txtDC"
      BuddyDispid     =   196612
      OrigLeft        =   332
      OrigTop         =   48
      OrigRight       =   349
      OrigBottom      =   73
      Max             =   11
      Min             =   8
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Image imgIDC 
      Height          =   360
      Left            =   4380
      MouseIcon       =   "frmAdvOptsVideo.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "frmAdvOptsVideo.frx":015E
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "DC precision"
      Height          =   195
      Left            =   3180
      TabIndex        =   12
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Image imgIMPEG2Copy 
      Height          =   360
      Left            =   2220
      MouseIcon       =   "frmAdvOptsVideo.frx":0C40
      MousePointer    =   99  'Custom
      Picture         =   "frmAdvOptsVideo.frx":0D92
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label lblPerc 
      Caption         =   "%"
      Height          =   195
      Left            =   4380
      TabIndex        =   11
      Top             =   720
      Width           =   435
   End
   Begin VB.Image imgIPSNR 
      Height          =   360
      Left            =   1800
      MouseIcon       =   "frmAdvOptsVideo.frx":1874
      MousePointer    =   99  'Custom
      Picture         =   "frmAdvOptsVideo.frx":19C6
      Top             =   240
      Width           =   360
   End
   Begin VB.Image imgIOverscanBorders 
      Height          =   360
      Left            =   5160
      MouseIcon       =   "frmAdvOptsVideo.frx":24A8
      MousePointer    =   99  'Custom
      Picture         =   "frmAdvOptsVideo.frx":25FA
      Top             =   240
      Width           =   360
   End
   Begin VB.Image imgIHalfRes 
      Height          =   360
      Left            =   2700
      MouseIcon       =   "frmAdvOptsVideo.frx":30DC
      MousePointer    =   99  'Custom
      Picture         =   "frmAdvOptsVideo.frx":322E
      Top             =   720
      Width           =   360
   End
End
Attribute VB_Name = "frmAdvOptsVideo"
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
'   File purpose: Interface for settings advanced video encoding options
'
Option Explicit
Option Compare Binary
Option Base 0


Private Sub chkOverscanBorders_Click()

    If chkOverscanBorders.Value = 1 Then
        txtOverscanSize.Enabled = True
    Else
        txtOverscanSize.Enabled = False
    End If

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."
    
End Sub


Private Sub cmdAccept_Click()

    Me.Hide

End Sub


Private Sub txtOverscanSize_Change()

    If IsNumeric(txtOverscanSize.Text) Then
        If CSng(txtOverscanSize.Text) < 0.1 Then txtOverscanSize.Text = CSng(0.1)
        If CSng(txtOverscanSize.Text) > 30 Then txtOverscanSize.Text = CSng(30)
    Else
        txtOverscanSize.Text = CSng(0.1)
    End If

End Sub


Private Sub imgIDC_Click()

    frmDialog.Display "Discrete Cosine precision. The higher this value, the better the quality, but at the cost of bitrate.", OkOnly Or Information

End Sub

Private Sub imgIMPEG2Copy_Click()

    frmDialog.Display "This option will copy MPEG-2 video streams instead of re-encoding them." & vbNewLine & vbNewLine & "ONLY use this option if you are CERTAIN that your MPEG-2 video streams are MPEG-2 DVD compliant, as well as intended for the right target format (NTSC or PAL).", OkOnly Or Information

End Sub

Private Sub imgIHalfRes_Click()

    frmDialog.Display "This will use half the horizontal resolution normally used for DVDs (352 pixels in width instead of 720). This will significantly reduce the quality and allow you to lower the bitrate. Subtitles will not work properly with this setting!", OkOnly Or Information

End Sub

Private Sub imgIOverscanBorders_Click()

    frmDialog.Display "Many older TV sets have a so-called overscan area which puts the borders of the image outside of view. Enabling this option will add a border to the encoded video so that approximately the entire image is displayed on screen.", OkOnly Or Information

End Sub

Private Sub imgIPSNR_Click()

    frmDialog.Display "Calculates the PSNR (Peak Signal to Noise Ratio). This indicates the objective measurement of quality of reconstruction of the video image. The final average is written to dvdflick.log in the project's destination folder." & vbNewLine & vbNewLine & "PSNR values range from 0 to 100 dB, where 100 is perfect reproduction (impossible to achieve with video compression). Typical values range from 40 to 50.", OkOnly Or Information

End Sub
