VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmSubtitle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subtitle"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   ControlBox      =   0   'False
   Icon            =   "frmSubtitle.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   724
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdOutlineCol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   10260
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   28
      Top             =   4620
      Width           =   315
   End
   Begin VB.PictureBox cmdBackCol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6900
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   27
      Top             =   5040
      Width           =   315
   End
   Begin VB.PictureBox cmdTextCol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6900
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   26
      Top             =   4620
      Width           =   315
   End
   Begin VB.PictureBox picAlignment 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   300
      ScaleHeight     =   1395
      ScaleWidth      =   1695
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5100
      Width           =   1695
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   60
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   60
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   15
         Top             =   60
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   20
         Top             =   1020
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   19
         Top             =   1020
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   21
         Top             =   540
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   17
         Top             =   540
         Width           =   255
      End
      Begin VB.OptionButton optAlign 
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   18
         Top             =   1020
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdAutoFit 
      Caption         =   "Auto-fit"
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
      Left            =   2580
      TabIndex        =   13
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CheckBox chkBold 
      Alignment       =   1  'Right Justify
      Caption         =   "Bold"
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
      Left            =   300
      TabIndex        =   14
      Top             =   4260
      Width           =   1695
   End
   Begin VB.ComboBox cmbCodepage 
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
      ItemData        =   "frmSubtitle.frx":000C
      Left            =   1680
      List            =   "frmSubtitle.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3420
      Width           =   2715
   End
   Begin VB.CheckBox chkTransBack 
      Caption         =   "Transparent background"
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
      Left            =   2280
      TabIndex        =   24
      Top             =   5580
      Width           =   2175
   End
   Begin VB.CheckBox chkAA 
      Caption         =   "Anti-alias"
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
      Left            =   2280
      TabIndex        =   23
      Top             =   5100
      Width           =   1995
   End
   Begin VB.TextBox txtOutline 
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
      Left            =   6900
      TabIndex        =   29
      Top             =   5580
      Width           =   375
   End
   Begin VB.CheckBox chkDisplay 
      Caption         =   "Display by default"
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
      Left            =   2280
      TabIndex        =   25
      Top             =   6060
      Width           =   1635
   End
   Begin VB.CommandButton cmdUseDefault 
      Caption         =   "Use as defaults"
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
      Left            =   6420
      TabIndex        =   31
      Top             =   6300
      Width           =   2055
   End
   Begin VB.TextBox txtFPS 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   660
      Width           =   915
   End
   Begin VB.ComboBox cmbLang 
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
      ItemData        =   "frmSubtitle.frx":0010
      Left            =   1680
      List            =   "frmSubtitle.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2715
   End
   Begin VB.ComboBox cmbFont 
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
      ItemData        =   "frmSubtitle.frx":0014
      Left            =   1680
      List            =   "frmSubtitle.frx":0016
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3000
      Width           =   2715
   End
   Begin ComCtl2.UpDown udTop 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      BuddyControl    =   "txtMarginTop"
      BuddyDispid     =   196628
      OrigLeft        =   176
      OrigTop         =   152
      OrigRight       =   192
      OrigBottom      =   173
      Max             =   100
      Min             =   -20
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtMarginRight 
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
      Left            =   1680
      TabIndex        =   8
      Top             =   2460
      Width           =   915
   End
   Begin VB.TextBox txtMarginLeft 
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
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   915
   End
   Begin VB.TextBox txtMarginBottom 
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1620
      Width           =   915
   End
   Begin VB.TextBox txtMarginTop 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   915
   End
   Begin VB.ComboBox cmbFontSize 
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
      ItemData        =   "frmSubtitle.frx":0018
      Left            =   1680
      List            =   "frmSubtitle.frx":004F
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
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
      Left            =   8640
      TabIndex        =   32
      Top             =   6300
      Width           =   2055
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4320
      Left            =   4920
      MouseIcon       =   "frmSubtitle.frx":0095
      MousePointer    =   99  'Custom
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   180
      Width           =   5760
   End
   Begin ComCtl2.UpDown udBottom 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1620
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      BuddyControl    =   "txtMarginBottom"
      BuddyDispid     =   196627
      OrigLeft        =   176
      OrigTop         =   180
      OrigRight       =   192
      OrigBottom      =   201
      Max             =   100
      Min             =   -20
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udLeft 
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   2040
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      BuddyControl    =   "txtMarginLeft"
      BuddyDispid     =   196626
      OrigLeft        =   176
      OrigTop         =   208
      OrigRight       =   192
      OrigBottom      =   229
      Max             =   100
      Min             =   -20
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udRight 
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   2460
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      BuddyControl    =   "txtMarginRight"
      BuddyDispid     =   196625
      OrigLeft        =   176
      OrigTop         =   236
      OrigRight       =   192
      OrigBottom      =   257
      Max             =   100
      Min             =   -20
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udOutline 
      Height          =   315
      Left            =   7320
      TabIndex        =   30
      Top             =   5580
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   327681
      Value           =   2
      BuddyControl    =   "txtOutline"
      BuddyDispid     =   196619
      OrigLeft        =   176
      OrigTop         =   152
      OrigRight       =   192
      OrigBottom      =   173
      Max             =   8
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Image imgIAutoFit 
      Height          =   360
      Left            =   4440
      MouseIcon       =   "frmSubtitle.frx":01E7
      MousePointer    =   99  'Custom
      Picture         =   "frmSubtitle.frx":0339
      Top             =   3840
      Width           =   360
   End
   Begin VB.Image imgICodePage 
      Height          =   360
      Left            =   4440
      MouseIcon       =   "frmSubtitle.frx":0E1B
      MousePointer    =   99  'Custom
      Picture         =   "frmSubtitle.frx":0F6D
      Top             =   3420
      Width           =   360
   End
   Begin VB.Label lblCodepage 
      Caption         =   "Character set"
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
      Left            =   300
      TabIndex        =   49
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblFPS 
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
      Height          =   195
      Left            =   2700
      TabIndex        =   48
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "Background color"
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
      Left            =   4980
      TabIndex        =   47
      Top             =   5100
      Width           =   1635
   End
   Begin VB.Label Label6 
      Caption         =   "Alignment"
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
      Left            =   300
      TabIndex        =   46
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Outline color"
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
      Left            =   8340
      TabIndex        =   45
      Top             =   4680
      Width           =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "Text color"
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
      Left            =   4980
      TabIndex        =   44
      Top             =   4680
      Width           =   1635
   End
   Begin VB.Label Label3 
      Caption         =   "pixels"
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
      Left            =   7620
      TabIndex        =   43
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Outline"
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
      Left            =   4980
      TabIndex        =   42
      Top             =   5640
      Width           =   855
   End
   Begin VB.Image imgIDisplay 
      Height          =   360
      Left            =   3960
      MouseIcon       =   "frmSubtitle.frx":1A4F
      MousePointer    =   99  'Custom
      Picture         =   "frmSubtitle.frx":1BA1
      Top             =   6060
      Width           =   360
   End
   Begin VB.Label lblFrameRate 
      Caption         =   "Framerate"
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
      Left            =   300
      TabIndex        =   41
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label lblLang 
      Caption         =   "Language"
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
      Left            =   300
      TabIndex        =   40
      Top             =   300
      Width           =   795
   End
   Begin VB.Label lblRightMargin 
      Caption         =   "Right margin"
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
      Left            =   300
      TabIndex        =   39
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label lblLeftMargin 
      Caption         =   "Left margin"
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
      Left            =   300
      TabIndex        =   38
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label lblBottomMargin 
      Caption         =   "Bottom margin"
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
      Left            =   300
      TabIndex        =   37
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label lblTopMargin 
      Caption         =   "Top margin"
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
      Left            =   300
      TabIndex        =   36
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label lblSize 
      Caption         =   "Size"
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
      Left            =   300
      TabIndex        =   35
      Top             =   3900
      Width           =   735
   End
   Begin VB.Label lblFont 
      Caption         =   "Font"
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
      Left            =   300
      TabIndex        =   34
      Top             =   3060
      Width           =   615
   End
End
Attribute VB_Name = "frmSubtitle"
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
'   File purpose: Subtitle edit dialog
'
Option Explicit
Option Compare Binary
Option Base 0


' Current subtitle
Private mySub As clsSubtitle
Private myFile As clsSubFile
Private myTitle As clsTitle
Private myVideo As clsVideo

' Misc.
Private noRender As Boolean
Private dispRight As Long
Private thumbPic As clsGDIImage


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private Sub renderThumbPic()

    Dim useThumb As clsGDIImage
    
    
    Set useThumb = myVideo.Thumbnail
    If useThumb Is Nothing Then Set useThumb = noThumb
    Set thumbPic = resizeToMatch(useThumb, picDisplay.ScaleWidth, picDisplay.ScaleHeight, myVideo.PAR)

End Sub


Private Sub renderDisplay()

    If noRender Then Exit Sub
    modRenderSub.renderPreview mySub, myFile, myTitle, 0.5, picDisplay, thumbPic

End Sub


Public Sub Setup(ByRef Subtitle As clsSubtitle, ByRef Title As clsTitle)

    Dim A As Long
    Dim Encode As Dictionary
    
    
    Set Encode = Title.encodeInfo
    Set mySub = Subtitle
    Set myTitle = Title
    Set myVideo = Title.Videos.Item(0)
    noRender = True
    
    ' Set preview display size to match target video size
    'picDisplay.Width = Encode("Height") * modUtil.getAspect(myTitle.targetAspect) / 2
    picDisplay.Width = Encode("Height") * (4 / 3) / 2
    picDisplay.Height = Encode("Height") / 2
    picDisplay.Left = dispRight - picDisplay.Width
    
    
    ' Open and parse subtitle file
    Set myFile = New clsSubFile
    myFile.openFrom mySub.fileName, mySub.codePage
    
    If myFile.codePage = -1 Then
        cmbCodepage.Visible = False
        lblCodepage.Visible = False
        imgICodePage.Visible = False
    Else
        cmbCodepage.Visible = True
        lblCodepage.Visible = True
        imgICodePage.Visible = True
    End If

    ' Set control info to match loaded subtitle
    With mySub
        cmbFontSize.Text = .fontSize
        chkBold.Value = .fontBold

        cmdTextCol.backColor = .colorText
        cmdOutlineCol.backColor = .colorOutline
        cmdBackCol.backColor = .ColorBack
        txtOutline.Text = .Outline
        chkAA.Value = .antiAlias
        chkTransBack.Value = .transBack

        txtMarginTop.Text = .marginTop
        txtMarginBottom.Text = .marginBottom
        txtMarginLeft.Text = .marginLeft
        txtMarginRight.Text = .marginRight
        optAlign(.Alignment).Value = True

        txtFPS.Text = .FPS
        cmbLang.ListIndex = .Language
        chkDisplay.Value = .displayDefault

        ' Select codepage
        For A = 0 To codePages.Count - 1
            If codePages.Keys(A) = .codePage Then
                cmbCodepage.ListIndex = A
                Exit For
            End If
        Next A
    End With

    ' Select font used
    For A = 1 To cmbFont.ListCount
        If mySub.Font = cmbFont.List(A) Then
            cmbFont.ListIndex = A
            Exit For
        End If
    Next A

    ' Framerate useful?
    lblFrameRate.Visible = CBool(mySub.frameBased)
    lblFPS.Visible = CBool(mySub.frameBased)
    txtFPS.Visible = CBool(mySub.frameBased)

    noRender = False
    renderThumbPic
    renderDisplay

End Sub


Private Sub chkAA_Click()

    mySub.antiAlias = chkAA.Value
    renderDisplay

End Sub


Private Sub chkDisplay_Click()

    mySub.displayDefault = chkDisplay.Value

End Sub


Private Sub chkTransBack_Click()

    mySub.transBack = chkTransBack.Value
    renderDisplay

End Sub


Private Sub cmbCodePage_Change()

    cmbCodePage_Click

End Sub

Private Sub cmbCodePage_Click()

    mySub.codePage = codePages.Keys(cmbCodepage.ListIndex)
    
    ' Open and parse subtitle file
    Set myFile = New clsSubFile
    myFile.openFrom mySub.fileName, mySub.codePage
    
    renderDisplay

End Sub


Private Sub cmbFont_Click()

    mySub.Font = cmbFont.List(cmbFont.ListIndex)
    renderDisplay

End Sub


Private Sub cmbFontSize_Change()

    If IsNumeric(cmbFontSize.Text) Then
        mySub.fontSize = CLng(cmbFontSize.Text)
        If mySub.fontSize < 2 Then cmbFontSize.Text = 2
        If mySub.fontSize > 512 Then cmbFontSize.Text = 512
    Else
        cmbFontSize.Text = "8"
    End If
    renderDisplay

End Sub


Private Sub cmbFontSize_Click()

    cmbFontSize_Change

End Sub


Private Sub cmbLang_Click()

    mySub.Language = cmbLang.ListIndex

End Sub


Private Sub chkBold_Click()

    mySub.fontBold = chkBold.Value
    renderDisplay

End Sub


Private Sub cmdAccept_Click()

    Dim A As Long
    
    
    ' Scan subtitle blocks to see if they all fit
    modRenderSub.initRender mySub
    For A = 0 To myFile.blockCount - 1
        If modRenderSub.getBlockWidth(myFile.getBlock(A)) > 720 Then
            frmDialog.Display "One or more subtitle lines are too wide and will be clipped. The font size will have to be reduced in order for all subtitle lines to fit into view.", Exclamation Or OkOnly
            Exit For
        End If
    Next A

    ' Clean up
    Set mySub = Nothing
    Set myFile = Nothing

    Me.Hide

End Sub


' Attempt to fit all subtitle lines in view
Private Sub cmdAutoFit_Click()

    Dim A As Long
    Dim newSize As Long
    Dim tooBig As Boolean
    
    
    newSize = 48
    Do
        newSize = newSize - 1
    
        mySub.fontSize = newSize
        modRenderSub.initRender mySub
        
        tooBig = False
        For A = 0 To myFile.blockCount - 1
            If modRenderSub.getBlockWidth(myFile.getBlock(A)) > 720 Then
                tooBig = True
                Exit For
            End If
        Next A

    Loop Until newSize <= 5 Or tooBig = False
    
    cmbFontSize.Text = newSize

End Sub


Private Sub cmdTextCol_Click()

    Dim Col As Long
    
    
    Col = fileDialog.getColor(Me.hWnd, CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN Or CC_RGBINIT, cmdTextCol.backColor)
    If Col <> -1 Then
        cmdTextCol.backColor = Col
        mySub.colorText = Col
    End If
    
    renderDisplay

End Sub


Private Sub cmdOutlineCol_Click()

    Dim Col As Long
    
    
    Col = fileDialog.getColor(Me.hWnd, CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN Or CC_RGBINIT, cmdOutlineCol.backColor)
    If Col <> -1 Then
        cmdOutlineCol.backColor = Col
        mySub.colorOutline = Col
    End If
    
    renderDisplay

End Sub


Private Sub cmdBackCol_Click()

    Dim Col As Long
    
    
    Col = fileDialog.getColor(Me.hWnd, CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN Or CC_RGBINIT, cmdBackCol.backColor)
    If Col <> -1 Then
        cmdBackCol.backColor = Col
        mySub.ColorBack = Col
    End If
    
    renderDisplay

End Sub


Private Sub cmdUseDefault_Click()

    Config.WriteSetting "subFont", cmbFont.List(cmbFont.ListIndex)
    Config.WriteSetting "subFontSize", CLng(cmbFontSize.Text)
    Config.WriteSetting "subFontBold", chkBold.Value
    Config.WriteSetting "subCodePage", cmbCodepage.ListIndex

    Config.WriteSetting "subColorText", cmdTextCol.backColor
    Config.WriteSetting "subColorOutline", cmdOutlineCol.backColor
    Config.WriteSetting "subColorBack", cmdBackCol.backColor
    Config.WriteSetting "subOutline", CLng(txtOutline.Text)
    Config.WriteSetting "subAA", chkAA.Value
    Config.WriteSetting "subTransBack", chkTransBack.Value

    Config.WriteSetting "subMarginTop", CLng(txtMarginTop.Text)
    Config.WriteSetting "subMarginBottom", CLng(txtMarginBottom.Text)
    Config.WriteSetting "subMarginLeft", CLng(txtMarginLeft.Text)
    Config.WriteSetting "subMarginRight", CLng(txtMarginRight.Text)
    Config.WriteSetting "subAlignment", mySub.Alignment
    
    Config.WriteSetting "subLanguage", cmbLang.ListIndex
    
End Sub


Private Sub Form_Load()

    Dim A As Long
    
    
    appLog.Add "Loading " & Me.Name & "..."
    
    dispRight = picDisplay.Left + picDisplay.Width
    
    ' Fill language combolist
    fillLangCombo cmbLang
    
    ' Fill font combolist
    cmbFont.Clear
    For A = 1 To Screen.FontCount - 1
        cmbFont.addItem Screen.Fonts(A)
    Next A
    
    ' Fill character set list
    cmbCodepage.Clear
    For A = 0 To codePages.Count - 1
        cmbCodepage.addItem codePages.Items(A)
    Next A
    
End Sub


Private Sub imgIAutoFit_Click()

    frmDialog.Display "This will adjust the font size so that all subtitle lines will fit on screen.", OkOnly Or Information

End Sub

Private Sub imgICodePage_Click()

    frmDialog.Display "Use this setting to indicate in which character set the subtitle file is stored in. For example, for Russian subtitles, choose the Cyrillic codepage.", OkOnly Or Information

End Sub

Private Sub imgIDisplay_Click()

    frmDialog.Display "This will enable this subtitle when starting to play the title. If other subtites are enabled as well, only the first will be displayed.", OkOnly Or Information

End Sub


Private Sub optAlign_Click(Index As Integer)

    mySub.Alignment = Index
    renderDisplay

End Sub


Private Sub picDisplay_Click()

    frmSubPreview.targetHeight = myTitle.encodeInfo("Height")
    frmSubPreview.manualActivate
    modRenderSub.renderPreview mySub, myFile, myTitle, 1, frmSubPreview.picDisplay, thumbPic
    frmSubPreview.Show 1

End Sub


Private Sub txtFPS_Change()

    If IsNumeric(txtFPS.Text) Then
        mySub.FPS = CSng(txtFPS.Text)
        If mySub.FPS < 1 Then txtFPS.Text = 1
        If mySub.FPS > 1000 Then txtFPS.Text = 1000
    Else
        txtFPS.Text = 1
    End If

End Sub


Private Sub txtMarginTop_Change()

    If IsNumeric(txtMarginTop.Text) Then
        mySub.marginTop = CLng(txtMarginTop.Text)
    Else
        txtMarginTop.Text = 0
    End If
    
    renderDisplay

End Sub

Private Sub txtMarginBottom_Change()

    If IsNumeric(txtMarginBottom.Text) Then
        mySub.marginBottom = CLng(txtMarginBottom.Text)
    Else
        txtMarginBottom.Text = 0
    End If
    
    renderDisplay

End Sub

Private Sub txtMarginLeft_Change()

    If IsNumeric(txtMarginLeft.Text) Then
        mySub.marginLeft = CLng(txtMarginLeft.Text)
    Else
        txtMarginLeft.Text = 0
    End If
    
    renderDisplay

End Sub

Private Sub txtMarginRight_Change()

    If IsNumeric(txtMarginRight.Text) Then
        mySub.marginRight = CLng(txtMarginRight.Text)
    Else
        txtMarginRight.Text = 0
    End If
    
    renderDisplay

End Sub


Private Sub txtOutline_Change()

    If IsNumeric(txtOutline.Text) Then
        mySub.Outline = CLng(txtOutline.Text)
        If mySub.Outline < 0 Then txtOutline.Text = 0
        If mySub.Outline > 8 Then txtOutline.Text = 8
    Else
        txtOutline.Text = 0
    End If
    
    renderDisplay

End Sub
