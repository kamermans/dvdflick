VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About DVD Flick"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin dvdflick.ctlSeparator ctlSeparator1 
      Height          =   210
      Left            =   -60
      Top             =   3840
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   476
   End
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   5100
      Width           =   4695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   5220
      TabIndex        =   0
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label lblDebugFlags 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   3360
      Width           =   5595
   End
   Begin VB.Label lblURL 
      Caption         =   "http://www.dvdflick.net/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "This site has monies!"
      Top             =   4680
      Width           =   2955
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   7155
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1620
      TabIndex        =   1
      Top             =   3180
      Width           =   5535
   End
End
Attribute VB_Name = "frmAbout"
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
'   File purpose: About dialog.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Playing As Boolean
Private Credits() As String

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub cmdClose_Click()

    Playing = False
    Me.Hide

End Sub


Private Sub Form_Activate()

    Dim A As Long
    Dim Offs As Long
    Dim lineHeight As Long
    Dim topHeight As Long
    Dim Ticker As Long

    
    lineHeight = picScroll.Font.Size + 10
    topHeight = -(UBound(Credits) * lineHeight + lineHeight)
    Offs = picScroll.Height - lineHeight
    
    Playing = True
    
    Do
        DoEvents
        
        Offs = Offs - 1
        If Offs <= topHeight Then Offs = picScroll.Height
        
        picScroll.Cls
        For A = 0 To UBound(Credits)
            TextOut picScroll.hDC, 0, Offs + (A * lineHeight), Credits(A), Len(Credits(A))
        Next A
        
        Ticker = GetTickCount
        Do
            DoEvents
        Loop Until GetTickCount - Ticker >= 55
    Loop Until Playing = False

End Sub


Private Sub addCredit(ByVal Data As String)

    ReDim Preserve Credits(UBound(Credits) + 1)
    Credits(UBound(Credits)) = Data

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

    ' Get picture from the splash form to save space
    Set Me.Picture = frmSplash.Picture
    
    ' Versioning captions
    lblVersion.Caption = versionString
    Me.Caption = "About " & App.Title
    
    ' Create credits
    ReDim Credits(0)
    addCredit "DVDAuthor by Scott Smith"
    addCredit "http://dvdauthor.sourceforge.net/"
    addCredit ""
    addCredit "FFmpeg by the FFmpeg team"
    addCredit "http://ffmpeg.mplayerhq.hu/"
    addCredit ""
    addCredit "MPLEX by the mjpegtools team"
    addCredit ""
    addCredit "Base icon art by Scrow's Icons"
    addCredit "http://www.virtualplastic.net/scrow/"
    addCredit ""
    addCredit "ImgBurn by LIGHTNING UK!"
    addCredit "http://www.imgburn.com/"
    addCredit ""
    addCredit "MKVExtract by Moritz Bunkus"
    addCredit "http://www.bunkus.org/videotools/mkvtoolnix/"
    addCredit ""
    addCredit "7-Zip by Igor Pavlov"
    addCredit "http://www.7-zip.org/"
    addCredit ""
    addCredit "DelayCut by jsoto"
    addCredit "http://jsoto.posunplugged.com/"
    addCredit ""
    addCredit "Pulldown by Brent Beyeler and Hard Code"
    addCredit "http://www.inwards.com/inwards/?id=36"
    addCredit ""
    addCredit "MPGTX by Laurent Alacoque"
    addCredit "http://mpgtx.sourceforge.net/"
    addCredit ""
    addCredit "RSP CPU detection by RSP Software"
    addCredit "http://rspsoftware.clic3.net/"
    addCredit ""
    addCredit "Tray Icon control by Robert Gelb"
    addCredit ""
    addCredit "BMP2PNG by Miyasaka Masaru"
    addCredit ""
    addCredit "Special thanks to"
    addCredit "Ken Bosward"
    addCredit "'Guzeppi'"
    addCredit "'WaltP'"
    addCredit "'Neal'"
    addCredit "Evandro A.Gouveia"
    addCredit "'XhmikosR'"
    addCredit "Melissa Kip"
    addCredit "guystooges '"
    addCredit "'Ruler'"
    addCredit "'ricardo0y'"
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""
    addCredit ""

End Sub


Private Sub lblCopyright_Click()

    Playing = False
    Me.Hide
    frmTetris.Show 1

End Sub


Private Sub lblURL_Click()

    Execute lblURL.Caption, "", EP_Normal, False

End Sub
