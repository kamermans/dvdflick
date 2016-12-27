VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8235
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
   Icon            =   "frmWelcome.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   549
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin dvdflick.ctlSeparator ctlSeparator1 
      Height          =   75
      Left            =   480
      Top             =   2160
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   132
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin dvdflick.ctlSeparator ctlSeparator2 
      Height          =   75
      Left            =   480
      Top             =   3360
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   132
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   480
      Picture         =   "frmWelcome.frx":000C
      Top             =   3600
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   480
      Picture         =   "frmWelcome.frx":0AEE
      Top             =   1200
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   480
      Picture         =   "frmWelcome.frx":15D0
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label lblVisitWebsite 
      Caption         =   "Visit the website"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1020
      MouseIcon       =   "frmWelcome.frx":20B2
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4140
      Width           =   2010
   End
   Begin VB.Label lblVisitForums 
      Caption         =   "Visit the forums"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1020
      MouseIcon       =   "frmWelcome.frx":2204
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2940
      Width           =   1995
   End
   Begin VB.Label lblReadGuide 
      Caption         =   "Read the guide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1020
      MouseIcon       =   "frmWelcome.frx":2356
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1740
      Width           =   1995
   End
   Begin VB.Label Label4 
      Caption         =   "For more information about DVD Flick and the latest versions, please visit the DVD Flick website."
      Height          =   435
      Left            =   1020
      TabIndex        =   7
      Top             =   3660
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "If you have questions regarding DVD Flick, you can visit the forums and post what is on your mind."
      Height          =   435
      Left            =   1020
      TabIndex        =   6
      Top             =   2460
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "If you want to know how to get started using DVD Flick, be sure to read the guide, which covers all the basics you need to know."
      Height          =   435
      Left            =   1020
      TabIndex        =   5
      Top             =   1260
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to DVD Flick!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   660
      Width           =   5175
   End
   Begin VB.Image imgTop 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmWelcome"
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
'   File purpose: Welcome dialog with some helpful links.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Sub cmdClose_Click()

    Me.Hide

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

End Sub


Private Sub lblReadGuide_Click()

    Execute "file://" & APP_PATH & "guide/index_en.html", "", WS_Normal, False

End Sub

Private Sub lblVisitForums_Click()

    Execute "http://www.dvdflick.net/forums", "", WS_Normal, False

End Sub

Private Sub lblVisitWebsite_Click()

    Execute "http://www.dvdflick.net", "", WS_Normal, False

End Sub
