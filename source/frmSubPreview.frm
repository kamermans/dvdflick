VERSION 5.00
Begin VB.Form frmSubPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subtitle preview (click to close)"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11520
   ControlBox      =   0   'False
   Icon            =   "frmSubPreview.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   768
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   8640
      Left            =   0
      MouseIcon       =   "frmSubPreview.frx":000C
      MousePointer    =   99  'Custom
      ScaleHeight     =   8640
      ScaleWidth      =   11520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11520
   End
End
Attribute VB_Name = "frmSubPreview"
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
'   File purpose: Subtitle large scale display.
'
Option Explicit
Option Compare Binary
Option Base 0


Public targetHeight As Long


Public Sub manualActivate()

    picDisplay.Width = targetHeight * (4 / 3)
    picDisplay.Height = targetHeight
    
    Me.Width = picDisplay.Width * 15
    Me.Height = (picDisplay.Height + modUtil.getTitleBarHeight) * 15
    
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

End Sub


Private Sub picDisplay_Click()

    Me.Hide

End Sub
