VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1560
      TabIndex        =   0
      Top             =   3180
      Width           =   5595
   End
End
Attribute VB_Name = "frmSplash"
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
'   File purpose: Splash screen
'
Option Explicit
Option Compare Binary
Option Base 0


Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000


Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Private Sub Form_Load()

    Dim formLong As Long
    
    
    appLog.Add "Loading " & Me.Name & "..."
    
    lblVersion.Caption = versionString
    
    ' Get current form style and add layered attribute
    formLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, formLong Or WS_EX_LAYERED

End Sub


Public Sub fadeIn()

    Dim Alpha As Long
    Dim Tick As Long
    
    
    For Alpha = 0 To 255 Step 20
        ' Set window alpha
        SetLayeredWindowAttributes Me.hWnd, 0, Alpha, LWA_ALPHA
        
        ' Wait a bit
        Tick = GetTickCount
        Do
            DoEvents
        Loop Until GetTickCount - Tick > 5
    Next Alpha
    
    ' Set alpha off entirely
    SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA

End Sub
