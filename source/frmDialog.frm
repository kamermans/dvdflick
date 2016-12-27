VERSION 5.00
Begin VB.Form frmDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialog"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8055
   ControlBox      =   0   'False
   Icon            =   "frmDialog.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNotAgain 
      Caption         =   "Do not display this again"
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
      Left            =   1020
      TabIndex        =   3
      Top             =   780
      Width           =   3615
   End
   Begin VB.CommandButton cmdButton1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   180
      TabIndex        =   1
      Top             =   1260
      Width           =   2055
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Ok"
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
      Left            =   5820
      TabIndex        =   0
      Top             =   1260
      Width           =   2055
   End
   Begin VB.Image imgTopRed 
      Height          =   315
      Left            =   0
      Picture         =   "frmDialog.frx":000C
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image imgTop 
      Height          =   315
      Left            =   0
      Picture         =   "frmDialog.frx":08AE
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image imgIconQuestion 
      Height          =   480
      Left            =   240
      Picture         =   "frmDialog.frx":1A9E
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIconExclamation 
      Height          =   480
      Left            =   240
      Picture         =   "frmDialog.frx":2368
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIconCritical 
      Height          =   480
      Left            =   240
      Picture         =   "frmDialog.frx":2C32
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIconInfo 
      Height          =   480
      Left            =   240
      Picture         =   "frmDialog.frx":34FC
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
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
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   6900
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDialog"
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
'   File purpose: Custom messagebox dialog.
'
Option Explicit
Option Compare Binary
Option Base 0


Public Enum dialogConstants
    OkOnly = 1
    okCancel = 2
    YesNo = 4
    retryCancel = 8
    Critical = 16
    Exclamation = 32
    Information = 64
    Question = 128
End Enum

Public Enum dialogResultConstants
    buttonOk = 1
    buttonCancel = 2
    buttonYes = 4
    buttonNo = 8
    buttonRetry = 16
    checkNotAgain = 32
End Enum


Private buttonResult As Long


Public Function Display(ByVal Message As String, ByVal Icon As dialogConstants, Optional ByVal notAgain As Boolean = False) As dialogResultConstants

    imgIconQuestion.Visible = False
    imgIconExclamation.Visible = False
    imgIconCritical.Visible = False
    imgIconInfo.Visible = False
    
    cmdButton1.Visible = False
    cmdButton2.Visible = False
    chkNotAgain.Visible = notAgain
    chkNotAgain.Value = 0

    If Icon And Critical Then imgIconCritical.Visible = True
    If Icon And Exclamation Then imgIconExclamation.Visible = True
    If Icon And Information Then imgIconInfo.Visible = True
    If Icon And Question Then imgIconQuestion.Visible = True
    
    If (Icon And Critical) Or (Icon And Exclamation) Then
        imgTopRed.Visible = True
        imgTop.Visible = False
    Else
        imgTopRed.Visible = False
        imgTop.Visible = True
    End If
    
    ' Only Ok button
    If Icon And OkOnly Then
        cmdButton2.Caption = "Ok"
        cmdButton2.Visible = True
    
    ' Ok and cancel button
    ElseIf Icon And okCancel Then
        cmdButton1.Caption = "Cancel"
        cmdButton2.Caption = "Ok"
        cmdButton1.Visible = True
        cmdButton2.Visible = True
        
    ' Yes and No
    ElseIf Icon And YesNo Then
        cmdButton1.Caption = "No"
        cmdButton2.Caption = "Yes"
        cmdButton1.Visible = True
        cmdButton2.Visible = True
        
    ' Retry and Cancel
    ElseIf Icon And retryCancel Then
        cmdButton1.Caption = "Cancel"
        cmdButton2.Caption = "Retry"
        cmdButton1.Visible = True
        cmdButton2.Visible = True
        
    End If
    
    ' Reorder some controls
    lblMessage.Caption = Message
    cmdButton1.Top = lblMessage.Top + lblMessage.Height + 32
    If notAgain Then cmdButton1.Top = cmdButton1.Top + chkNotAgain.Height
    cmdButton2.Top = cmdButton1.Top
    Me.Height = (cmdButton1.Top + cmdButton1.Height + 38) * 15
    If notAgain Then chkNotAgain.Top = cmdButton1.Top - chkNotAgain.Height - 16
    
    ' Show
    Me.Show 1
    
    ' Return button pressed result
    If (Icon And OkOnly) Or (Icon And okCancel) Then
        If buttonResult = 1 Then Display = buttonCancel
        If buttonResult = 2 Then Display = buttonOk
        
    ElseIf Icon And retryCancel Then
        If buttonResult = 1 Then Display = buttonCancel
        If buttonResult = 2 Then Display = buttonRetry
        
    ElseIf Icon And YesNo Then
        If buttonResult = 1 Then Display = buttonNo
        If buttonResult = 2 Then Display = buttonYes
        
    End If
    
    If notAgain And chkNotAgain.Value = 1 Then Display = Display Or checkNotAgain

End Function


Private Sub cmdButton1_Click()

    buttonResult = 1
    Me.Hide

End Sub

Private Sub cmdButton2_Click()

    buttonResult = 2
    Me.Hide

End Sub


Private Sub Form_Load()

    Me.Caption = App.Title

End Sub
