VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmEncodeError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DVD Flick Error"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ControlBox      =   0   'False
   Icon            =   "frmEncodeError.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to clipboard"
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
      TabIndex        =   2
      Top             =   5520
      Width           =   2055
   End
   Begin VB.ComboBox cmbLogs 
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
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   2715
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
      Left            =   6180
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtfLog 
      Height          =   3495
      Left            =   180
      TabIndex        =   1
      Top             =   1860
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6165
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   32000
      TextRTF         =   $"frmEncodeError.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Below you can find a log file which provides details as to how the error might have ocurred."
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
      Left            =   900
      TabIndex        =   6
      Top             =   1020
      Width           =   7335
   End
   Begin VB.Label lblError 
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   900
      TabIndex        =   5
      Top             =   420
      Width           =   7335
   End
   Begin VB.Label lblMessage 
      Caption         =   "An error occured during the encoding process"
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
      Left            =   900
      TabIndex        =   4
      Top             =   180
      Width           =   6855
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmEncodeError.frx":008C
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmEncodeError"
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
'   File purpose: Encoding error display + log
'
Option Explicit
Option Base 0
Option Compare Binary


Public Sub Setup(ByVal Error As String)

    Dim A As Long
    Dim myFolder As Folder
    Dim myFile As File
    

    lblError.Caption = Error
    
    ' Fill log combo
    cmbLogs.Clear
    Set myFolder = FS.GetFolder(Project.destinationDir)
    For Each myFile In myFolder.Files
        If LCase$(FS.GetExtensionName(myFile.Path)) = "txt" Then cmbLogs.addItem LCase$(myFile.Name)
    Next myFile
    
    ' Select errorlog
    cmbLogs.ListIndex = 0
    For A = 0 To cmbLogs.ListCount - 1
        If cmbLogs.List(A) = "errorlog.txt" Then
            cmbLogs.ListIndex = A
            Exit For
        End If
    Next A

End Sub


Private Sub cmbLogs_Change()

    cmbLogs_Click
    
End Sub


Private Sub cmbLogs_Click()
    
    On Error Resume Next
    
    rtfLog.Text = ""
    rtfLog.LoadFile Project.destinationDir & "\" & cmbLogs.List(cmbLogs.ListIndex), rtfText
    rtfLog.SelStart = 0
    
    On Error GoTo 0
    
End Sub


Private Sub cmdCopy_Click()

    Clipboard.SetText rtfLog.Text

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."
    
End Sub


Private Sub cmdClose_Click()

    Me.Hide

End Sub
