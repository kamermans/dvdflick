VERSION 5.00
Begin VB.Form frmMenuSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Settings"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11415
   ControlBox      =   0   'False
   Icon            =   "frmMenuSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin dvdflick.ctlFancyList flTemplates 
      Height          =   3435
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   6059
   End
   Begin VB.CheckBox chkShowSubFirst 
      Caption         =   "Show subtitle menu first"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   3780
      Width           =   2355
   End
   Begin VB.CheckBox chkShowAudioFirst 
      Caption         =   "Show audio menu first"
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
      Left            =   5520
      TabIndex        =   4
      Top             =   4140
      Width           =   2355
   End
   Begin VB.CheckBox chkMenuAutoPlay 
      Caption         =   "Auto-play menu"
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
      Left            =   3420
      TabIndex        =   2
      Top             =   3780
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   6600
      TabIndex        =   5
      Top             =   4740
      Width           =   2235
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
      Left            =   9000
      TabIndex        =   6
      Top             =   4740
      Width           =   2235
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
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
      Top             =   3720
      Width           =   2955
   End
   Begin VB.PictureBox picExample 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   3435
      Left            =   3300
      ScaleHeight     =   225
      ScaleMode       =   0  'User
      ScaleWidth      =   308.219
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   4560
   End
   Begin VB.Label lblCopyrights 
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
      Left            =   8100
      TabIndex        =   13
      Top             =   2580
      Width           =   3135
   End
   Begin VB.Label lblDescription 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8100
      TabIndex        =   12
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblAuthor 
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
      Left            =   8100
      TabIndex        =   11
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Copyrights"
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
      Left            =   8100
      TabIndex        =   10
      Top             =   2280
      Width           =   2595
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
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
      Left            =   8100
      TabIndex        =   9
      Top             =   1140
      Width           =   2595
   End
   Begin VB.Label Label2 
      Caption         =   "Author"
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
      Left            =   8100
      TabIndex        =   8
      Top             =   300
      Width           =   2595
   End
End
Attribute VB_Name = "frmMenuSettings"
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
'   File purpose: DVD menu template selection.
'
Option Explicit
Option Compare Binary
Option Base 0


Private cTemplate As clsMenuTemplate
Private templateName As String


Private Sub cmdAccept_Click()

    Project.menuTemplateName = flTemplates.selectedText
    Project.menuAutoPlay = chkMenuAutoPlay.Value
    Project.menuShowSubtitleFirst = chkShowSubFirst.Value
    Project.menuShowAudioFirst = chkShowAudioFirst.Value
    
    Me.Hide

End Sub


Private Sub cmdCancel_Click()

    Me.Hide

End Sub


Private Sub cmdPreview_Click()
    
    If Project.Titles.Count = 0 Then
        frmDialog.Display "There are no titles present in your project to generate a preview with.", Information Or OkOnly
        Exit Sub
    End If

    ' Nasty form juggling here, beware
    Me.Hide
    DoEvents
    frmStatus.setStatus "Rendering menus..."
    frmMenuPreview.Setup APP_PATH & "templates\" & templateName & "\template.cfg"
    frmStatus.Hide
    frmMenuPreview.Show 1
    Me.Show 1

End Sub


Public Sub Setup()

    Dim A As Long


    chkMenuAutoPlay.Value = Project.menuAutoPlay
    chkShowSubFirst.Value = Project.menuShowSubtitleFirst
    chkShowAudioFirst.Value = Project.menuShowAudioFirst
    
    ' Select template
    If flTemplates.Count > 0 Then flTemplates.selectedItem = 0
    For A = 0 To flTemplates.Count - 1
        If Project.menuTemplateName = flTemplates.itemText(A) Then
            flTemplates.selectedItem = A
            Exit For
        End If
    Next A

End Sub


Private Sub Form_Load()

    Dim A As Long
    Dim myFolder As Folder
    Dim cFolder As Folder
    
    
    appLog.Add "Loading " & Me.Name & "..."
    
    ' Fill templates list
    flTemplates.itemHeight = 24
    flTemplates.Padding = 4
    flTemplates.Thumbnails = False
    
    Set myFolder = FS.GetFolder(APP_PATH & "templates")
    flTemplates.addItem STR_DISABLED_MENU
    For Each cFolder In myFolder.SubFolders
        If FS.FileExists(APP_PATH & "templates\" & cFolder.Name & "\template.cfg") Then flTemplates.addItem cFolder.Name
    Next cFolder
    
    flTemplates.selectedItem = 0
    flTemplates_Click
    
End Sub


Private Sub flTemplates_Click()

    ' Disallow no selection
    If flTemplates.selectedItem = -1 Then
        flTemplates.selectedItem = 0
        Exit Sub
    End If
    
    ' None selected
    If flTemplates.selectedItem = 0 Then
        Set picExample.Picture = Nothing
        lblAuthor.Caption = "-"
        lblDescription.Caption = "-"
        lblCopyrights.Caption = "-"
        cmdPreview.Enabled = False
        Exit Sub
    
    ' Some selected
    Else
        cmdPreview.Enabled = True
        
    End If

    ' Load template and display
    templateName = flTemplates.selectedText
    
    Set cTemplate = New clsMenuTemplate
    cTemplate.openFrom APP_PATH & "templates\" & templateName & "\template.cfg"
    
    lblAuthor.Caption = cTemplate.Author
    lblDescription.Caption = cTemplate.Description
    lblCopyrights.Caption = cTemplate.Copyrights
    
    Set picExample.Picture = LoadPicture(APP_PATH & "templates\" & templateName & "\example.bmp")

End Sub

