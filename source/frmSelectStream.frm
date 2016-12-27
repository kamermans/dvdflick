VERSION 5.00
Begin VB.Form frmSelectTrack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select audio track(s)"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "frmSelectStream.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTracks 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      IntegralHeight  =   0   'False
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   780
      Width           =   6735
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
      Left            =   4860
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblDesc 
      Caption         =   "Please select one or more audio tracks you wish to add."
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
      Left            =   180
      TabIndex        =   3
      Top             =   420
      Width           =   6735
   End
   Begin VB.Label lblSource 
      Caption         =   "Filename"
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
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   6735
   End
End
Attribute VB_Name = "frmSelectTrack"
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
'   File purpose: Select an audio track from a list
'
Option Explicit
Option Compare Binary
Option Base 0


Public Tracks As Dictionary


Public Function Setup(ByRef Source As clsSource) As Boolean

    Dim A As Long
    Dim Info As Dictionary
    
    
    Setup = True
    
    lstTracks.Clear
    Set Tracks = New Dictionary
    
    For A = 0 To Source.streamCount - 1
        Set Info = Source.streamInfo(A)
        
        ' Only add the stream to the list if it's audio
        If Info("Type") = ST_Audio Then
            lstTracks.addItem "Track " & A + 1 & ", " & Info("Compression") & ", " & Info("bitRate") & " Kbit\s, " & Info("Channels") & " channels, " & visualTime(Info("Duration"))
            lstTracks.ItemData(lstTracks.ListCount - 1) = A
        End If
    Next A
    
    lblSource.Caption = FS.GetFileName(Source.fileName)
    
    ' No streams?
    If frmSelectTrack.lstTracks.ListCount = 0 Then
        frmDialog.Display "There are no audio tracks in the file or the audio tracks found are not supported.", Exclamation Or OkOnly
        Setup = False
    End If
    
End Function


Private Sub cmdAccept_Click()

    Dim A As Long
    
    
    Set Tracks = New Dictionary
    For A = 0 To lstTracks.ListCount - 1
        If lstTracks.Selected(A) = True Then Tracks.Add lstTracks.ItemData(A), lstTracks.List(A)
    Next A
    
    Me.Hide

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

End Sub
