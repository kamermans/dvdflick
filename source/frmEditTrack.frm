VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEditTrack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio track sources"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIgnoreDelay 
      Caption         =   "Ignore audio delay for this track"
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   4200
      Width           =   2955
   End
   Begin dvdflick.ctlFancyList flAudios 
      Height          =   3795
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   6694
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   4140
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ilEdit 
      Left            =   10440
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTrack.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTrack.frx":0AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTrack.frx":15E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTrack.frx":20D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditTrack.frx":2BC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbEdit 
      Height          =   2250
      Left            =   9840
      TabIndex        =   1
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   3969
      ButtonWidth     =   2752
      ButtonHeight    =   794
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilEdit"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add  "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove  "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Move up  "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Move down  "
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEditTrack"
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
'   File purpose: Audio track editor
'
Option Explicit
Option Compare Binary
Option Base 0


Private myTrack As clsAudioTrack


Public Sub Setup(ByRef Track As clsAudioTrack)

    Set myTrack = Track
    refreshSources
    
    flAudios.selectedItem = 0
    chkIgnoreDelay.Value = myTrack.ignoreDelay

End Sub


Private Sub flAudios_OLEDragDrop(Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim A As Long
    
    
    For A = 1 To Data.Files.Count
         addAudioSources Data.Files.Item(A)
    Next A
    
    refreshSources

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."
    
    ' Resize to fit toolbar
    With tlbEdit
        .Width = .ButtonWidth
        .Left = (Me.Width \ 15) - .Width - 16
        flAudios.Width = (Me.Width \ 15) - .Width - 32
    End With
    
End Sub


Private Sub refreshSources()

    Dim A As Long
    Dim Info As Dictionary
    Dim myAudio As clsAudio
    Dim itemText As String
    Dim oldSelect As Long
    
    
    oldSelect = flAudios.selectedItem
    flAudios.Refresh = False
    flAudios.Clear
    flAudios.Thumbnails = False
    flAudios.itemHeight = 60
    
    For A = 0 To myTrack.Sources.Count - 1
        Set myAudio = myTrack.Sources.Item(A)
        Set Info = myAudio.streamInfo
        
        itemText = myAudio.Source.fileName & ", track " & Info("streamIndex") + 1 & vbNewLine
        itemText = itemText & "Duration: " & visualTime(Info("Duration"))
        If Info("Delay") Then itemText = itemText & ", " & Info("Delay") & " ms delay"
        itemText = itemText & vbNewLine
        itemText = itemText & Info("Compression")
        If Info("VBR") Then
            itemText = itemText & " VBR"
        Else
            itemText = itemText & " CBR"
        End If
        itemText = itemText & ", " & Info("sampleRate") & " Hz, " & Info("Channels") & " channels"
        
        If Info("bitRate") Then itemText = itemText & ", " & Info("bitRate") & " kbit\s"
        
        flAudios.addItem itemText, Nothing
    Next A
    
    flAudios.Refresh = True
    If oldSelect < flAudios.Count And oldSelect >= 0 Then
        flAudios.selectedItem = oldSelect
    Else
        flAudios.selectedItem = 0
    End If
    
    Set myAudio = Nothing

End Sub


Private Sub cmdAccept_Click()

    myTrack.ignoreDelay = chkIgnoreDelay.Value

    Set myTrack = Nothing
    Me.Hide

End Sub


Private Sub audioAdd()
    Dim fileList As Dictionary
    

    Set fileList = fileDialog.openFile(Me.hWnd, "Select audio source", audioFiles, "", "", cdlOFNFileMustExist Or cdlOFNHideReadOnly)
    If fileList.Count Then addAudioSources fileList.Items(0)

End Sub


Private Sub addAudioSources(ByVal fileName As String)

    Dim A As Long
    Dim mySource As clsSource
    Dim Info0 As Dictionary
    Dim Info As Dictionary
    

    Me.Hide
        
    ' Get source
    Set mySource = Project.getSource(fileName)
    Set Info0 = myTrack.Sources.Item(0).streamInfo
    If Not (mySource Is Nothing) Then
        
        ' Display stream selection dialog
        If frmSelectTrack.Setup(mySource) Then frmSelectTrack.Show 1
        
        ' Add each selected stream
        With frmSelectTrack.Tracks
            For A = 0 To .Count - 1
                Set Info = mySource.streamInfo(.Keys(A))
                
                If Info("Compression") <> Info0("Compression") Or Info("sampleRate") <> Info0("sampleRate") Or Info("Channels") <> Info0("Channels") Then
                    frmDialog.Display .Items(A) & vbNewLine & "This audio source cannot be added to the track. You can only combine audio sources that are equal in compression method, samplerate and channel count.", Exclamation Or OkOnly
                Else
                    myTrack.addSource mySource, .Keys(A)
                End If
                
            Next A
        End With
        
        refreshSources
        
    
    ' Was not loaded properly
    Else
        frmDialog.Display "Unable to load " & fileName & ".", Exclamation Or OkOnly
    
    End If
    
    frmStatus.Hide
    Me.Show 1

End Sub


Private Sub audioRemove()

    If flAudios.selectedItem = -1 Then Exit Sub
    If flAudios.Count = 1 Then
        frmDialog.Display "You cannot delete all audio sources, there must be at least one.", Exclamation Or OkOnly
        Exit Sub
    End If
    
    If frmDialog.Display("Are you sure you want to remove this audio source?", Question Or YesNo) = buttonNo Then Exit Sub
    
    myTrack.Sources.Remove flAudios.selectedItem
    refreshSources

End Sub


Private Sub audioMoveUp()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flAudios.selectedItem
    If Selected <= 0 Then Exit Sub
    
    myTrack.Sources.moveBackward Selected
    
    refreshSources
    flAudios.selectedItem = Selected - 1
    flAudios.focusOn flAudios.selectedItem

End Sub


Private Sub audioMoveDown()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flAudios.selectedItem
    If Selected = myTrack.Sources.Count - 1 Then Exit Sub
    If Selected = -1 Then Exit Sub
    
    myTrack.Sources.moveForward Selected
    
    refreshSources
    flAudios.selectedItem = Selected + 1
    flAudios.focusOn flAudios.selectedItem

End Sub


Private Sub tlbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            audioAdd
        Case 2
            audioRemove
        
        Case 4
            audioMoveUp
        Case 5
            audioMoveDown
    End Select

End Sub
