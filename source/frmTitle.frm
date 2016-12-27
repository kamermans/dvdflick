VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTitle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title properties"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
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
   Icon            =   "frmTitle.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilEdit 
      Left            =   840
      Top             =   5040
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
            Picture         =   "frmTitle.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitle.frx":0AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitle.frx":15F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitle.frx":20E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitle.frx":2BD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin dvdflick.ctlSeparator ctlSeparator1 
      Height          =   30
      Left            =   180
      Top             =   4860
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdNextTitle 
      Caption         =   "Next title >"
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrevTitle 
      Caption         =   "< Previous title"
      Height          =   375
      Left            =   2340
      TabIndex        =   25
      Top             =   5040
      Width           =   2055
   End
   Begin dvdflick.ctlFancyList flTitleOptions 
      Height          =   4515
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7964
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   9180
      TabIndex        =   29
      Top             =   5040
      Width           =   2055
   End
   Begin VB.PictureBox framGeneral 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   2280
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   60
      Width           =   9015
      Begin VB.CheckBox chkCopyTS 
         Caption         =   "Copy timestamps"
         Height          =   315
         Left            =   4260
         TabIndex        =   5
         Top             =   4200
         Width           =   1575
      End
      Begin VB.ComboBox cmbAspect 
         Height          =   315
         ItemData        =   "frmTitle.frx":36C6
         Left            =   2520
         List            =   "frmTitle.frx":36D0
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2235
      End
      Begin VB.TextBox txtName 
         Height          =   795
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   180
         Width           =   5055
      End
      Begin VB.PictureBox picThumb 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2880
         Left            =   180
         ScaleHeight     =   188
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   252
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1740
         Width           =   3840
      End
      Begin MSComCtl2.DTPicker dtTimeIndex 
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   2040
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   65208323
         UpDown          =   -1  'True
         CurrentDate     =   39312
      End
      Begin VB.Image imgICopyTS 
         Height          =   360
         Left            =   5880
         MouseIcon       =   "frmTitle.frx":36F5
         MousePointer    =   99  'Custom
         Picture         =   "frmTitle.frx":3847
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label lblTargetAspect 
         Caption         =   "Target aspect ratio"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Image imgIAspectRatio 
         Height          =   360
         Left            =   4800
         MouseIcon       =   "frmTitle.frx":4329
         MousePointer    =   99  'Custom
         Picture         =   "frmTitle.frx":447B
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Thumbnail time index"
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   1740
         Width           =   2415
      End
   End
   Begin VB.PictureBox framVideoSources 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   2280
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   60
      Width           =   9015
      Begin VB.CheckBox chkInterlaced 
         Caption         =   "Interlaced"
         Height          =   315
         Left            =   4320
         TabIndex        =   9
         Top             =   4260
         Width           =   1455
      End
      Begin VB.ComboBox cmbPAR 
         Height          =   315
         ItemData        =   "frmTitle.frx":4F5D
         Left            =   1620
         List            =   "frmTitle.frx":4F79
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4260
         Width           =   1275
      End
      Begin dvdflick.ctlFancyList flVideos 
         Height          =   4035
         Left            =   60
         TabIndex        =   6
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7117
      End
      Begin MSComctlLib.Toolbar tlbEditVideos 
         Height          =   2250
         Left            =   7440
         TabIndex        =   7
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label3 
         Caption         =   "Pixel aspect ratio"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Image imgIPAR 
         Height          =   360
         Left            =   2940
         MouseIcon       =   "frmTitle.frx":4FB2
         MousePointer    =   99  'Custom
         Picture         =   "frmTitle.frx":5104
         Top             =   4260
         Width           =   360
      End
   End
   Begin VB.PictureBox framSubtitleTracks 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   2280
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   60
      Width           =   9015
      Begin dvdflick.ctlFancyList flSubs 
         Height          =   4515
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7964
      End
      Begin MSComctlLib.Toolbar tlbEditSubs 
         Height          =   2700
         Left            =   7440
         TabIndex        =   11
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   4763
         ButtonWidth     =   2752
         ButtonHeight    =   794
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ilEdit"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add  "
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit  "
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Remove  "
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Move up  "
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Move down  "
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox framAudioTracks 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   2280
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   60
      Width           =   9015
      Begin VB.ComboBox cmbTrackLang 
         Height          =   315
         ItemData        =   "frmTitle.frx":5BE6
         Left            =   1620
         List            =   "frmTitle.frx":5BF0
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4260
         Width           =   2055
      End
      Begin VB.CommandButton cmdUseDefaultAudio 
         Caption         =   "Use as default"
         Height          =   375
         Left            =   5340
         TabIndex        =   28
         Top             =   4260
         Width           =   2055
      End
      Begin dvdflick.ctlFancyList flTracks 
         Height          =   4035
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7117
      End
      Begin MSComctlLib.Toolbar tlbEditTracks 
         Height          =   2700
         Left            =   7440
         TabIndex        =   13
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   4763
         ButtonWidth     =   2752
         ButtonHeight    =   794
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ilEdit"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add  "
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit  "
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Remove  "
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Move up  "
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Move down  "
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTrackLang 
         Caption         =   "Track language"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   4320
         Width           =   1395
      End
   End
   Begin VB.PictureBox framChapters 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   2280
      ScaleHeight     =   4695
      ScaleWidth      =   9015
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   60
      Width           =   9015
      Begin VB.ComboBox cmbChapterCount 
         Height          =   315
         ItemData        =   "frmTitle.frx":5C0E
         Left            =   2640
         List            =   "frmTitle.frx":5C27
         TabIndex        =   20
         Top             =   720
         Width           =   915
      End
      Begin VB.CheckBox chkChaptersSource 
         Caption         =   "Create chapters on every video source"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   1140
         Width           =   5355
      End
      Begin VB.CheckBox chkChapterCount 
         Caption         =   "Create"
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1995
      End
      Begin VB.CommandButton cmdAllChapters 
         Caption         =   "Apply to all titles"
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Top             =   1800
         Width           =   2070
      End
      Begin VB.CommandButton cmdUseDefault 
         Caption         =   "Use as default"
         Height          =   375
         Left            =   300
         TabIndex        =   22
         Top             =   1800
         Width           =   2070
      End
      Begin VB.ComboBox cmbChapterInterval 
         Height          =   315
         ItemData        =   "frmTitle.frx":5C47
         Left            =   2640
         List            =   "frmTitle.frx":5C66
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
      Begin VB.CheckBox chkChapterInterval 
         Caption         =   "Create chapters every"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblChapters 
         Caption         =   "chapters"
         Height          =   195
         Left            =   3660
         TabIndex        =   33
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblMinutes 
         Caption         =   "minutes"
         Height          =   195
         Left            =   3660
         TabIndex        =   32
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmTitle"
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
'   File purpose: Title edit dialog
'
Option Explicit
Option Compare Binary
Option Base 0


' Currently loaded title items
Private myTitle As clsTitle
Private myAudio As clsAudioTrack
Private myVideo As clsVideo
Private myIndex As Long

' Do not update title info
Private noUpdate As Boolean


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public Sub Setup(ByVal Index As Long)
    
    ' Get title reference
    myIndex = Index
    Set myTitle = Project.Titles.Item(myIndex)
    Me.Caption = "Properties of title " & myIndex + 1 & " - " & myTitle.Name
    
    ' Update lists
    updateTrackList
    updateSubsList
    updateVideoList
    
    ' Update General tab
    txtName.Text = myTitle.Name
    cmbAspect.ListIndex = myTitle.targetAspect
    chkCopyTS.Value = myTitle.copyTS
    
    ' Update chapter controls states
    updateChapterControls
    
    ' Thumbnail
    setDTValue dtTimeIndex, myTitle.thumbTimeIndex
    
    ' Next and previous button availability
    If Index = 0 Then cmdPrevTitle.Enabled = False Else cmdPrevTitle.Enabled = True
    If Index = Project.Titles.Count - 1 Then cmdNextTitle.Enabled = False Else cmdNextTitle.Enabled = True
    
End Sub


' Chapter control states
Private Sub updateChapterControls()

    noUpdate = True

    If myTitle.chapterCount <> -1 Then
        cmbChapterCount.Enabled = True
        chkChapterCount.Value = 1
        cmbChapterCount.Text = myTitle.chapterCount
    Else
        cmbChapterCount.Enabled = False
        chkChapterCount.Value = 0
        cmbChapterCount.Text = ""
    End If
    
    If myTitle.chapterInterval <> -1 Then
        cmbChapterInterval.Enabled = True
        chkChapterInterval.Value = 1
        cmbChapterInterval.Text = myTitle.chapterInterval
    Else
        cmbChapterInterval.Enabled = False
        chkChapterInterval.Value = 0
        cmbChapterInterval.Text = ""
    End If
        
    If myTitle.chapterOnSource = 1 Then
        chkChaptersSource.Value = 1
    Else
        chkChaptersSource.Value = 0
    End If
    
    noUpdate = False

End Sub


Private Sub chkChaptersSource_Click()

    myTitle.chapterOnSource = chkChaptersSource.Value

End Sub


Private Sub chkCopyTS_Click()

    myTitle.copyTS = chkCopyTS.Value

End Sub


Private Sub chkInterlaced_Click()
    
    If noUpdate Then Exit Sub
    
    myVideo.Interlaced = chkInterlaced.Value
    updateVideoList

End Sub


Private Sub cmbChapterInterval_Change()

    cmbChapterInterval_Click

End Sub


Private Sub cmbChapterInterval_Click()
    
    If IsNumeric(cmbChapterInterval.Text) Then
        myTitle.chapterInterval = CLng(cmbChapterInterval.Text)
        If myTitle.chapterInterval <= 0 Then cmbChapterInterval.Text = 1
        
    ElseIf cmbChapterInterval.Text <> "" Then
        cmbChapterInterval.Text = 1
        
    End If

End Sub


Private Sub cmbChapterCount_Change()

    cmbChapterCount_Click

End Sub


Private Sub cmbChapterCount_Click()

    If IsNumeric(cmbChapterCount.Text) Then
        myTitle.chapterCount = CLng(cmbChapterCount.Text)
        If myTitle.chapterCount <= 0 Then cmbChapterCount.Text = 1
        If myTitle.chapterCount > 200 Then cmbChapterCount.Text = 200
        
    ElseIf cmbChapterCount.Text <> "" Then
        cmbChapterCount.Text = 1
        
    End If

End Sub


Private Sub videoAdd()

    Dim fileList As Dictionary


    Set fileList = fileDialog.openFile(Me.hWnd, "Select video file", titleFiles, "", "", cdlOFNFileMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly)
    If fileList.Count Then addVideos fileList

End Sub


Private Sub addVideos(ByRef fileList As Dictionary)

    Dim A As Long, B As Long, C As Long
    Dim Source As clsSource
    Dim myTrack As clsAudioTrack
    Dim myVideo As clsVideo
    Dim Info As Dictionary
    Dim Info0 As Dictionary
    Dim audioCount As Long
    

    Me.Hide
        
    ' Retrieve files
    Set Info0 = myTitle.Videos.Item(0).streamInfo
    For A = 0 To fileList.Count - 1
        Set Source = Project.getSource(fileList.Items(A))
        
        For B = 0 To Source.streamCount - 1
            Set Info = Source.streamInfo(B)
            If Info("Type") = ST_Video Then
            
                If Not (Info("Width") = Info0("Width") And Info("Height") = Info0("Height") And Info("FPS") = Info0("FPS") And Info("Compression") = Info0("Compression")) Then
                    frmDialog.Display fileList.Items(A) & vbNewLine & "This file cannot be added to the title. You can only combine video sources that are equal in width, height, framerate and compression method.", Exclamation Or OkOnly
                    Exit For
                End If
            
                Set myVideo = myTitle.addVideo(Source, B)
                myVideo.Source.createStreamThumb myVideo.streamIndex, 1
                
                ' Audio track merging
                For C = 0 To Source.streamCount - 1
                    Set Info = Source.streamInfo(C)
                    If Info("Type") = ST_Audio Then
                        
                        If myTitle.audioTracks.Count = 0 Then
                            myTitle.addTrack Source, C
                        Else
                            If audioCount < myTitle.audioTracks.Count Then
                                Set myTrack = myTitle.audioTracks.Item(audioCount)
                                audioCount = audioCount + 1
                            Else
                                Set myTrack = myTitle.audioTracks.Item(0)
                            End If
                            
                            myTrack.addSource Source, C
                        End If
                        
                    End If
                Next C
                
                Exit For
            End If
        Next B
    Next A
    
    updateVideoList
    updateTrackList
    
    ' Focus on last added item
    flVideos.selectedItem = flVideos.Count - 1
    flVideos.focusOn flVideos.selectedItem
    flVideos_Click
    
    frmStatus.Hide
    Me.Show 1

End Sub


Private Sub cmdAllChapters_Click()

    Dim A As Long
    Dim Result As dialogResultConstants
    
    
    If Config.ReadSetting("dialogApplyChapterSettings", 1) = 1 Then
        Result = frmDialog.Display("Are you sure you want to apply these chapter settings to all the titles in the project?", YesNo Or Question, True)
        If (Result And checkNotAgain) Then Config.WriteSetting "dialogApplyChapterSettings", 0
        If (Result And buttonNo) Then Exit Sub
    End If
    
    For A = 0 To Project.Titles.Count - 1
        With Project.Titles.Item(A)
            .chapterCount = myTitle.chapterCount
            .chapterInterval = myTitle.chapterInterval
            .chapterOnSource = myTitle.chapterOnSource
        End With
    Next A

End Sub


Private Sub trackEdit()

    If flTracks.selectedItem = -1 Then Exit Sub
    
    frmEditTrack.Setup myTitle.audioTracks.Item(flTracks.selectedItem)
    frmEditTrack.Show 1
    
    updateTrackList
    
End Sub


Private Sub cmdPrevTitle_Click()

    If myIndex = 0 Then Exit Sub
    
    Me.Setup myIndex - 1
    renderThumb

End Sub


Private Sub cmdNextTitle_Click()

    If myIndex = Project.Titles.Count - 1 Then Exit Sub
    
    Me.Setup myIndex + 1
    renderThumb

End Sub


Private Sub videoRemove()

    Dim Result As dialogResultConstants
    

    If flVideos.selectedItem = -1 Then Exit Sub
    If flVideos.Count = 1 Then
        frmDialog.Display "You cannot delete all video files, there must be at least one video source.", Exclamation Or OkOnly
        Exit Sub
    End If
    
    If Config.ReadSetting("dialogRemoveTitleVideo", 1) = 1 Then
        Result = frmDialog.Display("Are you sure you want to remove this video source?", Question Or YesNo, True)
        If (Result And checkNotAgain) Then Config.WriteSetting "dialogRemoveTitleVideo", 0
        If (Result And buttonNo) Then Exit Sub
    End If
    
    myTitle.Videos.Remove flVideos.selectedItem
    updateVideoList

End Sub


Private Sub cmdUseDefault_Click()

    Config.WriteSetting "titleChapterCount", myTitle.chapterCount
    Config.WriteSetting "titleChapterInterval", myTitle.chapterInterval
    Config.WriteSetting "titleChapterOnSource", myTitle.chapterOnSource

End Sub


Private Sub cmdUseDefaultAudio_Click()

    Config.WriteSetting "titleAudioLanguage", myAudio.Language

End Sub


Private Sub videoMoveUp()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flVideos.selectedItem
    If Selected <= 0 Then Exit Sub
    
    myTitle.Videos.moveBackward Selected
    
    updateVideoList
    updateTrackList
    flVideos.selectedItem = Selected - 1

End Sub


Private Sub videoMoveDown()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flVideos.selectedItem
    If Selected = myTitle.Videos.Count - 1 Then Exit Sub
    If Selected = -1 Then Exit Sub
    
    myTitle.Videos.moveForward Selected
    
    updateVideoList
    updateTrackList
    flVideos.selectedItem = Selected + 1

End Sub


Private Sub trackMoveUp()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flTracks.selectedItem
    If Selected <= 0 Then Exit Sub
    
    myTitle.audioTracks.moveBackward Selected
    
    updateTrackList
    flTracks.selectedItem = Selected - 1

End Sub


Private Sub trackMoveDown()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flTracks.selectedItem
    If Selected = myTitle.audioTracks.Count - 1 Then Exit Sub
    If Selected = -1 Then Exit Sub
    
    myTitle.audioTracks.moveForward Selected
    
    updateTrackList
    flTracks.selectedItem = Selected + 1

End Sub


Private Sub subMoveUp()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flSubs.selectedItem
    If Selected <= 0 Then Exit Sub
    
    myTitle.Subtitles.moveBackward Selected
    
    updateSubsList
    flSubs.selectedItem = Selected - 1

End Sub


Private Sub subMoveDown()

    Dim Selected As Long
    
    
    ' Valid selection
    Selected = flSubs.selectedItem
    If Selected = myTitle.Subtitles.Count - 1 Then Exit Sub
    If Selected = -1 Then Exit Sub
    
    myTitle.Subtitles.moveForward Selected
    
    updateSubsList
    flSubs.selectedItem = Selected + 1

End Sub


Private Sub updateVideoList()

    Dim A As Long
    Dim itemText As String
    Dim Info As Dictionary
    Dim thisVideo As clsVideo
    Dim useThumb As clsGDIImage
    Dim oldSelect As Long
    
    
    oldSelect = flVideos.selectedItem
    flVideos.Refresh = False
    flVideos.Clear
    flVideos.imageWidth = 64
    flVideos.imageHeight = 50
    flVideos.itemHeight = 66
    
    For A = 0 To myTitle.Videos.Count - 1
        Set thisVideo = myTitle.Videos.Item(A)
        Set Info = thisVideo.streamInfo
        
        itemText = FS.GetFileName(thisVideo.Source.fileName) & vbNewLine
        itemText = itemText & "Duration: " & visualTime(Info("Duration")) & ", " & Info("FPS") & " FPS"
        If Info("bitRate") Then itemText = itemText & ", " & Info("bitRate") & " Kbit/s"
        itemText = itemText & ", " & thisVideo.guessedFormat
        itemText = itemText & vbNewLine
        
        itemText = itemText & Info("Compression") & ", " & Info("Width") & "x" & Info("Height") & ", " & visualAspectRatio(Info("sourceAR")) & " SAR, " & visualAspectRatio(thisVideo.PAR) & " PAR"
        If thisVideo.Interlaced = 1 Then itemText = itemText & ", Interlaced"
        
        Set useThumb = thisVideo.Thumbnail
        If useThumb Is Nothing Then Set useThumb = noThumb
        flVideos.addItem itemText, resizeToMatch(useThumb, flVideos.imageWidth, flVideos.imageHeight, thisVideo.PAR)
    Next A
    
    flVideos.Refresh = True
    
    ' Select first
    If oldSelect < flVideos.Count And oldSelect >= 0 Then
        flVideos.selectedItem = oldSelect
    Else
        flVideos.selectedItem = 0
    End If
    flVideos.focusOn flVideos.selectedItem
    flVideos_Click
    
    Set thisVideo = Nothing

End Sub


Private Sub updateTrackList()

    Dim A As Long, B As Long
    Dim Audio As clsAudioTrack
    Dim itemText As String
    Dim oldSelect As Long
    
    
    oldSelect = flTracks.selectedItem
    flTracks.Refresh = False
    flTracks.Thumbnails = False
    flTracks.itemHeight = 44
    flTracks.Clear
    
    For A = 0 To myTitle.audioTracks.Count - 1
        Set Audio = myTitle.audioTracks.Item(A)
        
        itemText = Audio.Sources.Item(0).Source.fileName
        If Audio.Sources.Count > 1 Then
            For B = 1 To Audio.Sources.Count - 1
                itemText = itemText & " + " & FS.GetFileName(Audio.Sources.Item(B).Source.fileName)
            Next B
        End If
        itemText = itemText & vbNewLine
        
        itemText = itemText & Audio.Sources.Count & " sources, " & visualTime(Audio.Duration)
        
        flTracks.addItem itemText, Nothing
    Next A
    
    flTracks.Refresh = True
    
    ' Select first
    If flTracks.Count Then
        If oldSelect < flTracks.Count And oldSelect >= 0 Then
            flTracks.selectedItem = oldSelect
        Else
            flTracks.selectedItem = 0
        End If
    End If
    flTracks.focusOn flTracks.selectedItem
    flTracks_Click
    
    Set Audio = Nothing

End Sub


Private Sub cmbAspect_Click()

    myTitle.targetAspect = cmbAspect.ListIndex

End Sub


Private Sub cmbTrackLang_Click()

    If noUpdate = False Then myAudio.Language = cmbTrackLang.ListIndex

End Sub


Private Sub cmdAccept_Click()
    
    Project.Modified = True
    Me.Hide

End Sub


' Add a new audio stream
Private Sub trackAdd()

    Dim fileList As Dictionary
    

    ' Max. audio streams
    If myTitle.audioTracks.Count = 8 Then
        frmDialog.Display "The maximum number of allowed audio tracks has been reached.", Exclamation Or OkOnly
        
    Else
        Set fileList = fileDialog.openFile(Me.hWnd, "Select audio source", audioFiles, "", "", cdlOFNFileMustExist Or cdlOFNHideReadOnly)
        If fileList.Count Then addTrackFile fileList.Items(0)
        
    End If

End Sub


Private Sub addTrackFile(ByVal fileName As String)

    Dim A As Long
    Dim mySource As clsSource
    Dim tempTrack As clsAudioTrack
    

    Me.Hide

    ' Get source class
    Set mySource = Project.getSource(fileName)
    frmStatus.Hide
    
    If Not (mySource Is Nothing) Then
        
        ' Display stream selection dialog
        frmSelectTrack.Setup mySource
        
        ' No streams?
        If frmSelectTrack.lstTracks.ListCount = 0 Then
            frmDialog.Display "There are no audio tracks in the file or the audio tracks found are not supported.", Exclamation Or OkOnly
        
        Else
            frmSelectTrack.Show 1
            
            ' Add each selected stream
            With frmSelectTrack.Tracks
                For A = 0 To .Count - 1
                    Set tempTrack = myTitle.addTrack(mySource, .Keys(A))
                    tempTrack.Language = Config.ReadSetting("titleAudioLanguage", tempTrack.Language)
                Next A
            End With
            
            updateTrackList
            
            ' Focus on last added item
            flTracks.selectedItem = flTracks.Count - 1
            flTracks.focusOn flTracks.selectedItem
            
        End If
        
    
    ' Was not loaded properly
    Else
        frmDialog.Display "Unable to load " & fileName & ".", Exclamation Or OkOnly
    
    End If
    
    Set mySource = Nothing
    Me.Show 1

End Sub


Private Sub subAdd()

    Dim fileList As Dictionary
    

    ' Max. subtitles reached
    If myTitle.Subtitles.Count = 32 Then
        frmDialog.Display "The maximum number of allowed subtitles has been reached.", Exclamation Or OkOnly
        Exit Sub
    End If

    Set fileList = fileDialog.openFile(Me.hWnd, "Select subtitle file", subFiles, "", "", cdlOFNFileMustExist Or cdlOFNHideReadOnly)
    If fileList.Count Then addSubFile fileList.Items(0)

End Sub


Private Sub addSubFile(ByVal fileName As String)

    Dim subFile As clsSubFile
    
    
    ' Warn for binary (thus possibly graphical) subtitle files
    If isBinaryFile(fileName) Then
        frmDialog.Display "The subtitle file you selected appears to be graphical. DVD Flick only supports textual subtitles.", OkOnly Or Information
        Exit Sub
    End If
    
    ' Check if the file is valid by attempting to load it
    Set subFile = New clsSubFile
    If Not subFile.openFrom(fileName) Or subFile.blockCount = 0 Then
        frmDialog.Display "The subtitle file is corrupt or is unsupported by DVD Flick.", OkOnly Or Exclamation
        Exit Sub
    End If
    
    ' Check for overlaps
    If subFile.fixOverlaps(False) > 0 Then
        frmDialog.Display "WARNING: the subtitle has overlapping timestamps. DVD Flick will attempt to fix them, but the resulting DVD may end up truncated.", Exclamation Or OkOnly
    End If
        
    Me.Enabled = False
    
    myTitle.addSub subFile
    updateSubsList
    
    ' Focus on last added item
    flSubs.selectedItem = flSubs.Count - 1
    flSubs.focusOn flSubs.selectedItem
    
    Me.Enabled = True

End Sub


Private Sub subRemove()
    Dim Result As dialogResultConstants
    
    
    If flSubs.selectedItem = -1 Then Exit Sub
    
    If Config.ReadSetting("dialogRemoveTitleSub", 1) = 1 Then
        Result = frmDialog.Display("Are you sure you want to remove this subtitle?", Question Or YesNo, True)
        If (Result And checkNotAgain) Then Config.WriteSetting "dialogRemoveTitleSub", 0
        If (Result And buttonNo) Then Exit Sub
    End If

    myTitle.Subtitles.Remove flSubs.selectedItem
    updateSubsList

End Sub


Private Sub updateSubsList()

    Dim A As Long
    Dim itemText As String
    Dim mySub As clsSubtitle
    Dim oldSelect As Long
    
    
    oldSelect = flSubs.selectedItem
    flSubs.Refresh = False
    flSubs.Thumbnails = False
    flSubs.itemHeight = 60
    flSubs.Clear
    
    For A = 0 To myTitle.Subtitles.Count - 1
        Set mySub = myTitle.Subtitles.Item(A)
        
        itemText = mySub.fileName & vbNewLine
        itemText = itemText & "Format: " & mySub.fileFormat
        If mySub.frameBased = 1 Then itemText = itemText & ", " & mySub.FPS & " FPS"
        itemText = itemText & vbNewLine
        itemText = itemText & "Language: " & langCodes.Items(mySub.Language) & vbNewLine
        flSubs.addItem itemText, Nothing
    Next A
    
    flSubs.Refresh = True
    
    ' Select first
    If flSubs.Count Then
        If oldSelect < flSubs.Count And oldSelect >= 0 Then
            flSubs.selectedItem = oldSelect
        Else
            flSubs.selectedItem = 0
        End If
    End If
    
    flSubs.focusOn flSubs.selectedItem
    
    Set mySub = Nothing

End Sub


Private Sub trackRemove()

    Dim Result As dialogResultConstants
    
    
    If flTracks.selectedItem = -1 Then Exit Sub
    
    If Config.ReadSetting("dialogRemoveTitleTrack", 1) = 1 Then
        Result = frmDialog.Display("Are you sure you want to remove this track?", Question Or YesNo, True)
        If (Result And checkNotAgain) Then Config.WriteSetting "dialogRemoveTitleTrack", 0
        If (Result And buttonNo) Then Exit Sub
    End If
    
    myTitle.audioTracks.Remove flTracks.selectedItem
    updateTrackList

End Sub


Private Sub subEdit()

    If flSubs.selectedItem = -1 Then Exit Sub
    
    ' Display subtitle editing form
    frmSubtitle.Setup myTitle.Subtitles.Item(flSubs.selectedItem), myTitle
    frmSubtitle.Show 1
    
    updateSubsList

End Sub


Private Sub dtTimeIndex_Change()

    Dim timeIndex As Long
    Dim Duration As Long
    
    
    Duration = myTitle.Videos.Item(0).streamInfo("Duration")
    timeIndex = getDTValue(dtTimeIndex)
    
    If timeIndex >= Duration Then setDTValue dtTimeIndex, Duration
    
    myTitle.thumbTimeIndex = timeIndex
    renderThumb

End Sub


Private Sub flSubs_KeyUp(ByVal Key As Long, ByVal Shift As Boolean)

    If Key = 46 And Shift = False Then subRemove

End Sub


Private Sub flSubs_OLEDragDrop(Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    Dim A As Long
    
    
    For A = 1 To Data.Files.Count
         addSubFile Data.Files.Item(A)
    Next A
    
    updateSubsList

End Sub


Private Sub flTitleOptions_Click()

    If flTitleOptions.selectedItem = -1 Then Exit Sub

    framGeneral.Visible = False
    framChapters.Visible = False
    framVideoSources.Visible = False
    framAudioTracks.Visible = False
    framSubtitleTracks.Visible = False
    
    Select Case flTitleOptions.selectedText
    Case "General"
        framGeneral.Visible = True
    Case "Chapters"
        framChapters.Visible = True
    Case "Video sources"
        framVideoSources.Visible = True
    Case "Audio tracks"
        framAudioTracks.Visible = True
    Case "Subtitle tracks"
        framSubtitleTracks.Visible = True
    End Select

End Sub


Private Sub flTracks_DblClick()

    trackEdit

End Sub


Private Sub flTracks_KeyUp(ByVal Key As Long, ByVal Shift As Boolean)

    If Key = 46 And Shift = False Then trackRemove

End Sub


Private Sub flTracks_OLEDragDrop(Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    Dim A As Long
    
    
    For A = 1 To Data.Files.Count
         addTrackFile Data.Files.Item(A)
    Next A
    
    updateTrackList

End Sub


Private Sub flVideos_KeyUp(ByVal Key As Long, ByVal Shift As Boolean)

    If Key = 46 And Shift = False Then videoRemove

End Sub


Private Sub flVideos_OLEDragDrop(Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    Dim A As Long
    Dim fileList As Dictionary
    
    
    Set fileList = New Dictionary
    For A = 1 To Data.Files.Count
        fileList.Add A, Data.Files.Item(A)
    Next A
    addVideos fileList
    
    updateVideoList

End Sub


Private Sub Form_Activate()

    renderThumb

End Sub


Private Sub renderThumb()

    Dim Pic As clsGDIImage
    Dim myVideo As clsVideo
    Dim useThumb As clsGDIImage
    
    
    Set myVideo = myTitle.Videos.Item(0)
    
    Set useThumb = myVideo.Thumbnail
    If useThumb Is Nothing Then Set useThumb = noThumb
    
    Set Pic = resizeToMatch(useThumb, picThumb.ScaleWidth, picThumb.ScaleHeight, myVideo.PAR)
    BitBlt picThumb.hDC, 0, 0, Pic.Width, Pic.Height, Pic.hDC, 0, 0, vbSrcCopy
    picThumb.Refresh

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    ' Alt+L or R moves through titles
    If KeyCode = 37 And Shift = 4 Then cmdPrevTitle_Click
    If KeyCode = 39 And Shift = 4 Then cmdNextTitle_Click

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

    ' Setup menu list
    flTitleOptions.Thumbnails = False
    flTitleOptions.Padding = 8
    flTitleOptions.itemHeight = 32
    
    flTitleOptions.addItem "General"
    flTitleOptions.addItem "Chapters"
    flTitleOptions.addItem "Video sources"
    flTitleOptions.addItem "Audio tracks"
    flTitleOptions.addItem "Subtitle tracks"
    
    flTitleOptions.selectedItem = 0
    flTitleOptions_Click
    
    ' Setup language combobox
    fillLangCombo cmbTrackLang
    
    
    ' Resize to fit edit toolbars
    With tlbEditVideos
        .Width = .ButtonWidth
        .Left = framVideoSources.Width - .Width
        flVideos.Width = framVideoSources.Width - .Width - 8
    End With
    
    With tlbEditSubs
        .Width = .ButtonWidth
        .Left = framSubtitleTracks.Width - .Width
        flSubs.Width = framSubtitleTracks.Width - .Width - 8
    End With
    
    With tlbEditTracks
        .Width = .ButtonWidth
        .Left = framAudioTracks.Width - .Width
        flTracks.Width = framAudioTracks.Width - .Width - 8
        cmdUseDefaultAudio.Left = flTracks.Left + flTracks.Width - cmdUseDefaultAudio.Width
    End With

End Sub


Private Sub flVideos_Click()

    noUpdate = True
    
    ' Clear stream info if none in list
    If flVideos.selectedItem = -1 Then
        cmbPAR.Enabled = False
        chkInterlaced.Enabled = False
        chkInterlaced.Value = 0
        cmbPAR.ListIndex = -1
        noUpdate = False
        Exit Sub
    End If
    
    ' Get video object
    Set myVideo = myTitle.Videos.Item(flVideos.selectedItem)
    
    ' Update displayed stream info
    If myVideo.PAR = myVideo.streamInfo("pixelAR") Then
        cmbPAR.ListIndex = 0
    Else
        Select Case Round(myVideo.PAR, 2)
            Case 1: cmbPAR.ListIndex = 1
            Case Round(3 / 4, 2): cmbPAR.ListIndex = 2
            Case Round(4 / 3, 2): cmbPAR.ListIndex = 3
            Case Round(16 / 9, 2): cmbPAR.ListIndex = 4
            Case Round(16 / 10, 2): cmbPAR.ListIndex = 5
            Case 2.35: cmbPAR.ListIndex = 6
        End Select
    End If
    chkInterlaced.Value = myVideo.Interlaced
    
    cmbPAR.Enabled = True
    chkInterlaced.Enabled = True
    
    noUpdate = False
    
End Sub


Private Sub flTracks_Click()

    noUpdate = True
    
    ' Clear stream info if none in list
    If flTracks.Count = 0 Or flTracks.selectedItem = -1 Then
        cmbTrackLang.Enabled = False
        cmbTrackLang.ListIndex = -1
        noUpdate = False
        Exit Sub
    End If
    
    ' Get audio class
    Set myAudio = myTitle.audioTracks.Item(flTracks.selectedItem)
    
    ' Update displayed stream info
    cmbTrackLang.ListIndex = myAudio.Language
    cmbTrackLang.Enabled = True
    
    noUpdate = False
    
End Sub


Private Sub flSubs_DblClick()

    subEdit

End Sub


Private Sub imgIAspectRatio_Click()

    frmDialog.Display "This setting controls the aspect ratio of the resulting video; whether it is made to fit widescreen TVs (16:9) or regular TVs (4:3). This is autodetected from the source video's dimensions.", OkOnly Or Information

End Sub

Private Sub imgICopyTS_Click()

    frmDialog.Display "When enabled, the timestamps present in the original source files will be copied, and not recalculated.", OkOnly Or Information

End Sub

Private Sub imgIPAR_Click()

    frmDialog.Display "If the video thumbnail appears to be stretched, use this setting to adjust the pixel aspect ratio. If it appears OK, leave it set to Default.", OkOnly Or Information

End Sub


' Chapter by count
Private Sub chkChapterCount_Click()

    If noUpdate Then Exit Sub

    If chkChapterCount.Value = 1 Then
        cmbChapterCount.Text = Config.ReadSetting("titleChapterCount", Default_ChapterCount)
        cmbChapterCount.Enabled = True
    Else
        cmbChapterCount.Text = ""
        cmbChapterCount.Enabled = False
        myTitle.chapterCount = -1
    End If
    
End Sub


' Chapter by interval
Private Sub chkChapterInterval_Click()

    If noUpdate Then Exit Sub

    If chkChapterInterval.Value = 1 Then
        cmbChapterInterval.Text = Config.ReadSetting("titleChapterInterval", Default_ChapterInterval)
        cmbChapterInterval.Enabled = True
    Else
        cmbChapterInterval.Text = ""
        cmbChapterInterval.Enabled = False
        myTitle.chapterInterval = -1
    End If

End Sub


Private Sub tlbEditVideos_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            videoAdd
        Case 2
            videoRemove
        
        Case 4
            videoMoveUp
        Case 5
            videoMoveDown
    
    End Select

End Sub


Private Sub tlbEditSubs_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1
            subAdd
        Case 2
            subEdit
        Case 3
            subRemove
        
        Case 5
            subMoveUp
        Case 6
            subMoveDown
    End Select

End Sub


Private Sub tlbEditTracks_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    
        Case 1
            trackAdd
        Case 2
            trackEdit
        Case 3
            trackRemove
        
        Case 5
            trackMoveUp
        Case 6
            trackMoveDown
        
    End Select

End Sub


Private Sub txtName_Change()

    myTitle.Name = txtName.Text

End Sub


Private Sub cmbPAR_Click()

    cmbPAR_Change

End Sub

Private Sub cmbPAR_Change()

    If noUpdate Then Exit Sub
    
    If cmbPAR.ListIndex = 0 Then myVideo.PAR = myVideo.streamInfo("pixelAR")
    If cmbPAR.ListIndex = 1 Then myVideo.PAR = 1
    If cmbPAR.ListIndex = 2 Then myVideo.PAR = 3 / 4
    If cmbPAR.ListIndex = 3 Then myVideo.PAR = 4 / 3
    If cmbPAR.ListIndex = 4 Then myVideo.PAR = 16 / 9
    If cmbPAR.ListIndex = 5 Then myVideo.PAR = 16 / 10
    If cmbPAR.ListIndex = 6 Then myVideo.PAR = 1.85
    If cmbPAR.ListIndex = 7 Then myVideo.PAR = 2.35
    
    updateVideoList

End Sub
