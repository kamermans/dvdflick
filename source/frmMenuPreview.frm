VERSION 5.00
Begin VB.Form frmMenuPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Preview"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11520
   Icon            =   "frmMenuPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   616
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   768
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbTitles 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2715
   End
   Begin VB.ComboBox cmbSubMenus 
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
      Left            =   2940
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2715
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      ForeColor       =   &H80000008&
      Height          =   8640
      Left            =   0
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   768
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   11520
   End
   Begin VB.Label lblLocation 
      Caption         =   "Main menu"
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
      Left            =   5880
      TabIndex        =   3
      Top             =   180
      Width           =   5475
   End
End
Attribute VB_Name = "frmMenuPreview"
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
'   File purpose: Display an interactive preview of a menu.
'
Option Explicit
Option Compare Binary
Option Base 0


Private cTemplate As clsMenuTemplate

Private cButtonIndex As Long
Private cMenu As clsMenu
Private cButton As clsMenuButton
Private cTitle As clsTitle

Private buttonHighlight As Boolean
Private imgTemp As clsGDIImage

Private drawBorders As Boolean


Private Declare Function TransBlt Lib "msimg32" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nDestLeft As Long, ByVal nDestTop As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hSrcDC As Long, ByVal nSrcLeft As Long, ByVal nSrcTop As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal stretchMode As Long) As Long


Public Sub Setup(ByVal Path As String)
    
    Dim A As Long
    
    
    SetStretchBltMode picPreview.hDC, STRETCH_DELETESCANS
    
    ' Resize window to fit target format
    picPreview.Width = Project.menuHeight * MENU_ASPECT
    picPreview.Height = Project.menuHeight
    Me.Width = picPreview.Width * 15
    Me.Height = (picPreview.Top + picPreview.Height + modUtil.getTitleBarHeight) * 15
    
    ' Center window
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    
    ' Open template
    Set cTemplate = New clsMenuTemplate
    cTemplate.openFrom Path
    
    ' Generate and render menus
    Project.generateMenus cTemplate.Templates
    Project.renderMenus cTemplate
    Project.rescaleMenus
    
    ' Reset preview state
    Set cButton = Nothing
    Set cTitle = Nothing
    Set cMenu = Project.Menus.Items(0)
    cButtonIndex = -1
    buttonHighlight = False
    
    ' Stretched background image to speed up blitting
    Set imgTemp = New clsGDIImage
    imgTemp.createNew picPreview.Width, picPreview.Height, cMenu.imgBack.BPP
    
    ' Create list of titles
    cmbTitles.Clear
    cmbTitles.addItem "Main menu"
    For A = 0 To Project.Titles.Count - 1
        cmbTitles.addItem "Title " & A + 1 & ": " & Project.Titles.Item(A).Name
    Next A
    cmbTitles.ListIndex = 0
    
End Sub


Private Sub updateMenuImage()

    imgTemp.renderImage cMenu.imgBack, 0, 0, imgTemp.Width, imgTemp.Height, Render_Copy, STRETCH_DELETESCANS

End Sub


Private Sub renderPreview()

    Dim x As Long, y As Long
    Dim Width As Long, Height As Long
    Dim Modif As Single
    Dim vertScan As Long, horzScan As Long


    BitBlt picPreview.hDC, 0, 0, picPreview.Width, picPreview.Height, imgTemp.hDC, 0, 0, vbSrcCopy
    
    If Not cButton Is Nothing Then
        Modif = (picPreview.Width / 720)
    
        x = cButton.Left
        y = cButton.Top
        Width = cButton.Right - cButton.Left
        Height = cButton.Bottom - cButton.Top
    
        ' This is supposed to be the other way around logically. Highlight = was just clicked.
        ' TODO: Rename highlight to select.
        If buttonHighlight Then
            TransBlt picPreview.hDC, x * Modif, y, Width * Modif, Height, cMenu.imgSelect.hDC, x, y, Width, Height, 0
        Else
            TransBlt picPreview.hDC, x * Modif, y, Width * Modif, Height, cMenu.imgHighlight.hDC, x, y, Width, Height, 0
        End If
    End If

    horzScan = picPreview.Width * 0.067
    vertScan = picPreview.Width * 0.05
    If drawBorders Then picPreview.Line (horzScan, vertScan)-(picPreview.Width - horzScan, picPreview.Height - vertScan), RGB(255, 0, 255), B

    picPreview.Refresh
    
End Sub


Private Sub cmbSubMenus_Click()

    ' VMGM
    If cTitle Is Nothing Then
        Set cMenu = Project.Menus.Items(cmbSubMenus.ListIndex)
    
    ' Titleset
    Else
        Set cMenu = cTitle.Menus.Items(cmbSubMenus.ListIndex)
    
    End If
    
    buttonHighlight = False
    updateMenuImage
    renderPreview

End Sub


Private Sub cmbTitles_Click()

    Dim Index As Long
    
    
    Index = cmbTitles.ListIndex
    
    ' Titleset
    If Index > 0 Then
        Set cTitle = Project.Titles.Item(Index - 1)
        updateSubMenus cTitle.Menus
        cmbSubMenus.ListIndex = 0
        
    ' VMGM
    Else
        Set cTitle = Nothing
        updateSubMenus Project.Menus
        cmbSubMenus.ListIndex = 0
        
    End If
    
    lblLocation.Caption = cmbTitles.List(cmbTitles.ListIndex)

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    Dim Key As String
    
    
    Key = LCase$(Chr$(KeyAscii))
    
    If KeyAscii = 27 Then Me.Hide
    If Key = "s" Then
        drawBorders = Not drawBorders
        renderPreview
    End If

End Sub


Private Sub Form_Load()

    appLog.Add "Loading " & Me.Name & "..."

End Sub


Private Sub picPreview_Click()

    Dim startCount As Long
    
    
    buttonHighlight = True
    renderPreview
    
    startCount = GetTickCount
    Do
        DoEvents
    Loop Until GetTickCount - startCount >= 400
    
    parseButtonAction

End Sub


' Parse a button's actions; crudely
Private Sub parseButtonAction()

    Dim A As Long
    Dim Actions() As String
    Dim Title As Long, Page As Long
    

    If cButton Is Nothing Then Exit Sub
    
    Actions = Split(cButton.Action, ";")
    For A = 0 To UBound(Actions)
        Actions(A) = Trim$(Actions(A))
    
        If Left$(Actions(A), 5) = "jump " Then
            
            ' Jump to root menu
            If Mid$(Actions(A), 6, 10) = "vmgm menu " Then
                Set cTitle = Nothing
                Page = Val(Mid$(Actions(A), 16)) - 1
                
                cmbTitles.ListIndex = 0
                lblLocation.Caption = "Main menu"
                updateSubMenus Project.Menus
                cmbSubMenus.ListIndex = Page
            
            ' Jump to title
            ElseIf Right$(Actions(A), 13) = "jump vmgm fpc" Then
                Set cTitle = Project.Titles.Item(Title)
                
                cmbTitles.ListIndex = Title + 1
                updateSubMenus cTitle.Menus
                cmbSubMenus.ListIndex = 0

            End If
        
        ' Select title
        ElseIf Left$(Actions(A), 5) = "g0 = " Then
            Title = Val(Mid$(Actions(A), 6)) - 1
        
        End If
    Next A
    
    buttonHighlight = False

End Sub


Private Sub updateSubMenus(ByRef Dict As Dictionary)

    Dim A As Long
    
    
    cmbSubMenus.Clear
    For A = 0 To Dict.Count - 1
        cmbSubMenus.addItem Dict.Keys(A)
    Next A

End Sub


Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim A As Long
    Dim myButton As clsMenuButton
    
    
    x = x * (720 / (Project.menuHeight * MENU_ASPECT))
    For A = 0 To cMenu.Buttons.Count - 1
        Set myButton = cMenu.Buttons.Items(A)
        If x >= myButton.Left And x <= myButton.Right Then
            If y >= myButton.Top And y <= myButton.Bottom Then
                cButtonIndex = A
                Set cButton = myButton
                renderPreview
            End If
        End If
    Next A

End Sub


Private Sub picPreview_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then Me.Hide

End Sub
