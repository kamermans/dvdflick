VERSION 5.00
Object = "{1C4E972B-90D2-46D1-A2F1-DAF5A8DE9F08}#26.0#0"; "mousewheel.ocx"
Begin VB.UserControl ctlFancyList 
   BackStyle       =   0  'Transparent
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   Begin MouseWheel.ctlMouseWheel mwWheel 
      Height          =   480
      Left            =   6900
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5520
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.VScrollBar scrlBar 
         Height          =   5355
         LargeChange     =   100
         Left            =   6180
         SmallChange     =   25
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "ctlFancyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'   File purpose: A listview-style control to display text with images
'
'   TODO: Rewrite from scratch. More flexible in placements, more OO, less ugly.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Type fancyItem
    Text As String
    Image As clsGDIImage
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Items() As fancyItem
Private nItems As Long

Private mSelectedItem As Long
Private mItemHeight As Long
Private mPadding As Long
Private mImageWidth As Long
Private mImageHeight As Long
Private mRefresh As Boolean
Private mThumbnails As Boolean
Private mNoThumbBlack As Boolean

Private hasFocus As Boolean


Event Click()
Event DblClick()
Event OLEDragDrop(ByRef Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Event keyUp(ByVal Key As Long, ByVal Shift As Boolean)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal OLE_COLOR As Long, ByVal hPal As Long, dwRGB As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long

Private Declare Function getRed Lib "dvdflick" (ByVal Value As Long) As Long
Private Declare Function getGreen Lib "dvdflick" (ByVal Value As Long) As Long
Private Declare Function getBlue Lib "dvdflick" (ByVal Value As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal Idx As Long) As Long


' Return item text
Public Function itemText(ByVal Index As Long) As String

    itemText = Items(Index).Text

End Function


' Move an item up
Public Sub moveUp(ByVal Index As Long)

    Dim Temp As fancyItem
    

    If Index = nItems - 1 Then Exit Sub

    Temp = Items(Index + 1)
    Items(Index + 1) = Items(Index)
    Items(Index) = Temp
    
    If mSelectedItem = Index Then mSelectedItem = mSelectedItem + 1
    
    doRefresh

End Sub


' Move an item down
Public Sub moveDown(ByVal Index As Long)

    Dim Temp As fancyItem
    

    If Index = 0 Then Exit Sub

    Temp = Items(Index - 1)
    Items(Index - 1) = Items(Index)
    Items(Index) = Temp
    
    If mSelectedItem = Index Then mSelectedItem = mSelectedItem - 1
    
    doRefresh
    
End Sub


' Make sure that an element is within view
Public Sub focusOn(ByVal Index As Long)

    Dim sVal As Long
    Dim Offs As Long
    Dim itemSize As Long
    
    
    itemSize = mItemHeight
    Offs = Index * itemSize
    sVal = -1
    
    ' Bottom
    If Offs + itemSize >= scrlBar.Value + picList.Height - 1 Then
        sVal = Offs - picList.Height + itemHeight + itemHeight
    
    ' Top
    ElseIf Offs <= scrlBar.Value + 1 Then
        sVal = Offs - itemHeight
    
    End If
    
    If sVal <> -1 Then
        If sVal < 0 Then sVal = 0
        If sVal > scrlBar.Max Then sVal = scrlBar.Max
        scrlBar.Value = sVal
    End If

End Sub


' Render a thumbnail on a listitem
Public Sub renderThumb(ByVal hDC As Long, ByRef Image As clsGDIImage, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)

    Dim Top As Long, Left As Long


    ' Black background
    If Not mNoThumbBlack Then fillBox hDC, x, y, x + Width, y + Height, 0
    
    If Image.Height < mImageHeight Then Top = (mImageHeight - Image.Height) / 2
    If Image.Width < imageWidth Then Left = (mImageWidth - Image.Width) / 2
    
    BitBlt hDC, x + Left, y + Top, Image.Width, Image.Height, Image.hDC, 0, 0, vbSrcCopy

End Sub


' Font property
Public Property Get Font() As StdFont

    Set Font = picList.Font

End Property

Public Property Set Font(ByRef newFont As StdFont)

    Set picList.Font = newFont

End Property


' Add a listitem
Public Function addItem(ByVal Text As String, Optional ByRef Image As clsGDIImage = Nothing) As Long

    Dim Aspect As Single
    Dim rndWidth As Long
    Dim rndHeight As Long


    ReDim Preserve Items(nItems)
    nItems = nItems + 1

    With Items(nItems - 1)
        .Text = Text
        
        ' Copy image
        If Not (Image Is Nothing) Then
            Set .Image = New clsGDIImage
            .Image.copyFrom Image
            
            ' Resize keeping aspect ratio
            Aspect = getResizeValue(Image.Width, Image.Height, mImageWidth, mImageHeight)
            rndWidth = Image.Width * Aspect
            rndHeight = Image.Height * Aspect
            
            .Image.Resize rndWidth, rndHeight, STRETCH_HALFTONE
        End If
    End With
    
    doRefresh

End Function


' Clear the entire list
Public Sub Clear()

    Erase Items
    nItems = 0
    
    mSelectedItem = -1
    
    doRefresh

End Sub


Private Sub doRefresh()

    If mRefresh Then
        adjustScrollBar
        UserControl_Paint
    End If

End Sub


' Do not draw a black background behind thumbnails
Public Property Get noThumbBlack() As Boolean

    noThumbBlack = mNoThumbBlack

End Property

Public Property Let noThumbBlack(ByVal Value As Boolean)

    mNoThumbBlack = Value
    doRefresh

End Property


' Refresh (redraw) the list
Public Property Get Refresh() As Boolean

    Refresh = mRefresh

End Property

Public Property Let Refresh(ByVal Value As Boolean)

    mRefresh = Value
    
    adjustScrollBar
    UserControl_Paint

End Property


' Enable\disable thumbnails
Public Property Get Thumbnails() As Boolean

    Thumbnails = mThumbnails

End Property

Public Property Let Thumbnails(ByVal Value As Boolean)

    mThumbnails = Value

End Property


' Get\set selected listitem
Public Property Get selectedItem() As Long

    selectedItem = mSelectedItem

End Property

Public Property Let selectedItem(ByVal Value As Long)

    If Value < 0 Or Value > nItems - 1 Then Exit Property
    
    mSelectedItem = Value
    doRefresh
    
    RaiseEvent Click

End Property


' Get selected listitem's text
Public Property Get selectedText() As String

    If mSelectedItem = -1 Then Exit Property
    selectedText = Items(mSelectedItem).Text

End Property


' Get\set listitem height
Public Property Get itemHeight() As Long

    itemHeight = mItemHeight

End Property

Public Property Let itemHeight(ByVal Value As Long)

    mItemHeight = Value
    doRefresh

End Property


' Get\set thumbnail height
Public Property Get imageHeight() As Long

    imageHeight = mImageHeight

End Property

Public Property Let imageHeight(ByVal Value As Long)

    mImageHeight = Value
    doRefresh

End Property


' Get\set thumbnail width
Public Property Get imageWidth() As Long

    imageWidth = mImageWidth

End Property

Public Property Let imageWidth(ByVal Value As Long)

    mImageWidth = Value
    doRefresh

End Property


' Get\set padding between image, borders and text
Public Property Get Padding() As Long

    Padding = mPadding

End Property

Public Property Let Padding(ByVal Value As Long)

    mPadding = Value
    doRefresh

End Property


' Remove a listitem
Public Sub removeItem(ByVal Index As Long)

    Dim A As Long
    
    
    If Index < 0 Or Index > nItems - 1 Then Exit Sub
    mSelectedItem = -1
    
    For A = 0 To nItems - 2
        Items(A) = Items(A + 1)
    Next A
    
    nItems = nItems - 1
    If nItems = 0 Then
        Erase Items
    Else
        ReDim Preserve Items(nItems - 1)
    End If
    
    doRefresh

End Sub


Private Sub mwWheel_Wheel(ByVal Delta As Long)

    Me.Scroll Delta

End Sub


Private Sub picList_Click()

    RaiseEvent Click

End Sub


Private Sub picList_DblClick()

    RaiseEvent DblClick

End Sub


Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Select next\previous items
    If KeyCode = 38 And mSelectedItem > 0 Then
        mSelectedItem = mSelectedItem - 1
        focusOn mSelectedItem
        doRefresh
        RaiseEvent Click
        
    ElseIf KeyCode = 40 And mSelectedItem < nItems - 1 Then
        mSelectedItem = mSelectedItem + 1
        focusOn mSelectedItem
        doRefresh
        RaiseEvent Click
        
    End If

End Sub


Private Sub picList_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent keyUp(KeyCode, CBool(Shift))

End Sub


Private Sub picList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim Sel As Long
    

    If Button = 1 Then
        Sel = (y + scrlBar.Value) \ itemHeight
        If Sel < nItems Then
            mSelectedItem = Sel
        Else
            mSelectedItem = -1
        End If
        
        doRefresh
    End If

End Sub


' Get number of listitems
Public Property Get Count() As Long

    Count = nItems

End Property


Private Sub picList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub


Private Sub picList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)

End Sub


Private Sub scrlBar_Change()

    scrlBar_Scroll

End Sub

Private Sub scrlBar_Scroll()

    UserControl_Paint

End Sub


Private Sub UserControl_EnterFocus()

    hasFocus = True
    doRefresh

End Sub


Private Sub UserControl_ExitFocus()

    hasFocus = False
    doRefresh

End Sub


Private Sub UserControl_Initialize()

    mSelectedItem = -1
    mItemHeight = 64
    mPadding = 8
    mRefresh = True
    mThumbnails = True
    mNoThumbBlack = False
    
    mwWheel.Init UserControl.hWnd
    
End Sub


Private Sub UserControl_Paint()

    Dim A As Long
    Dim textRect As RECT
    Dim backColor As Long
    Dim Offs As Long
    Dim startItem As Long
    Dim endItem As Long
    
    
    picList.Cls
    SetStretchBltMode picList.hDC, 4
    Offs = scrlBar.Value
    
    ' Calculate start and ending item that is visible
    startItem = (Offs \ mItemHeight)
    endItem = startItem + (picList.Height \ mItemHeight) + 1
    
    ' Hide or show scrollbar
    If endItem > nItems Then
        scrlBar.Visible = False
    Else
        scrlBar.Visible = True
    End If
    
    ' Make sure item limit is not reached
    If endItem > nItems - 1 Then endItem = nItems - 1
    
    
    For A = startItem To endItem
        
        ' Text rectangle
        With textRect
            .Top = A * mItemHeight + mPadding - Offs
            .Bottom = .Top + mItemHeight - mPadding
            .Left = mPadding
            .Right = picList.Width - mPadding
        End With
        If mThumbnails Then textRect.Left = textRect.Left + mImageWidth + mPadding + mPadding
            
        ' Background
        If A = mSelectedItem Then
            'backColor = &H8000000D
            renderGradient picList.hDC, A * mItemHeight - Offs, picList.Width, mItemHeight - 1, Brighten(GetSysColor(13), 1.25), Brighten(GetSysColor(13), 0.75)
            picList.ForeColor = &H8000000E
            
        Else
            backColor = &H80000005
            fillBox picList.hDC, 0, A * mItemHeight - Offs, picList.Width, A * mItemHeight + mItemHeight - 1 - Offs, backColor
            picList.ForeColor = &H80000012
            
            ' Separator line
            drawLine picList.hDC, 0, A * mItemHeight + mItemHeight - Offs - 1, picList.Width, A * mItemHeight + mItemHeight - Offs - 1, &H8000000F
        
        End If
        
        ' Thumbnail black background and image
        If mThumbnails And Not (Items(A).Image Is Nothing) Then
            renderThumb picList.hDC, Items(A).Image, mPadding, A * mItemHeight + mPadding - Offs, mImageWidth, mImageHeight
        End If
        
        ' Text
        DrawText picList.hDC, StrPtr(Items(A).Text), -1, textRect, &H800
    Next A

End Sub


Private Function Brighten(ByVal Val As Long, ByVal Amount As Single) As Long

    Brighten = RGB(getRed(Val) * Amount, getGreen(Val) * Amount, getBlue(Val) * Amount)

End Function


Private Sub fillBox(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Color As Long)

    Dim Col As Long
    Dim Brush As Long
    Dim myRect As RECT
    
    
    OleTranslateColor Color, 0, Col
    Brush = CreateSolidBrush(Col)
    With myRect
        .Top = Y1
        .Bottom = Y2
        .Left = X1
        .Right = X2
    End With
    
    FillRect hDC, myRect, Brush
    
    DeleteObject Brush

End Sub


Private Sub drawLine(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

    Dim Pen As Long
    Dim Col As Long
    
    
    OleTranslateColor Color, 0, Col
    
    Pen = CreatePen(0, 1, Col)
    SelectObject hDC, Pen
    
    MoveToEx hDC, X1, Y1, ByVal 0&
    LineTo hDC, X2, Y2
    
    SelectObject hDC, 0
    DeleteObject Pen

End Sub


Private Function renderGradient(hDC As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startCol As Long, ByVal endCol As Long)

    Dim A As Long
    Dim Perc As Single
    Dim vRed As Long, vGreen As Long, vBlue As Long
    
    
    For A = 0 To Height
        Perc = 1 - (A / Height)
        
        vRed = (getRed(startCol) * Perc) + (getRed(endCol) * (1 - Perc))
        vGreen = (getGreen(startCol) * Perc) + (getGreen(endCol) * (1 - Perc))
        vBlue = (getBlue(startCol) * Perc) + (getBlue(endCol) * (1 - Perc))
        
        drawLine hDC, 0, y + A, Width, y + A, RGB(vRed, vGreen, vBlue)
    Next A

End Function


Private Sub UserControl_Resize()

    On Error Resume Next
    
    picList.Width = UserControl.Width / Screen.TwipsPerPixelX
    picList.Height = UserControl.Height / Screen.TwipsPerPixelY
    
    scrlBar.Left = picList.Width - scrlBar.Width - 4
    scrlBar.Height = picList.Height - 4
    
    doRefresh
    
    On Error GoTo 0

End Sub


Private Sub adjustScrollBar()

    Dim Value As Long
    
    
    If nItems = 0 Then
        scrlBar.Max = 0
    Else
        Value = ((nItems - 1) * mItemHeight + mItemHeight) - picList.Height
        If Value > 0 Then scrlBar.Max = Value Else scrlBar.Max = 0
    End If

End Sub


Public Sub Scroll(ByVal Value As Long)

    Dim newValue As Long
    
    
    newValue = scrlBar.Value - (Value * 12)
    If newValue < 0 Then newValue = 0
    If newValue > scrlBar.Max Then newValue = scrlBar.Max
    
    scrlBar.Value = newValue
    doRefresh

End Sub
