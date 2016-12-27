VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDIImage"
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
'   File purpose: GDI DIB manipulation class.
'
Option Explicit
Option Compare Binary
Option Base 0


' Bitmap object data
Private Type BitmapData
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPerPixel As Integer
    bmBitsPointer As Long
End Type

' OLE GUID
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' OLE picture object
Private Type OLEPictureType
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

' Rectangle structure
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


' LoadImage constants
Private Const LR_DEFAULTCOLOR = 0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000


' Properties
Private mHandle As Long
Private mWidth As Long
Private mHeight As Long
Private mBPP As Long
Private mDC As Long


' GDI
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function getObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal startIndex As Long, ByVal Entries As Long, ByVal ptrColors As Long) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal startIndex As Long, ByVal Entries As Long, ByVal ptrColors As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, ByRef pBMI As bitmapInfo, ByVal iUsage As Long, ByRef ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nDestLeft As Long, ByVal nDestTop As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hSrcDC As Long, ByVal nSrcLeft As Long, ByVal nSrcTop As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal stretchMode As Long) As Long
Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Long, ByVal dwLayout As Long) As Long

' User
Private Declare Function LoadImage Lib "user32" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

' Kernel
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

' DVD Flick
Private Declare Function GDIImageSelect Lib "dvdflick" (ByVal hBitmap As Long) As Long
Private Declare Function GDIColorFill Lib "dvdflick" (ByVal Color As Long) As Long
Private Declare Function GDIColorReplace Lib "dvdflick" (ByVal colorReplace As Long, ByVal Color As Long) As Long
Private Declare Function GDIAlphaBlit Lib "dvdflick" (ByVal hSource As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GDIRenderOutline Lib "dvdflick" (ByVal outlineSize As Long, ByVal baseColor As Long, ByVal backColor As Long, ByVal outlineColor As Long) As Long

' Misc.
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As OLEPictureType, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Function TransBlt Lib "msimg32" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long


' Crop this image
Public Function Crop(ByVal newWidth As Long, ByVal newHeight As Long) As Boolean

    Dim imgTemp As clsGDIImage
    Dim Pal(255) As rgbQuad
    Dim oldWidth As Long, oldHeight As Long
    
    
    Dbg.addLog DM_GDI, "Cropping " & mWidth & ", " & mHeight & " to " & newWidth & ", " & newHeight
    
    ' Check dimensions
    If newWidth >= mWidth Then Exit Function
    If newHeight >= mHeight Then Exit Function
    
    ' Create copy of this image
    Set imgTemp = New clsGDIImage
    If Not imgTemp.copyFrom(Me) Then Exit Function

    ' Copy it to this one cropped
    oldWidth = mWidth
    oldHeight = mHeight
    If imgTemp.BPP = 8 Then
        imgTemp.getPalette VarPtr(Pal(0))
        If Not Me.createNew(newWidth, newHeight, imgTemp.BPP, VarPtr(Pal(0))) Then Exit Function
    Else
        If Not Me.createNew(newWidth, newHeight, imgTemp.BPP) Then Exit Function
    End If
    If BitBlt(Me.hDC, (oldWidth / 2) - (newWidth / 2), (oldHeight / 2) - (newHeight / 2), newWidth, newHeight, imgTemp.hDC, 0, 0, vbSrcCopy) = 0 Then Exit Function
    
    Crop = True

End Function


' Render an outline on this bitmap
Public Function renderOutline(ByVal outlineSize As Long, ByVal colorForeground As Long, ByVal colorBackground As Long, ByVal colorOutline As Long) As Boolean

    Dbg.addLog DM_GDI, "Rendering outline " & outlineSize & ", " & colorForeground & ", " & colorBackground & ", " & colorOutline
    
    If GDIImageSelect(mHandle) Then Exit Function
    If GDIRenderOutline(outlineSize, colorForeground, colorBackground, colorOutline) Then Exit Function

    renderOutline = True

End Function


' Fill the bitmap with a color
Public Function colorFill(ByVal Color As Long) As Boolean

    Dbg.addLog DM_GDI, "Filling with color " & Color
    
    If GDIImageSelect(mHandle) Then Exit Function
    If GDIColorFill(Color) Then Exit Function
    
    colorFill = True

End Function


' Replace a color
Public Function colorReplace(ByVal replaceColor As Long, ByVal Color As Long) As Boolean

    Dbg.addLog DM_GDI, "Replacing color " & replaceColor & " with " & Color
    
    If GDIImageSelect(mHandle) Then Exit Function
    If GDIColorReplace(replaceColor, Color) Then Exit Function
    
    colorReplace = True

End Function


' Resize this image
' TODO: Faster.
Public Function Resize(ByVal Width As Long, ByVal Height As Long, ByVal stretchMode As stretchModes) As Boolean

    Dim imgTemp As clsGDIImage
    Dim Pal(255) As rgbQuad
    
    
    Dbg.addLog DM_GDI, "Resizing " & mWidth & ", " & mHeight & " to " & Width & ", " & Height
    
    Set imgTemp = New clsGDIImage
    If Not imgTemp.copyFrom(Me) Then Exit Function

    If imgTemp.BPP = 8 Then
        imgTemp.getPalette VarPtr(Pal(0))
        If Not Me.createNew(Width, Height, imgTemp.BPP, VarPtr(Pal(0))) Then Exit Function
    Else
        If Not Me.createNew(Width, Height, imgTemp.BPP) Then Exit Function
    End If
    If Not Me.renderImage(imgTemp, 0, 0, Width, Height, Render_Copy, stretchMode) Then Exit Function
    
    Resize = True

End Function


' Copy another GDI image into this one, allowing for different color depth
Public Function copyFrom(ByRef Image As clsGDIImage, Optional ByVal BPP As Long = -1) As Boolean

    Dim Pal(255) As rgbQuad
    
    
    Dbg.addLog DM_GDI, "Copying image"
    
    If Image Is Nothing Then Exit Function
    If BPP = -1 Then BPP = Image.BPP
    
    If Not createNew(Image.Width, Image.Height, BPP) Then Exit Function
    If Image.BPP = 8 Then
        Image.getPalette VarPtr(Pal(0))
        setPalette VarPtr(Pal(0))
    End If
    
    If BitBlt(hDC, 0, 0, Width, Height, Image.hDC, 0, 0, vbSrcCopy) = 0 Then Exit Function
    copyFrom = True

End Function


' Modify this image's palette
' Uses pointers to circumvent VB's incapability of handling UDTs as class method parameters
Public Function getPalette(ByVal palPointer As Long) As Boolean

    If palPointer = 0 Then Exit Function
    
    If GetDIBColorTable(mDC, 0, 256, palPointer) = 0 Then Exit Function
    getPalette = True

End Function

Public Function setPalette(ByVal palPointer As Long) As Boolean

    If palPointer = 0 Then Exit Function
    
    If SetDIBColorTable(mDC, 0, 256, palPointer) = 0 Then Exit Function
    setPalette = True

End Function


' Return width and height of a text
Public Function getTextWidth(ByRef Font As clsGDIFont, ByVal Text As String) As Long

    Dim Area As RECT
    
    
    If Font Is Nothing Then Exit Function
    If Font.Handle = 0 Then Exit Function
    
    SelectObject mDC, Font.Handle
    If DrawText(mDC, StrPtr(Text), -1, Area, DT_CALCRECT Or DT_NOPREFIX) = 0 Then Exit Function
    getTextWidth = Area.Right * 1.05

End Function

Public Function getTextHeight(ByRef Font As clsGDIFont, ByVal Text As String) As Long

    Dim Area As RECT
    
    
    If Font Is Nothing Then Exit Function
    If Font.Handle = 0 Then Exit Function
    
    SelectObject mDC, Font.Handle
    If DrawText(mDC, StrPtr(Text), -1, Area, DT_CALCRECT Or DT_NOPREFIX) = 0 Then Exit Function
    getTextHeight = Area.Bottom

End Function


' Render a solid box shape
Public Function renderBox(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long) As Boolean

    Dim Area As RECT
    Dim hBrush As Long
    
    
    Dbg.addLog DM_GDI, "Rendering box " & x & ", " & y & ", " & Width & ", " & Height & ", " & Color
    
    With Area
        .Top = y
        .Bottom = y + Height
        .Left = x
        .Right = x + Width
    End With
    
    hBrush = CreateSolidBrush(Color)
    If hBrush = 0 Then Exit Function
    If FillRect(mDC, Area, hBrush) = 0 Then Exit Function
    
    DeleteObject hBrush
    renderBox = True

End Function


' Create an image from a Picture object
Public Function createFromPicture(ByRef Pic As Picture) As Boolean

    Dim tempBitmap As BitmapData
    Dim tempDC As Long
    
    
    If Pic Is Nothing Then Exit Function
    Dbg.addLog DM_GDI, "Creating image from picture object"
    
    Free

    ' Create a temporary DC with Picture handle in it
    tempDC = CreateCompatibleDC(0)
    If tempDC = 0 Then Exit Function
    SelectObject tempDC, Pic.Handle

    ' Get Picture's properties and create a new image to match
    If getObject(Pic.Handle, Len(tempBitmap), tempBitmap) = 0 Then Exit Function
    If Not createNew(tempBitmap.bmWidth, tempBitmap.bmHeight, tempBitmap.bmBitsPerPixel) Then Exit Function
    If BitBlt(mDC, 0, 0, mWidth, mHeight, tempDC, 0, 0, vbSrcCopy) = 0 Then Exit Function
    
    DeleteDC tempDC
    createFromPicture = True

End Function


' Render another image onto this one
Public Function renderImage(ByRef Image As clsGDIImage, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Method As renderImageMethods, ByVal stretchMode As stretchModes, Optional ByVal transColor As Long = -1) As Boolean

    If Image Is Nothing Then Exit Function
    Dbg.addLog DM_GDI, "Rendering image"
    
    SetStretchBltMode mDC, stretchMode
    
    ' Straight copy
    If Method = Render_Copy Then
        If Width = Image.Width And Height = Image.Height Then
            If BitBlt(mDC, x, y, Width, Height, Image.hDC, 0, 0, vbSrcCopy) = 0 Then Exit Function
        Else
            If StretchBlt(mDC, x, y, Width, Height, Image.hDC, 0, 0, Image.Width, Image.Height, vbSrcCopy) = 0 Then Exit Function
        End If
    
    ' One color will not be copied
    ElseIf Method = Render_Trans Then
        If TransBlt(mDC, x, y, Width, Height, Image.hDC, 0, 0, Image.Width, Image.Height, transColor) = 0 Then Exit Function
    
    ' Alpha blend
    ' Requires both images to be 32 bit
    ElseIf Method = Render_Alpha Then
        
        If GDIImageSelect(mHandle) <> 0 Then Exit Function
        If GDIAlphaBlit(Image.Handle, x, y) <> 0 Then Exit Function
    
    End If
    
    renderImage = True

End Function


' Render text to this image
Public Function renderText(ByRef Font As clsGDIFont, ByVal Text As String, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal textColor As Long, ByVal Alignment As textAlignment, Optional ByVal outlineSize As Long = 0, Optional ByVal outlineColor As Long = 0) As Boolean

    Dim textArea As RECT
    Dim renderArea As RECT
    Dim imgText As clsGDIImage
    Dim imgOutline As clsGDIImage
    Dim avgColor As Long
    Dim textPal(255) As rgbQuad
    
    Const baseFlags As Long = DT_NOPREFIX
    
    
    Dbg.addLog DM_GDI, "Rendering text " & Text & ", " & x & ", " & y & ", " & Width & ", " & Height & ", " & textColor & ", " & Alignment & ", " & outlineSize & ", " & outlineColor
    
    If Font Is Nothing Then Exit Function
    If Font.Handle = 0 Then Exit Function
    
    If textColor = 0 Then textColor = 16
    If outlineColor = 0 Then outlineColor = 8
    
    ' Setup font and text color
    If SelectObject(mDC, Font.Handle) = 0 Then Exit Function
    If SetTextColor(mDC, textColor) = &HFFFFFFFF Then Exit Function
    
    ' Get text size
    If DrawText(mDC, StrPtr(Text), -1, textArea, baseFlags Or DT_CALCRECT) = 0 Then Exit Function
    
    ' Align text
    If Alignment And DT_CENTER Then x = x + (Width / 2) - (textArea.Right / 2)
    If Alignment And DT_RIGHT Then x = (x + Width) - textArea.Right
    If Alignment And DT_VCENTER Then y = y + (Height / 2) - (textArea.Bottom / 2)
    If Alignment And DT_BOTTOM Then y = (y + Height) - textArea.Bottom
    
    ' Create textarea
    With renderArea
        .Left = x
        .Right = x + textArea.Right
        .Top = y
        .Bottom = y + textArea.Bottom
    End With
    
    ' Outline?
    If outlineSize Then

        Set imgText = New clsGDIImage
        Set imgOutline = New clsGDIImage

        ' Setup palette
        avgColor = colorAverage(textColor, outlineColor)
        colorToQuad textPal(0), 0
        colorToQuad textPal(1), avgColor
        colorToQuad textPal(2), textColor
        colorToQuad textPal(3), outlineColor

        ' Create temporary images
        If Not imgText.createNew(textArea.Right, textArea.Bottom, 32) Then Exit Function
        If Not imgOutline.createNew(textArea.Right + outlineSize * 2, textArea.Bottom + outlineSize * 2, 8, VarPtr(textPal(0)), 0) Then Exit Function

        ' Setup font and text color
        If SetTextColor(imgText.hDC, textColor) = &HFFFFFFFF Then Exit Function
        If SelectObject(imgText.hDC, Font.Handle) = 0 Then Exit Function

        ' Render the text then transfer it to the paletted bitmap to get the anti-aliasing
        If DrawText(imgText.hDC, StrPtr(Text), -1, textArea, baseFlags Or Alignment) = 0 Then Exit Function
        If Not imgOutline.renderImage(imgText, outlineSize, outlineSize, imgText.Width, imgText.Height, Render_Copy, STRETCH_DELETESCANS) Then Exit Function

        ' Render the outline
        If GDIImageSelect(imgOutline.Handle) Then Exit Function
        If GDIRenderOutline(outlineSize, 2, 0, 3) Then Exit Function

        ' Copy the image onto this image
        If Not Me.renderImage(imgOutline, renderArea.Left, renderArea.Top, imgOutline.Width, imgOutline.Height, Render_Trans, 0) Then Exit Function

        ' Adjust text render area for outline
        With renderArea
            .Top = .Top + outlineSize
            .Bottom = .Bottom + outlineSize
            .Left = .Left + outlineSize
            .Right = .Right + outlineSize
        End With
    End If

    ' Render text
    If DrawText(mDC, StrPtr(Text), -1, renderArea, baseFlags Or Alignment) = 0 Then Exit Function

    renderText = True
    
End Function


' Create a new empty bitmap
Public Function createNew(ByVal Width As Long, ByVal Height As Long, ByVal BPP As Long, Optional ByVal palPointer As Long = 0, Optional ByVal fillColor As Byte = 0) As Boolean

    Dim tempBitmap As bitmapInfo
    Dim tempHandle As Long
    Dim dataPointer As Long
    
    
    Free
    Dbg.addLog DM_GDI, "Creating new image"
    
    ' Set DIB header info
    With tempBitmap.biHeader
        .biSize = Len(tempBitmap.biHeader)
        .biWidth = Width
        .biHeight = Height
        .biBitCount = BPP
        .biPlanes = 1
        .biSizeImage = alignScan(Width, BPP) * Height
    End With
    
    ' Set palette
    If BPP = 8 And palPointer <> 0 Then CopyMemory ByVal VarPtr(tempBitmap.biColors(0)), ByVal palPointer, 256 * 4

    ' Create the DIB
    tempHandle = CreateDIBSection(mDC, tempBitmap, 0, dataPointer, 0, 0)
    If tempHandle = 0 Then Exit Function
    
    ' Fill with default color byte
    FillMemory ByVal dataPointer, tempBitmap.biHeader.biSizeImage, fillColor
    
    ' Select into this DC
    If SelectObject(mDC, tempHandle) = 0 Then Exit Function
    
    ' Set properties
    mWidth = Width
    mHeight = Height
    mBPP = BPP
    mHandle = tempHandle
    
    createNew = True

End Function


' Property exposers
Public Property Get Handle() As Long

    Handle = mHandle

End Property

Public Property Get Width() As Long

    Width = mWidth

End Property

Public Property Get Height() As Long

    Height = mHeight

End Property

Public Property Get BPP() As Long

    BPP = mBPP

End Property

Public Property Get hDC() As Long

    hDC = mDC

End Property


' Open from a file
Public Function openFrom(ByVal fileName As String) As Boolean

    Dim tempBitmap As BitmapData
    Dim tempHandle As Long
    Dim tempPic As StdPicture
    

    fileName = fileName
    Dbg.addLog DM_GDI, "Loading image from " & fileName
    
    ' Load as JPEG through StdPicture object
    If Right$(fileName, 4) = ".jpg" Then
        Set tempPic = LoadPicture(fileName)
        If tempPic Is Nothing Then Exit Function
        Me.createFromPicture tempPic
    
    ' Load as bitmap
    Else
        tempHandle = LoadImage(0, StrPtr(fileName), 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
        If tempHandle = 0 Then
            Dbg.addLog DM_GDI, "Unable to load file."
            Exit Function
        End If
    
        Free
    
        ' Get new bitmap's properties
        If getObject(tempHandle, Len(tempBitmap), tempBitmap) = 0 Then Exit Function
        mWidth = tempBitmap.bmWidth
        mHeight = tempBitmap.bmHeight
        mBPP = tempBitmap.bmBitsPerPixel
        mHandle = tempHandle
        
        If SelectObject(mDC, mHandle) = 0 Then Exit Function
        
    End If

    openFrom = True
    
End Function


' Return a Picture object of this image
Public Function getPicture() As Picture
    
    Dim Pic As OLEPictureType
    Dim IID_IDispatch As GUID
    Dim Value As Long
    
    
    Dbg.addLog DM_GDI, "Getting picture object from image"
    
    ' Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    ' Fill picture info
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = mHandle
        .hPal = 0
    End With
    
    ' Create the picture
    Value = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, getPicture)
    If Value <> 0 Then
        frmDialog.Display "OleCreatePictureIndirect return value " & Value, OkOnly Or Critical
        Set getPicture = Nothing
    End If
    
End Function


' Free used handles
Private Sub Free()

    DeleteObject mHandle

End Sub


' Class
Private Sub Class_Initialize()

    mDC = CreateCompatibleDC(0)
    If mDC = 0 Then Err.Raise -1, "clsGDIImage", "Cannot create compatible device context at initialization."

    SetBkMode mDC, 1

End Sub

Private Sub Class_Terminate()

    DeleteDC mDC
    Free

End Sub
