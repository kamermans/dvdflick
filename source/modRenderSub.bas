Attribute VB_Name = "modRenderSub"
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
'   File purpose: Subtitle renderer.
'
Option Explicit
Option Compare Binary
Option Base 0


' 4 color subtitle palette
Private Const PAL_BACK = 0
Private Const PAL_TEXT = 1
Private Const PAL_AA = 2
Private Const PAL_OUTLINE = 3


' Subtitle appearance
Public marginTop As Long, marginBottom As Long
Public marginLeft As Long, marginRight As Long
Public outlineSize As Long
Public antiAlias As Boolean, transparentBack As Boolean
Public fontFace As String, fontSize As Long, fontBold As Boolean
Public colorBackground As Long, colorText As Long, colorOutline As Long
Private cSub As clsSubtitle

' Bitmap data
Private Palette(255) As rgbQuad
Private subFont As clsGDIFont


' API calls
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nDestLeft As Long, ByVal nDestTop As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal hSrcDC As Long, ByVal nSrcLeft As Long, ByVal nSrcTop As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function TransBlt Lib "msimg32" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal stretchMode As Long) As Long

Private Declare Function GDIImageSelect Lib "dvdflick" (ByVal hBitmap As Long) As Long
Private Declare Function GDIRenderOutline Lib "dvdflick" (ByVal outlineSize As Long, ByVal baseColor As Long, ByVal backColor As Long, ByVal outlineColor As Long) As Long


' Return the background color in use
Public Function backgroundColor() As Long

    backgroundColor = quadToColor(Palette(PAL_BACK))

End Function


' Initialise rendering
Public Sub initRender(ByRef mySub As clsSubtitle)

    Dim refColor As Long
    
    
    Dbg.addLog DM_Subtitles, "Initializing subtitle render"
    
    ' Set subtitle appearance properties
    fontFace = mySub.Font
    fontSize = mySub.fontSize
    fontBold = mySub.fontBold
    
    colorText = mySub.colorText
    colorBackground = mySub.ColorBack
    colorOutline = mySub.colorOutline
    antiAlias = mySub.antiAlias
    outlineSize = mySub.Outline
    
    marginTop = mySub.marginTop
    marginBottom = mySub.marginBottom
    marginLeft = mySub.marginLeft
    marginRight = mySub.marginRight
    
    Set cSub = mySub
    

    ' Create palette
    If colorText = 0 Then colorText = 12
    If colorOutline = 0 Then colorOutline = 6
    colorToQuad Palette(PAL_BACK), colorBackground
    colorToQuad Palette(PAL_TEXT), colorText
    colorToQuad Palette(PAL_OUTLINE), colorOutline
    
    ' Calculate anti-alias color (average between either text and background or text and outline)
    If outlineSize > 0 Then
        refColor = PAL_OUTLINE
    Else
        refColor = PAL_BACK
    End If
    colorToQuad Palette(PAL_AA), colorAverage(quadToColor(Palette(PAL_TEXT)), quadToColor(Palette(refColor)))
    
    
    ' Create font
    Set subFont = New clsGDIFont
    
    subFont.Size = fontSize
    subFont.Name = fontFace
    
    If antiAlias Then
        subFont.Quality = ANTIALIASED_QUALITY
    Else
        subFont.Quality = NONANTIALIASED_QUALITY
    End If
    
    If fontBold = True Then subFont.Weight = 700
    
    
    ' Halve sizes for half resolution
    If Project.halfRes = 1 Then
        subFont.Size = subFont.Size / 2
        If outlineSize > 1 Then outlineSize = outlineSize / 2
        marginLeft = marginLeft / 2
        marginRight = marginRight / 2
    End If
    
End Sub


' Cleanup
Public Sub endRender()
    
    Set subFont = Nothing
    
End Sub


' Render all subtitles
Public Sub renderSubFile(ByRef subFile As clsSubFile, ByRef myTitle As clsTitle, ByRef mySub As clsSubtitle, ByVal preFix As String, Optional ByRef prgBar As ComctlLib.ProgressBar = Nothing)

    Dim A As Long
    Dim Block As clsSubBlock
    
    
    Dbg.addLog DM_Subtitles, "Rendering entire subtitle file"
    
    initRender mySub
    If Not (prgBar Is Nothing) Then prgBar.Max = subFile.blockCount

    ' Render subtitles to bitmaps
    For A = 0 To subFile.blockCount - 1
        Set Block = subFile.getBlock(A)
        
        SavePicture renderBlock(Block, myTitle).getPicture, preFix & A & ".bmp"
        Dbg.addLog DM_Subtitles, "Rendered block " & A
        
        If Not (prgBar Is Nothing) Then prgBar.Value = A
        
        DoEvents
    Next A
    
    endRender

End Sub


' Render a single subtitle block
Public Function renderBlock(ByRef Block As clsSubBlock, ByRef myTitle As clsTitle) As clsGDIImage

    Dim A As Long
    Dim refColor As Long
    Dim imgRender As clsGDIImage
    Dim imgSub As clsGDIImage
    Dim Width As Long, Height As Long
    Dim bmpData() As Byte, bmpRender() As Byte
    Dim x As Long, y As Long
    Dim Encode As Dictionary
    Dim Scaler As Single
        

    Set Encode = myTitle.encodeInfo

    ' Create render images
    Set imgRender = New clsGDIImage
    Set imgSub = New clsGDIImage

    ' Calculate width and height of subtitle block
    Width = imgRender.getTextWidth(subFont, Block.Text) + (outlineSize * 2)
    Height = imgRender.getTextHeight(subFont, Block.Text) + (outlineSize * 2)

    ' Round off width and height to be multiples of 2
    Width = Width + (Width Mod 2)
    Height = Height + (Height Mod 2)

    ' Create the render bitmap
    If Not imgRender.createNew(Width, Height, 32) Then Exit Function

    ' Fill the render bitmap with either the background or outline color for proper anti-aliasing
    If outlineSize Then refColor = PAL_OUTLINE Else refColor = PAL_BACK
    If Not imgRender.colorFill(refColor) Then Exit Function
    
    ' Now really render it
    If Not imgRender.renderText(subFont, Block.Text, 0, 0, imgRender.Width, imgRender.Height, quadToColor(Palette(PAL_TEXT)), DT_CENTER Or DT_VCENTER) Then Exit Function

    ' Store aligned block coordinates
    Block.oX = getXAlign(cSub, Width, Encode("Width"))
    Block.oY = getYAlign(cSub, Height, Encode("Height"))

    ' Create a new empty paletted bitmap
    If Not imgSub.createNew(Width, Height, 8, VarPtr(Palette(0))) Then Exit Function
    imgSub.setPalette VarPtr(Palette(0))

    ' Copy 32 bit subtitle to paletted bitmap to keep antialiasing
    If Not imgSub.renderImage(imgRender, 0, 0, imgSub.Width, imgSub.Height, Render_Copy, STRETCH_DELETESCANS) Then Exit Function

    ' Swap the outline color in the bitmap to the background color and render outline
    If Not imgSub.colorReplace(PAL_OUTLINE, PAL_BACK) Then Exit Function
    If Not imgSub.renderOutline(outlineSize, PAL_TEXT, PAL_BACK, PAL_OUTLINE) Then Exit Function
    
    ' Crop if necessary
    If Width > Encode("Width") Then imgSub.Crop Encode("Width"), imgSub.Height
    If Height > Encode("Height") Then imgSub.Crop imgSub.Width, Encode("Height")

    ' Return picture
    Set renderBlock = imgSub
    
End Function


' Return the width of a block
Public Function getBlockWidth(ByRef Block As clsSubBlock) As Long

    Dim imgRender As clsGDIImage
    
    
    Set imgRender = New clsGDIImage

    getBlockWidth = imgRender.getTextWidth(subFont, Block.Text) + (outlineSize * 2)
    getBlockWidth = getBlockWidth + (getBlockWidth Mod 2)
    
    Set imgRender = Nothing

End Function


' Render a subtitle preview to a picturebox
Public Sub renderPreview(ByRef mySub As clsSubtitle, ByRef myFile As clsSubFile, ByRef myTitle As clsTitle, ByVal mScale As Single, ByRef picBox As PictureBox, Optional ByRef picBack As clsGDIImage)

    Dim Pic As clsGDIImage
    Dim x As Long, y As Long
    Dim Encode As Dictionary

    
    Dbg.addLog DM_Subtitles, "Rendering subtitle preview"
    
    Set Encode = myTitle.encodeInfo
    
    ' Render the subtitle block
    modRenderSub.initRender mySub
    Set Pic = modRenderSub.renderBlock(myFile.getBlock(myFile.blockCount / 2), myTitle)
    If Pic Is Nothing Then Exit Sub

    x = getXAlign(mySub, Pic.Width, Encode("Height") * (4 / 3))
    y = getYAlign(mySub, Pic.Height, Encode("Height"))

    ' Set appropriate scaling mode
    If mScale <> 1 Then
        SetStretchBltMode picBox.hDC, STRETCH_HALFTONE
    Else
        SetStretchBltMode picBox.hDC, STRETCH_DELETESCANS
    End If

    ' Blit the picture
    picBox.Cls
    If Not picBack Is Nothing Then StretchBlt picBox.hDC, 0, 0, picBox.Width, picBox.Height, picBack.hDC, 0, 0, picBack.Width, picBack.Height, vbSrcCopy
    If mySub.transBack Then
        TransBlt picBox.hDC, x * mScale, y * mScale, Pic.Width * mScale, Pic.Height * mScale, Pic.hDC, 0, 0, Pic.Width, Pic.Height, quadToColor(Palette(PAL_BACK))
    Else
        StretchBlt picBox.hDC, x * mScale, y * mScale, Pic.Width * mScale, Pic.Height * mScale, Pic.hDC, 0, 0, Pic.Width, Pic.Height, vbSrcCopy
    End If

    picBox.Refresh

End Sub


' Block X alignment
Private Function getXAlign(ByRef mySub As clsSubtitle, ByVal Width As Long, ByVal targetWidth As Long) As Long

    ' X alignment
    If mySub.Alignment = SA_BottomCenter Or mySub.Alignment = SA_TopCenter Then
        getXAlign = (targetWidth / 2) - (Width / 2)
    End If
    If mySub.Alignment = SA_BottomLeft Or mySub.Alignment = SA_CenterLeft Or mySub.Alignment = SA_TopLeft Then
        getXAlign = mySub.marginLeft
    End If
    If mySub.Alignment = SA_BottomRight Or mySub.Alignment = SA_CenterRight Or mySub.Alignment = SA_TopRight Then
        getXAlign = targetWidth - mySub.marginRight - Width
    End If
    
End Function


' Block Y alignment
Private Function getYAlign(ByRef mySub As clsSubtitle, ByVal Height As Long, ByVal targetHeight) As Long

    If mySub.Alignment = SA_CenterLeft Or mySub.Alignment = SA_CenterRight Then
        getYAlign = (targetHeight / 2) - (Height / 2)
    End If
    If mySub.Alignment = SA_BottomCenter Or mySub.Alignment = SA_BottomLeft Or mySub.Alignment = SA_BottomRight Then
        getYAlign = targetHeight - mySub.marginBottom - Height
    End If
    If mySub.Alignment = SA_TopCenter Or mySub.Alignment = SA_TopLeft Or mySub.Alignment = SA_TopRight Then
        getYAlign = mySub.marginTop
    End If

End Function
