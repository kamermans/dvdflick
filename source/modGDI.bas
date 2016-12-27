Attribute VB_Name = "modGDI"
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
'   File purpose: Windows GDI definitions and utility functions.
'
Option Explicit
Option Compare Binary
Option Base 0


' RGB bitmap palette
Public Type rgbQuad
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

' Bitmap Info Header
Public Type bitmapInfoHeader
     biSize As Long
     biWidth As Long
     biHeight As Long
     biPlanes As Integer
     biBitCount As Integer
     biCompression As Long
     biSizeImage As Long
     biXPelsPerMeter As Long
     biYPelsPerMeter As Long
     biClrUsed As Long
     biClrImportant As Long
End Type

Public Type bitmapInfo
    biHeader As bitmapInfoHeader
    biColors(0 To 255) As rgbQuad
End Type


' Font character sets
Public Enum fontCharSets
    ANSI_CHARSET = 0
    DEFAULT_CHARSET = 1
    SYMBOL_CHARSET = 2
    MAC_CHARSET = 77
    SHIFTJIS_CHARSET = 128
    HANGEUL_CHARSET = 129
    JOHAB_CHARSET = 130
    GB2312_CHARSET = 134
    CHINESEBIG5_CHARSET = 136
    GREEK_CHARSET = 161
    TURKISH_CHARSET = 162
    VIETNAMESE_CHARSET = 163
    HEBREW_CHARSET = 177
    ARABIC_CHARSET = 178
    BALTIC_CHARSET = 186
    RUSSIAN_CHARSET = 204
    THAI_CHARSET = 222
    EASTEUROPE_CHARSET = 238
    OEM_CHARSET = 255
End Enum

' Font weights
Public Enum fontWeights
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
End Enum

' Font qualities
Public Enum fontQualities
    DEFAULT_QUALITY = 0
    DRAFT_QUALITY = 1
    PROOF_QUALITY = 2
    NONANTIALIASED_QUALITY = 3
    ANTIALIASED_QUALITY = 4
    CLEARTYPE_QUALITY = 5
End Enum

' DrawText alignment
Public Enum textAlignment
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

' Stretch modes
Public Enum stretchModes
    STRETCH_NEAREST = 0
    STRETCH_ANDSCANS = 1
    STRETCH_ORSCANS = 2
    STRETCH_DELETESCANS = 3
    STRETCH_HALFTONE = 4
End Enum

' Methods that GDI images can be rendered with
Public Enum renderImageMethods
    Render_Copy = 0
    Render_Trans
    Render_Alpha
End Enum


' DWord-alignment for bitmap data
Public Function alignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    
    alignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
    
End Function


' Return the average color of 2 other colors
Public Function colorAverage(ByVal Color1 As Long, ByVal Color2 As Long) As Long

    Dim Col1 As rgbQuad
    Dim Col2 As rgbQuad
    Dim valR As Long, valG As Long, valB As Long
    
    
    colorToQuad Col1, Color1
    colorToQuad Col2, Color2
    
    valR = (CLng(Col1.rgbRed) + CLng(Col2.rgbRed)) / 2
    valG = (CLng(Col1.rgbGreen) + CLng(Col2.rgbGreen)) / 2
    valB = (CLng(Col1.rgbBlue) + CLng(Col2.rgbBlue)) / 2
    
    colorAverage = RGB(valR, valG, valB)

End Function


' Return an RGB color value from a quad
Public Function quadToColor(ByRef Quad As rgbQuad) As Long

    quadToColor = RGB(Quad.rgbRed, Quad.rgbGreen, Quad.rgbBlue)

End Function


' Set a quad to an RGB color value
Public Sub colorToQuad(ByRef Quad As rgbQuad, ByVal Color As Long)
    
    Quad.rgbRed = Color Mod 256
    Quad.rgbGreen = ((Color And &HFF00) / 256&) Mod 256&
    Quad.rgbBlue = (Color And &HFF0000) / 65536
  
End Sub
