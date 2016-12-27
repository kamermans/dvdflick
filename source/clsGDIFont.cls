VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDIFont"
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
'   File purpose: GDI font wrapper class.
'
Option Explicit
Option Compare Binary
Option Base 0


' Font structure
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(31) As Byte
End Type


' Font output precision
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_DEVICE_PRECIS = 5

' Font clipping precision
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2

' Font pitch and family
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2


' Properties
Private mHandle As Long
Private mName As String
Private mSize As Long
Private mWeight As fontWeights
Private mItalic As Boolean
Private mUnderlined As Boolean
Private mQuality As fontQualities

' Require updating when handle is requested
Private requireUpdate As Boolean


' External calls
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


' Update\create the current font
Private Function updateFont() As Boolean

    Dim fontDesc As LOGFONT
    
    
    ' Fill font structure
    With fontDesc
        .lfCharSet = 0
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        stringToByteArray mName, .lfFaceName
        .lfHeight = mSize
        .lfItalic = mItalic
        .lfOutPrecision = OUT_DEFAULT_PRECIS
        .lfQuality = mQuality
        .lfUnderline = mUnderlined
        .lfWeight = mWeight
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
    End With
    
    ' Create the new font
    DeleteObject mHandle
    mHandle = CreateFontIndirect(fontDesc)
    If mHandle = 0 Then Err.Raise -1, "clsGDIFont", "Cannot update font."
    
    requireUpdate = False
    updateFont = True

End Function


' Font handle
Public Property Get Handle() As Long

    If requireUpdate Then updateFont
    Handle = mHandle

End Property


' Font face name
Public Property Get Name() As String

    Name = mName

End Property

Public Property Let Name(ByVal newValue As String)

    mName = newValue
    requireUpdate = True

End Property


' Font size
Public Property Get Size() As Long

    Size = mSize

End Property

Public Property Let Size(ByVal newValue As Long)

    mSize = newValue
    requireUpdate = True

End Property


' Font weight
Public Property Get Weight() As fontWeights

    Weight = mWeight

End Property

Public Property Let Weight(ByVal newValue As fontWeights)

    mWeight = newValue
    requireUpdate = True

End Property


' Italicized
Public Property Get Italic() As Boolean

    Italic = mItalic

End Property

Public Property Let Italic(ByVal newValue As Boolean)

    mItalic = newValue
    requireUpdate = True

End Property


' Underlined
Public Property Get Underlined() As Boolean

    Underlined = mUnderlined

End Property

Public Property Let Underlined(ByVal newValue As Boolean)

    mUnderlined = newValue
    requireUpdate = True

End Property


' Quality
Public Property Get Quality() As fontQualities

    Quality = mQuality

End Property

Public Property Let Quality(ByVal newValue As fontQualities)

    mQuality = newValue
    requireUpdate = True

End Property


' Class
Private Sub Class_Initialize()

    mName = "Arial"
    mSize = 16
    mQuality = ANTIALIASED_QUALITY
    requireUpdate = True

End Sub

Private Sub Class_Terminate()

    DeleteObject mHandle

End Sub
