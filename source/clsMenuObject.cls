VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuObject"
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
'   File purpose: Generic menu object.
'
Option Explicit
Option Compare Binary
Option Base 0


' Properties
Public Name As String
Public typeName As String
Public Parent As Boolean
Public thumbID As Long
Public subID As Long
Public audioID As Long

Public x As Long
Public y As Long
Public Width As Long
Public Height As Long

Public Font As String
Public fontSize As Long
Public textColor As Long
Public Bold As Boolean
Public Italic As Boolean
Public Underlined As Boolean

Public Outline As Long
Public outlineColor As Long

Public Text As String
Public Image As String
Public selectImage As String
Public xAlign As String
Public yAlign As String

Public selectMethod As Long

' Menu-private properties
Public imageRef As Boolean
Public imageLD As clsGDIImage
Public selectImageLD As clsGDIImage


' Return DrawText compatible alignment of this object
Public Function getAlignment() As textAlignment

    If xAlign = "left" Then getAlignment = getAlignment Or DT_LEFT
    If xAlign = "center" Then getAlignment = getAlignment Or DT_CENTER
    If xAlign = "right" Then getAlignment = getAlignment Or DT_RIGHT
    
    If yAlign = "top" Then getAlignment = getAlignment Or DT_TOP
    If yAlign = "center" Then getAlignment = getAlignment Or DT_VCENTER
    If yAlign = "bottom" Then getAlignment = getAlignment Or DT_BOTTOM

End Function


' Create and return a GDIFont object which describes object's font
Public Function getGDIFont() As clsGDIFont

    Set getGDIFont = New clsGDIFont

    With getGDIFont
        .Name = Font
        .Size = fontSize
        If Bold Then .Weight = FW_BOLD
        .Italic = Italic
        .Underlined = .Underlined
    End With

End Function


' Read this obejct's properties from a dictionary
Public Sub readFrom(ByRef Dict As Dictionary, ByVal myName As String)

    Name = myName
    typeName = Dict("Type")
    Parent = Dict("Parent")
    thumbID = Dict("thumbID")
    subID = Dict("subID")
    audioID = Dict("audioID")
    
    x = Dict("X")
    y = Dict("Y")
    Width = Dict("Width")
    Height = Dict("Height")
    
    Font = Dict("Font")
    fontSize = Dict("fontSize")
    textColor = stringToColor(Dict("textColor"))
    Bold = Dict("Bold")
    Italic = Dict("Italic")
    Underlined = Dict("Underlined")
    
    Outline = Dict("Outline")
    outlineColor = stringToColor(Dict("outlineColor"))
    
    Text = Dict("Text")
    Image = Dict("Image")
    selectImage = Dict("selectImage")
    xAlign = Dict("xAlign")
    yAlign = Dict("yAlign")

    selectMethod = getSelectMethod(Dict("selectMethod"))

End Sub


' Copy this object's properties from another
Public Sub copyFrom(ByRef Obj As clsMenuObject)

    With Obj
        Name = .Name
        typeName = .typeName
        Parent = .Parent
        thumbID = .thumbID
        subID = .subID
        audioID = .audioID
    
        x = .x
        y = .y
        Width = .Width
        Height = .Height
        
        Font = .Font
        fontSize = .fontSize
        textColor = .textColor
        Bold = .Bold
        Italic = .Italic
        Underlined = .Underlined
        
        Outline = .Outline
        outlineColor = .outlineColor
        
        Text = .Text
        Image = .Image
        selectImage = .selectImage
        xAlign = .xAlign
        yAlign = .yAlign
        
        selectMethod = .selectMethod
    End With

End Sub


' Return a proper selection method value from a string
Private Function getSelectMethod(ByVal Method As String) As Long

    Select Case LCase$(Method)
        Case "outline"
            getSelectMethod = selMethod_Outline
        Case "image"
            getSelectMethod = selMethod_Image
        Case "text"
            getSelectMethod = selMethod_Text
        Case Else
            getSelectMethod = selMethod_None
    End Select
    
End Function
