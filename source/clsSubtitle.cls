VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubtitle"
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
'   File purpose: Subtitle for titles
'
Option Explicit
Option Compare Binary
Option Base 0


Public fileName As String

' Subtitle properties
Public Font As String
Public fontSize As Long
Public fontBold As Byte
Public codePage As Long

Public colorText As Long
Public colorOutline As Long
Public ColorBack As Long
Public Outline As Long
Public antiAlias As Byte
Public transBack As Byte

Public marginTop As Long
Public marginBottom As Long
Public marginLeft As Long
Public marginRight As Long
Public Alignment As enumSubAlignment

Public FPS As Single
Public displayDefault As Byte
Public Language As Long
Public fileFormat As String
Public frameBased As Byte


' Write to dictionary
Public Function Serialize() As Dictionary

    Set Serialize = New Dictionary
    
    Serialize.Add "fileName", fileName
    
    Serialize.Add "Font", Font
    Serialize.Add "fontSize", fontSize
    Serialize.Add "fontBold", fontBold
    Serialize.Add "codePage", codePage
    
    Serialize.Add "colorText", colorText
    Serialize.Add "colorOutline", colorOutline
    Serialize.Add "colorBack", ColorBack
    Serialize.Add "Outline", Outline
    Serialize.Add "antiAlias", antiAlias
    Serialize.Add "transBack", transBack
    
    Serialize.Add "marginTop", marginTop
    Serialize.Add "marginBottom", marginBottom
    Serialize.Add "marginLeft", marginLeft
    Serialize.Add "marginRight", marginRight
    Serialize.Add "Alignment", Alignment
    
    Serialize.Add "FPS", FPS
    Serialize.Add "displayDefault", displayDefault
    Serialize.Add "Language", Language
    Serialize.Add "fileFormat", fileFormat
    Serialize.Add "frameBased", frameBased

End Function


' Read from dictionary
Public Function unSerialize(ByRef myDict As Dictionary) As Boolean

    unSerialize = False
    
    fileName = myDict("fileName")
    
    Font = myDict("Font")
    fontSize = myDict("fontSize")
    fontBold = myDict("fontBold")
    codePage = myDict("codePage")
    
    colorText = myDict("colorText")
    colorOutline = myDict("colorOutline")
    ColorBack = myDict("colorBack")
    Outline = myDict("Outline")
    antiAlias = myDict("antiAlias")
    transBack = myDict("transBack")
    
    marginTop = myDict("marginTop")
    marginBottom = myDict("marginBottom")
    marginLeft = myDict("marginLeft")
    marginRight = myDict("marginRight")
    Alignment = myDict("Alignment")

    FPS = myDict("FPS")
    displayDefault = myDict("displayDefault")
    Language = myDict("Language")
    fileFormat = myDict("fileFormat")
    frameBased = myDict("frameBased")
    
    ' Read and check if source file still exists
    If Not FS.FileExists(fileName) Then
        frmDialog.Display "Unable to find the subtitle file " & fileName & ".", Exclamation Or OkOnly
        Exit Function
    End If
    
    unSerialize = True
    
End Function


Private Sub Reset()

    Font = Config.ReadSetting("subFont", Default_SubFont)
    fontSize = Config.ReadSetting("subFontSize", Default_SubFontSize)
    fontBold = Config.ReadSetting("subFontBold", default_SubFontBold)
    codePage = Config.ReadSetting("subCodePage", default_SubCodePage)
    
    colorText = Config.ReadSetting("subColorText", Default_SubColorText)
    colorOutline = Config.ReadSetting("subColorOutline", Default_SubColorOutline)
    ColorBack = Config.ReadSetting("subColorBack", Default_SubColorBack)
    Outline = Config.ReadSetting("subOutline", Default_SubOutline)
    antiAlias = Config.ReadSetting("subAA", Default_SubAA)
    transBack = Config.ReadSetting("subTransBack", Default_SubTransBack)
    
    marginTop = Config.ReadSetting("subMarginTop", Default_SubMarginTop)
    marginBottom = Config.ReadSetting("subMarginBottom", Default_SubMarginBottom)
    marginLeft = Config.ReadSetting("subMarginLeft", Default_SubMarginLeft)
    marginRight = Config.ReadSetting("subMarginRight", Default_SubMarginRight)
    Alignment = Config.ReadSetting("subAlignment", Default_SubAlignment)
    
    Language = Config.ReadSetting("subLanguage", Default_SubLanguage)
    FPS = Default_SubFramerate
    displayDefault = 0
    fileFormat = "Unknown"
    frameBased = 0

End Sub


Private Sub Class_Initialize()

    Reset

End Sub

