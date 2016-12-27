Attribute VB_Name = "modRenderMenu"
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
'   File purpose: Menu template rendering and definitions.
'
Option Explicit
Option Compare Binary
Option Base 0


' Selection methods
Public Enum selectMethods
    selMethod_None = 0
    
    selMethod_Text
    selMethod_Image
    selMethod_Outline
End Enum


' Menu image dimensions
Public Const MENU_WIDTH As Long = 768
Public Const MENU_HEIGHT As Long = 576
Public Const MENU_ASPECT As Single = 4 / 3


Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceW" (ByVal ptrFilename As Long) As Long


' Register fonts found in template directories only for this Windows session
Public Sub registerTemplateFonts()

    Dim topFolder As Folder, tempFolder As Folder
    Dim myFile As File
    Dim fileName As String
    
    
    Set topFolder = FS.GetFolder(APP_PATH & "templates")
    For Each tempFolder In topFolder.SubFolders
        For Each myFile In tempFolder.Files
        
            If LCase$(FS.GetExtensionName(myFile.Name)) = "ttf" Then
                appLog.Add "Adding font resource " & myFile.Name, 1
                fileName = APP_PATH & "templates\" & tempFolder.Name & "\" & myFile.Name
                AddFontResource StrPtr(fileName)
            End If
            
        Next myFile
    Next tempFolder

End Sub


' Convert string "xxx,xxx,xxx" to a color
Public Function stringToColor(ByVal colorString) As Long

    Dim Temp() As String
    
    
    Temp = Split(colorString, ",")
    If UBound(Temp) <> 2 Then Exit Function
    
    stringToColor = RGB(CLng(Trim$(Temp(0))), CLng(Trim$(Temp(1))), CLng(Trim$(Temp(2))))

End Function
