VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubBlock"
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
'   File purpose: A subtitle block.
'
Option Explicit
Option Compare Binary
Option Base 0


Public Text As String

Public startTime As Single
Public endTime As Single

Public startFrame As Long
Public endFrame As Long

Public oX As Long
Public oY As Long


' Strip HTML style tags
Public Sub stripHTML()

    Dim A As Long
    Dim Char As String
    Dim inTag As Boolean
    Dim newText As String
    
    
    ' <i>How now brow cow</i>
    For A = 1 To Len(Text)
        Char = Mid$(Text, A, 1)
        
        If Char = "<" Then inTag = True
        If Char = ">" And inTag = True Then
            inTag = False
            Char = vbNullString
        End If
        
        If inTag = False Then newText = newText & Char
    Next A
    
    Text = newText

End Sub


' Strip RTF style tags
Public Sub stripRTF()

    Dim A As Long
    Dim Char As String
    Dim inTag As Boolean
    Dim newText As String
    

    ' {\3c&H572C1A&} {\bord2}
    For A = 1 To Len(Text)
        Char = Mid$(Text, A, 1)
        
        If Mid$(Text, A, 2) = "{\" Then inTag = True
        If Char = "}" And inTag = True Then
            inTag = False
            Char = vbNullString
        End If
        
        If inTag = False Then newText = newText & Char
    Next A
    
    Text = newText

End Sub
