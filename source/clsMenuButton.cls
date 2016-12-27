VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuButton"
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
'   File purpose: Menu DVD button object.
'
Option Explicit
Option Compare Binary
Option Base 0


' Properties
Public Action As String

Public Top As Long
Public Bottom As Long
Public Left As Long
Public Right As Long

' Whether a surface has already been set or not
Private noSurface As Boolean


' Add the surface area of a menu object to this one's area
Public Sub addSurface(ByRef Obj As clsMenuObject)

    Dim nTop As Long, nBottom As Long
    Dim nLeft As Long, nRight As Long
    
    
    nTop = Obj.y - Obj.Outline
    nBottom = Obj.y + Obj.Height + Obj.Outline
    nLeft = Obj.x - Obj.Outline
    nRight = Obj.x + Obj.Width + Obj.Outline

    If noSurface Then
        Left = nLeft
        Right = nRight
        Top = nTop
        Bottom = nBottom
        noSurface = False
    
    Else
        If nLeft < Left Then Left = nLeft
        If nRight > Right Then Right = nRight
        If nTop < Top Then Top = nTop
        If nBottom > Bottom Then Bottom = nBottom
        
    End If
    
End Sub


' Class
Private Sub Class_Initialize()

    noSurface = True

End Sub
