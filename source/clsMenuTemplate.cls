VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuTemplate"
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
'   File purpose: Menu template class.
'
Option Explicit
Option Compare Binary
Option Base 0


' Properties
Public Path As String
Public transColor As Long

Public Name As String
Public Author As String
Public Description As String
Public Copyrights As String

' Menu templates
Public Templates As New Dictionary


' Read this template from a file
Public Function openFrom(ByVal fileName As String) As Boolean

    Dim A As Long
    Dim Dict As Dictionary
    Dim menuDict As Dictionary
    Dim Conf As clsConfiguration
    Dim Menu As clsMenu

    
    ' Get dictionary from configuration file
    Set Conf = New clsConfiguration
    Conf.LoadConfiguration fileName
    Set Dict = Conf.Root(True)
    
    For A = 0 To Dict.Count - 1
        
        ' Properties
        If Dict.Keys(A) = "transColor" Then
            transColor = stringToColor(Dict.Items(A))
        ElseIf Dict.Keys(A) = "Name" Then
            Name = Dict.Items(A)
        ElseIf Dict.Keys(A) = "Author" Then
            Author = Dict.Items(A)
        ElseIf Dict.Keys(A) = "Description" Then
            Description = Dict.Items(A)
        ElseIf Dict.Keys(A) = "Copyrights" Then
            Copyrights = Dict.Items(A)
            
        ' Menu
        Else
            Set Menu = New clsMenu
            Set menuDict = Dict.Items(A)
            Menu.readFrom menuDict
            Templates.Add Dict.Keys(A), Menu
            
        End If
    Next A
    
    Path = Left$(fileName, InStrRev(fileName, "\"))

End Function
