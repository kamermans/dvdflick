VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterList"
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
'   File purpose: Little class that generates a filter list for use
'                 with common dialogs.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Type filterItem
    Description As String
    Extension As String
End Type


Private Items() As filterItem
Private nItems As Long


Public Sub addType(ByVal Description As String, ByVal Extension As String)

    ReDim Preserve Items(nItems)
    nItems = nItems + 1
    
    With Items(nItems - 1)
        .Description = Description
        .Extension = Extension
    End With

End Sub


Public Sub Clear()

    Erase Items
    nItems = 0

End Sub


Public Sub Sort()

    Dim A As Long
    Dim Sort As Boolean
    Dim tempItem As filterItem
    
    
    Do
        Sort = False
        For A = 0 To nItems - 2
            If Items(A).Description > Items(A + 1).Description Then
                tempItem = Items(A)
                Items(A) = Items(A + 1)
                Items(A + 1) = tempItem
                Sort = True
            End If
        Next A
    Loop Until Sort = False

End Sub


Public Function fullString() As String

    Dim A As Long
    
    
    Me.Sort
    
    fullString = "All supported types|"
    For A = 0 To nItems - 1
        fullString = fullString & Items(A).Extension & ";"
    Next A
    fullString = Left$(fullString, Len(fullString) - 1) & "|"
    
    For A = 0 To nItems - 1
        fullString = fullString & Items(A).Description & " (" & Items(A).Extension & ")|" & Items(A).Extension & "|"
    Next A
    
    fullString = fullString & "All files (*.*)|*.*"

End Function
