VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectArray"
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
'   File purpose: Basic object array management functionality. Not all of this
'                 is available in a Collection.
'
Option Explicit
Option Compare Binary
Option Base 0


Private Objects() As Variant
Private mObjectCount As Long


' Clear the array
Public Sub Clear()

    Erase Objects
    mObjectCount = 0

End Sub


' Return the number of objects in the array
Public Property Get Count() As Long

    Count = mObjectCount

End Property


' Move an object forward in the array
Public Sub moveForward(ByVal Index As Long)

    Dim Temp As Object
    

    If Index = mObjectCount - 1 Then Exit Sub

    Set Temp = Objects(Index + 1)
    Set Objects(Index + 1) = Objects(Index)
    Set Objects(Index) = Temp

End Sub


' Move an object backward in the array
Public Sub moveBackward(ByVal Index As Long)

    Dim Temp As Object
    

    If Index = 0 Then Exit Sub

    Set Temp = Objects(Index - 1)
    Set Objects(Index - 1) = Objects(Index)
    Set Objects(Index) = Temp

End Sub


' Add an object to the array
Public Sub Add(ByRef Obj As Object)

    ReDim Preserve Objects(mObjectCount)
    mObjectCount = mObjectCount + 1
    
    Set Objects(mObjectCount - 1) = Obj

End Sub


' Return an object from the array
Public Function Item(ByVal Index As Long) As Object

    Set Item = Objects(Index)

End Function


' Remove an object from the array
Public Sub Remove(ByVal Index)

    Dim A As Long
    
    
    ' Move objects
    For A = Index To mObjectCount - 2
        Set Objects(A) = Objects(A + 1)
    Next A
    
    ' Decrease array size
    mObjectCount = mObjectCount - 1
    If mObjectCount > 0 Then
        ReDim Preserve Objects(mObjectCount)
    Else
        Erase Objects
    End If

End Sub
