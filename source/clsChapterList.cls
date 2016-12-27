VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChapterList"
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
'   File purpose: Chapter list creation and output.
'
Option Explicit
Option Compare Binary
Option Base 0


' Entry list
Private Entries() As Single
Private nEntries As Long

' Duration of the list in time
Private mDuration As Single


' Duration property
Public Property Get Duration() As Single

    Duration = mDuration

End Property

Public Property Let Duration(ByVal newValue As Single)

    mDuration = newValue

End Property


' Return number of entries in the list
Public Function Count() As Long

    Count = nEntries

End Function


' Return a string of this chapter list ready to be used by DVD Author
Public Function getAuthorString() As String

    Dim A As Long
    
    
    ' Add start and end chapter marks and create the string
    getAuthorString = "0:00"
    For A = 0 To nEntries - 1
        getAuthorString = getAuthorString & ", " & timeToString(Entries(A))
    Next A

End Function


' Return a time in seconds as a minutes:seconds.ms display
Private Function timeToString(ByVal Duration As Single) As String

    Dim Hours As Long
    Dim Minutes As Long
    Dim Seconds As Long
    Dim mS As Long
    
    
    Do
        If Duration >= 60 Then
            Minutes = Minutes + 1
            Duration = Duration - 60
        ElseIf Duration >= 1 Then
            Seconds = Seconds + 1
            Duration = Duration - 1
        Else
            mS = Fix(Duration * 100)
            Exit Do
        End If
    Loop

    timeToString = Minutes & ":"
    If Seconds > 9 Then
        timeToString = timeToString & Seconds & "."
    Else
        timeToString = timeToString & "0" & Seconds & "."
    End If
    If mS > 9 Then
        timeToString = timeToString & mS
    Else
        timeToString = timeToString & "0" & mS
    End If
    
End Function


' Add a chapter every interval (0 excluded)
Public Sub addByInterval(ByVal Interval As Single)

    Dim A As Single
    
    
    For A = Interval To mDuration Step Interval
        addEntry A
    Next A

End Sub


' Add a set number of chapters (0 and end excluded)
Public Sub addByCount(ByVal Count As Long)

    Dim A As Single
    Dim Interval As Single
    
    
    Interval = mDuration / Count
    For A = Interval To mDuration Step Interval
        addEntry A
    Next A

End Sub


' Bubble sort entry list
Public Sub Sort()

    Dim A As Long
    Dim tempEntry As Single
    Dim Sorted As Boolean
    
    
    Do
        Sorted = False
    
        For A = 0 To nEntries - 2
            
            If Entries(A + 1) < Entries(A) Then
                tempEntry = Entries(A)
                Entries(A) = Entries(A + 1)
                Entries(A + 1) = tempEntry
                
                Sorted = True
            End If
            
        Next A
    Loop Until Sorted = False

End Sub


' Remove any duplicates from the list
' Only works if the list has been sorted beforehand
Public Sub removeDuplicates()

    Dim A As Long
    Dim Removed As Boolean
    
    
    Do
        Removed = False
        For A = 0 To nEntries - 2
            If Entries(A) = Entries(A + 1) Then
                removeEntry A
                Removed = True
                Exit For
            End If
        Next A
    Loop Until Removed = False

End Sub


' Remove an entry from the list
Private Sub removeEntry(ByVal Index As Long)

    Dim A As Long
    
    
    For A = Index To nEntries - 2
        Entries(A) = Entries(A + 1)
    Next A
    
    nEntries = nEntries - 1
    If nEntries > 0 Then
        ReDim Preserve Entries(nEntries - 1)
    Else
        Erase Entries
    End If

End Sub


' Add a single entry
Public Sub addEntry(ByVal timeIndex As Single)

    ReDim Preserve Entries(nEntries)
    nEntries = nEntries + 1
    
    Entries(nEntries - 1) = timeIndex

End Sub
