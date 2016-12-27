VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParameters"
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
'   File purpose: Tiny class to parse out commandline parameters.
'
Option Explicit
Option Compare Binary
Option Base 0


Public Switches As Dictionary


Private Sub Class_Initialize()

    Dim A As Long
    Dim Char As String
    Dim Args() As String
    Dim nArgs As Long
    Dim inString As Boolean
    
    
    Set Switches = New Dictionary
    
    ' Parse commandline into space-separated args quoting too
    nArgs = 1
    ReDim Args(0)
    For A = 1 To Len(Command)
        Char = Mid$(Command, A, 1)
        
        ' String start\end
        If Char = vbQuote Then
            If inString = False Then inString = True Else inString = False
            Char = vbNullString
        
        ' Space for next arg
        ElseIf Char = " " And inString = False Then
            ReDim Preserve Args(nArgs)
            nArgs = nArgs + 1
            Char = vbNullString
        
        End If
        
        Args(nArgs - 1) = Args(nArgs - 1) & Char
    Next A

    
    ' Parse args
    For A = 0 To nArgs - 1
        If Left$(Args(A), 1) = "-" Then
            If A <= nArgs - 2 Then
                Switches.Add LCase$(Mid$(Args(A), 2)), Args(A + 1)
            Else
                Switches.Add LCase$(Mid$(Args(A), 2)), -1
            End If
        End If
    Next A

End Sub
