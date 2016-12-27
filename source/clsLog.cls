VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
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
'   File purpose: Log text to a file.
'
Option Explicit
Option Compare Binary
Option Base 0


' Handle to the log file
Private logStream As clsTextFile
Private logFileName As String

' Add timestamps
Private mTimeStamps As Boolean

' Unicode enabled
Private useUnicode As Boolean


' Begin logging
Public Sub Start(ByVal fileName As String, Optional ByVal timeStamps As Boolean = True, Optional ByVal Unicode As Boolean = False)

    Set logStream = New clsTextFile
    If Not logStream.fileCreate(fileName, encodeASCIIorUTF8, CODEPAGE_LATIN1) Then
        Err.Raise -1, "clsLog", "Unable to create text file."
        Exit Sub
    End If
    
    logFileName = fileName
    mTimeStamps = timeStamps
    useUnicode = Unicode

End Sub


' Add a line to the log
Public Sub Add(ByVal Text As String, Optional ByVal Indent As Long = 0)

    Dim Data As String
    
    
    ' Data to add
    If mTimeStamps Then Data = Time & ": "
    Data = Data & String(Indent * 4, ChrW$(32)) & Replace(Text, vbNullChar, ChrW$(32))
    
    ' Write to file
    logStream.writeText Data, True
    
End Sub


' Class
Private Sub Class_Terminate()

    Set logStream = Nothing
    logFileName = vbNullString

End Sub
