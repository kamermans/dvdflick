VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
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
'   File purpose: Audio source for audio track.
'
Option Explicit
Option Compare Binary
Option Base 0


' Source object
Public Source As clsSource
Private mStreamIndex As Long


' Save to dictionary
Public Function Serialize() As Dictionary

    Set Serialize = New Dictionary

    Serialize.Add "sourceFile", Source.fileName
    Serialize.Add "sourceStream", mStreamIndex

End Function


' Return the index of the stream in the source file
Public Property Get streamIndex() As Long

    streamIndex = mStreamIndex

End Property


' Read from dictionary
Public Function unSerialize(ByRef myDict As Dictionary) As Boolean

    Dim fileName As String
    Dim Stream As Long
    
    
    unSerialize = False
    
    fileName = myDict("sourceFile")
    mStreamIndex = myDict("sourceStream")
    
    ' Get source from project
    Set Source = Project.getSource(fileName)
    If Source Is Nothing Then Exit Function
    
    ' Check if streamtype has not changed
    If Source.streamInfo(mStreamIndex)("Type") <> ST_Audio Then Exit Function
    
    unSerialize = True
    
End Function


' Open an audio stream from a file
Public Function openFrom(ByRef mySource As clsSource, ByVal streamIndex As Long) As Boolean
    
    openFrom = False
    If mySource Is Nothing Then Exit Function
    
    Reset
    Set Source = mySource
    mStreamIndex = streamIndex
    
    openFrom = True
    
End Function


' Return the source stream's information
Public Function streamInfo() As Dictionary

    Set streamInfo = Source.streamInfo(mStreamIndex)

End Function


' Return the size when this audio source is encoded in bytes
Public Function encodedSize(ByVal bitRate As Long) As Long

    encodedSize = (bitRate * Source.streamInfo(mStreamIndex)("Duration")) / 8

End Function


' Reset all properties
Private Sub Reset()
    
    Set Source = Nothing
    mStreamIndex = -1

End Sub


' Class
Private Sub Class_Initialize()

    Reset

End Sub

Private Sub Class_Terminate()

    Reset

End Sub
