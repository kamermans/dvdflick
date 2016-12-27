VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideo"
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
'   File purpose: Project video source
'
Option Explicit
Option Compare Binary
Option Base 0


Public Source As clsSource
Public PAR As Single
Public Interlaced As Byte

Private mStreamIndex As Long


' Return guessed format that this title is in it's source form
Public Function guessedFormat() As String

    Dim FPS As Single
    
    
    FPS = Source.streamInfo(mStreamIndex)("FPS")
    
    If FPS > 27 Or FPS <= 24 Then
        guessedFormat = "NTSC"
    Else
        guessedFormat = "PAL"
    End If

End Function


' streamINdex property
Public Property Get streamIndex() As Long

    streamIndex = mStreamIndex

End Property


' Thumbnail property
Public Property Get Thumbnail() As clsGDIImage

    Set Thumbnail = Source.streamImage(mStreamIndex)

End Property


' streamInfo property
Public Function streamInfo() As Dictionary

    Set streamInfo = Source.streamInfo(mStreamIndex)

End Function


' Write to file
Public Function Serialize() As Dictionary

    Set Serialize = New Dictionary
    
    Serialize.Add "sourceFile", Source.fileName
    Serialize.Add "sourceStream", mStreamIndex
    Serialize.Add "PAR", PAR
    Serialize.Add "Interlaced", Interlaced
    
End Function


' Read from file
Public Function unSerialize(ByVal myDict As Dictionary) As Boolean

    Dim fileName As String
    
    
    unSerialize = False
    
    fileName = myDict("sourceFile")
    mStreamIndex = myDict("sourceStream")
    PAR = myDict("PAR")
    Interlaced = myDict("Interlaced")
    
    ' Get source from project
    Set Source = Project.getSource(fileName)
    If Source Is Nothing Then Exit Function
    
    ' Check if streamtype has not changed
    If Source.streamInfo(mStreamIndex)("Type") <> ST_Video Then Exit Function
    
    unSerialize = True
    
End Function


' Read a video stream from a source
Public Function openFrom(ByRef mySource As clsSource, ByVal streamIndex As Long) As Boolean

    Dim A As Long
    Dim Info As Dictionary
    
    
    If mySource Is Nothing Then Exit Function
    
    Set Source = mySource
    
    ' First video stream is the one we will use
    Set Info = Source.streamInfo(streamIndex)

    If Info("Type") = ST_Video Then
        mStreamIndex = streamIndex
        PAR = Info("pixelAR")
        Interlaced = Info("Interlaced")
        
        If unsupportedVideo(Info("Compression")) Then
            frmDialog.Display "The video codec """ & Info("Compression") & """ is unsupported.", Exclamation Or OkOnly
            Exit Function
        End If
    End If
    
    openFrom = True

End Function


' Return the size of this video when encoded
Public Function encodedSize(ByVal bitRate As Long) As Long

    encodedSize = (bitRate * Source.streamInfo(mStreamIndex)("Duration")) / 8

End Function


Private Sub Class_Terminate()

    Set Source = Nothing
    mStreamIndex = 0
    PAR = 0
    Interlaced = 0

End Sub
