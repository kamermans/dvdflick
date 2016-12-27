VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudioTrack"
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
'   File purpose: Audio track for a title.
'
Option Explicit
Option Compare Binary
Option Base 0


' Audio track sources
Public Sources As New clsObjectArray

Public Language As Long
Public ignoreDelay As Byte


' Return the total duration of this track
Public Property Get Duration() As Single

    Dim A As Long
    
    
    For A = 0 To Sources.Count - 1
        Duration = Duration + Sources.Item(A).streamInfo("Duration")
    Next A

End Property


' Write to dictionary
Public Function Serialize() As Dictionary

    Dim A As Long
    
    
    Set Serialize = New Dictionary
    Serialize.Add "Language", Language
    Serialize.Add "ignoreDelay", ignoreDelay
    
    For A = 0 To Sources.Count - 1
        Serialize.Add "Audio" & A, Sources.Item(A).Serialize
    Next A
    
End Function


' Read from dictionary
Public Function unSerialize(ByRef myDict As Dictionary) As Boolean

    Dim A As Long
    Dim newAudio As clsAudio
    Dim myData As Dictionary
    
    
    unSerialize = False
    
    Language = myDict("Language")
    ignoreDelay = myDict("ignoreDelay")
    
    For A = 0 To myDict.Count - 1
        If Left$(myDict.Keys(A), 5) = "Audio" Then
            Set newAudio = New clsAudio
            Set myData = myDict.Items(A)
            If newAudio.unSerialize(myData) Then Sources.Add newAudio
        End If
    Next A
    
    unSerialize = True

End Function


' Return the target bitrate of this track
Public Function targetBitrate() As Long

    Dim targetCount As Long
    Dim Info As Dictionary
    
    
    targetCount = Me.targetChannels
    Set Info = Sources.Item(0).streamInfo
    
    ' Do not recompress
    If Info("Compression") = "ac3" And Info("sampleRate") = 48000 And Info("Channels") = targetCount And Project.volumeMod = 100 Then
        targetBitrate = Info("bitRate")
        Exit Function
    
    ' Auto-fit
    ElseIf Project.audioBitRate = AB_Auto Then
        
        Select Case targetCount
            Case 1
                targetBitrate = 96
            Case 2
                targetBitrate = 192
            Case 5 Or 6
                targetBitrate = 448
            Case Else
                targetBitrate = targetCount * 96
        End Select
       
    ' Custom bitrate
    Else
        If Project.audioBitRate = AB_64 Then targetBitrate = 64
        If Project.audioBitRate = AB_128 Then targetBitrate = 128
        If Project.audioBitRate = AB_256 Then targetBitrate = 256
        If Project.audioBitRate = AB_384 Then targetBitrate = 384
        If Project.audioBitRate = AB_448 Then targetBitrate = 448
        
    End If
    
    If targetBitrate > MAX_AUDIO_BITRATE Then targetBitrate = MAX_AUDIO_BITRATE

End Function


' Return number of target channels
Public Function targetChannels() As Long

    Dim sourceChannels As Long
    
    
    sourceChannels = Sources.Item(0).streamInfo("Channels")
    
    Select Case Project.channelCount
        Case AC_Auto
            targetChannels = sourceChannels
        Case AC_Mono
            targetChannels = 1
        Case AC_Stereo
            targetChannels = 2
        Case AC_Surround
            targetChannels = 6
    End Select
    
    ' Cannot resample input with more than 2 channels to 2 or less channels, so use input amount
    If Sources.Item(0).streamInfo("Compression") = "ac3" Then
        If sourceChannels > 2 And targetChannels <= 2 Then targetChannels = sourceChannels
    Else
        If targetChannels <> sourceChannels Then targetChannels = sourceChannels
    End If

End Function


' Adds a new audio source to this track
Public Function addSource(ByRef Source As clsSource, ByVal Stream As Long) As clsAudio

    Dim mySource As New clsAudio
    Dim Info As Dictionary
    
    
    If Not mySource.openFrom(Source, Stream) Then
        Set mySource = Nothing
        Exit Function
    End If
    
    ' Resampling with input channels greater than 2 unsupported by FFmpeg. Grr.
    Set Info = mySource.streamInfo
    If Info("sampleRate") <> 48000 And Info("Channels") > 2 Then
        frmDialog.Display "Cannot add audio track " & Stream + 1 & " from " & Source.fileName & ". It needs to be resampled but has more than 2 channels, which is not yet possible with input using more than 2 channels. The audio source will not be loaded.", OkOnly Or Information
        
        Set mySource = Nothing
        Exit Function
    End If
    
    Set addSource = mySource
    Sources.Add mySource

End Function


' Return the size of this track when encoded in bytes
Public Function encodedSize() As Long

    encodedSize = (Me.targetBitrate * Me.Duration) / 8

End Function


' Class
Private Sub Class_Initialize()

    Language = 28
    ignoreDelay = 0

End Sub

Private Sub Class_Terminate()

    Sources.Clear

End Sub
