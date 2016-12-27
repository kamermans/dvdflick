VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDebug"
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
'   File purpose: Debug message logging.
'
Option Explicit
Option Compare Binary
Option Base 0


Private mDebugMode As enumDebugModes


Public Property Get debugMode() As enumDebugModes

    debugMode = mDebugMode

End Property

Public Property Let debugMode(ByVal newMode As enumDebugModes)

    mDebugMode = newMode

End Property


Public Sub setMode(ByVal modeFlags As String)

    Dim A As Long
    Dim flags() As String
    
    
    modeFlags = LCase$(modeFlags)
    flags = Split(modeFlags, "+")
    
    For A = 0 To UBound(flags)
        If flags(A) = "gdi" Then mDebugMode = mDebugMode Or DM_GDI
        If flags(A) = "encoding" Then mDebugMode = mDebugMode Or DM_Encoding
        If flags(A) = "menus" Then mDebugMode = mDebugMode Or DM_Menus
        If flags(A) = "pipes" Then mDebugMode = mDebugMode Or DM_Pipes
        If flags(A) = "sourceparser" Then mDebugMode = mDebugMode Or DM_SourceParser
        If flags(A) = "subtitles" Then mDebugMode = mDebugMode Or DM_Subtitles
        If flags(A) = "all" Then mDebugMode = &HFFFFFFFF
    Next A
    
    appLog.Add "Debugmode flags: " & mDebugMode

End Sub


Public Sub addLog(ByVal Mode As enumDebugModes, ByVal Msg As String)

    Dim Pre As String
    
    
    If mDebugMode And Mode Then
        If Mode = DM_Encoding Then Pre = "ENCODE"
        If Mode = DM_GDI Then Pre = "GDI"
        If Mode = DM_Menus Then Pre = "MENU"
        If Mode = DM_Pipes Then Pre = "PIPE"
        If Mode = DM_SourceParser Then Pre = "SOURCE"
        If Mode = DM_Subtitles Then Pre = "SUBTITLE"
        
        appLog.Add Pre & ": " & Msg
    End If

End Sub
