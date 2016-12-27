Attribute VB_Name = "modShellEx"
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
'   File purpose: Shell Execute function and utils
'
Option Explicit
Option Compare Binary
Option Base 0


' ShellExecute constants
' http://msdn.microsoft.com/library/en-us/shellcc/platform/Shell/reference/functions/shellexecute.asp
Public Enum ENUM_SHELLWINDOWSTYLE
    WS_HIDE = 0
    WS_SHOW = 5
    WS_DEFAULT = 10
    WS_MAXIMIZED = 3
    WS_MINIMIZED = 2
    WS_MINIMIZED_NOACTIVE = 7
    WS_NA = 8
    WS_NOACTIVE = 4
    WS_Normal = 1
End Enum

' ShellExecuteEx mask constants
Public Enum ENUM_SHELLEXECUTEMASK
    SEE_MASK_CLASSKEY = &H3
    SEE_MASK_CLASSNAME = &H1
    SEE_MASK_CONNECTNETDRV = &H80
    SEE_MASK_DOENVSUBST = &H200
    SEE_MASK_FLAG_DDEWAIT = &H100
    SEE_MASK_FLAG_NO_UI = &H400
    SEE_MASK_HOTKEY = &H20
    SEE_MASK_ICON = &H10
    SEE_MASK_IDLIST = &H4
    SEE_MASK_INVOKEIDLIST = &HC
    SEE_MASK_NOCLOSEPROCESS = &H40
End Enum

' ShellExecuteEx structure
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/structures/shellexecuteinfo.asp
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    
    ' Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (ByRef execInfo As SHELLEXECUTEINFO) As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Function Execute(ByRef fileName As String, ByRef Parameters As String, ByVal windowStyle As ENUM_SHELLWINDOWSTYLE, ByVal waitForProcess As Boolean, Optional ByVal pauseWait As Boolean = True) As Boolean
     
    Dim execInfo As SHELLEXECUTEINFO
    
    
    ' Check if we should add local path
    If (InStr(fileName, "\") = 0) And _
        (InStr(fileName, "/") = 0) And _
        (InStr(fileName, ":") = 0) Then
     
        ' Add local path to filename
        fileName = APP_PATH & fileName
    End If
    
    ' Fill structure
    With execInfo
        .cbSize = Len(execInfo)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .lpFile = fileName
        .lpParameters = Parameters
        .lpDirectory = FS.GetParentFolderName(fileName)
        .nShow = windowStyle
    End With
    
    haltShellEx = False
    Dbg.addLog DM_Pipes, "Executing (ShellExecuteEx) " & fileName & " " & Parameters
    Dbg.addLog DM_Pipes, "Options: " & windowStyle & ", " & waitForProcess & ", " & pauseWait
    Execute = ShellExecuteEx(execInfo)
    
    ' Check if we should wait
    If (waitForProcess) Then
         
        ' Wait for the process to end
        Do While WaitForSingleObject(execInfo.hProcess, 10)
            If haltShellEx Then
                TerminateProcess execInfo.hProcess, 0
                Exit Do
            End If
            If pauseWait Then DoEvents
        Loop
        
    End If
     
End Function
