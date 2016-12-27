Attribute VB_Name = "modPipes"
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
'   File purpose: Input\output pipe related functions. Crude and soon to be replaced.
'
Option Explicit
Option Compare Binary
Option Base 0


' Where do we send it to?
' Silly Visual Basic without function pointery things
Public Enum enumSendModes
    SM_Nothing = 0
    SM_EncodeVideo
    SM_EncodeAudio
    SM_TCMPlex
    SM_DVDAuthor
    SM_Pulldown
    SM_MKVExtract
End Enum


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    wFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type


' Process priority classes
Public Const PRIORITY_CLASS_REALTIME As Long = &H100
Public Const PRIORITY_CLASS_HIGH As Long = &H80
Public Const PRIORITY_CLASS_ABOVE_NORMAL As Long = &H8000
Public Const PRIORITY_CLASS_NORMAL As Long = &H20
Public Const PRIORITY_CLASS_BELOW_NORMAL As Long = &H4000
Public Const PRIORITY_CLASS_IDLE As Long = &H40

' Pipe constants
Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const PIPE_WAIT = &H0
Private Const PIPE_NOWAIT = &H1
Private Const PIPE_READMODE_BYTE = &H0
Private Const PIPE_READMODE_MESSAGE = &H2
Private Const PIPE_TYPE_BYTE = &H0
Private Const PIPE_TYPE_MESSAGE = &H4

' Process constants
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

' Public modifiers
Public newPriorityClass As Long


' API calls
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByVal lpStartupInfo As Long, ByVal lpProcessInformation As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetNamedPipeHandleState Lib "kernel32" (ByVal hNamedPipe As Long, lpMode As Long, lpMaxCollectionCount As Long, lpCollectDataTimeout As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal hClass As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function executeToFile(ByVal App As String, ByVal Parameters As String, ByVal outFile As String, ByVal sendMode As enumSendModes, ByVal Priority As enumEncodePriorities, ByVal startDir As String) As Boolean

    Dim Result As Long
    Dim readPipe As Long
    Dim writePipe As Long
    
    Const BUF_SIZE = 256
    
    Dim startInfo As STARTUPINFO
    Dim procInfo As PROCESS_INFORMATION
    Dim tempRead As String * BUF_SIZE
    Dim bytesRead As Long
    Dim openFile As Long
    Dim Sec As SECURITY_ATTRIBUTES
    Dim procPriority As Long

    
    executeToFile = False
    If LenB(startDir) = 0 Then startDir = APP_PATH
    newPriorityClass = -1
    
    ' Valid app and parameters
    If Left$(Parameters, 1) <> " " Then Parameters = " " & Parameters
    
    ' Set encoder process priority default
    If Priority = EP_AboveNormal Then
        procPriority = PRIORITY_CLASS_ABOVE_NORMAL
    ElseIf Priority = EP_Normal Then
        procPriority = PRIORITY_CLASS_NORMAL
    ElseIf Priority = EP_BelowNormal Then
        procPriority = PRIORITY_CLASS_BELOW_NORMAL
    ElseIf Priority = EP_Idle Then
        procPriority = PRIORITY_CLASS_IDLE
    Else
        procPriority = PRIORITY_CLASS_NORMAL
    End If
    
    Dbg.addLog DM_Pipes, "Executing " & App & Parameters
    Dbg.addLog DM_Pipes, "Options: " & startDir & ", " & outFile
    
    ' Create the pipe and set it up
    With Sec
        .bInheritHandle = True
        .nLength = Len(Sec)
    End With
    
    If CreatePipe(readPipe, writePipe, Sec, 0) = 0 Then
        Err.Raise Err_CreatePipe, , "Could not create pipe LastError " & Err.LastDllError
        Exit Function
    End If
    
    If SetNamedPipeHandleState(readPipe, PIPE_NOWAIT, 0, 0) <> 0 Then
        Err.Raise Err_PipeState, , "Could not set pipe state LastError " & Err.LastDllError
        Exit Function
    End If
    
    
    ' Create process and attach pipe to it
    With startInfo
        .cb = Len(startInfo)
        .wFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .hStdOutput = writePipe
        .hStdError = writePipe
    End With
    
    haltShellEx = False
    Result = CreateProcess(StrPtr(App), StrPtr(Parameters), Sec, Sec, True, 0, 0, StrPtr(startDir), VarPtr(startInfo), VarPtr(procInfo))
    If Result = 0 Then
        Err.Raise Err_PipeProcess, , "Could not create piped process. System error code " & Err.LastDllError
        Exit Function
    End If
    SetPriorityClass procInfo.hProcess, procPriority

    
    openFile = FreeFile
    Open outFile For Output As openFile
    
        Print #openFile, App & Parameters
    
        Do
            DoEvents
            
            Result = ReadFile(readPipe, tempRead, BUF_SIZE, bytesRead, 0&)
            If Result <> 0 And bytesRead > 0 Then
                Print #openFile, Left$(tempRead, bytesRead);
                If sendMode Then sendData Left$(tempRead, bytesRead), sendMode
            End If
            
            If haltShellEx Then
                TerminateProcess procInfo.hProcess, 0
                
                Do: DoEvents: Loop While WaitForSingleObject(procInfo.hProcess, 400)
                Exit Do
            End If
            
            If newPriorityClass <> -1 Then
                SetPriorityClass procInfo.hProcess, newPriorityClass
                newPriorityClass = -1
            End If
            
        Loop While WaitForSingleObject(procInfo.hProcess, 400)
        
        
        ' Read the last data that might be left in the buffer
        Do
            Result = ReadFile(readPipe, tempRead, BUF_SIZE, bytesRead, 0&)
            If Result <> 0 And bytesRead > 0 Then
                Print #openFile, Left$(tempRead, bytesRead);
                If sendMode Then sendData Left$(tempRead, bytesRead), sendMode
            End If
            DoEvents
        Loop Until Result = 0
    
    Close openFile

    CloseHandle procInfo.hProcess
    CloseHandle procInfo.hThread
    CloseHandle readPipe
    CloseHandle writePipe
    
    ' Ensure it's terminated... a possible hack to make a possible fix.
    WaitForSingleObject procInfo.hProcess, 400
    TerminateProcess procInfo.hProcess, 0

    executeToFile = True

End Function


Private Sub sendData(ByVal Data As String, ByVal sendMode As enumSendModes)

    Select Case sendMode
    
        Case SM_EncodeVideo, SM_EncodeAudio
            modEncode.ffmpegPipe Data
            
        Case SM_TCMPlex
            modEncode.tcmplexPipe Data
            
        Case SM_DVDAuthor
            modEncode.dvdauthorPipe Data
            
        Case SM_Pulldown
            modEncode.pulldownPipe Data
            
        Case SM_MKVExtract
            modEncode.mkvExtractPipe Data
                    
    End Select

End Sub
