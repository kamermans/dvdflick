VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPUDetect"
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
'   File purpose: CPU detection DLL function calls and wrappers.
'                 rspcpu.dll is copyright RSP Software - http://rspsoftware.clic3.net
'
Option Explicit
Option Compare Binary
Option Base 0


' rspcpu.dll calls
Private Declare Function getprocessorvendor Lib "rspcpu" (ByVal processorname As String) As Long
Private Declare Function getcpubrandstring Lib "rspcpu" (ByVal processorstring As String) As Long
Private Declare Function getprocessorspeed Lib "rspcpu" (ByVal processorstring As String) As Long
Private Declare Function getcputype Lib "rspcpu" (ByRef Family As Long, ByRef Model As Long, ByRef Stepping As Long) As Long

Private Declare Function hasMMXDll Lib "rspcpu" Alias "hasMMX" () As Long
Private Declare Function hasSSEDll Lib "rspcpu" Alias "hasSSE" () As Long
Private Declare Function hasSSE2Dll Lib "rspcpu" Alias "hasSSE2" () As Long
Private Declare Function has3DNOWDll Lib "rspcpu" Alias "has3DNOW" () As Long
Private Declare Function has3DNOWExtensionsDll Lib "rspcpu" Alias "has3DNOWExtensions" () As Long
Private Declare Function IsAMD64Dll Lib "rspcpu" Alias "UsAMD64" () As Long
Private Declare Function hasCMOVDll Lib "rspcpu" Alias "hasCMOV" () As Long
Private Declare Function hasFPUDll Lib "rspcpu" Alias "hasFPU" () As Long
Private Declare Function OShasSSEsupportDll Lib "rspcpu" Alias "OShasSSEsupport" () As Long
Private Declare Function hasHyperThreadingDll Lib "rspcpu" Alias "hasHyperThreading" () As Long
Private Declare Function getcpuextendedfeature Lib "rspcpu.dll" (ByRef eisst As Long, ByRef tm2 As Long, ByRef ci As Long) As Long

' Win32 API calls
Private Declare Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, ByRef lpProcessAffinityMask As Long, ByRef lpSystemAffinityMask As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


' Returns the number of available logical CPUs on the system
' Supports "only" 31 CPUs
Public Function getLogicalCPUCount() As Long

    Dim A As Long
    Dim procMask As Long
    Dim sysMask As Long
    
    
    GetProcessAffinityMask GetCurrentProcess, procMask, sysMask
    
    For A = 0 To 30
        If (sysMask And (2 ^ A)) Then getLogicalCPUCount = getLogicalCPUCount + 1
    Next A

End Function


' Feature retrieval functions
Public Function hasHyperThreading() As Boolean

    hasHyperThreading = CBool(hasHyperThreadingDll)

End Function

Public Function osHasSSESupport() As Boolean

    osHasSSESupport = CBool(OShasSSEsupportDll)

End Function

Public Function hasMMX() As Boolean

    hasMMX = CBool(hasMMXDll)

End Function

Public Function hasSSE() As Boolean

    hasSSE = CBool(hasSSEDll)

End Function

Public Function hasSSE2() As Boolean

    hasSSE2 = CBool(hasSSE2Dll)

End Function

Public Function has3DNOW() As Boolean

    has3DNOW = CBool(has3DNOWDll)

End Function

Public Function hasCMOW() As Boolean

    hasCMOW = CBool(hasCMOVDll)

End Function

Public Function hasFPU() As Boolean

    hasFPU = CBool(hasFPUDll)

End Function

Public Function has3DNOWExtensions() As Boolean

    has3DNOWExtensions = CBool(has3DNOWExtensionsDll)

End Function

Public Function isAMD64() As Boolean

    isAMD64 = CBool(IsAMD64Dll)

End Function


' CPU vendor string
Public Function getCPUVendor() As String
    
    Dim CPU As String * 255
    Dim Ret As Long
    
    
    Ret = getprocessorvendor(CPU)
    If InStr(CPU, vbNullChar) <> 0 Then CPU = Left$(CPU, InStr(CPU, vbNullChar) - 1)
    
    getCPUVendor = CPU
    
End Function


' CPU processor brand string
Public Function getProcessorBrandString() As String
    
    Dim CPU As String * 255
    Dim Ret As Long
    
    
    Ret = getcpubrandstring(CPU)
    If InStr(CPU, vbNullChar) <> 0 Then CPU = Left$(CPU, InStr(CPU, vbNullChar) - 1)
    
    getProcessorBrandString = CPU
    
End Function


' CPU speed as string
Public Function getCPUSpeed() As String
    
    Dim CPU As String * 255
    Dim Ret As Long
    
    
    Ret = getprocessorspeed(CPU)
    If InStr(CPU, vbNullChar) <> 0 Then CPU = Left$(CPU, InStr(CPU, vbNullChar) - 1)
    
    getCPUSpeed = CPU
    
End Function


' CPU family value
Public Function getCPUFamily() As Long

    Dim Family As Long
    Dim Model As Long
    Dim Stepping As Long
    
    
    getcputype Family, Model, Stepping
    getCPUFamily = Family

End Function


' CPU model value
Public Function getCPUModel() As Long

    Dim Family As Long
    Dim Model As Long
    Dim Stepping As Long
    
    
    getcputype Family, Model, Stepping
    getCPUModel = Model

End Function


' CPU stepping value
Public Function getCPUStepping() As Long

    Dim Family As Long
    Dim Model As Long
    Dim Stepping As Long
    
    
    getcputype Family, Model, Stepping
    getCPUStepping = Stepping

End Function

