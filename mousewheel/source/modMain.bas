Attribute VB_Name = "modMain"
Option Explicit
Option Compare Binary
Option Base 0


Public Type Point
    X As Long
    Y As Long
End Type


Public Type Message
    hWnd As Long
    Msg As Long
    wParam As Long
    lParam As Long
    Time As Long
    Pt As Point
End Type


Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal ptrHookProc As Long, ByVal hMod As Long, ByVal threadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal Code As Long, ByVal wParam As Long, ByRef lParam As Message) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long


Public Sub initHook(ByVal hWnd As Long)
    
    Dim hHook As Long
    
 
    hHook = SetWindowsHookEx(3, AddressOf msgHook, 0, GetCurrentThreadId)
    SetProp hWnd, "hookMouseWheel", hHook

End Sub


Public Sub destroyHook()

    UnhookWindowsHookEx hHook

End Sub


Public Function msgHook(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Message) As Long
    
    If lParam.Msg = 522 Then Owner.WheelMoved lParam.wParam / 7864320, 0, lParam.Pt.X, lParam.Pt.Y
    msgHook = CallNextHookEx(hHook, nCode, wParam, lParam)

End Function


