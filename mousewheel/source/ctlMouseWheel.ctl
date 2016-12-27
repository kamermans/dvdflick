VERSION 5.00
Begin VB.UserControl ctlMouseWheel 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   HasDC           =   0   'False
   Picture         =   "ctlMouseWheel.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "ctlMouseWheel.ctx":1142
End
Attribute VB_Name = "ctlMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Binary
Option Base 0


Private Const WM_MOUSEWHEEL As Long = &H20A

Private pHwnd As Long
Private doneInit As Boolean
Private m_emr As EMsgResponse


Implements ISubclass

Event Wheel(ByVal Delta As Long)


Private Property Let ISubClass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)

    m_emr = RHS

End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
 
    ISubClass_MsgResponse = emrConsume
    
End Property


Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    RaiseEvent Wheel(wParam / &H780000)

End Function


Public Sub Init(ByVal hWnd As Long)

    ' Init here to get Parent hWnd
    If Not doneInit Then
        If Ambient.UserMode Then
            Set UserControl.Picture = Nothing
            UserControl.BackStyle = 0
        End If
        pHwnd = hWnd
        AttachMessage Me, pHwnd, WM_MOUSEWHEEL
        doneInit = True
    End If

End Sub


Private Sub UserControl_Resize()

    UserControl.Width = 32 * 15
    UserControl.Height = 32 * 15

End Sub


Private Sub UserControl_Terminate()

    If doneInit Then DetachMessage Me, pHwnd, WM_MOUSEWHEEL
    
End Sub
