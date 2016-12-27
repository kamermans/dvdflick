VERSION 5.00
Object = "{1C4E972B-90D2-46D1-A2F1-DAF5A8DE9F08}#25.0#0"; "mousewheel.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MouseWheel.ctlMouseWheel ctlMouseWheel1 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ctlMouseWheel1_Wheel(ByVal Delta As Long)

    Debug.Print Delta

End Sub


Private Sub Form_Load()

    ctlMouseWheel1.Init Me.hWnd

End Sub
