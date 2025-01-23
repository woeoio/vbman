VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub TopMost(ByVal hwnd As Long, Optional Cancel As Boolean)
    Call ToolsWindow.TopMost(hwnd, Cancel)
End Sub

Public Sub SwitchToThis(ByVal hwnd As Long, Optional IsAltTab As Boolean = True)
    Call ToolsWindow.SwitchToThis(frm.hwnd, IsAltTab)
End Sub
