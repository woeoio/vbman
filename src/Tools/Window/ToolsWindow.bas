Attribute VB_Name = "ToolsWindow"
Option Explicit
    
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
    
Private Declare Function SwitchToThisWindow Lib "user32" (ByVal hwnd As Long, ByVal fAltTab As Long) As Long

Public Sub TopMost(ByVal hwnd As Long, Optional Cancel As Boolean)
    If Cancel = True Then
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Public Sub SwitchToThis(ByVal hwnd As Long, Optional IsAltTab As Boolean = True)
    SwitchToThisWindow hwnd, IsAltTab
End Sub
