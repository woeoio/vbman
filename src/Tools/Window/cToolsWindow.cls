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

    
'---------------------
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) As Long

Private Const VK_MENU = &H12                                                    ' Alt键
'---------------------
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
    
Private Declare Function SwitchToThisWindow Lib "user32" (ByVal hWnd As Long, ByVal fAltTab As Long) As Long

Public Sub TopMost(ByVal hWnd As Long, Optional Cancel As Boolean)
    If Cancel = True Then
        SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Public Sub SwitchToThis(ByVal hWnd As Long, Optional IsAltTab As Boolean = True)
    SwitchToThisWindow hWnd, IsAltTab
End Sub


Public Sub ActiveForm(ByVal hWnd As Long)
    '模拟 Alt 键按下（绕过系统限制强制获取焦点解决方案）
    '建议在 Form_LostFocus 事件执行，可以达到用不失焦的效果
    keybd_event VK_MENU, 0, 0, 0
    ' 设置前景窗口
    SetForegroundWindow hWnd
    ' 模拟释放Alt键
    keybd_event VK_MENU, 0, &H2, 0
End Sub

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub
