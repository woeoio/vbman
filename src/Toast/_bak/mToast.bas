Attribute VB_Name = "mToast"
Option Explicit

'===========================================================================
' Toast 提示框便捷函数模块
' 提供全局快速调用方法
'===========================================================================

' 创建一个新的 Toast 实例（链式调用的起点）
Public Function Toast() As cToast
    Set Toast = New cToast
End Function

' 快速显示成功提示
Public Sub ToastSuccess(ByVal Message As String, Optional ByVal Duration As Long = 3000)
    With New cToast
        .Style ToastStyle_Success
        .Text Message
        .Duration Duration
        .Show
    End With
End Sub

' 快速显示错误提示
Public Sub ToastDanger(ByVal Message As String, Optional ByVal Duration As Long = 3000)
    With New cToast
        .Style ToastStyle_Error
        .Text Message
        .Duration Duration
        .Show
    End With
End Sub

' 快速显示警告提示
Public Sub ToastWarning(ByVal Message As String, Optional ByVal Duration As Long = 3000)
    With New cToast
        .Style ToastStyle_Warning
        .Text Message
        .Duration Duration
        .Show
    End With
End Sub

' 快速显示信息提示
Public Sub ToastInfo(ByVal Message As String, Optional ByVal Duration As Long = 3000)
    With New cToast
        .Style ToastStyle_Info
        .Text Message
        .Duration Duration
        .Show
    End With
End Sub

' 在指定位置显示提示
Public Sub ToastAt(ByVal Message As String, ByVal x As Long, ByVal Y As Long, Optional ByVal Duration As Long = 3000)
    With New cToast
        .Pos x, Y
        .Text Message
        .Duration Duration
        .Show
    End With
End Sub

' 显示自定义样式提示
Public Sub ToastCustom(ByVal Message As String, _
                       ByVal BackColor As Long, _
                       ByVal TextColor As Long, _
                       Optional ByVal Position As ToastPosition = Toast_BottomCenter, _
                       Optional ByVal Duration As Long = 3000)
    With New cToast
        Select Case Position
            Case Toast_TopLeft: .TopLeft
            Case Toast_TopCenter: .TopCenter
            Case Toast_TopRight: .TopRight
            Case Toast_MiddleLeft: .MiddleLeft
            Case Toast_MiddleCenter: .MiddleCenter
            Case Toast_MiddleRight: .MiddleRight
            Case Toast_BottomLeft: .BottomLeft
            Case Toast_BottomCenter: .BottomCenter
            Case Toast_BottomRight: .BottomRight
        End Select
        .BackColor BackColor
        .ForeColor TextColor
        .Text Message
        .Duration Duration
        .Show
    End With
End Sub
