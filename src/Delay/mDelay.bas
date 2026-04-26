Attribute VB_Name = "mDelay"
'===========================================================================
' 名称:    cDelay
' 描述:    延时对象类 - 支持事件触发、回调函数和同步等待三种模式
' 作者:    邓伟，QQ: 215879458
' 网站:    https://vb6.pro
' 日期:    2026-03-31
' 参考:    cTimer.cls 架构
'===========================================================================

Option Explicit

Public Delay As New cDelay

'----------------------------------
' Windows API 声明
'----------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'----------------------------------
' 全局集合 - 存储所有活动的延迟对象
'----------------------------------
Public Delays As Collection

'----------------------------------
' 实例计数 - 用于自动初始化和清理
'----------------------------------
Private m_InstanceCount As Long

'----------------------------------
' 函数: GetInstanceCount (获取实例计数)
'----------------------------------
Public Function GetInstanceCount() As Long
    GetInstanceCount = m_InstanceCount
End Function

'----------------------------------
' 过程: SetInstanceCount (设置实例计数)
'----------------------------------
Public Sub SetInstanceCount(ByVal count As Long)
    m_InstanceCount = count
End Sub

'----------------------------------
' 模块初始化
'----------------------------------
Public Sub InitDelaySystem()
    Set Delays = New Collection
End Sub

'----------------------------------
' 模块清理
'----------------------------------
Public Sub TermDelaySystem()
    Dim Delay As cDelay
    
    If Not Delays Is Nothing Then
        '停止所有活动的定时器
        For Each Delay In Delays
            Delay.Cancel
        Next Delay
        
        Set Delays = Nothing
    End If
End Sub

'----------------------------------
' 定时器回调过程
' 注意: 必须是在标准模块中的Public Sub，否则AddressOf会失败
'----------------------------------
Public Sub DelayTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    
    Dim Delay As cDelay
    
    On Error Resume Next
    
    '从集合中获取对应的cDelay对象
    Set Delay = Delays("id:" & idEvent)
    
    If ERR.Number = 0 And Not Delay Is Nothing Then
        '停止定时器
        KillTimer 0, idEvent
        '从集合中移除
        Delays.Remove "id:" & idEvent
        '执行动作
        Delay.ExecuteAction
    End If
    
    On Error GoTo 0
    
End Sub
