VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===========================================================================
' 名称:    cDelay
' 描述:    延时对象类 - 支持事件触发、回调函数和同步等待三种模式
' 作者:    邓伟，QQ: 215879458
' 网站:    https://vb6.pro
' 日期:    2026-03-31
' 参考:    cTimer.cls 架构
'===========================================================================

'----------------------------------
' Windows API 声明
'----------------------------------
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long

Private Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    ptX As Long
    ptY As Long
End Type

Private Const PM_REMOVE = &H1
Private Const WM_TIMER = &H113

'----------------------------------
' 枚举类型
'----------------------------------
Public Enum DelayMode
    dmEvent = 0     '事件触发模式
    dmCallback = 1  '回调函数模式
    dmSync = 2      '同步等待模式
End Enum

'----------------------------------
' 事件声明
'----------------------------------
Public Event OnTime()
Public Event OnCancel()

'----------------------------------
' 成员变量
'----------------------------------
Private m_TimerID As Long           '定时器ID
Private m_IsActive As Boolean       '是否激活
Private m_Mode As DelayMode         '当前模式
Private m_DelayMs As Long           '延迟毫秒数
Private m_StartTick As Long         '开始时间戳
Private m_Cancelled As Boolean      '取消标志(同步模式用)

'回调模式参数
Private m_CallbackObj As Object     '回调对象
Private m_CallbackProc As String    '回调过程名
Private m_ParamCount As Integer     '参数数量
Private m_P1 As Variant             '参数1
Private m_P2 As Variant             '参数2
Private m_P3 As Variant             '参数3
Private m_P4 As Variant             '参数4
Private m_P5 As Variant             '参数5
Private m_P6 As Variant             '参数6
Private m_P7 As Variant             '参数7
Private m_P8 As Variant             '参数8
Private m_P9 As Variant             '参数9

'集合管理用
Private m_Index As Long
Private m_ParentsColKey As Variant

'----------------------------------
' 属性: Index (用于集合管理)
'----------------------------------
Friend Property Let Index(lIndex As Long)
    m_Index = lIndex
End Property

Friend Property Get Index() As Long
    Index = m_Index
End Property

Friend Property Let ParentsColKey(key As Variant)
    m_ParentsColKey = key
End Property

Friend Property Get ParentsColKey() As Variant
    ParentsColKey = m_ParentsColKey
End Property

'----------------------------------
' 属性: Mode (当前模式)
'----------------------------------
Public Property Get Mode() As DelayMode
    Mode = m_Mode
End Property

'----------------------------------
' 属性: IsActive (是否激活)
'----------------------------------
Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

'----------------------------------
' 属性: IsCancelled (是否被取消)
' 说明: 同步模式下使用，用于判断 CountDown 是被取消还是正常完成
'----------------------------------
Public Property Get IsCancelled() As Boolean
    IsCancelled = m_Cancelled
End Property

'----------------------------------
' 属性: DelayMs (延迟毫秒数)
'----------------------------------
Public Property Get DelayMs() As Long
    DelayMs = m_DelayMs
End Property

'----------------------------------
' 方法: Callback (设置回调模式)
' 用法: Delay.Callback(Me, "MyFunc", arg1, arg2, ...).CountDown(3000)
' 说明: 使用 ParamArray 支持变长参数，最多9个
'----------------------------------
Public Function Callback(ByVal CallbackObject As Object, ByVal ProcName As String, ParamArray P()) As cDelay
    
    Dim i As Long
    
    '如果定时器正在运行，先停止并清理
    If m_IsActive Then
        StopTimer
        ClearCallback
    End If
    
    Set m_CallbackObj = CallbackObject
    m_CallbackProc = ProcName
    
    '获取参数数量 (最多9个)
    m_ParamCount = UBound(P) + 1
    If m_ParamCount > 9 Then m_ParamCount = 9
    
    '存储参数
    For i = 0 To m_ParamCount - 1
        Select Case i
            Case 0: m_P1 = P(i)
            Case 1: m_P2 = P(i)
            Case 2: m_P3 = P(i)
            Case 3: m_P4 = P(i)
            Case 4: m_P5 = P(i)
            Case 5: m_P6 = P(i)
            Case 6: m_P7 = P(i)
            Case 7: m_P8 = P(i)
            Case 8: m_P9 = P(i)
        End Select
    Next i
    
    m_Mode = dmCallback
    Set Callback = Me
End Function

'----------------------------------
' 方法: Sync (设置同步等待模式)
' 用法: Delay.Sync().CountDown(3000)
'----------------------------------
Public Function Sync() As cDelay
    m_Mode = dmSync
    Set Sync = Me
End Function

'----------------------------------
' 方法: CountDown (开始倒计时)
'----------------------------------
Public Sub CountDown(ByVal Milliseconds As Long)
    
    '停止现有定时器
    If m_IsActive Then StopTimer
    
    m_DelayMs = Milliseconds
    m_StartTick = GetTickCount()
    m_Cancelled = False
    
    '同步模式特殊处理
    If m_Mode = dmSync Then
        m_IsActive = True
        DoSyncWait
        m_IsActive = False
    Else
        '事件或回调模式 - 使用Windows定时器
        StartTimer
    End If
    
End Sub

'----------------------------------
' 方法: Cancel (取消延迟)
' 说明: 停止定时器并清理所有回调信息
'       同步模式下设置取消标志，使消息循环提前退出
'----------------------------------
Public Sub Cancel()
    If m_IsActive Then
        '同步模式特殊处理: 设置取消标志
        If m_Mode = dmSync Then
            m_Cancelled = True
        Else
            StopTimer
        End If
        ClearCallback
        RaiseEvent OnCancel
    End If
End Sub

'----------------------------------
' 私有方法: ClearCallback (清理回调信息)
'----------------------------------
Private Sub ClearCallback()
    Set m_CallbackObj = Nothing
    m_CallbackProc = ""
    m_ParamCount = 0
    m_P1 = Empty: m_P2 = Empty: m_P3 = Empty
    m_P4 = Empty: m_P5 = Empty: m_P6 = Empty
    m_P7 = Empty: m_P8 = Empty: m_P9 = Empty
End Sub

'----------------------------------
' 私有方法: StartTimer (启动定时器)
'----------------------------------
Private Sub StartTimer()
    
    If m_DelayMs <= 0 Then Exit Sub
    
    m_TimerID = SetTimer(0&, 0&, m_DelayMs, AddressOf DelayTimerProc)
    
    If m_TimerID <> 0 Then
        m_IsActive = True
        '添加到全局集合
        Delays.Add Me, "id:" & m_TimerID
    Else
        m_IsActive = False
        m_TimerID = 0
        ERR.Raise vbObjectError + 1, "cDelay", "无法创建定时器"
    End If
    
End Sub

'----------------------------------
' 私有方法: StopTimer (停止定时器)
'----------------------------------
Private Sub StopTimer()
    
    If m_TimerID <> 0 Then
        KillTimer 0, m_TimerID
        '从全局集合移除
        On Error Resume Next
        Delays.Remove "id:" & m_TimerID
        On Error GoTo 0
    End If
    
    m_IsActive = False
    m_TimerID = 0
    
End Sub

'----------------------------------
' 私有方法: DoSyncWait (同步等待 - 消息泵)
' 说明: 进入消息循环等待，不卡UI，支持取消
'----------------------------------
Private Sub DoSyncWait()
    
    Dim uMsg As msg
    Dim EndTick As Long
    Dim bDone As Boolean
    
    EndTick = GetTickCount() + m_DelayMs
    bDone = False
    '注意: m_Cancelled 在 CountDown 中已经重置为 False
    '不要在循环开始前重置，否则可能覆盖 Cancel 方法设置的 True
    
    '消息循环 - 不卡UI的等待
    Do While Not bDone
        '检查是否被取消
        If m_Cancelled Then
            bDone = True
            Exit Do
        End If
        
        '检查是否到达目标时间
        If GetTickCount() >= EndTick Then
            bDone = True
            Exit Do
        End If
        
        '处理消息队列
        Do While PeekMessage(uMsg, 0&, 0&, 0&, PM_REMOVE) <> 0
            TranslateMessage uMsg
            DispatchMessage uMsg
        Loop
        
        '短暂休眠，避免CPU占用过高
        Sleep 1

    Loop
    
    '同步模式下直接返回，不执行回调或事件
    '调用者会继续执行下一行代码
    
End Sub

'----------------------------------
' 公共方法: ExecuteAction (执行动作 - 由定时器回调调用)
'----------------------------------
Public Sub ExecuteAction()
    
    m_IsActive = False
    m_TimerID = 0
    
    Select Case m_Mode
        Case dmEvent
            '事件模式
            RaiseEvent OnTime
            
        Case dmCallback
            '回调模式 - 使用CallByName
            If Not m_CallbackObj Is Nothing And m_CallbackProc <> "" Then
                ExecuteCallback
            End If
            
        Case dmSync
            '同步模式已在DoSyncWait中处理
            '这里不需要额外操作
            
    End Select
    
End Sub

'----------------------------------
' 私有方法: ExecuteCallback (执行回调)
'----------------------------------
Private Sub ExecuteCallback()
    
    On Error Resume Next
    
    Select Case m_ParamCount
        Case 0
            CallByName m_CallbackObj, m_CallbackProc, VbMethod
        Case 1
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1
        Case 2
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2
        Case 3
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3
        Case 4
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3, m_P4
        Case 5
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3, m_P4, m_P5
        Case 6
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3, m_P4, m_P5, m_P6
        Case 7
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3, m_P4, m_P5, m_P6, m_P7
        Case 8
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3, m_P4, m_P5, m_P6, m_P7, m_P8
        Case 9
            CallByName m_CallbackObj, m_CallbackProc, VbMethod, m_P1, m_P2, m_P3, m_P4, m_P5, m_P6, m_P7, m_P8, m_P9
    End Select
    
    On Error GoTo 0
    
End Sub

'----------------------------------
' 类初始化
'----------------------------------
Private Sub Class_Initialize()
    m_Mode = dmEvent      '默认事件模式
    m_IsActive = False
    m_TimerID = 0
    m_DelayMs = 0
    m_ParamCount = 0
    m_Cancelled = False
    
    '自动初始化：增加引用计数，首次创建时初始化全局资源
    AutoInit
End Sub

'----------------------------------
' 私有方法: AutoInit (自动初始化)
'----------------------------------
Private Sub AutoInit()
    On Error Resume Next
    
    '增加实例计数
    Dim count As Long
    count = GetInstanceCount()
    
    '如果是第一个实例，初始化全局资源
    If count = 0 Then
        InitDelaySystem
    End If
    
    count = count + 1
    SetInstanceCount count
    
    On Error GoTo 0
End Sub

'----------------------------------
' 类终止
'----------------------------------
Private Sub Class_Terminate()
    StopTimer
    Set m_CallbackObj = Nothing
    
    '自动清理：减少引用计数，如果为0则清理全局资源
    AutoCleanup
End Sub

'----------------------------------
' 私有方法: AutoCleanup (自动清理)
'----------------------------------
Private Sub AutoCleanup()
    On Error Resume Next
    
    '减少实例计数
    Dim count As Long
    count = GetInstanceCount()
    count = count - 1
    SetInstanceCount count
    
    '如果计数为0，清理全局资源
    If count <= 0 Then
        TermDelaySystem
    End If
    
    On Error GoTo 0
End Sub
