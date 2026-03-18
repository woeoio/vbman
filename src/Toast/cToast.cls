VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ========== 事件声明 ==========
Public Event OnToastCountChange(ByVal TagName As String, ByVal IsDelete As Boolean, ByVal CurrentCount As Long)
Public Event OnCloseAll(ByVal closedCount As Long)

Public Enum EnumPos
    LeftTop = 10
    LeftCenter = 12
    LeftBottom = 14
    CenterTop = 20
    Center = 22
    CenterBottom = 24
    RightTop = 30
    RightCenter = 32
    RightBottom = 34
End Enum

Public Enum EnumTheme
    Light = 1
    Dark = 2
End Enum

Public Enum EnumState
    Info = 1
    Success = 2
    Warning = 3
    Danger = 4
End Enum

' 私有成员
Private m_Pos As EnumPos
Private m_State As EnumState
Private m_Theme As EnumTheme
Private m_Index As Long
Private m_Tag As String
Private m_ManualStack As Boolean

' 9个方位独立集合
Private CollCenter As Collection
Private CollCenterTop As Collection
Private CollCenterBottom As Collection
Private CollLeftTop As Collection
Private CollLeftBottom As Collection
Private CollLeftCenter As Collection
Private CollRightTop As Collection
Private CollRightBottom As Collection
Private CollRightCenter As Collection

' 当前窗体实例引用
Private m_InstRef As Object

' 父级 cToast 引用（供窗体回调）
Private m_ParentToast As cToast

' 所有活动弹窗的 Key 集合（方便快速获取）
Private m_AllKeys As Collection

' 初始化
Private Sub Class_Initialize()
    m_Pos = Center
    m_State = Info
    m_Theme = Light
    m_Index = -1
    m_Tag = ""
    m_ManualStack = False
    Set m_InstRef = Nothing
    
    ' 初始化所有集合
    Set CollCenter = New Collection
    Set CollCenterTop = New Collection
    Set CollCenterBottom = New Collection
    Set CollLeftTop = New Collection
    Set CollLeftBottom = New Collection
    Set CollLeftCenter = New Collection
    Set CollRightTop = New Collection
    Set CollRightBottom = New Collection
    Set CollRightCenter = New Collection
    Set m_AllKeys = New Collection
End Sub

' 清理
Private Sub Class_Terminate()
    CloseAll
    Set CollCenter = Nothing
    Set CollCenterTop = Nothing
    Set CollCenterBottom = Nothing
    Set CollLeftTop = Nothing
    Set CollLeftBottom = Nothing
    Set CollLeftCenter = Nothing
    Set CollRightTop = Nothing
    Set CollRightBottom = Nothing
    Set CollRightCenter = Nothing
    Set m_AllKeys = Nothing
End Sub

' ========== 链式调用方法 ==========

Public Function Pos(p As EnumPos) As cToast
    Set Pos = Me
    m_Pos = p
End Function

Public Function State(s As EnumState) As cToast
    Set State = Me
    m_State = s
End Function

Public Function Theme(t As EnumTheme) As cToast
    Set Theme = Me
    m_Theme = t
End Function

Public Function InstIndex(ByVal i As Long) As cToast
    Set InstIndex = Me
    ' 居中/覆盖位置不支持堆叠
    If m_Pos = Center Or m_Pos = LeftCenter Or m_Pos = RightCenter Then Exit Function
    m_Index = i
    m_ManualStack = True
End Function

Public Function Tag(ByVal Name As String) As cToast
    Set Tag = Me
    m_Tag = Name
End Function

Public Function Show(ByVal Content As String, Optional ByVal Delay As Long = 3000, Optional ByVal Title As String = "提示") As cToast
    Dim Inst As Object
    Dim Coll As Collection
    Dim TagName As String
    
    Set Show = Me
    
    ' 生成Tag名称
    If m_Tag = "" Then
        TagName = RandomTag()
    Else
        TagName = m_Tag
    End If
    
    ' 检查是否已存在同名实例
    On Error Resume Next
    m_AllKeys.Item TagName
    If ERR.Number = 0 Then
        ' 同名实例已存在，不创建新实例，直接退出
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' 根据位置获取对应集合和创建实例
    Set Coll = GetCollectionByPos(m_Pos)
    
    ' 自动堆叠管理
    If Not m_ManualStack Then
        ' 用户没有手动调用InstIndex，且集合已有实例，自动堆叠
        If Coll.Count > 0 Then
            m_Index = Coll.Count
        End If
    End If
    
    ' 创建实例
    If IsCenter(m_Pos) Then
        Set Inst = New FToastCenter
        With Inst
            .Init Me, TagName  ' 传入Parent和TagName
            .Pos m_Pos
            .State m_State
            .Theme m_Theme
            If m_Index >= 0 Then .InstIndex m_Index
            .ShowMe Content, Delay, Title
        End With
    Else
        Set Inst = New FToastDrawer
        With Inst
            .Init Me, TagName  ' 传入Parent和TagName
            .Pos m_Pos
            .State m_State
            .Theme m_Theme
            If m_Index >= 0 Then .InstIndex m_Index
            .ShowMe Content, Delay, Title
        End With
    End If
    
    ' 存入集合
    Coll.Add Inst, TagName
    
    ' 同时将 TagName 加入 Keys 集合
    On Error Resume Next
    m_AllKeys.Add TagName, TagName
    On Error GoTo 0
    
    ' 触发数量变化事件（新增）
    RaiseEvent OnToastCountChange(TagName, False, m_AllKeys.Count)
    
    ' 重置状态（为下次Show做准备）
    ResetState
End Function

' ========== 管理方法 ==========

Public Function CloseMe(ByVal Name As String) As Boolean
    Dim Coll As Collection
    Dim i As Integer
    
    ' 遍历所有集合查找
    For i = 1 To 9
        Select Case i
            Case 1: Set Coll = CollCenter
            Case 2: Set Coll = CollCenterTop
            Case 3: Set Coll = CollCenterBottom
            Case 4: Set Coll = CollLeftTop
            Case 5: Set Coll = CollLeftBottom
            Case 6: Set Coll = CollLeftCenter
            Case 7: Set Coll = CollRightTop
            Case 8: Set Coll = CollRightBottom
            Case 9: Set Coll = CollRightCenter
        End Select
        
        On Error Resume Next
        Dim Inst As Object
        Set Inst = Coll(Name)
        On Error GoTo 0
        
        If Not Inst Is Nothing Then
            Inst.CloseMe
            Coll.Remove Name
            ' 同时从 Keys 集合中移除
            m_AllKeys.Remove Name
            ' 所有集合清理完毕后再触发事件，确保 Count 属性准确
            RaiseEvent OnToastCountChange(Name, True, m_AllKeys.Count)
            CloseMe = True
            Exit Function
        End If
    Next i
    
    CloseMe = False
End Function

Public Function CloseAll() As Boolean
    Dim closedCount As Long
    
    ' 记录关闭前的数量
    closedCount = m_AllKeys.Count
    
    ' 关闭Center相关
    CloseCollection CollCenter
    CloseCollection CollCenterTop
    CloseCollection CollCenterBottom
    
    ' 关闭Left相关
    CloseCollection CollLeftTop
    CloseCollection CollLeftBottom
    CloseCollection CollLeftCenter
    
    ' 关闭Right相关
    CloseCollection CollRightTop
    CloseCollection CollRightBottom
    CloseCollection CollRightCenter
    
    ' 清空 Keys 集合
    Set m_AllKeys = New Collection
    
    ' 批量关闭完成后，只触发一次 OnCloseAll 事件
    If closedCount > 0 Then
        RaiseEvent OnCloseAll(closedCount)
    End If
    
    CloseAll = True
End Function

Public Function Exists(ByVal Name As String) As Boolean
    Dim Coll As Collection
    Dim i As Integer
    
    For i = 1 To 9
        Select Case i
            Case 1: Set Coll = CollCenter
            Case 2: Set Coll = CollCenterTop
            Case 3: Set Coll = CollCenterBottom
            Case 4: Set Coll = CollLeftTop
            Case 5: Set Coll = CollLeftBottom
            Case 6: Set Coll = CollLeftCenter
            Case 7: Set Coll = CollRightTop
            Case 8: Set Coll = CollRightBottom
            Case 9: Set Coll = CollRightCenter
        End Select
        
        On Error Resume Next
        Dim Inst As Object
        Set Inst = Coll(Name)
        On Error GoTo 0
        
        If Not Inst Is Nothing Then
            Exists = True
            Exit Function
        End If
    Next i
    
    Exists = False
End Function

' ========== 统计属性 ==========

Public Property Get Count() As Long
    ' 返回所有方位集合的实例总数
    Count = CollCenter.Count + CollCenterTop.Count + CollCenterBottom.Count + _
            CollLeftTop.Count + CollLeftBottom.Count + CollLeftCenter.Count + _
            CollRightTop.Count + CollRightBottom.Count + CollRightCenter.Count
End Property

Public Property Get ActiveKeys() As Collection
    ' 直接返回维护好的 Key 集合
    Set ActiveKeys = m_AllKeys
End Property

' ========== 暂停/恢复方法（供窗体调用） ==========

Public Function Pause(ByVal Name As String) As Boolean
    Dim Inst As Object
    Set Inst = FindInstance(Name)
    If Not Inst Is Nothing Then
        ' 假设窗体有 Timer 控件，这里暂停需要窗体配合
        ' 实际实现需要在 FToastCenter/FToastDrawer 中添加 PauseTimer 方法
        Pause = True
    Else
        Pause = False
    End If
End Function

Public Function ResumeTimer(ByVal Name As String) As Boolean
    Dim Inst As Object
    Set Inst = FindInstance(Name)
    If Not Inst Is Nothing Then
        ResumeTimer = True
    Else
        ResumeTimer = False
    End If
End Function

' ========== 私有辅助方法 ==========

Private Function RandomTag() As String
    Dim chars As String
    Dim result As String
    Dim i As Integer
    Dim idx As Integer
    
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    result = ""
    
    Randomize
    For i = 1 To 8
        idx = Int(Rnd() * Len(chars)) + 1
        result = result & Mid(chars, idx, 1)
    Next i
    
    RandomTag = result
End Function

Private Function GetCollectionByPos(p As EnumPos) As Collection
    Select Case p
        Case Center: Set GetCollectionByPos = CollCenter
        Case CenterTop: Set GetCollectionByPos = CollCenterTop
        Case CenterBottom: Set GetCollectionByPos = CollCenterBottom
        Case LeftTop: Set GetCollectionByPos = CollLeftTop
        Case LeftBottom: Set GetCollectionByPos = CollLeftBottom
        Case LeftCenter: Set GetCollectionByPos = CollLeftCenter
        Case RightTop: Set GetCollectionByPos = CollRightTop
        Case RightBottom: Set GetCollectionByPos = CollRightBottom
        Case RightCenter: Set GetCollectionByPos = CollRightCenter
        Case Else: Set GetCollectionByPos = CollCenter
    End Select
End Function

Private Sub CloseCollection(ByRef Coll As Collection)
    Dim Inst As Object
    On Error Resume Next
    For Each Inst In Coll
        Inst.CloseMe
    Next Inst
    On Error GoTo 0
    Set Coll = New Collection
End Sub

Private Sub ResetState()
    ' 保留m_Pos作为下次的默认位置，其他重置
    m_State = Info
    m_Theme = Light
    m_Index = -1
    m_Tag = ""
    m_ManualStack = False
    Set m_InstRef = Nothing
    Set m_ParentToast = Nothing
End Sub

' ========== 窗体回调方法（Friend，供 FToastCenter/FToastDrawer 调用） ==========

Friend Sub RegisterToast(ByVal TagName As String, ByRef ToastInst As Object, ByRef Parent As cToast)
    ' 窗体初始化时调用，保存TagName和Parent引用
    ' 实际存储在窗体内部，这里只是约定接口
End Sub

Friend Sub UnloadToastForm(ByVal TagName As String)
    ' 窗体销毁时回调，从对应集合中移除
    On Error Resume Next
    
    ' 遍历所有方位集合查找并移除
    Dim Coll As Collection
    Dim i As Integer
    
    For i = 1 To 9
        Select Case i
            Case 1: Set Coll = CollCenter
            Case 2: Set Coll = CollCenterTop
            Case 3: Set Coll = CollCenterBottom
            Case 4: Set Coll = CollLeftTop
            Case 5: Set Coll = CollLeftBottom
            Case 6: Set Coll = CollLeftCenter
            Case 7: Set Coll = CollRightTop
            Case 8: Set Coll = CollRightBottom
            Case 9: Set Coll = CollRightCenter
        End Select
        
        ' 尝试移除，如果存在则成功
        Coll.Remove TagName
    Next i
    
    ' 从 Keys 集合中移除
    m_AllKeys.Remove TagName
    
    ' 所有集合清理完毕后再触发事件，确保 Count 属性准确
    RaiseEvent OnToastCountChange(TagName, True, m_AllKeys.Count)
    
    On Error GoTo 0
End Sub

Friend Sub PauseToast(ByVal TagName As String)
    ' 鼠标悬停时暂停计时器
    Dim Inst As Object
    Set Inst = FindInstance(TagName)
    If Not Inst Is Nothing Then
        ' 假设窗体有 PauseTimer 方法
        On Error Resume Next
        Inst.PauseTimer
        On Error GoTo 0
    End If
End Sub

Friend Sub ResumeToast(ByVal TagName As String)
    ' 鼠标移出时恢复计时器
    Dim Inst As Object
    Set Inst = FindInstance(TagName)
    If Not Inst Is Nothing Then
        On Error Resume Next
        Inst.ResumeTimer
        On Error GoTo 0
    End If
End Sub

Private Function FindInstance(ByVal Name As String) As Object
    Dim Coll As Collection
    Dim i As Integer
    
    For i = 1 To 9
        Select Case i
            Case 1: Set Coll = CollCenter
            Case 2: Set Coll = CollCenterTop
            Case 3: Set Coll = CollCenterBottom
            Case 4: Set Coll = CollLeftTop
            Case 5: Set Coll = CollLeftBottom
            Case 6: Set Coll = CollLeftCenter
            Case 7: Set Coll = CollRightTop
            Case 8: Set Coll = CollRightBottom
            Case 9: Set Coll = CollRightCenter
        End Select
        
        On Error Resume Next
        Dim Inst As Object
        Set Inst = Coll(Name)
        On Error GoTo 0
        
        If Not Inst Is Nothing Then
            Set FindInstance = Inst
            Exit Function
        End If
    Next i
    
    Set FindInstance = Nothing
End Function
