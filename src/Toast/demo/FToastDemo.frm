VERSION 5.00
Begin VB.Form FToastDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cToast 专业演示 - VBMAN Toast 组件"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10335
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame6 
      Caption         =   "外观设置"
      Height          =   855
      Left            =   5760
      TabIndex        =   40
      Top             =   1920
      Width           =   4455
      Begin VB.OptionButton optTheme 
         Caption         =   "浅色 (Light)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optTheme 
         Caption         =   "深色 (Dark)"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   41
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "一键满屏"
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton btnCloseAll 
      Caption         =   "关闭所有弹窗"
      Height          =   855
      Left            =   1680
      TabIndex        =   38
      Top             =   960
      Width           =   3975
   End
   Begin VB.Frame Frame4 
      Caption         =   "已命名弹窗管理"
      Height          =   2535
      Left            =   5760
      TabIndex        =   25
      Top             =   4200
      Width           =   4455
      Begin VB.ListBox lstActiveToasts 
         Height          =   1320
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton btnCloseSelected 
         Caption         =   "关闭选中"
         Height          =   495
         Left            =   2880
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblCount 
         Caption         =   "当前: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "快速测试"
      Height          =   1815
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   5535
      Begin VB.CommandButton btnQuickStack 
         Caption         =   "自动堆叠测试 (5个)"
         Height          =   495
         Left            =   2880
         TabIndex        =   24
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton btnQuickBottom 
         Caption         =   "快速底部倒序"
         Height          =   495
         Left            =   2880
         TabIndex        =   23
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton btnQuickTop 
         Caption         =   "快速顶部顺序"
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton btnQuickCenter 
         Caption         =   "快速居中显示"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "消息设置"
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   5535
      Begin VB.TextBox txtTitle 
         Height          =   375
         Left            =   960
         TabIndex        =   18
         Text            =   "提示"
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtContent 
         Height          =   855
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "FToastDemo.frx":0000
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtDelay 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Text            =   "3000"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtTag 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox chkAutoTag 
         Caption         =   "自动生成命名"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoStack 
         Caption         =   "自动堆叠"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtManualIndex 
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "标题:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "内容:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "延迟(ms)"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "命名:"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "手动索引"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "位置选择"
      Height          =   1695
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton optPos 
         Caption         =   "右中"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "右下"
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "右上"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "左中"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "左下"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "左上"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "居中"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "居下"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optPos 
         Caption         =   "居上"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   32
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "外观设置"
      Height          =   1215
      Left            =   5760
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
      Begin VB.OptionButton optState 
         Caption         =   "危险 (Danger)"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optState 
         Caption         =   "警告 (Warning)"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optState 
         Caption         =   "成功 (Success)"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   35
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optState 
         Caption         =   "信息 (Info)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton btnShow 
      Caption         =   "显示弹窗"
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FToastDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' cToast 实例（使用 WithEvents 监听事件）
Private WithEvents Toast As cToast
Attribute Toast.VB_VarHelpID = -1

' 位置枚举数组（与 optPos Index 对应）
Private PosArray(0 To 8) As EnumPos

Private Sub Command1_Click()
    '===========================================
    ' 一键满屏显示 - 全量展示所有9个方位
    '===========================================
    
    '---------- Center 系列（使用 FToastCenter）----------
    
    ' 居中覆盖 - 信息状态
    Toast.Pos(Center).State(Info).Theme(Light).Tag("center_info").Show "居中覆盖 - 信息提示", 0, "Center"
    
    ' 居上顺序堆叠 - 成功状态（第1条）
    Toast.Pos(CenterTop).State(Success).Theme(Light).Tag("ctop_1").Show "居上顺序 - 第1条", 0, "CenterTop"
    ' 居上顺序堆叠 - 警告状态（第2条）
    Toast.State(Warning).Theme(Dark).Tag("ctop_2").Show "居上顺序 - 第2条", 0, "CenterTop"
    
    ' 居下倒序堆叠 - 危险状态（第1条，实际在最下）
    Toast.Pos(CenterBottom).State(Danger).Theme(Light).Tag("cbottom_1").Show "居下倒序 - 第1条(下)", 0, "CenterBottom"
    ' 居下倒序堆叠 - 信息状态（第2条，实际在上）
    Toast.State(Info).Tag("cbottom_2").Show "居下倒序 - 第2条(上)", 0, "CenterBottom"
    
    '---------- Left 系列（使用 FToastDrawer）----------
    
    ' 左上堆叠 - 成功状态
    Toast.Pos(LeftTop).State(Success).Theme(Light).Tag("ltop_1").Show "左上堆叠 - 成功", 0, "LeftTop"
    Toast.State(Warning).Tag("ltop_2").Show "左上堆叠 - 警告", 0, "LeftTop"
    
    ' 左下堆叠 - 信息状态
    Toast.Pos(LeftBottom).State(Info).Theme(Dark).Tag("lbottom_1").Show "左下堆叠 - 深色主题", 0, "LeftBottom"
    Toast.State(Danger).Tag("lbottom_2").Show "左下堆叠 - 危险", 0, "LeftBottom"
    
    ' 左中覆盖 - 警告状态
    Toast.Pos(LeftCenter).State(Warning).Theme(Light).Tag("lcenter").Show "左中覆盖 - 警告提示", 0, "LeftCenter"
    
    '---------- Right 系列（使用 FToastDrawer）----------
    
    ' 右上堆叠 - 信息状态
    Toast.Pos(RightTop).State(Info).Theme(Light).Tag("rtop_1").Show "右上堆叠 - 信息", 0, "RightTop"
    Toast.State(Success).Tag("rtop_2").Show "右上堆叠 - 成功", 0, "RightTop"
    
    ' 右下堆叠 - 危险状态
    Toast.Pos(RightBottom).State(Danger).Theme(Light).Tag("rbottom_1").Show "右下堆叠 - 危险", 0, "RightBottom"
    Toast.State(Info).Theme(Dark).Tag("rbottom_2").Show "右下堆叠 - 深色", 0, "RightBottom"
    
    ' 右中覆盖 - 成功状态
    Toast.Pos(RightCenter).State(Success).Theme(Light).Tag("rcenter").Show "右中覆盖 - 成功提示", 0, "RightCenter"
    
    ' 触发一次刷新（部分弹窗可能因命名重复未创建）
    RefreshActiveList
End Sub

Private Sub Form_Load()
    ' 初始化 cToast
    Set Toast = New cToast
    
    ' 初始化位置数组
    PosArray(0) = CenterTop
    PosArray(1) = CenterBottom
    PosArray(2) = Center
    PosArray(3) = LeftTop
    PosArray(4) = LeftBottom
    PosArray(5) = LeftCenter
    PosArray(6) = RightTop
    PosArray(7) = RightBottom
    PosArray(8) = RightCenter
    
    '全局单实例对象 VBMAN 中可以在任何地方免New调用
'    VBMAN.Toast.Theme(Dark).Show "你好，来自任何地方的问候", , "VBMAN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 关闭所有弹窗
    If Not Toast Is Nothing Then
        Toast.CloseAll
    End If
    Set Toast = Nothing
End Sub

' ========== 事件处理 ==========

Private Sub Toast_OnToastCountChange(ByVal TagName As String, ByVal IsDelete As Boolean, ByVal CurrentCount As Long)
    ' 单个弹窗数量变化时自动刷新列表
    ' 实际可以通过 TagName 进行更细粒度的控制列表增改，这里偷懒
    RefreshActiveList
End Sub

Private Sub Toast_OnCloseAll(ByVal closedCount As Long)
    ' 批量关闭完成后触发（只触发一次，避免频繁刷新）
    RefreshActiveList
    ' 可选：显示提示
    ' Debug.Print "批量关闭了 " & ClosedCount & " 个弹窗"
End Sub

' ========== 主要按钮 ==========

Private Sub btnShow_Click()
    On Error GoTo ErrorHandler
    
    Dim posIdx As Integer
    Dim stateIdx As Integer
    Dim themeIdx As Integer
    Dim delayMs As Long
    Dim Content As String
    Dim Title As String
    Dim TagName As String
    Dim useAutoStack As Boolean
    Dim useAutoTag As Boolean
    Dim manualIdx As Long
    
    ' 获取选中的位置
    For posIdx = 0 To 8
        If optPos(posIdx).Value Then Exit For
    Next posIdx
    If posIdx > 8 Then posIdx = 0
    
    ' 获取选中的状态
    For stateIdx = 0 To 3
        If optState(stateIdx).Value Then Exit For
    Next stateIdx
    If stateIdx > 3 Then stateIdx = 0
    
    ' 获取选中的主题
    For themeIdx = 0 To 1
        If optTheme(themeIdx).Value Then Exit For
    Next themeIdx
    If themeIdx > 1 Then themeIdx = 0
    
    ' 获取其他参数
    delayMs = val(txtDelay.Text)
    If delayMs < 0 Then delayMs = 3000
    Content = txtContent.Text
    If Content = "" Then Content = "测试消息"
    Title = txtTitle.Text
    If Title = "" Then Title = "提示"
    
    useAutoStack = (chkAutoStack.Value = vbChecked)
    useAutoTag = (chkAutoTag.Value = vbChecked)
    
    ' 设置位置
    Toast.Pos PosArray(posIdx)
    
    ' 设置状态
    Select Case stateIdx
        Case 0: Toast.State Info
        Case 1: Toast.State Success
        Case 2: Toast.State Warning
        Case 3: Toast.State Danger
    End Select
    
    ' 设置主题
    Select Case themeIdx
        Case 0: Toast.Theme Light
        Case 1: Toast.Theme Dark
    End Select
    
    ' 设置Tag（如果不使用自动生成）
    If Not useAutoTag Then
        TagName = Trim(txtTag.Text)
        If TagName <> "" Then
            Toast.Tag TagName
        End If
    End If
    
    ' 设置手动堆叠索引
    If Not useAutoStack Then
        manualIdx = val(txtManualIndex.Text)
        Toast.InstIndex manualIdx
    End If
    
    ' 显示弹窗（事件会自动触发刷新）
    Toast.Show Content, delayMs, Title
    
    Exit Sub
    
ErrorHandler:
    MsgBox "显示弹窗时出错: " & Err.Description, vbCritical, "错误"
End Sub

Private Sub btnCloseAll_Click()
    ' 关闭所有弹窗（事件会自动触发刷新）
    If Not Toast Is Nothing Then
        Toast.CloseAll
    End If
End Sub

' ========== 快速测试按钮 ==========

Private Sub btnQuickCenter_Click()
    ' 快速居中显示（事件会自动触发刷新）
    Toast.Pos(Center).State(Info).Theme(Light).Show "快速居中测试消息", 3000, "居中"
End Sub

Private Sub btnQuickTop_Click()
    ' 快速顶部顺序堆叠（事件会自动触发刷新）
    Toast.Pos(CenterTop).State(Success).Show "顶部第1条", 0, "顺序"
    Toast.State(Success).Show "顶部第2条", 0, "顺序"
    Toast.State(Warning).Show "顶部第3条", 0, "顺序"
End Sub

Private Sub btnQuickBottom_Click()
    ' 快速底部倒序堆叠（事件会自动触发刷新）
    Toast.Pos(CenterBottom).State(Info).Show "底部第1条 (最下)", 0, "倒序"
    Toast.State(Warning).Show "底部第2条", 0, "倒序"
    Toast.State(Danger).Show "底部第3条 (最上)", 0, "倒序"
End Sub

Private Sub btnQuickStack_Click()
    ' 自动堆叠测试 - 5个消息（事件会自动触发刷新）
    Dim i As Integer
    Dim titles(0 To 4) As String
    Dim contents(0 To 4) As String
    
    titles(0) = "第1条"
    titles(1) = "第2条"
    titles(2) = "第3条"
    titles(3) = "第4条"
    titles(4) = "第5条"
    
    contents(0) = "这是自动堆叠的第一条消息"
    contents(1) = "第二条消息会自动堆叠"
    contents(2) = "第三条继续往上堆"
    contents(3) = "第四条"
    contents(4) = "第五条完成堆叠"
    
    Toast.Pos(RightTop).State (Info)
    For i = 0 To 4
        Toast.State Choose(i + 1, Info, Success, Warning, Danger, Info)
        Toast.Show contents(i), 5000, titles(i)
    Next i
End Sub

' ========== 弹窗列表管理 ==========

Private Sub btnCloseSelected_Click()
    Dim idx As Integer
    idx = lstActiveToasts.ListIndex
    If idx < 0 Then
        MsgBox "请先选择一个弹窗", vbInformation, "提示"
        Exit Sub
    End If
    
    Dim TagName As String
    TagName = lstActiveToasts.List(idx)
    
    ' 关闭选中弹窗（事件会自动触发刷新）
    If Not Toast.CloseMe(TagName) Then
'        MsgBox "关闭失败，弹窗可能已自动关闭", vbExclamation, "提示"
'        RefreshActiveList
    End If
End Sub

Private Sub RefreshActiveList()
    ' 从 cToast 获取活动弹窗列表
    Dim Keys As Collection
    Dim Key As Variant
    Dim Count As Long
    
    lstActiveToasts.Clear
    
    If Toast Is Nothing Then
        lblCount.Caption = "当前: 0"
        Exit Sub
    End If
    
    Count = Toast.Count
    lblCount.Caption = "当前: " & Count
    
    If Count = 0 Then
        lstActiveToasts.AddItem "(暂无活动弹窗)"
        Exit Sub
    End If
    
    ' 获取所有 Key
    Set Keys = Toast.ActiveKeys
    
    ' 通过 Keys.Count 判断是否为空
    If Keys.Count > 0 Then
        For Each Key In Keys
            lstActiveToasts.AddItem CStr(Key)
        Next Key
    End If
End Sub

' ========== 控件事件 ==========

Private Sub chkAutoStack_Click()
    ' 自动堆叠选中时禁用手动索引输入
    txtManualIndex.Enabled = (chkAutoStack.Value = vbUnchecked)
End Sub

Private Sub chkAutoTag_Click()
    ' 自动生成Tag选中时禁用Tag输入
    txtTag.Enabled = (chkAutoTag.Value = vbUnchecked)
End Sub

Private Sub optPos_Click(Index As Integer)
    ' 居中位置时不支持堆叠
    If Index = 2 Then ' Center
        chkAutoStack.Value = vbUnchecked
        chkAutoStack.Enabled = False
        txtManualIndex.Enabled = False
    Else
        chkAutoStack.Enabled = True
        chkAutoStack.Value = vbChecked
    End If
End Sub

Private Sub txtDelay_Change()
    ' 确保输入的是数字
    Dim val As Long
    On Error Resume Next
    val = val(txtDelay.Text)
    If Err.Number <> 0 Then txtDelay.Text = "3000"
    On Error GoTo 0
End Sub

Private Sub txtManualIndex_Change()
    ' 确保输入的是非负数字
    Dim val As Long
    On Error Resume Next
    val = val(txtManualIndex.Text)
    If val < 0 Then txtManualIndex.Text = "0"
    On Error GoTo 0
End Sub
