VERSION 5.00
Begin VB.Form FToastDrawer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2760
      Top             =   960
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FToastDrawer.frx":0000
      Top             =   720
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C58B36&
      Height          =   420
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "提示"
      Top             =   240
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00C58B36&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "FToastDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const GWL_EXSTYLE = -20
Private Const GWL_STYLE = -16
Private Const WS_EX_NOACTIVATE = &H8000000
Private Const WS_DISABLED = &H8000000

Dim Sad As cShadow
Dim m_Theme As EnumTheme
Dim m_State As EnumState

Dim PosVal As EnumPos
Dim TopVal As Long
Dim LeftVal As Long

' 父级 cToast 引用和 TagName（用于销毁回调）
Dim m_ParentToast As cToast
Dim m_TagName As String

Public Property Get Self() As FToastDrawer
    Set Self = Me
End Property

' ========== 初始化方法（由 cToast 调用） ==========

Public Sub Init(ByRef Parent As cToast, ByVal TagName As String)
    Set m_ParentToast = Parent
    m_TagName = TagName
End Sub

' ========== Timer 控制方法（供 cToast 调用） ==========

Public Sub PauseTimer()
    Timer1.Enabled = False
End Sub

Public Sub ResumeTimer()
    If Timer1.Interval > 0 Then
        Timer1.Enabled = True
    End If
End Sub

    
    
    Public Function State(s As EnumState) As FToastDrawer
    Set State = Me
    m_State = s
    ' 应用主题颜色
    Select Case m_State
    Case Success
        Text1.ForeColor = &HC000&                                               ' 标题文字
        Picture1.BackColor = &HC000&                                            ' 绿色
    Case Warning
        Text1.ForeColor = &HC0C0&                                               ' 标题文字
        Picture1.BackColor = &HC0C0&                                            ' 黄色
    Case Danger
        Text1.ForeColor = &HC0&                                                 ' 标题文字
        Picture1.BackColor = &HC0&                                              ' 红色
    Case Else
        ' 默认
        Text1.ForeColor = &HC58B36                                              ' 标题文字
        Picture1.BackColor = &HC58B36                                           ' 左侧条蓝色
    End Select
End Function

Public Function Theme(t As EnumTheme) As FToastDrawer
    Set Theme = Me
    m_Theme = t
    
    ' 应用主题颜色
    Select Case m_Theme
    Case Dark
        ' Dark 主题
        Me.BackColor = RGB(45, 45, 48)
        Text1.BackColor = RGB(45, 45, 48)                                       ' 标题背景
        '        Text1.ForeColor = &HFF&                                                 ' 标题文字
        Text2.BackColor = RGB(45, 45, 48)                                       ' 内容背景
        Text2.ForeColor = RGB(200, 200, 200)                                    ' 内容文字
        '        Picture1.BackColor = &HC0&                                              ' 橙色
    Case Else
        ' Light 主题（默认）
        Me.BackColor = &H8000000F
        Text1.BackColor = &H8000000F                                            ' 标题背景
        '        Text1.ForeColor = &HC58B36                                              ' 标题文字
        Text2.BackColor = &H8000000F                                            ' 内容背景
        Text2.ForeColor = &H80000008                                            ' 内容文字浅灰
        '        Picture1.BackColor = &HC58B36                                           ' 左侧条蓝色
    End Select
End Function


Public Function InstIndex(ByVal i As Long) As FToastDrawer
    Set InstIndex = Me
    ' 居中位置不支持堆叠
    If PosVal = LeftCenter Or PosVal = RightCenter Then Exit Function
    
    Dim ItemHeight As Long
    ItemHeight = Me.Height + 200 ' 每项高度 + 间距
    
    If PosVal = LeftTop Or PosVal = RightTop Then
        ' 从上往下排列：i=0 在最上面（y 最小）
        TopVal = 800 + (ItemHeight * i)
    ElseIf PosVal = LeftBottom Or PosVal = RightBottom Then
        ' 从下往上排列：i 越大，位置越靠上（y 越小）
        ' 修正：直接使用 i 计算，确保 i=0 在最底部，i=1 在其上方
        TopVal = Screen.Height - Me.Height - 800 - (i * ItemHeight)
    End If
End Function

Public Function Pos(p As EnumPos) As FToastDrawer
    Set Pos = Me
    PosVal = p
    
    ' 设置 LeftVal
    If PosVal = LeftTop Or PosVal = LeftCenter Or PosVal = LeftBottom Then
        LeftVal = 0
    ElseIf PosVal = RightTop Or PosVal = RightCenter Or PosVal = RightBottom Then
        LeftVal = Screen.Width - Me.Width
    End If
    
    ' 设置初始 TopVal
    If PosVal = LeftTop Or PosVal = RightTop Then
        TopVal = 800
    ElseIf PosVal = LeftBottom Or PosVal = RightBottom Then
        TopVal = Screen.Height - Me.Height - 800
    ElseIf PosVal = LeftCenter Or PosVal = RightCenter Then
        ' 居中：TopVal = 0 表示由 ShowMe 计算屏幕中心
        TopVal = 0
    End If
End Function

Public Sub ShowMe(ByVal Content As String, Optional ByVal Delay As Long = 3000, Optional ByVal Title As String = "提示")
    Dim ShowTop As Long
    
    Text1.Text = Title
    Text2.Text = Content
    Timer1.Interval = Delay
    If Delay > 0 Then Timer1.Enabled = True
    If IsLeft(PosVal) = True Then
        Picture1.Align = 4
    End If
    If IsRight(PosVal) = True Then
        Picture1.Align = 3
    End If
    
    ' 如果 TopVal = 0（居中位置），计算屏幕中心
    If TopVal = 0 Then
        ShowTop = (Screen.Height - Me.Height) \ 2
    Else
        ShowTop = TopVal
    End If
    
    Me.Move LeftVal, ShowTop
    
    ' 使用 SetWindowPos 显示窗口（无焦点 + 顶置）
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
    ' 初始化阴影（Form_Activate 未触发时手动调用）
    Sad.ShowBorders Me.hwnd, False
End Sub
Public Sub CloseMe()
    Unload Me
End Sub

' Private Sub Form_Activate()
'     Sad.ShowBorders Me.hwnd, False
' End Sub

Private Sub Form_Initialize()
    PosVal = RightTop
    TopVal = 800
    LeftVal = Screen.Width - Me.Width
End Sub

Private Sub Form_Load()
    
    Set Sad = New cShadow
    
    With Sad
        .BackColor = vbBlack
        '        .BorderColor = vbRed
        .BorderRadius = 0
        .BorderWidth = 0
        '        .ContainerBkColor = vbRed
        .ShadowColor = &H0&
        .ShadowOffsetX = 0
        .ShadowOffsetY = 0
        .ShadowSize = 5
    End With
    
    ' 设置无焦点样式 + 禁用窗口
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_NOACTIVATE
    SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Or WS_DISABLED
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 通知父级 cToast 此窗体正在销毁
    If Not m_ParentToast Is Nothing Then
        If m_TagName <> "" Then
            m_ParentToast.UnloadToastForm m_TagName
        End If
        Set m_ParentToast = Nothing
    End If
    Set Sad = Nothing
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub

' ========== 鼠标悬停事件（暂停/恢复倒计时） ==========

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' 鼠标进入文本框，暂停计时器
    If Not m_ParentToast Is Nothing And m_TagName <> "" Then
        m_ParentToast.PauseToast m_TagName
    End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' 鼠标进入内容文本框，也暂停计时器
    If Not m_ParentToast Is Nothing And m_TagName <> "" Then
        m_ParentToast.PauseToast m_TagName
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' 检测鼠标是否离开窗体区域
    If x < 0 Or y < 0 Or x > Me.ScaleWidth Or y > Me.ScaleHeight Then
        If Not m_ParentToast Is Nothing And m_TagName <> "" Then
            m_ParentToast.ResumeToast m_TagName
        End If
    End If
End Sub
