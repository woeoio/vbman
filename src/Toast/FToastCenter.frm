VERSION 5.00
Begin VB.Form FToastCenter 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   ScaleHeight     =   630
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C58B36&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   5325
      TabIndex        =   1
      Top             =   510
      Width           =   5325
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3600
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "FToastCenter"
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
'Dim TopVal As Long
Dim TopVal As Long

' 父级 cToast 引用和 TagName（用于销毁回调）
Dim m_ParentToast As cToast
Dim m_TagName As String

Public Property Get Self() As FToastCenter
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

Public Function State(s As EnumState) As FToastCenter
    Set State = Me
    m_State = s
    ' 应用主题颜色
    Select Case m_State
    Case Success
        Picture1.BackColor = &HC000&                                            ' 绿色
    Case Warning
        Picture1.BackColor = &HC0C0&                                            ' 黄色
    Case Danger
        Picture1.BackColor = &HC0&                                              ' 红色
    Case Else
        ' 默认
        Picture1.BackColor = &HC58B36                                           ' 左侧条蓝色
    End Select
End Function

Public Function Theme(t As EnumTheme) As FToastCenter
    Set Theme = Me
    m_Theme = t
    
    ' 应用主题颜色
    Select Case m_Theme
    Case Dark
        ' Dark 主题：深色背景 + 白色文字
        Me.BackColor = RGB(45, 45, 48)
        Text1.BackColor = RGB(45, 45, 48)
        Text1.ForeColor = RGB(255, 255, 255)
    Case Else
        ' Light 主题（默认）：系统颜色
        Me.BackColor = &H8000000F
        Text1.BackColor = &H8000000F
        Text1.ForeColor = &H80000008
    End Select
End Function

Public Function InstIndex(ByVal i As Long) As FToastCenter
    Set InstIndex = Me
    If PosVal = Center Then Exit Function
    
    Dim ItemHeight As Long
    ItemHeight = Me.Height + 200 ' 每项高度 + 间距
    
    If PosVal = CenterTop Then
        ' 从上往下排列：i=0 在最上面（y 最小）
        TopVal = 800 + (ItemHeight * i)
    ElseIf PosVal = CenterBottom Then
        ' 从下往上排列：i 越大，位置越靠上（y 越小）
        ' 修正：直接使用 i 计算，确保 i=0 在最底部，i=1 在其上方
        TopVal = Screen.Height - Me.Height - 800 - (i * ItemHeight)
    End If
End Function

Public Function Pos(p As EnumPos) As FToastCenter
    Set Pos = Me
    PosVal = p
    
    If PosVal = Center Then
        ' 居中：TopVal = 0 表示由 ShowMe 计算屏幕中心
        TopVal = 0
    ElseIf PosVal = CenterTop Then
        ' 顶部起始位置
        TopVal = 800
    ElseIf PosVal = CenterBottom Then
        ' 底部起始位置（i=0 时的位置，与 InstIndex(0) 一致）
        TopVal = Screen.Height - Me.Height - 800
    End If
End Function

Public Sub ShowMe(ByVal Content As String, Optional ByVal Delay As Long = 3000, Optional ByVal Title As String = "")
    Dim CenterX As Long, CenterY As Long, YPos As Long
    '自动宽度
    Dim w As Long: w = Len(Content) * 240 + 1400
    If w < 160 * 15 Then w = 160 * 15
    If w > Screen.Width Then w = Screen.Width * 0.9
    Me.Width = w
    
    If Title <> "" Then Title = "[" & Title & "] "
    Text1.Text = Title & Content
    Timer1.Interval = Delay
    If Delay > 0 Then Timer1.Enabled = True
    
    ' 计算屏幕中心位置（转换为像素）
    CenterX = (Screen.Width - Me.Width) \ 2 \ Screen.TwipsPerPixelX
    CenterY = (Screen.Height - Me.Height) \ 2 \ Screen.TwipsPerPixelY
    
    ' 如果设置了 TopVal，使用 TopVal 转换为像素坐标，否则使用屏幕中心
    If TopVal > 0 Then
        YPos = TopVal \ Screen.TwipsPerPixelY
    Else
        YPos = CenterY
    End If
    
    ' 移动到屏幕中心并显示窗口（无焦点 + 顶置）
    SetWindowPos Me.hwnd, HWND_TOPMOST, CenterX, YPos, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
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
    '    PosVal = RightTop
    '    TopVal = 800
    '    TopVal = Screen.Width - Me.Width
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

Private Sub Form_Resize()
    Text1.Width = Me.ScaleWidth - 240
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' 检测鼠标是否离开窗体区域
    ' 简化为：如果鼠标在窗体外，恢复计时器
    If x < 0 Or y < 0 Or x > Me.ScaleWidth Or y > Me.ScaleHeight Then
        If Not m_ParentToast Is Nothing And m_TagName <> "" Then
            m_ParentToast.ResumeToast m_TagName
        End If
    End If
End Sub
