VERSION 5.00
Begin VB.UserControl ucLogs 
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ScaleHeight     =   6030
   ScaleWidth      =   8850
   Begin VB.TextBox Tlogs 
      Height          =   3495
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "ucLogs.ctx":0000
      Top             =   840
      Width           =   6495
   End
End
Attribute VB_Name = "ucLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim Logs As New cLogs

Const LOGO_1 As String = " __     __  ____    __  __      _      _   _ " & vbCrLf & _
    " \ \   / / | __ )  |  \/  |    / \    | \ | |" & vbCrLf & _
    "  \ \ / /  |  _ \  | |\/| |   / _ \   |  \| |" & vbCrLf & _
    "   \ V /   | |_) | | |  | |  / ___ \  | |\  |" & vbCrLf & _
    "    \_/    |____/  |_|  |_| /_/   \_\ |_| \_|" & vbCrLf & _
    "                                             " & vbCrLf

Public Enum EnumLevel
    Debugger = 0
    INFO = 1
    Warn = 2
    Danger = 3
    Errors = 4
    CUSTOM = 9
End Enum
Dim EnumLevelNames As Variant

''日志文件名规则
Public Enum EnumLogFileNameRule2
    None = 0
    ByMonth = 1
    ByDay = 2
    byUser = 3
End Enum
'日志文件根目录，为空则使用当前程序的目录
Public LogDir As String
'日志文件子目录，一般用于区分业务模块的日志
Public LogSubDir As String
'日志文件名称规则，默认按月份细分，也可以按日，或者用户定义
Public LogFileNameRule As EnumLogFileNameRule2
'日志文件名称用户定义，支持任意细分深层目录和文件名
Public LogFileNameByUer As String

Public LogLevel As EnumLevel
Public MaxContentShow As Long
Public LogLevelOnly As Boolean
Public IsSaveToFile As Boolean

Public Property Get EnumLevelName(Level As EnumLevel) As String
    If Level <= CUSTOM Then
        EnumLevelName = EnumLevelNames(Level)
    Else
        EnumLevelName = "CUSTOM+" & (Level - EnumLevel.CUSTOM)
    End If
End Property

Public Sub SetLogLevelMenu(MenuObj As Object, Index As Integer)
    Dim i As Long
    Const CUSTOM As Long = EnumLevel.CUSTOM
    For i = 0 To EnumLevel.Errors
        MenuObj(i).Checked = False
    Next
    MenuObj(CUSTOM).Checked = False
    '
    Dim n As Long, S As String
    If Index = CUSTOM Then
        S = InputBox("请输入自定义日志级别（正整数）", "设置日志级别", "0")
        If S <> "" Then n = CLng(S)
        If n = 0 Then
            MenuObj(CUSTOM).Caption = "Custom"
        Else
            MenuObj(CUSTOM).Caption = "Custom+" & S
        End If
    End If
    LogLevel = Index + n
    MenuObj(Index).Checked = True
End Sub

Public Sub Add(Level As EnumLevel, ByVal Title As String, ParamArray Contents() As Variant)
    If LogLevelOnly = True Then
        If Level <> LogLevel Then Exit Sub
    End If
    If Level < LogLevel Then Exit Sub
    '
    Dim Content As String: Content = Format(Now(), "yyyy-MM-ddThh:nn:ss") & " [" & EnumLevelName(Level) & "] " & Title ' & vbCrLf
    If UBound(Contents) > -1 Then
        Content = Content & vbCrLf & Join(Contents, vbCrLf)
    End If
    '写入文件
    If IsSaveToFile = True Then Save Content
    '显示到控件
    Content = Content & vbCrLf
    With Tlogs
        ' 使用API直接追加文本，避免字符串连接
        Const EM_SETSEL = &HB1
        Const EM_REPLACESEL = &HC2
        
        ' 移动到文本末尾
        SendMessage .hwnd, EM_SETSEL, ByVal LenB(.Text), ByVal LenB(.Text)
        
        ' 插入新文本
        SendMessage .hwnd, EM_REPLACESEL, 0, ByVal Content
        
        ' 滚动到底部
        Const WM_VSCROLL = &H115
        Const SB_BOTTOM = 7
        SendMessage .hwnd, WM_VSCROLL, SB_BOTTOM, 0
        
    End With
End Sub
Private Function Save(Content As String) As Boolean
    Dim LogPath As String
    Dim LogFile As String
    '默认值处理
    If LogSubDir = "" Then LogSubDir = "defalut"
    If LogFileNameByUer = "" Then LogFileNameByUer = "vb.txt"
    If LogFileNameRule = None Then LogFileNameRule = ByDay
    
    '    If MakeLogData(LogStr, LogLevel) = False Then Exit Function
    
    If LogFileNameRule = byUser Then
        LogPath = LogSubDir & "\"
        LogFile = LogFileNameByUer
    Else
        Select Case LogFileNameRule
        Case EnumLogFileNameRule.ByMonth
            LogPath = Format$(Now, "yyyy") & "\"
            LogFile = "\" & Format$(Now, "yyyyMM") & ".txt"
        Case EnumLogFileNameRule.ByDay
            LogPath = Format$(Now, "yyyy") & "\" & Format$(Now, "MM") & "\"
            LogFile = Format$(Now, "yyyyMMdd") & ".txt"
        Case Else
            LogPath = Format$(Now, "yyyy") & "\" & Format$(Now, "MM") & "\"
            LogFile = LogFileNameByUer & ".txt"
        End Select
        LogPath = "\logs\" + LogSubDir + "\" + LogPath
    End If
    If LogDir = "" Then
        LogPath = App.Path & "\" & LogPath
    Else
        LogPath = LogDir & "\" & LogPath
    End If
    If Dir(LogPath, vbDirectory) = "" Then                                      '判断文件夹是否存在
        MakeSureDirectoryPathExists LogPath                                     '创建文件夹
    End If
    Dim FileNumber As Long: FileNumber = FreeFile
    Open LogPath & "\" & LogFile For Append As #FileNumber
    Print #FileNumber, Content
    Close #FileNumber
    
    Save = True
End Function

Private Sub UserControl_Initialize()
    MaxContentShow = 65535
    EnumLevelNames = Array("DEBUGGER", "INFO", "WARN", "DANGER", "ERRORS", "", "", "", "", "CUSTOM")
    'Tlogs.Text = LOGO_1
    If App.LogMode = 0 Then
        LogDir = App.Path & "\..\dist\EXE\"
    Else
        LogLevel = INFO
    End If
    '
    Add INFO, "欢迎使用VBMAN服务器"
    Add INFO, "当前版本：" & Common.Version()
    Add INFO, "开发文档：http://doc.vb6.pro"
End Sub

Private Sub UserControl_Resize()
    Tlogs.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Function IsInDesignMode() As Boolean
    On Error Resume Next
    IsInDesignMode = (Not Ambient.UserMode)
    If ERR Then IsInDesignMode = False
End Function
