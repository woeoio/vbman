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


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Public Enum EnumLevel
    Debugger = 0
    Info = 1
    Warn = 2
    Danger = 3
    Errors = 4
    CUSTOM = 9
End Enum
Dim EnumLevelNames As Variant

Public LogLevel As EnumLevel
Public MaxContentShow As Long
Public LogLevelOnly As Boolean

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
    Dim n As Long, s As String
    If Index = CUSTOM Then
        s = InputBox("请输入自定义日志级别（正整数）", "设置日志级别", "0")
        If s <> "" Then n = CLng(s)
        If n = 0 Then
            MenuObj(CUSTOM).Caption = "Custom"
        Else
            MenuObj(CUSTOM).Caption = "Custom+" & s
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
    Dim Content As String: Content = Format(Now(), "yyyy-MM-ddThh:nn:ss") & " [" & EnumLevelName(Level) & "] " & Title & vbCrLf
    If UBound(Contents) > -1 Then
        Content = Content & Join(Contents, vbCrLf) & vbCrLf
    End If
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


Private Sub UserControl_Initialize()
    MaxContentShow = 65535
    EnumLevelNames = Array("DEBUGGER", "INFO", "WARN", "DANGER", "ERRORS", "", "", "", "", "CUSTOM")
    If App.LogMode <> 0 Then LogLevel = Info
End Sub

Private Sub UserControl_Resize()
    Tlogs.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub
