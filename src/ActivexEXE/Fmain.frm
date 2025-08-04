VERSION 5.00
Begin VB.Form Fmain 
   Caption         =   "VBMAN"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   9450
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5400
      Top             =   4080
   End
   Begin VBMANLIB.ucLogs Logs 
      Height          =   3015
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      _extentx        =   9763
      _extenty        =   5318
   End
   Begin VB.Menu Install 
      Caption         =   "安装"
      Begin VB.Menu InstallService 
         Caption         =   "系统服务"
         Enabled         =   0   'False
      End
      Begin VB.Menu InstallDirMenu 
         Caption         =   "右键启动"
         Enabled         =   0   'False
      End
      Begin VB.Menu InstallUrlCall 
         Caption         =   "URL协议启动"
         Enabled         =   0   'False
      End
      Begin VB.Menu InstallCOM 
         Caption         =   "COM组件"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Servers 
      Caption         =   "服务器"
      Begin VB.Menu ServerHttp 
         Caption         =   "Http/SSE服务器"
      End
      Begin VB.Menu ServerTcp 
         Caption         =   "Tcp服务器"
         Enabled         =   0   'False
      End
      Begin VB.Menu ServerWebsocket 
         Caption         =   "Websocket服务器"
         Enabled         =   0   'False
      End
      Begin VB.Menu ServerTelnet 
         Caption         =   "Telnet服务器"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu LogLevels 
      Caption         =   "日志级别"
      Begin VB.Menu LogLevel 
         Caption         =   "Debugger"
         Index           =   0
      End
      Begin VB.Menu LogLevel 
         Caption         =   "Info"
         Index           =   1
      End
      Begin VB.Menu LogLevel 
         Caption         =   "Warn"
         Index           =   2
      End
      Begin VB.Menu LogLevel 
         Caption         =   "Danger"
         Index           =   3
      End
      Begin VB.Menu LogLevel 
         Caption         =   "Errors"
         Index           =   4
      End
      Begin VB.Menu LogLevel 
         Caption         =   "Custom"
         Index           =   9
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu LogLevelOnly 
         Caption         =   "精确日志级别"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助"
      Begin VB.Menu HelpDoc 
         Caption         =   "使用手册"
      End
      Begin VB.Menu HelpAbout 
         Caption         =   "关于软件"
      End
   End
End
Attribute VB_Name = "Fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HttpServer As New cHttpServer



Private Sub Form_Load()
    Me.Caption = "VBMAN " & Common.Version
    Logs.Add Info, "欢迎使用VBMAN服务器"
    LogLevel(Logs.LogLevel).Checked = True
End Sub

Private Sub Form_Resize()
    Logs.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub HelpAbout_Click()
    With FLayer
        .IsBlurClose = True
        .ShowTo "作者：邓伟，邮箱：215879458@qq.com"
    End With
End Sub

Private Sub HelpDoc_Click()
    Shell "explorer https://doc.vb6.pro/vbsman"
End Sub

Private Sub LogLevel_Click(Index As Integer)
    Logs.SetLogLevelMenu LogLevel, Index
End Sub

Private Sub LogLevelOnly_Click()
    LogLevelOnly.Checked = Not LogLevelOnly.Checked
    Logs.LogLevelOnly = LogLevelOnly.Checked
End Sub

Private Sub ServerHttp_Click()
    FSettingHttp.Show 1
End Sub

Private Sub Timer1_Timer()
    Logs.Add Danger, "HHHHH"
    Logs.Add CUSTOM + 1002, "EEEEE"
End Sub
