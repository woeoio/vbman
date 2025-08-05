VERSION 5.00
Begin VB.Form FSettingHttp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "配置Http/SSE服务器"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7365
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "配置"
      Height          =   4575
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   4455
      Begin VB.TextBox TSSE 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Text            =   "/sse"
         Top             =   3360
         Width           =   4200
      End
      Begin VB.TextBox TWebRoot 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   "www"
         Top             =   2400
         Width           =   4200
      End
      Begin VB.TextBox TPort 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "80"
         Top             =   1560
         Width           =   4200
      End
      Begin VB.CheckBox ChkIsLogSave 
         Caption         =   "写入日志到文件"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3960
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "启动"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox TAddress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "0.0.0.0"
         Top             =   720
         Width           =   4200
      End
      Begin VB.Label Label1 
         Caption         =   "SSE路径：(为空泽不启用）"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   4200
      End
      Begin VB.Label Label4 
         Caption         =   "目录：(支持绝对路径，可为空）"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   4200
      End
      Begin VB.Label Label3 
         Caption         =   "监听端口："
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   4200
      End
      Begin VB.Label Label2 
         Caption         =   "监听地址："
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "服务器列表"
      Enabled         =   0   'False
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   3990
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FSettingHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Fmain.StartServer(TAddress.Text, TPort.Text, TWebRoot.Text, TSSE.Text, ChkIsLogSave.Value) = True Then Unload Me
End Sub

Private Sub Form_Load()
    Lang.Render Me
End Sub
