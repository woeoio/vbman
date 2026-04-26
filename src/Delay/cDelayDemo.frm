VERSION 5.00
Begin VB.Form frmDelayDemo 
   Caption         =   "cDelay 演示"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6015
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btnSync 
      Caption         =   "同步等待3秒"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "取消延迟"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton btnCallback 
      Caption         =   "回调模式(3秒)"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton btnEvent 
      Caption         =   "事件模式(3秒)"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "准备就绪"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDelayDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===========================================================================
' 名称:    cDelay
' 描述:    延时对象类 - 支持事件触发、回调函数和同步等待三种模式
' 作者:    邓伟，QQ: 215879458
' 网站:    https://vb6.pro
' 日期:    2026-03-31
' 参考:    cTimer.cls 架构
'===========================================================================

Option Explicit

Private WithEvents m_Delay As cDelay
Attribute m_Delay.VB_VarHelpID = -1
Private m_DelayCallback As cDelay
Private m_DelaySync As cDelay           '同步模式需要模块级变量才能取消

'----------------------------------
' 窗体加载
'----------------------------------
Private Sub Form_Load()
    '系统会自动初始化，无需手动调用 InitDelaySystem
    Set m_Delay = New cDelay
    Set m_DelayCallback = New cDelay
    Set m_DelaySync = New cDelay       '同步模式对象
End Sub

'----------------------------------
' 窗体卸载
'----------------------------------
Private Sub Form_Unload(Cancel As Integer)
    m_Delay.Cancel
    m_DelayCallback.Cancel
    m_DelaySync.Cancel
    Set m_Delay = Nothing
    Set m_DelayCallback = Nothing
    Set m_DelaySync = Nothing
    '系统会自动清理，无需手动调用 TermDelaySystem
End Sub

'----------------------------------
' 按钮: 事件模式
'----------------------------------
Private Sub btnEvent_Click()
    lblStatus.Caption = "事件模式: 开始3秒倒计时..." & vbCrLf & _
                       "时间: " & Now
    m_Delay.CountDown 3000
End Sub

'----------------------------------
' 按钮: 回调模式
'----------------------------------
Private Sub btnCallback_Click()
    lblStatus.Caption = "回调模式: 开始3秒倒计时..." & vbCrLf & _
                       "将调用 MyCallbackFunc 并传递参数" & vbCrLf & _
                       "时间: " & Now
    
    '使用回调模式，传递3个参数
    m_DelayCallback.Callback Me, "MyCallbackFunc", "Hello", 123, Now
    m_DelayCallback.CountDown 3000
End Sub

'----------------------------------
' 按钮: 取消延迟
'----------------------------------
Private Sub btnCancel_Click()
    m_Delay.Cancel
    m_DelayCallback.Cancel
    m_DelaySync.Cancel              '同步模式也可以取消
    lblStatus.Caption = "延迟已取消" & vbCrLf & "时间: " & Now
End Sub

'----------------------------------
' 按钮: 同步等待
'----------------------------------
Private Sub btnSync_Click()
    lblStatus.Caption = "同步模式: 开始3秒等待..." & vbCrLf & _
    "UI不会卡住，可以移动窗口" & vbCrLf & _
    "点击【取消延迟】可以提前结束" & vbCrLf & _
    "时间 " & Now
    lblStatus.Refresh
    
    '使用模块级变量，这样其他按钮可以调用 Cancel
    m_DelaySync.Sync().CountDown 3000
    
    MsgBox "时间到"
    
    '3秒后或被取消后才会执行到这里
    If m_DelaySync.IsCancelled Then
        lblStatus.Caption = "同步等待被取消!" & vbCrLf & _
        "时间    " & Now
    Else
        lblStatus.Caption = "同步等待完成!" & vbCrLf & _
        "时间    " & Now
    End If
End Sub

'----------------------------------
' 延迟事件处理 - 事件模式
'----------------------------------
Private Sub m_Delay_OnTime()
    lblStatus.Caption = "事件触发! 3秒倒计时结束!" & vbCrLf & _
                       "这是 OnTime 事件处理程序" & vbCrLf & _
                       "时间: " & Now
    Beep
End Sub

'----------------------------------
' 延迟取消事件
'----------------------------------
Private Sub m_Delay_OnCancel()
    lblStatus.Caption = "延迟被取消!" & vbCrLf & "时间: " & Now
End Sub

'----------------------------------
' 回调函数 - 供回调模式调用
'----------------------------------
Public Sub MyCallbackFunc(ByVal msg As String, ByVal num As Long, ByVal dt As Date)
    lblStatus.Caption = "回调函数被调用!" & vbCrLf & _
                       "参数1: " & msg & vbCrLf & _
                       "参数2: " & num & vbCrLf & _
                       "参数3: " & dt & vbCrLf & _
                       "时间: " & Now
    Beep
End Sub
