VERSION 5.00
Begin VB.Form frmCaptchaDemo 
   Caption         =   "阿里云验证码2.0 - 演示"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   9015
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Caption         =   "验证结果"
      Height          =   3135
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   8535
      Begin VB.TextBox txtResult 
         Height          =   2535
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   360
         Width           =   8055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "配置"
      Height          =   4095
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   8535
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   2520
         Width           =   6735
      End
      Begin VB.ComboBox cboRegion 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtSceneId 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   6735
      End
      Begin VB.TextBox txtSecretKey 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   6735
      End
      Begin VB.TextBox txtAccessKey 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label5 
         Caption         =   "前端验证结果:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "场景ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Region:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "SecretKey:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "AccessKey:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "清空"
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton btnVerify 
      Caption         =   "开始验证"
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      Caption         =   "状态: 就绪"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8280
      Width           =   8535
   End
End
Attribute VB_Name = "frmCaptchaDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' frmCaptchaDemo - 阿里云验证码2.0演示窗体
'
' Purpose: 演示如何使用cAliyunCaptcha类进行验证码验证
'
' Author: Auto
' Date: 2026-02-08
'
'=========================================================================

Option Explicit

' 声明带事件的验证码对象
Private WithEvents m_Captcha As cAliyunCaptcha
Attribute m_Captcha.VB_VarHelpID = -1

' 窗体加载
Private Sub Form_Load()
    On Error Resume Next
    
    ' 初始化地域下拉框
    With cboRegion
        .Clear
        .AddItem "中国内地 (cn)"
        .ItemData(.NewIndex) = ALIYUN_CAPTCHA_REGION_CN
        .AddItem "新加坡 (sgp)"
        .ItemData(.NewIndex) = ALIYUN_CAPTCHA_REGION_SGP
        .ListIndex = 0
    End With
    
    ' 创建验证码对象
    Set m_Captcha = New cAliyunCaptcha
    
    ' 加载配置（示例）
    txtAccessKey.Text = GetSetting("AliyunCaptcha", "Config", "AccessKey", "")
    txtSecretKey.Text = GetSetting("AliyunCaptcha", "Config", "SecretKey", "")
    txtSceneId.Text = GetSetting("AliyunCaptcha", "Config", "SceneId", "")
    cboRegion.ListIndex = GetSetting("AliyunCaptcha", "Config", "RegionIndex", 0)
    
    lblStatus.Caption = "状态: 就绪"
    
    On Error GoTo 0
End Sub

' 窗体卸载
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' 保存配置
    SaveSetting "AliyunCaptcha", "Config", "AccessKey", txtAccessKey.Text
    SaveSetting "AliyunCaptcha", "Config", "SecretKey", txtSecretKey.Text
    SaveSetting "AliyunCaptcha", "Config", "SceneId", txtSceneId.Text
    SaveSetting "AliyunCaptcha", "Config", "RegionIndex", cboRegion.ListIndex
    
    ' 释放对象
    Set m_Captcha = Nothing
    
    On Error GoTo 0
End Sub

' 开始验证按钮
Private Sub btnVerify_Click()
    '    On Error GoTo ErrorHandler
    
    ' 参数验证
    If LenB(txtAccessKey.Text) = 0 Then
        MsgBox "请输入AccessKey", vbExclamation, "提示"
        txtAccessKey.SetFocus
        Exit Sub
    End If
    
    If LenB(txtSecretKey.Text) = 0 Then
        MsgBox "请输入SecretKey", vbExclamation, "提示"
        txtSecretKey.SetFocus
        Exit Sub
    End If
    
    If LenB(txtSceneId.Text) = 0 Then
        MsgBox "请输入SceneId", vbExclamation, "提示"
        txtSceneId.SetFocus
        Exit Sub
    End If
    
    If LenB(Text1.Text) = 0 Then
        MsgBox "请输入前端验证值", vbExclamation, "提示"
        txtSceneId.SetFocus
        Exit Sub
    End If
    
    ' 配置验证码对象
    With m_Captcha
        .AccessKeyId(txtAccessKey.Text) _
        .AccessKeySecret(txtSecretKey.Text) _
        .Region(cboRegion.ItemData(cboRegion.ListIndex)) _
        .Timeout(30000) _
        .EnableDebug (True)
    End With
    
    lblStatus.Caption = "状态: 正在验证..."
    
    ' 显示输入对话框获取CaptchaVerifyParam
    Dim CaptchaVerifyParam As String
    CaptchaVerifyParam = Text1.Text
    
    If LenB(CaptchaVerifyParam) = 0 Then
        lblStatus.Caption = "状态: 已取消"
        Exit Sub
    End If
    
    ' 执行验证
    Dim Result As AliyunCaptchaVerifyResult
    Result = m_Captcha.Verify(CaptchaVerifyParam, txtSceneId.Text)
    
    ' 显示结果
    Call DisplayResult(Result)
    
    lblStatus.Caption = "状态: 验证完成"
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "状态: 错误 - " & Err.Description
    MsgBox "验证失败: " & Err.Description, vbCritical, "错误"
End Sub

' 清空按钮
Private Sub btnClear_Click()
    txtResult.Text = ""
    Text1.Text = ""
    lblStatus.Caption = "状态: 已清空"
End Sub

' 验证成功事件
Private Sub m_Captcha_OnVerifySuccess(ByVal VerifyCode As String, _
                                     ByVal CertifyID As String, _
                                     ByVal Result As Boolean)
    Dim Msg As String
    Msg = "验证成功！" & vbCrLf & _
          "VerifyCode: " & VerifyCode & vbCrLf & _
          "CertifyID: " & CertifyID & vbCrLf & _
          "Result: " & IIf(Result, "True", "False")
    MsgBox Msg, vbInformation, "验证结果"
End Sub

' 验证失败事件
Private Sub m_Captcha_OnVerifyFailure(ByVal ErrorCode As String, _
                                     ByVal ErrorMessage As String)
    Dim Msg As String
    Msg = "验证失败！" & vbCrLf & _
          "ErrorCode: " & ErrorCode & vbCrLf & _
          "Message: " & ErrorMessage
    MsgBox Msg, vbExclamation, "验证失败"
End Sub

' 错误事件
Private Sub m_Captcha_OnError(ByVal Description As String)
    MsgBox "发生错误：" & Description, vbCritical, "错误"
    lblStatus.Caption = "状态: 错误 - " & Description
End Sub

' 显示验证结果
Private Sub DisplayResult(ByRef Result As AliyunCaptchaVerifyResult)
    Dim Output As String
    
    Output = "========== 验证结果 ==========" & vbCrLf & _
"请求成功: " & IIf(Result.Success, "是", "否") & vbCrLf & _
"验证结果: " & IIf(Result.Result, "通过", "未通过") & vbCrLf & _
"VerifyCode: " & Result.VerifyCode & vbCrLf & _
"CertifyID: " & Result.CertifyID & vbCrLf & _
"RequestId: " & Result.requestID & vbCrLf
    
    If LenB(Result.Message) > 0 Then
        Output = Output & "错误信息: " & Result.Message & vbCrLf
    End If
    
    ' 添加错误码描述
    If LenB(Result.VerifyCode) > 0 Then
        Output = Output & "错误描述: " & Result.VerifyCode & vbCrLf
    End If
    
    Output = Output & "=============================="
    
    txtResult.Text = Output
End Sub

'=========================================================================
' 示例：使用同步验证
'=========================================================================
Private Sub DemoSyncVerify()
    On Error GoTo EH
    
    ' 使用链式调用配置并验证
    With New cAliyunCaptcha
        .AccessKeyId("your-access-key-id") _
         .AccessKeySecret("your-access-key-secret") _
         .Region(ALIYUN_CAPTCHA_REGION_CN) _
         .Timeout (30000)
        
        ' 验证（从客户端获取的CaptchaVerifyParam）
        Dim Result As Boolean
        Result = .VerifySync("eyJjZXxxxxxxxxxxxxxxnVlfQ==", "your-scene-id")
        
        If Result = True Then
            MsgBox "验证通过！"
        Else
            MsgBox "验证失败！"
        End If
    End With
    Exit Sub
    
EH:
    MsgBox "错误: " & Err.Description
End Sub

'=========================================================================
' 示例：使用带重试的验证
'=========================================================================
Private Sub DemoVerifyWithRetry()
    On Error GoTo EH
    
    With New cAliyunCaptcha
        .AccessKeyId("your-access-key-id") _
         .AccessKeySecret("your-access-key-secret") _
         .Region (ALIYUN_CAPTCHA_REGION_CN)
        
        ' 带重试的验证（最多3次）
        Dim Result As Boolean
        Result = .VerifyWithRetry("eyJjZXxxxxxxxxxxxxxxnVlfQ==", "your-scene-id", 3)
        
        If Result = True Then
            MsgBox "验证通过（可能经过重试）！"
        Else
            MsgBox "验证失败，已达到最大重试次数！"
        End If
    End With
    Exit Sub
    
EH:
    MsgBox "错误: " & Err.Description
End Sub
