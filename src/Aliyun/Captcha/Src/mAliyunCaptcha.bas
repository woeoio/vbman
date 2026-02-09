Attribute VB_Name = "mAliyunCaptcha"
'=========================================================================
'
' mAliyunCaptcha - 阿里云验证码2.0公共模块
'
' Purpose: 提供阿里云验证码2.0的枚举、常量、类型定义
'
' Author: Auto
' Date: 2026-02-08
'
'=========================================================================

Option Explicit


'=========================================================================
' 常量定义
'=========================================================================

' API常量
Public Const ALIYUN_CAPTCHA_ENDPOINT_CN As String = "captcha.cn-shanghai.aliyuncs.com"
Public Const ALIYUN_CAPTCHA_ENDPOINT_SGP As String = "captcha.ap-southeast-1.aliyuncs.com"
Public Const ALIYUN_CAPTCHA_ACTION_VERIFY As String = "VerifyIntelligentCaptcha"

' 默认配置
Public Const ALIYUN_CAPTCHA_DEFAULT_TIMEOUT As Long = 30000  ' 30秒
Public Const ALIYUN_CAPTCHA_DEFAULT_REGION As Long = 1       ' 中国内地

' 验证码常量
Public Const ALIYUN_CAPTCHA_VERIFY_SUCCESS_CODE As String = "T001"
Public Const ALIYUN_CAPTCHA_VERIFY_TEST_SUCCESS_CODE As String = "T005"

' 错误码常量
Public Const ALIYUN_CAPTCHA_ERROR_ATTACK As String = "F001"
Public Const ALIYUN_CAPTCHA_ERROR_PARAM_EMPTY As String = "F002"
Public Const ALIYUN_CAPTCHA_ERROR_PARAM_INVALID As String = "F003"
Public Const ALIYUN_CAPTCHA_ERROR_TEST_FAIL As String = "F004"
Public Const ALIYUN_CAPTCHA_ERROR_SCENE_INVALID_PARAM As String = "F005"
Public Const ALIYUN_CAPTCHA_ERROR_SCENE_INVALID_FRONTEND As String = "F006"
Public Const ALIYUN_CAPTCHA_ERROR_DUPLICATE As String = "F008"
Public Const ALIYUN_CAPTCHA_ERROR_VIRTUAL_DEVICE As String = "F009"
Public Const ALIYUN_CAPTCHA_ERROR_IP_LIMIT As String = "F010"
Public Const ALIYUN_CAPTCHA_ERROR_DEVICE_LIMIT As String = "F011"
Public Const ALIYUN_CAPTCHA_ERROR_SCENE_MISMATCH As String = "F012"
Public Const ALIYUN_CAPTCHA_ERROR_PARAM_MISSING As String = "F013"
Public Const ALIYUN_CAPTCHA_ERROR_NO_INIT As String = "F014"
Public Const ALIYUN_CAPTCHA_ERROR_VERIFY_FAIL As String = "F015"
Public Const ALIYUN_CAPTCHA_ERROR_URL_VERIFY_FAIL As String = "F016"
Public Const ALIYUN_CAPTCHA_ERROR_ATTACK_PARAM As String = "F017"
Public Const ALIYUN_CAPTCHA_ERROR_V3_REUSE As String = "F018"
Public Const ALIYUN_CAPTCHA_ERROR_V3_TIMEOUT As String = "F019"
Public Const ALIYUN_CAPTCHA_ERROR_V3_MISMATCH As String = "F020"
Public Const ALIYUN_CAPTCHA_ERROR_V2_INIT_FAIL As String = "F023"
Public Const ALIYUN_CAPTCHA_ERROR_AUTO_SCRIPT As String = "F024"

'=========================================================================
' 公共函数
'=========================================================================

' 获取地域的Endpoint
Public Function GetRegionEndpoint(Region As AliyunCaptchaRegion, _
                                  Optional DualStack As Boolean = False) As String
    Select Case Region
        Case ALIYUN_CAPTCHA_REGION_CN
            If DualStack Then
                GetRegionEndpoint = "https://captcha-dualstack.cn-shanghai.aliyuncs.com"
            Else
                GetRegionEndpoint = "https://captcha.cn-shanghai.aliyuncs.com"
            End If
        Case ALIYUN_CAPTCHA_REGION_SGP
            If DualStack Then
                GetRegionEndpoint = "https://captcha-dualstack.ap-southeast-1.aliyuncs.com"
            Else
                GetRegionEndpoint = "https://captcha.ap-southeast-1.aliyuncs.com"
            End If
    End Select
End Function

' 获取错误码描述
Public Function GetVerifyCodeDescription(ByVal VerifyCode As String) As String
    Select Case VerifyCode
        Case "T001": GetVerifyCodeDescription = "服务端校验通过"
        Case "T005": GetVerifyCodeDescription = "测试模式，配置了验证通过"
        Case "F001": GetVerifyCodeDescription = "疑似攻击请求，风险策略不通过"
        Case "F002": GetVerifyCodeDescription = "CaptchaVerifyParam参数为空"
        Case "F003": GetVerifyCodeDescription = "CaptchaVerifyParam格式不合法"
        Case "F004": GetVerifyCodeDescription = "测试模式，配置了验证不通过"
        Case "F005": GetVerifyCodeDescription = "场景ID不合法（参数异常）"
        Case "F006": GetVerifyCodeDescription = "场景ID不合法（前端配置错误）"
        Case "F008": GetVerifyCodeDescription = "验证数据重复提交"
        Case "F009": GetVerifyCodeDescription = "检测到虚拟设备环境"
        Case "F010": GetVerifyCodeDescription = "同IP访问频率超出限制"
        Case "F011": GetVerifyCodeDescription = "同设备访问频率超出限制"
        Case "F012": GetVerifyCodeDescription = "服务端SceneID与前端不一致"
        Case "F013": GetVerifyCodeDescription = "CaptchaVerifyParam缺少参数"
        Case "F014": GetVerifyCodeDescription = "无初始化记录"
        Case "F015": GetVerifyCodeDescription = "验证交互不通过"
        Case "F016": GetVerifyCodeDescription = "自定义策略URL验证不通过"
        Case "F017": GetVerifyCodeDescription = "疑似攻击请求，协议或参数异常"
        Case "F018": GetVerifyCodeDescription = "V3架构：CaptchaVerifyParam重复使用"
        Case "F019": GetVerifyCodeDescription = "V3架构：请求间隔超出90秒"
        Case "F020": GetVerifyCodeDescription = "V3架构：参数不匹配"
        Case "F023": GetVerifyCodeDescription = "V2架构：初始化失败自动触发"
        Case "F024": GetVerifyCodeDescription = "检测到自动化脚本操作"
        Case Else: GetVerifyCodeDescription = "未知错误：" & VerifyCode
    End Select
End Function

' 获取验证结果是否通过
Public Function IsVerifySuccess(ByVal VerifyCode As String) As Boolean
    Select Case VerifyCode
        Case "T001", "T005"
            IsVerifySuccess = True
        Case Else
            IsVerifySuccess = False
    End Select
End Function
