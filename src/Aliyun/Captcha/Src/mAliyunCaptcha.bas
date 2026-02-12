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
Public Const ALIYUN_CAPTCHA_API_VERSION As String = "2023-03-05"  ' API版本号

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
