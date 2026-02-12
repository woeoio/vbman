VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAliyunCaptcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'=========================================================================
'
' cAliyunCaptcha - 阿里云验证码2.0客户端类
'
' Purpose: 提供阿里云验证码2.0服务端验证功能
'          - 调用VerifyIntelligentCaptcha接口
'          - 实现阿里云RPC V1签名算法（HMAC-SHA1）
'          - 支持中国内地和新加坡地域
'
' Author: Auto
' Date: 2026-02-08
'
' 重构说明：
' 本版本基于文档 03.md 的指导进行重构：
' 1. 使用 cCollection 替代 Scripting.Dictionary，支持 SortByKey() 排序
' 2. 使用 ToolsDateTime.GetIso8601Timestamp() 生成 UTC 时间戳
' 3. 使用 cCryptoHMAC 链式调用简化签名计算
' 4. 严格遵循阿里云 RPC V1 签名流程
' 5. 正确处理 Query 和 Body 参数合并签名
'
'=========================================================================

Option Explicit

'=========================================================================
' 依赖引用
'=========================================================================
' - cHttpClient.cls: HTTP请求客户端（含内置JSON处理）
' - cCollection.cls: 带排序功能的集合类（Key字典序排序）
' - cCryptoHMAC.cls: HMAC-SHA1签名计算（支持链式调用）
' - cToolsHttp.cls: URL编码和Form编码工具
' - ToolsDateTime.bas: ISO8601时间戳和UUID生成
'
' 注意：HttpClient自带RequestDataJson和ResponseJson方法，
'       无需额外引用cJson。
'=========================================================================

'=========================================================================
' API 常量
'=========================================================================
Private Const API_VERSION As String = "2023-03-05"
Private Const API_FORMAT As String = "json"
Private Const API_SIGNATURE_METHOD As String = "HMAC-SHA1"
Private Const API_SIGNATURE_VERSION As String = "1.0"
Private Const API_ACTION As String = "VerifyIntelligentCaptcha"

'=========================================================================
' 私有变量
'=========================================================================
Private m_AccessKeyId As String                                                 ' AccessKey ID
Private m_AccessKeySecret As String                                             ' AccessKey Secret
Private m_Region As AliyunCaptchaRegion                                         ' 地域
Private m_Endpoint As String                                                    ' API Endpoint
Private m_Timeout As Long                                                       ' 超时时间（毫秒）
Private m_EnableDebug As Boolean                                                ' 是否启用调试

'=========================================================================
' 枚举定义
'=========================================================================

' 阿里云验证码地域枚举
Public Enum AliyunCaptchaRegion
    ALIYUN_CAPTCHA_REGION_CN = 1                                                ' 中国内地
    ALIYUN_CAPTCHA_REGION_SGP = 2                                               ' 新加坡
End Enum

' 验证码模式枚举（客户端使用）
Public Enum AliyunCaptchaMode
    ALIYUN_CAPTCHA_MODE_POPUP = 1                                               ' 弹出式
    ALIYUN_CAPTCHA_MODE_EMBED = 2                                               ' 嵌入式
End Enum

'=========================================================================
' 类型定义
'=========================================================================

' 验证结果结构
Public Type AliyunCaptchaVerifyResult
    VerifyResult As Boolean                                                     ' 验证结果（true=通过，false=未通过）
    VerifyCode As String                                                        ' 验证结果代码
    CertifyID As String                                                         ' 认证ID
    Success As Boolean                                                          ' 请求是否成功
    requestID As String                                                         ' 请求ID
    Message As String                                                           ' 错误信息
End Type

' 验证码配置结构
Public Type AliyunCaptchaConfig
    AccessKeyId As String
    AccessKeySecret As String
    Region As AliyunCaptchaRegion
    Timeout As Long
End Type

'=========================================================================
' 事件定义
'=========================================================================
' 验证成功事件
Public Event OnVerifySuccess(ByVal VerifyCode As String, _
    ByVal CertifyID As String, _
    ByVal Result As Boolean)

    ' 验证失败事件
Public Event OnVerifyFailure(ByVal ErrorCode As String, _
    ByVal ErrorMessage As String)

    ' 错误事件
Public Event OnError(ByVal Description As String)

'=========================================================================
' 公共方法 - 配置（链式调用）
'=========================================================================

' 设置AccessKeyId
Public Function AccessKeyId(Key As String) As cAliyunCaptcha
    m_AccessKeyId = Key
    Set AccessKeyId = Me
End Function

' 设置AccessKeySecret
Public Function AccessKeySecret(Secret As String) As cAliyunCaptcha
    m_AccessKeySecret = Secret
    Set AccessKeySecret = Me
End Function

' 设置地域
Public Function Region(RegionCode As AliyunCaptchaRegion) As cAliyunCaptcha
    m_Region = RegionCode
    Call UpdateEndpoint
    Set Region = Me
End Function

' 设置地域（直接使用地域字符串）
' @param Area 地域字符串，如 "cn-shanghai" 或 "ap-southeast-1"
' @return 返回当前对象，支持链式调用
Public Function RegionArea(ByVal Area As String) As cAliyunCaptcha
    ' 根据地域字符串自动映射到枚举值
    Select Case LCase(Trim(Area))
    Case "cn-shanghai"
        m_Region = ALIYUN_CAPTCHA_REGION_CN
    Case "ap-southeast-1"
        m_Region = ALIYUN_CAPTCHA_REGION_SGP
    Case Else
        ' 未知地域，保持不变
        Debug.Print "[cAliyunCaptcha] Warning: Unknown RegionArea '" & Area & "', region unchanged"
    End Select
    Call UpdateEndpoint
    Set RegionArea = Me
End Function

' 设置超时时间
Public Function Timeout(Milliseconds As Long) As cAliyunCaptcha
    m_Timeout = Milliseconds
    Set Timeout = Me
End Function

' 设置调试模式
Public Function EnableDebug(Value As Boolean) As cAliyunCaptcha
    m_EnableDebug = Value
    Set EnableDebug = Me
End Function

'=========================================================================
' 公共方法 - 验证
'=========================================================================

' 验证验证码（带阿里云RPC签名）
'@param CaptchaVerifyParam 客户端验证参数
'@param SceneId 场景ID（必填）
'@return 验证结果对象
Public Function Verify(ByVal CaptchaVerifyParam As String, _
    ByVal SceneId As String) As AliyunCaptchaVerifyResult
    
    On Error GoTo ErrorHandler
    
    Dim Result As AliyunCaptchaVerifyResult
    
    ' 参数验证
    If LenB(CaptchaVerifyParam) = 0 Then
        Result.Success = False
        Result.Message = "CaptchaVerifyParam不能为空"
        Call RaiseError(Result.Message)
        Verify = Result
        Exit Function
    End If
    
    ' 检查配置
    If LenB(m_Endpoint) = 0 Then
        Result.Success = False
        Result.Message = "Endpoint未配置，请先调用Region方法设置地域"
        Call RaiseError(Result.Message)
        Verify = Result
        Exit Function
    End If
    
    If LenB(m_AccessKeyId) = 0 Or LenB(m_AccessKeySecret) = 0 Then
        Result.Success = False
        Result.Message = "AccessKey未配置"
        Call RaiseError(Result.Message)
        Verify = Result
        Exit Function
    End If
    
    '=========================================================================
    ' 步骤1：构建系统参数（Query参数）
    '=========================================================================
    Dim QueryParams As New cCollection
    
    ' 系统级参数
    QueryParams.Add API_ACTION, "Action"
    QueryParams.Add API_VERSION, "Version"
    QueryParams.Add API_FORMAT, "Format"
    QueryParams.Add ToolsDateTime.GetIso8601Timestamp(), "Timestamp"
    QueryParams.Add GenerateNonce(), "SignatureNonce"
    QueryParams.Add API_SIGNATURE_METHOD, "SignatureMethod"
    QueryParams.Add API_SIGNATURE_VERSION, "SignatureVersion"
    QueryParams.Add m_AccessKeyId, "AccessKeyId"
    
    '=========================================================================
    ' 步骤2：构建业务参数（Body参数）
    '=========================================================================
    Dim BodyParams As New cCollection
    BodyParams.Add CaptchaVerifyParam, "CaptchaVerifyParam"
    BodyParams.Add SceneId, "SceneId"
    
    '=========================================================================
    ' 步骤3：计算签名（合并Query和Body参数，统一排序）
    '=========================================================================
    ' 创建用于签名的参数集合（合并Query和Body参数）
    Dim SignParams As New cCollection
    
    ' 复制Query参数
    Dim Keys() As String
    Dim Key As Variant
    Keys = QueryParams.Keys()
    For Each Key In Keys
        SignParams.Add QueryParams.Item(Key), Key
    Next Key
    
    ' 复制Body参数
    Keys = BodyParams.Keys()
    For Each Key In Keys
        SignParams.Add BodyParams.Item(Key), Key
    Next Key
    
    ' 按Key字典序排序（ASCII升序）
    SignParams.SortByKey
    
    ' 生成规范化查询字符串（参数名和值都已URL编码）
    Dim ToolsHttp As New cToolsHttp
    Dim CanonicalQueryString As String
    CanonicalQueryString = ToolsHttp.ToWwwFormUrlencoded(SignParams, True)
    
    ' 构造待签名字符串
    ' StringToSign = HTTPMethod + "&" + percentEncode("/") + "&" + percentEncode(CanonicalQueryString)
    Dim StringToSign As String
    StringToSign = "POST" & "&" & ToolsHttp.UrlEncodeUtf8("/") & "&" & ToolsHttp.UrlEncodeUtf8(CanonicalQueryString)
    
    ' 使用HMAC-SHA1计算签名（链式调用）
    Dim HMAC As New cCryptoHMAC
    Dim Signature As String
    
    Signature = HMAC.Mode(HMAC_ALG_SHA1) _
    .Secret(m_AccessKeySecret & "&", SECRET_KEY_STRING, HMAC_ENCODING_UTF8) _
    .DataString(StringToSign, HMAC_ENCODING_UTF8) _
    .ReturnBase64()
    
    Set HMAC = Nothing
    
    ' 最后添加Signature到Query参数（不参与签名计算）
    QueryParams.Add Signature, "Signature"
    
    '=========================================================================
    ' 步骤4：构建请求URL和Body
    '=========================================================================
    Dim RequestUrl As String
    Dim RequestBody As String
    
    ' URL包含所有Query参数（包括Signature）
    RequestUrl = m_Endpoint & "/?" & ToolsHttp.ToWwwFormUrlencoded(QueryParams, True)
    
    ' Body只包含业务参数
    RequestBody = ToolsHttp.ToWwwFormUrlencoded(BodyParams, True)
    
    ' 调试输出
    If m_EnableDebug Then
        Debug.Print "[cAliyunCaptcha] === 请求详情 ==="
        Debug.Print "[cAliyunCaptcha] Request URL: " & RequestUrl
        Debug.Print "[cAliyunCaptcha] Request Body: " & RequestBody
        Debug.Print "[cAliyunCaptcha] CanonicalQueryString: " & CanonicalQueryString
        Debug.Print "[cAliyunCaptcha] StringToSign: " & StringToSign
        Debug.Print "[cAliyunCaptcha] Signature: " & Signature
    End If
    
    '=========================================================================
    ' 步骤5：发送HTTPS POST请求
    '=========================================================================
    Dim Http As New cHttpClient
    
    With Http
        ' 设置超时
        .RequestTimeOut = m_Timeout
        
        ' 启用调试
        .DebugStart = m_EnableDebug
        
        ' 发送POST请求（参数已在URL中，Body为form-urlencoded格式）
        .SendPost RequestUrl, RequestBody
        
        '=========================================================================
        ' 步骤6：解析响应
        '=========================================================================
        With .ReturnJson()
            ' 解析基本响应字段
            Result.Success = .Root("Success")
            Result.requestID = .Root("RequestId")
            
            ' 解析Result对象（包含VerifyResult和VerifyCode）
            Result.VerifyResult = .Root("Result")("VerifyResult")
            Result.VerifyCode = .Root("Result")("VerifyCode")
            
            ' 解析其他字段
            Result.CertifyID = .Root("CertifyID")
            Result.Message = .Root("Message")
            
            ' 调试输出响应
            If m_EnableDebug Then
                Debug.Print "[cAliyunCaptcha] Response: " & .Encode(, 2, True)
            End If
        End With
        
        ' 调试信息输出
        If m_EnableDebug Then
            Debug.Print "[cAliyunCaptcha] Debug Info: " & .DebugInfo.Encode(, 2, True)
        End If
        
        ' 触发事件
        If Result.VerifyResult = True Then
            RaiseEvent OnVerifySuccess(Result.VerifyCode, Result.CertifyID, Result.VerifyResult)
        Else
            RaiseEvent OnVerifyFailure(Result.VerifyCode, GetVerifyMessage(Result.VerifyCode))
        End If
        
    End With
    
    Set Http = Nothing
    Set ToolsHttp = Nothing
    
    Verify = Result
    Exit Function
    
ErrorHandler:
    Call RaiseError("Verify失败: " & ERR.Description)
    Result.Success = False
    Result.Message = "Verify失败: " & ERR.Description
    Verify = Result
End Function

' 同步验证（简化接口，只返回布尔结果）
Public Function VerifySync(ByVal CaptchaVerifyParam As String, _
    ByVal SceneId As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim Result As AliyunCaptchaVerifyResult
    Result = Me.Verify(CaptchaVerifyParam, SceneId)
    
    VerifySync = Result.VerifyResult
    Exit Function
    
ErrorHandler:
    Call RaiseError("VerifySync失败: " & ERR.Description)
    VerifySync = False
End Function

' 重试验证
Public Function VerifyWithRetry(ByVal CaptchaVerifyParam As String, _
                               ByVal SceneId As String, _
                               Optional ByVal MaxRetries As Long = 3) As Boolean
    Dim RetryCount As Long
    Dim Result As Boolean

    For RetryCount = 1 To MaxRetries
        Result = Me.VerifySync(CaptchaVerifyParam, SceneId)

        If Result = True Then
            VerifyWithRetry = True
            Exit Function
        End If
    Next RetryCount

    VerifyWithRetry = False
End Function

'=========================================================================
' 公共方法 - 获取配置
'=========================================================================

' 获取当前Endpoint
Public Property Get Endpoint() As String
    Endpoint = m_Endpoint
End Property

' 获取当前地域
Public Property Get CurrentRegion() As AliyunCaptchaRegion
    CurrentRegion = m_Region
End Property

' 获取当前超时时间
Public Property Get CurrentTimeout() As Long
    CurrentTimeout = m_Timeout
End Property

'=========================================================================
' 私有方法
'=========================================================================

' 更新Endpoint
Private Sub UpdateEndpoint()
    Select Case m_Region
        Case ALIYUN_CAPTCHA_REGION_CN
            m_Endpoint = "https://captcha.cn-shanghai.aliyuncs.com"
        Case ALIYUN_CAPTCHA_REGION_SGP
            m_Endpoint = "https://captcha.ap-southeast-1.aliyuncs.com"
        Case Else
            m_Endpoint = ""
    End Select
End Sub

' 生成随机nonce（UUID格式）
Private Function GenerateNonce() As String
    ' 使用 cToolsStr.GetGUID 生成标准UUID
    Dim ToolsStr As New cToolsStr
    GenerateNonce = ToolsStr.GetGUID()
    Set ToolsStr = Nothing
End Function

' 获取验证码错误信息
Private Function GetVerifyMessage(ByVal VerifyCode As String) As String
    Select Case VerifyCode
        Case "100": GetVerifyMessage = "验证通过"
        Case "T001": GetVerifyMessage = "服务端校验通过"
        Case "T005": GetVerifyMessage = "测试模式，配置了验证通过"
        Case "F001": GetVerifyMessage = "疑似攻击请求，风险策略不通过"
        Case "F002": GetVerifyMessage = "CaptchaVerifyParam参数为空"
        Case "F003": GetVerifyMessage = "CaptchaVerifyParam格式不合法"
        Case "F004": GetVerifyMessage = "测试模式，配置了验证不通过"
        Case "F005": GetVerifyMessage = "场景ID不合法（参数异常）"
        Case "F006": GetVerifyMessage = "场景ID不合法（前端配置错误）"
        Case "F008": GetVerifyMessage = "验证数据重复提交"
        Case "F009": GetVerifyMessage = "检测到虚拟设备环境"
        Case "F010": GetVerifyMessage = "同IP访问频率超出限制"
        Case "F011": GetVerifyMessage = "同设备访问频率超出限制"
        Case "F012": GetVerifyMessage = "服务端SceneID与前端不一致"
        Case "F013": GetVerifyMessage = "CaptchaVerifyParam缺少参数"
        Case "F014": GetVerifyMessage = "无初始化记录"
        Case "F015": GetVerifyMessage = "验证交互不通过"
        Case "F016": GetVerifyMessage = "自定义策略URL验证不通过"
        Case "F017": GetVerifyMessage = "疑似攻击请求，协议或参数异常"
        Case "F018": GetVerifyMessage = "V3架构：CaptchaVerifyParam重复使用"
        Case "F019": GetVerifyMessage = "V3架构：请求间隔超过90秒"
        Case "F020": GetVerifyMessage = "V3架构：参数不匹配"
        Case "F023": GetVerifyMessage = "V2架构：初始化失败自动触发"
        Case "F024": GetVerifyMessage = "检测到自动化脚本操作"
        Case Else: GetVerifyMessage = "未知错误码: " & VerifyCode
    End Select
End Function

' 触发错误事件
Private Sub RaiseError(ByVal Description As String)
    Debug.Print "[cAliyunCaptcha] Error: " & Description
    RaiseEvent OnError(Description)
End Sub

'=========================================================================
' 类事件
'=========================================================================

Private Sub Class_Initialize()
    ' 初始化默认值
    m_Region = ALIYUN_CAPTCHA_REGION_CN
    m_Timeout = 30000                                                           ' 默认30秒超时
    m_EnableDebug = False
    Call UpdateEndpoint
    
    ' 初始化随机数种子
    Randomize
End Sub

Private Sub Class_Terminate()
    ' 清理资源
End Sub



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
