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
'          - 验证码验证
'          - 支持中国内地和新加坡地域
'
' Author: Auto
' Date: 2026-02-08
'
'=========================================================================

Option Explicit

'=========================================================================
' 依赖引用
'=========================================================================
' - cHttpClient.cls: HTTP请求客户端（含内置JSON处理）
' - mAliyunCaptcha.bas: 枚举、常量、类型定义
'
' 注意：HttpClient自带RequestDataJson和ResponseJson方法，
'       无需额外引用cJson。
'=========================================================================

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
' API 常量
'=========================================================================
Private Const API_VERSION As String = "2023-03-05"
Private Const API_FORMAT As String = "json"
Private Const API_SIGNATURE_METHOD As String = "HMAC-SHA1"
Private Const API_SIGNATURE_VERSION As String = "1.0"

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
    Result As Boolean                                                           ' 验证结果
    VerifyCode As String                                                        ' 验证码
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
' 公共方法 - 配置
'=========================================================================

' 设置AccessKey（链式调用）
Public Function AccessKeyId(Key As String) As cAliyunCaptcha
    m_AccessKeyId = Key
    Set AccessKeyId = Me
End Function

' 设置AccessKeySecret（链式调用）
Public Function AccessKeySecret(Secret As String) As cAliyunCaptcha
    m_AccessKeySecret = Secret
    Set AccessKeySecret = Me
End Function

' 设置地域（链式调用）
Public Function Region(RegionCode As AliyunCaptchaRegion) As cAliyunCaptcha
    m_Region = RegionCode
    Call UpdateEndpoint
    Set Region = Me
End Function

' 设置超时时间（链式调用）
Public Function Timeout(Milliseconds As Long) As cAliyunCaptcha
    m_Timeout = Milliseconds
    Set Timeout = Me
End Function

' 设置调试模式（链式调用）
Public Function EnableDebug(Value As Boolean) As cAliyunCaptcha
    m_EnableDebug = Value
    Set EnableDebug = Me
End Function

'=========================================================================
' 公共方法 - 验证
'=========================================================================

' 验证验证码（带阿里云RPC签名）
'@param CaptchaVerifyParam 客户端验证参数
'@param SceneId 场景ID（可选，建议传入）
'@return 验证结果对象
Public Function Verify(ByVal CaptchaVerifyParam As String, _
    Optional ByVal SceneId As String = "") As AliyunCaptchaVerifyResult
    On Error GoTo ErrorHandler
    
    Dim Result As AliyunCaptchaVerifyResult
    
    ' 参数验证
    If LenB(CaptchaVerifyParam) = 0 Then
        Call RaiseError("CaptchaVerifyParam不能为空")
        Result.Success = False
        Result.Message = "CaptchaVerifyParam不能为空"
        Verify = Result
        Exit Function
    End If
    
    ' 检查配置
    If LenB(m_Endpoint) = 0 Then
        Call RaiseError("Endpoint未配置，请先调用Region方法设置地域")
        Result.Success = False
        Result.Message = "Endpoint未配置"
        Verify = Result
        Exit Function
    End If
    
    If LenB(m_AccessKeyId) = 0 Or LenB(m_AccessKeySecret) = 0 Then
        Call RaiseError("AccessKey未配置")
        Result.Success = False
        Result.Message = "AccessKey未配置"
        Verify = Result
        Exit Function
    End If
    
    ' 构建签名参数
    Dim Params As Object
    Set Params = CreateObject("Scripting.Dictionary")
    
    ' 系统参数
    Params.Add "Action", "VerifyIntelligentCaptcha"
    Params.Add "Version", API_VERSION
    Params.Add "Format", API_FORMAT
    Params.Add "Timestamp", GetUTCTimestamp()
    Params.Add "SignatureNonce", GenerateNonce()
    Params.Add "SignatureMethod", API_SIGNATURE_METHOD
    Params.Add "SignatureVersion", API_SIGNATURE_VERSION
    Params.Add "AccessKeyId", m_AccessKeyId
    
    ' 业务参数
    Params.Add "CaptchaVerifyParam", CaptchaVerifyParam
    If LenB(SceneId) > 0 Then
        Params.Add "SceneId", SceneId
    End If
    
    ' 计算签名
    Dim Signature As String
    Signature = CalculateRPCSignature(Params, "POST", m_AccessKeySecret)
    Params.Add "Signature", Signature
    
    ' 构建请求URL（带所有参数）
    Dim RequestUrl As String
    RequestUrl = m_Endpoint & "/?" & BuildQueryString(Params)
    
    ' 调试输出
    If m_EnableDebug Then
        Debug.Print "[cAliyunCaptcha] Request URL: " & RequestUrl
        Debug.Print "[cAliyunCaptcha] Signature: " & Signature
    End If
    
    ' 使用HttpClient发送POST请求（form-urlencoded格式，Body为空）
    With New cHttpClient
        ' 设置超时时间
        .RequestTimeOut = m_Timeout
        
        ' 启用调试信息
        .DebugStart = m_EnableDebug
        
        ' 发送POST请求（RPC签名方式，参数已在URL中）
        .SendPost RequestUrl
        
        ' 解析响应
        With .ReturnJson()
            Result.Success = .Root("Success")
            Result.Result = .Root("Result")
            ' 注意：阿里云返回的是 Result.VerifyResult 和 Result.VerifyCode
            If .Root.Exists("Result") Then
                Dim ResultObj As Object
                Set ResultObj = .Root("Result")
                If TypeOf ResultObj Is cJson Then
                    Result.Result = ResultObj.Item("VerifyResult")
                    Result.VerifyCode = ResultObj.Item("VerifyCode")
                End If
            End If
            Result.CertifyID = .Root("CertifyID")
            Result.requestID = .Root("RequestId")
            
            ' 获取错误信息
            If .Root.Exists("Message") Then
                Result.Message = .Root("Message")
            End If
            
            ' 调试输出响应数据
            If m_EnableDebug Then
                Debug.Print "[cAliyunCaptcha] Response Data: " & .Encode(, 2, True)
            End If
        End With
        
        ' 调试信息输出
        If m_EnableDebug Then
            Debug.Print "[cAliyunCaptcha] Debug Info: " & .DebugInfo.Encode(, 2, True)
        End If
        
        ' 触发事件
        If Result.Result = True Then
            RaiseEvent OnVerifySuccess(Result.VerifyCode, Result.CertifyID, Result.Result)
        Else
            RaiseEvent OnVerifyFailure(Result.VerifyCode, GetVerifyMessage(Result.VerifyCode))
        End If
        
        Verify = Result
        
QH:
        If m_EnableDebug = True Then
            Debug.Print .DebugInfo.Encode(, 2, True)
        End If
    End With
    Exit Function
    
ErrorHandler:
    Call RaiseError("Verify失败: " & ERR.Description)
    Result.Success = False
    Result.Message = "Verify失败: " & ERR.Description
    Verify = Result
    Resume QH
End Function

' 同步验证（简化接口）
Public Function VerifySync(ByVal CaptchaVerifyParam As String, _
    Optional ByVal SceneId As String = "") As Boolean
    '    On Error GoTo ErrorHandler
    
    Dim Result As AliyunCaptchaVerifyResult
    Result = Me.Verify(CaptchaVerifyParam, SceneId)
    
    VerifySync = Result.Result
    Exit Function
    
ErrorHandler:
    Call RaiseError("VerifySync失败: " & ERR.Description)
    VerifySync = False
End Function

' 重试验证
Public Function VerifyWithRetry(ByVal CaptchaVerifyParam As String, _
                               Optional ByVal SceneId As String = "", _
                               Optional ByVal MaxRetries As Long = 3) As Boolean
    Dim RetryCount As Long
    Dim Result As Boolean

    For RetryCount = 1 To MaxRetries
        Result = Me.VerifySync(CaptchaVerifyParam, SceneId)

        If Result = True Then
            VerifyWithRetry = True
            Exit Function
        End If

        ' 等待一段时间后重试
'        Call Sleep(1000 * RetryCount)
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

' 获取UTC时间戳（ISO8601格式）
Private Function GetUTCTimestamp() As String
    Dim dt As Date
    dt = DateAdd("h", -8, Now) ' 本地时间转UTC（假设本地为北京时间）
    GetUTCTimestamp = Format$(dt, "yyyy-mm-ddThh:nn:ssZ")
End Function

' 生成随机nonce
Private Function GenerateNonce() As String
    ' 使用系统时间和随机数生成唯一标识
    Randomize
    GenerateNonce = Replace(CStr(Now), " ", "") & CStr(Int(Rnd * 1000000))
End Function

' URL编码（使用cToolsHttp的UTF-8编码方法）
Private Function UrlEncode(ByVal Text As String) As String
    UrlEncode = ToolsHttp.UrlEncodeUtf8(Text)
End Function

' 计算RPC签名（HMAC-SHA1）
Private Function CalculateRPCSignature(ByRef Params As Object, ByVal Method As String, ByVal Secret As String) As String
    Dim SortedKeys() As Variant
    Dim KeyCount As Long
    Dim i As Long
    Dim CanonicalQueryString As String
    Dim StringToSign As String
    Dim HMAC As cCryptoHMAC
    
    ' 获取所有键并排序
    KeyCount = Params.Count
    '    ReDim SortedKeys(KeyCount - 1)
    '    For i = 0 To KeyCount - 1
    SortedKeys = Params.Keys()
    '    Next i
    Call QuickSort(SortedKeys, 0, KeyCount - 1)
    
    ' 构建规范化的查询字符串
    For i = 0 To KeyCount - 1
        If i > 0 Then CanonicalQueryString = CanonicalQueryString & "&"
        CanonicalQueryString = CanonicalQueryString & UrlEncode(SortedKeys(i)) & "=" & UrlEncode(CStr(Params(SortedKeys(i))))
    Next i
    
    ' 构建待签名字符串
    StringToSign = Method & "&%2F&" & UrlEncode(CanonicalQueryString)
    
    ' 使用HMAC-SHA1签名
    Set HMAC = New cCryptoHMAC
    HMAC.Algorithm = HMAC_ALG_SHA1
    HMAC.SetKey Secret & "&", HMAC_ENCODING_UTF8                                ' 阿里云密钥格式: Secret&
    
    Dim DataBytes() As Byte
    DataBytes = StrConv(StringToSign, vbFromUnicode)
    
    Dim SignatureBytes() As Byte
    SignatureBytes = HMAC.ComputeBytes(DataBytes, StrConv(Secret & "&", vbFromUnicode))
    
    ' 返回Base64编码的签名
    CalculateRPCSignature = BytesToBase64(SignatureBytes)
    
    Set HMAC = Nothing
End Function

' 构建查询字符串
Private Function BuildQueryString(ByRef Params As Object) As String
    Dim Result As String
    Dim i As Long
    Dim Key As Variant
    
    For Each Key In Params.Keys
        If Len(Result) > 0 Then Result = Result & "&"
        Result = Result & UrlEncode(CStr(Key)) & "=" & UrlEncode(CStr(Params(Key)))
    Next Key
    
    BuildQueryString = Result
End Function

' 快速排序（用于参数排序）
Private Sub QuickSort(ByRef Arr() As Variant, ByVal Low As Long, ByVal High As Long)
    Dim i As Long, j As Long
    Dim Pivot As String
    Dim Temp As String
    
    If Low < High Then
        i = Low
        j = High
        Pivot = Arr((Low + High) \ 2)
        
        Do While i <= j
            Do While Arr(i) < Pivot And i < High
                i = i + 1
            Loop
            Do While Arr(j) > Pivot And j > Low
                j = j - 1
            Loop
            If i <= j Then
                Temp = Arr(i)
                Arr(i) = Arr(j)
                Arr(j) = Temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If Low < j Then QuickSort Arr, Low, j
        If i < High Then QuickSort Arr, i, High
    End If
End Sub

' 字节数组转Base64
Private Function BytesToBase64(ByRef Data() As Byte) As String
    Dim oXML As Object
    Dim oNode As Object
    
    On Error Resume Next
    Set oXML = CreateObject("MSXML2.DOMDocument")
    Set oNode = oXML.createElement("b64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Data
    BytesToBase64 = oNode.Text
    On Error GoTo 0
    
    Set oNode = Nothing
    Set oXML = Nothing
End Function

' 获取验证码错误信息
Private Function GetVerifyMessage(ByVal VerifyCode As String) As String
    Select Case VerifyCode
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
        Case Else: GetVerifyMessage = "未知错误码" & VerifyCode
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
    m_Timeout = 30000  ' 30秒
    m_EnableDebug = False
    Call UpdateEndpoint
End Sub

Private Sub Class_Terminate()
    ' 清理资源（如果需要）
    ' HttpClient在With块中使用，自动释放
End Sub
