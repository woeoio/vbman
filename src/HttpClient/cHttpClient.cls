VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'作者：邓伟，QQ：215879458，可自由传播，但请勿删本签名

'引用“Microsoft WinHTTP Services, version 5.1”组件
'引用了Microsoft Scripting Runtime组件
'引用 cJson.cls 类文件

'增加了通过Content-Type自动识别 json 数据和强制返回json数据的参数

'增加了"长轮询"功能
'疑问:winhttp组件支持证书,,或许可以参考实现winsock的证书服务?
'疑问:winhttp返回值类型还有个"流",,这个是可以做直播等流媒体?---- 了解了一下,可以做文件上传和下载的,不过直播应该也可以

'2023-03-15 增加了对json的content-type 的处理 : AjaxInst.postContentType = Json

'2023-03-30 改大了同步请求的默认超时值为5分钟，后续增加超时参数给用户设置

'todo 增加集合对象, 实现并发异步请求池子
'todo 增加集合对象, 实现串行异步请求队列

Dim WithEvents Inst As WinHttp.WinHttpRequest
Attribute Inst.VB_VarHelpID = -1
Dim Json As New cJson
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


    '缓存请求的返回内容
Public ResponseRaw As Variant
'模仿前几天写的php版ajax，做个完整处理过的 res.header+res.cookies+res.data
Public Event OnResponseFinished()

Public LastError As String
Dim pvRequestData() As Byte
Public RequestDataJson As New cJson
Public RequestDataBody As New Scripting.Dictionary                              '废弃
Public RequestDataForm As New Scripting.Dictionary
Public RequestDataQuery As New Scripting.Dictionary
Public RequestHeaders As New Scripting.Dictionary
Public RequestContentType As String
Public RequestChartSet As String
Public RequestTimeOut As Long

Public ResponseHeaders As New Scripting.Dictionary
Public Cookies As New Scripting.Dictionary

Public DebugStart As Boolean
Public DebugInfo As New cJson
'链式调用函数参数变量区
Dim IsAsync As Boolean

Public Function MapRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
    MapRequestContentType = SetRequestContentType(ReqType, ContentType)
End Function
Public Function SetRequestContentType(ReqType As EnumRequestContentType, Optional ContentType As String) As String
    Dim Arr As Variant: Arr = Array("", _
    Consts.TOOLS_HTTP_CONTENT_TYPE_JSON, _
    Consts.TOOLS_HTTP_CONTENT_TYPE_FORM_URLENCODED, _
    Consts.TOOLS_HTTP_CONTENT_TYPE_FORM_MULTIPART, _
    Consts.TOOLS_HTTP_CONTENT_TYPE_TEXT_PLAIN, _
    Consts.TOOLS_HTTP_CONTENT_TYPE_TEXT_HTML _
    )
    RequestContentType = Arr(ReqType)
    SetRequestContentType = RequestContentType
End Function

Public Sub ShowPage(url As String)
    ShellExecute 0, "open", url, 0, 0, 1
End Sub

Public Function SetCookies(ByVal Value As String) As cHttpClient
    Set SetCookies = Me
    If Trim(Value) = "" Then Exit Function
    'todo 此处应该实现字符串解析到 Cookies 对象,
    '并且返回头如果有 set-cookies 也要自动处理, 而且可能有多个 set-Cookies
    RequestHeaders("Cookie") = Value
End Function

Public Function Async(Bool As Boolean) As cHttpClient
    Set Async = Me
    IsAsync = Bool
End Function

Public Function SendPost(ByVal url As String, Optional Body As String) As cHttpClient
    Set SendPost = Fetch(ReqPost, url, Body)
End Function
Public Function SendGet(ByVal url As String, Optional Body As String) As cHttpClient
    Set SendGet = Fetch(ReqGet, url, Body)
End Function
Public Function SendPut(ByVal url As String, Optional Body As String) As cHttpClient
    Set SendPut = Fetch(ReqPut, url, Body)
End Function
Public Function SendDelete(ByVal url As String, Optional Body As String) As cHttpClient
    Set SendDelete = Fetch(ReqDelete, url, Body)
End Function
Public Function SendOptions(ByVal url As String, Optional Body As String) As cHttpClient
    Set SendOptions = Fetch(ReqOptions, url, Body)
End Function

'alias
Public Function Send(Method As EnumRequestMethod, ByVal url As String, Optional Body As String) As cHttpClient
    Set Send = Fetch(Method, url, Body)
End Function
Public Function Fetch(Method As EnumRequestMethod, ByVal url As String, Optional Body As String) As cHttpClient
    Set Fetch = Me
    On Error GoTo ERR:
    Dim x As Variant: LastError = "": Set DebugInfo = Nothing
    '检查url
    If url = "" Then ERR.Raise 500, , "请求的URL地址为空,请检查!"
    '处理请求参数
    If RequestDataQuery.Count > 0 Then url = ToolsHttp.AddToQueryString(url, ToolsHttp.MakeContent(RequestDataQuery))
    '是否添加URL资源版本号
    '    If addVer Then Url = ToolsHttp.AddToQueryString(Url, Rnd())
    '打开链路
    Dim MethodName As String: MethodName = ToolsHttp.MapMethodName(Method)
    Inst.Open MethodName, url, True
    '处理Header
    With RequestHeaders
        If .Exists("Cookie") = False Then
            'todo
        End If
        If .Exists("Content-Type") = False Then
            If RequestContentType = "" Then RequestContentType = Consts.TOOLS_HTTP_CONTENT_TYPE_FORM_URLENCODED
            .Add Consts.TOOLS_HTTP_HEADER_FIELD_CONTENT_TYPE, RequestContentType
        End If
        If .Count > 0 Then
            For Each x In .Keys()
                If .Item(x) <> "" Then Inst.SetRequestHeader x, .Item(x)
            Next
        End If
    End With
    '忽略错误SSL证书
    Inst.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    '处理POSTBODY
    If Body = "" Then
        If RequestContentType = Consts.TOOLS_HTTP_CONTENT_TYPE_JSON Then
            Body = RequestDataJson.Encode(, , True)
        ElseIf RequestContentType = Consts.TOOLS_HTTP_CONTENT_TYPE_FORM_URLENCODED Then
            Body = ToolsHttp.MakeContent(RequestDataForm)
        End If
    End If
    If Body = "" Then
        '发送数据
        Inst.Send
    Else
        Dim BodyByte() As Byte
        If ToolsStr.HasStr("utf", RequestChartSet) > 0 Then
            BodyByte = ToolsUtf8.Encode(Body)
        Else
            BodyByte = ToolsArray.StringToByteArray(Body)
        End If
        Inst.SetRequestHeader Consts.TOOLS_HTTP_HEADER_FIELD_CONTENT_LENGHT, UBound(BodyByte) + 1
        '发送数据
        Inst.Send BodyByte
    End If
    '返回
    
    If IsAsync = False Then
        '设置等待超时, '由于使用异步请求,如果用户参数改为同步请求,此处则等待返回,无需使用do循环+doevents
        If RequestTimeOut < 30 Then RequestTimeOut = 30
        If Inst.WaitForResponse(RequestTimeOut) = False Then ERR.Raise 900, , "请求超时: " & RequestTimeOut
        '处理网络不可用
        '        If Inst.Status = WinHttpRequestStatus_Sending Then ERR.Raise 901, , "网络不可用"
        '        Do While httpRequest.Status = WinHttpRequestStatus_Sending
        '            DoEvents                                                            '允许其他事件处理，防止卡死
        '        Loop
        '
        If Inst.Status <> 200 Then ERR.Raise Inst.Status, , Inst.Status & "#" & Inst.StatusText & "#" & Left(Inst.ResponseText, 1024)
        '解析 返回头
        Call ToolsHttp.ParseKeyValue(Inst.GetAllResponseHeaders(), ResponseHeaders)
    End If
    '清理body,但保留头部和get请求参数,头部和参数基本上是不变的,需要清理可以手工执行字典清楚方法
    GoSub CLEAN
    Exit Function
ERR:
    Rem todo 这里可以优化一下，对超时错误代码进行准确描述
    If DebugStart = True Then
        With DebugInfo
            With .NewItem("Error")
                .Item("Number") = ERR.Number
                .Item("Source") = ERR.Source
                .Item("Description") = ERR.Description
            End With
        End With
    End If
    LastError = ERR.Description
    GoSub CLEAN
    ERR.Raise ERR.Number, ERR.Source, ERR.Description, ERR.HelpFile, ERR.HelpContext
    Exit Function
CLEAN:
    If DebugStart = True Then
        With DebugInfo
            With .NewItem("Request")
                .Item("IsAsync") = IsAsync
                .Item("Method") = MethodName
                .Item("Url") = url
                .Item("Body") = Body
                Set .Item("Headers") = RequestHeaders
                .Item("TimeOut") = RequestTimeOut
                .Item("ChartSet") = RequestChartSet
            End With
            If IsAsync = False Then
                With .NewItem("Response")
                    .Item("Status") = Inst.Status
                    .Item("StatusText") = Inst.StatusText
                    Set .Item("Headers") = ResponseHeaders
                    .Item("Content") = ReturnText(True)
                End With
            End If
        End With
    End If
    IsAsync = False
    DebugStart = False
    
    RequestDataJson.Clear
    RequestDataBody.RemoveAll
    RequestDataForm.RemoveAll
    RequestDataQuery.RemoveAll
    '    RequestHeaders.RemoveAll
    '    Cookies.RemoveAll
    Return
End Function

Public Function ReturnStream() As Variant
    ReturnStream = Inst.ResponseStream
End Function

Public Function ReturnBody() As Byte()
    ReturnBody = Inst.ResponseBody
End Function

Public Function ReturnJson(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As cJson
    Set ReturnJson = New cJson
    Dim S As String: S = ReturnText(IsUtf8, IsConvert)
    Call ReturnJson.Decode(S)
End Function

Public Function ReturnText(Optional IsUtf8 As Boolean = True, Optional IsConvert As Boolean) As String
    If IsUtf8 = True Then
        ReturnText = ToolsUtf8.Decode(Inst.ResponseBody)
        Exit Function
    End If
    If IsConvert = True Then
        '---------------------------二进制转字符串[直接返回字串出现乱码时尝试]
        ReturnText = StrConv(Inst.ResponseBody, vbUnicode)
    Else
        ReturnText = Inst.ResponseText
    End If
End Function

Private Sub Class_Initialize()
    Set Inst = New WinHttp.WinHttpRequest
    '    oHttp.SetProxy 1
    '    oHttp.SetTimeouts 2000, 2000, 6000, 20000
    'WinHttp.SetTimeouts 60000, 60000, 60000, 3000 '设置操作超时时间
    'WinHttp.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
    'resolveTimeout = 10000 '解析 DNS 名字的超时时间，10000 毫秒。
    'connectTimeout = 10000 '建立 Winsock 连接的超时时间，10000 毫秒。
    'sendTimeout = 120000 '发送数据的超时时间，120000 毫秒。
    'receiveTimeout = 60000 '接收 response 的超时时间，60000 毫秒。
    
    '    reqHeaders.Add Consts.TOOLS_HTTP_HEADER_FIELD_CONTENT_TYPE, Consts.TOOLS_HTTP_CONTENT_TYPE_FORM_URLENCODED
    '    reqHeaders.Add "req-from", "vb ajax class by 215879458@qq.com"
    '    postBody.Add "ver", 1#
    RequestDataBody.CompareMode = TextCompare
    RequestDataForm.CompareMode = TextCompare
    RequestDataQuery.CompareMode = TextCompare
    RequestHeaders.CompareMode = TextCompare
    Cookies.CompareMode = TextCompare
    ResponseHeaders.CompareMode = TextCompare
    RequestTimeOut = 30
    RequestChartSet = "utf-8"
End Sub

Private Sub Class_Terminate()
    Set Inst = Nothing
End Sub

'WinHttpRequest 事件处理程序
Private Sub Inst_OnResponseFinished()
    '解析返回头
    Call ToolsHttp.ParseKeyValue(Inst.GetAllResponseHeaders(), ResponseHeaders)
    '自动处理 Set-Cookie 响应头
    If ResponseHeaders.Exists("Set-Cookie") Then
        Call ParseSetCookie(ResponseHeaders("Set-Cookie"))
    End If
    '触发自定义事件
    RaiseEvent OnResponseFinished
End Sub

Private Sub Inst_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    LastError = ErrorDescription & " (错误代码: " & ErrorNumber & ")"
    '记录错误到调试信息
    If DebugStart = True Then
        With DebugInfo.NewItem("AsyncError")
            .Item("Number") = ErrorNumber
            .Item("Description") = ErrorDescription
        End With
    End If
End Sub

'解析 Set-Cookie 响应头
Private Sub ParseSetCookie(ByVal SetCookieValue As String)
    If Trim(SetCookieValue) = "" Then Exit Sub
    
    Dim CookieParts() As String
    Dim CookieItem As String
    Dim CookieParts2() As String
    Dim CookieName As String
    Dim CookieValue As String
    
    '处理可能存在多个 Set-Cookie 的情况（使用逗号分隔或换行符）
    Dim MultiCookies() As String
    MultiCookies = Split(SetCookieValue, ",")
    
    Dim i As Long
    For i = LBound(MultiCookies) To UBound(MultiCookies)
        CookieItem = Trim(MultiCookies(i))
        If CookieItem <> "" Then
            '找到第一个等号分隔 name=value
            Dim EqPos As Long
            EqPos = InStr(CookieItem, "=")
            If EqPos > 0 Then
                CookieName = Trim(Left(CookieItem, EqPos - 1))
                '只取到分号或行尾
                Dim SemicolonPos As Long
                SemicolonPos = InStr(EqPos + 1, CookieItem, ";")
                If SemicolonPos > 0 Then
                    CookieValue = Mid(CookieItem, EqPos + 1, SemicolonPos - EqPos - 1)
                Else
                    CookieValue = Mid(CookieItem, EqPos + 1)
                End If
                Cookies(CookieName) = CookieValue
            End If
        End If
    Next i
End Sub




