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

Dim WithEvents oHttp As WinHttp.WinHttpRequest
Attribute oHttp.VB_VarHelpID = -1
Public JsonInst As New cJson
Public resJson As New Dictionary
'缓存请求的返回内容
Public Res As Variant
'模仿前几天写的php版ajax，做个完整处理过的 res.header+res.cookies+res.data
Public Event HttpRep(ByVal Dat)

Public Enum ResDataEnum
    Header = 0
    Cookies = 1
    Body = 2
End Enum

Public Enum DataEnum
    ResponseJson = 0
    ResponseBody = 1
    ResponseNoChange = 2
    ResponseText = 3
End Enum

'Public Enum ActEnum
'    optPOST = 0
'    optGET = 1
'End Enum

Public Enum AsyncEnum
    optSync = 0
    optAsync = 1
    optNoChange = 2
End Enum

Public Enum EnumPostContentType
    FormData = 0
    Multipart = 1
    Json = 2
End Enum

Public waitSeconds As Integer
Public LastErr As String
    Rem 存储上次地址
Public ioUrl As String
    Rem 存储基础请求地址
Public mBaseUrl As String
Public mContetType As String
Public repType As DataEnum
Public isAsync As AsyncEnum
Public addVer As Boolean
Public reqQuery As New Scripting.Dictionary
Private reqHeaders As New Scripting.Dictionary
Public postBody As New Scripting.Dictionary
Public postFilesPath As New Scripting.Dictionary
Public postFilesBin As New Scripting.Dictionary
Public postContentType As EnumPostContentType

Private isLongAjax As Boolean
Private LongAjaxReqType As String

Public Function getPostContentType(Optional ByVal cType As EnumPostContentType) As String
    On Error Resume Next
    Dim a: a = Array( _
    "application/x-www-form-urlencoded", _
    "multipart/form-data", _
    "application/json" _
    )
    getPostContentType = a(cType)
End Function
Public Function makeReqData(Dic As Scripting.Dictionary) As String
    If Dic.Count > 0 Then
        If postContentType = Json Then
            makeReqData = JsonInst.Encode(Dic)
        Else
            Dim Arr() As String
            ReDim Arr(Dic.Count - 1)
            Dim i%
            For Each ii In Dic
                Arr(i) = ii & "=" & Dic(ii)
                i = i + 1
            Next
            makeReqData = Join(Arr, "&")
        End If
    Else
        makeReqData = ""
    End If
End Function

Public Function Cancel()
    If isLongAjax Then
        Call LongAjax_stop
    Else
        oHttp.Abort
    End If
    d "请求被用户终止..."
End Function
'Public Function LongAjax_send(Dat) As Boolean  '更新一个版本,,独立设置一个长轮询对象,和短连接请求分开,以便一个类对象即可实现长轮询和短连接,用于发送和接受
'
'    If isLongAjax Then
'        oHttp.Send Dat
'        LongAjax_send = True
'    Else
'        LongAjax_send = False
'    End If
'    D "长轮询发送数据..." & LongAjax_send
'End Function

Public Function LongAjax_stop()
    LongAjaxReqType = ""
    isLongAjax = False
    Call oHttp.Abort
    d "长轮询终止..."
End Function

Public Function LongAjax_start(Optional ByVal isPost As Boolean, Optional ByVal Url As String, Optional ByVal Dat As String, Optional ByVal DataStic As DataEnum = ResponseNoChange) As Boolean
    
    If LongAjaxReqType = "" Then LongAjaxReqType = IIf(isPost, "POST", "GET")
    '    waitSeconds = 28 '长轮询默认主动不关闭,,如果需要更长时间,请使用 ohttp.SetTimeouts(),,默认大约32秒主动关闭,,可修改
    LongAjax_start = IO(LongAjaxReqType, Url, Dat, 0, DataStic, optAsync)
    If LongAjax_start = True Then
        isLongAjax = True
        d "长轮询开始..."
    Else
        d "长轮询错误..." & LastErr
    End If
End Function

Public Function GetData(Optional ByVal Url As String, Optional ByVal Dat As String, Optional ByVal waitSec As Integer, Optional ByVal DataStic As DataEnum = ResponseNoChange, Optional ByVal Async As AsyncEnum = optNoChange) As Boolean
Attribute GetData.VB_UserMemId = 0
    
    GetData = IO("GET", Url, "", waitSec, DataStic, Async)
    
End Function

Public Function PostData(Optional ByVal Url As String, Optional ByVal Dat As String, Optional ByVal waitSec As Integer, Optional ByVal DataStic As DataEnum = ResponseNoChange, Optional ByVal Async As AsyncEnum = optNoChange) As Boolean
    
    PostData = IO("POST", Url, Dat, waitSec, DataStic, Async)
    
End Function

Private Function getContentType() As String
    Select Case mContetType
    Case "json"
        getContentType = getPostContentType(Json)
    Case "form"
        getContentType = getPostContentType(FormData)
    Case "file"
        getContentType = getPostContentType(Multipart)
    Case Else
        getContentType = mContetType
    End Select
End Function

Private Function IO(Act$, Url$, Dat$, waitSec%, DataStic As DataEnum, Async As AsyncEnum) As Boolean
    On Error GoTo ERR:
    reqHeaders("CONTENT-TYPE") = getContentType()
    'reqHeaders("CONTENT-TYPE") = getPostContentType(postContentType)
    LastErr = ""
    '    JsonInst.Dic_de.RemoveAll
    '    JsonInst.Dic_en.RemoveAll
    '长轮询模式下不要在请求其他数据了
    If isLongAjax Then
        IO = False
        LastErr = "请先停止长轮询模式后再请求!或者另外实例化一个cAjax对象处理!"
        d LastErr
        Exit Function
    End If
    '检查url
    If Url = "" Then
        If ioUrl = "" Then
            IO = False
            LastErr = "请求的URL地址为空,请检查!"
            d LastErr
            Exit Function
        Else
            Url = ioUrl
        End If
    Else
        If mBaseUrl = "" Then
            '
        Else
            Url = mBaseUrl & Url
        End If
        ioUrl = Url
    End If
    
    '    D "开始请求..."
    '参数设置
    If Async <> optNoChange Then isAsync = Async
    If DataStic <> ResponseNoChange Then repType = DataStic
    '处理请求参数
    If reqQuery.Count > 0 Then
        If InStr(Url, "?") > 0 Then
            Url = Url & "&" & makeReqData(reqQuery)
        Else
            Url = Url & "?" & makeReqData(reqQuery)
        End If
    End If
    '是否添加URL资源版本号
    If addVer Then
        If InStr(Url, "?") > 0 Then
            Url = Url & "&ver=" & Rnd()                                         ' Timer()
        Else
            Url = Url & "?ver=" & Rnd()                                         ' Timer()
        End If
    End If
    '打开链路
    oHttp.Open Act, Url, True
    '处理Header
    If reqHeaders.Count > 0 Then
        For Each ii In reqHeaders
            If reqHeaders(ii) <> "" Then oHttp.SetRequestHeader ii, reqHeaders(ii)
        Next
    End If
    '处理POSTBODY
    If Dat = "" Then
        Dat = makeReqData(postBody)
    End If
    '忽略错误SSL证书
    oHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    '发送数据
    oHttp.Send Dat
    '发送文件
    '    If postFilesPath.Count > 0 Then
    '        Dim xx
    '        For Each xx In postFilesPath
    '
    '            oHttp.Send Dat
    '        Next
    '    End If
    
    '返回
    
    If isAsync = optSync Then
        '设置等待超时
        If waitSeconds < 5 Then waitSeconds = 300
        If waitSec = 0 Then waitSec = waitSeconds
        oHttp.WaitForResponse waitSec                                           '由于使用异步请求,如果用户参数改为同步请求,此处则等待返回,无需使用do循环+doevents
        Res = repData()
        If oHttp.Status <> 200 Or Res = False Then ERR.Raise oHttp.Status, , oHttp.Status & "#" & oHttp.StatusText & "#"
    End If
    IO = True
CLS:
    '清理body,但保留头部和get请求参数,头部和参数基本上是不变的,需要清理可以手工执行字典清楚方法
    postFilesPath.RemoveAll
    postFilesBin.RemoveAll
    postBody.RemoveAll
    reqQuery.RemoveAll
    '    reqHeaders.RemoveAll
    Exit Function
ERR:
    Rem todo 这里可以优化一下，对超时错误代码进行准确描述
    LastErr = ERR.Description
    IO = False                                                                  '仅限同步模式
    d "IO错误:" & LastErr
    GoTo CLS
End Function

Public Function SetCookies(ByVal Val$) As Boolean
    SetCookies = Me.SetResHeader("Cookie", Val)
End Function

Public Function SetResHeader(ByVal Key$, ByVal Val$) As Boolean
    If reqHeaders.Exists(Key) = False Then
        reqHeaders.Add Key, Val
    Else
        reqHeaders(Key) = Val
    End If
    SetResHeader = True
End Function


Public Function GetResHeader(Key$, Optional isCookie As Boolean) As String
    GetResHeader = oHttp.GetResponseHeader(Key)
End Function

Private Function repData() As Variant
    On Error Resume Next
    
    '    If (oHttp.GetResponseHeader("Content-Type") Like "*application/json*" Or repType = ResponseJson) And repType <> ResponseNoChange Then
    '        Call jsonInst.Decode(oHttp.ResponseText)
    '        Set res = jsonInst.Dic_de
    '        'D oHttp.ResponseText                                                    'jsonInst.Dic_de("ip")
    '        repData = oHttp.ResponseText                                            '此处可以copy-dic
    '    Else
    
    '    If repType = ResponseJson Then
    '        If JsonInst.Decode(oHttp.ResponseText) = True Then Set resJson = JsonInst.Dic_de
    '    End If
    '
    On Error GoTo ERR:
    '-----------------------------函数返回,,
    Select Case repType
    Case ResponseText
        '--------------------------------直接返回字符串
        repData = oHttp.ResponseText
    Case ResponseBody
        '--------------------------------直接返回二进制
        repData = oHttp.ResponseBody
    Case ResponseBody + ResponseText
        '---------------------------二进制转字符串[直接返回字串出现乱码时尝试]
        repData = StrConv(oHttp.ResponseBody, vbUnicode)
    Case Else
        '--------------------------------无效的返回
        repData = oHttp.ResponseText
    End Select
    
    'End If
    
    Exit Function
ERR:
    LastErr = "请求已经超时,请调整等待时间!" & ERR.Description
    repData = False                                                             '仅限同步模式
    d "rep错误:" & LastErr
End Function
'Private Function BytesToStr(ByVal vIn) As String
'  strReturn = ""
'  For i = 1 To LenB(vIn)
'    ThisCharCode = AscB(MidB(vIn, i, 1))
'    If ThisCharCode < &H80 Then
'      strReturn = strReturn & Chr(ThisCharCode)
'    Else
'      NextCharCode = AscB(MidB(vIn, i + 1, 1))
'      strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
'      i = i + 1
'    End If
'  Next
'  BytesToStr = strReturn
'End Function

Private Sub Class_Initialize()
    Set oHttp = New WinHttp.WinHttpRequest
    oHttp.SetProxy 1
    '    oHttp.SetTimeouts 2000, 2000, 6000, 20000
    'WinHttp.SetTimeouts 60000, 60000, 60000, 3000 '设置操作超时时间
    'WinHttp.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
    'resolveTimeout = 10000 '解析 DNS 名字的超时时间，10000 毫秒。
    'connectTimeout = 10000 '建立 Winsock 连接的超时时间，10000 毫秒。
    'sendTimeout = 120000 '发送数据的超时时间，120000 毫秒。
    'receiveTimeout = 60000 '接收 response 的超时时间，60000 毫秒。
    '    reqHeaders.Add "Content-Length", Len(strdata)
    mContetType = getPostContentType(FormData)
    reqHeaders.Add "CONTENT-TYPE", mContetType
    reqHeaders.Add "req-from", "vb ajax class by 215879458@qq.com"
    '    postBody.Add "ver", 1#
End Sub

Private Sub Class_Terminate()
    Set oHttp = Nothing
End Sub

Private Sub oHttp_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    d "WinHttp错误:[" & ErrorNumber & "]" & ErrorDescription
    Call LongAjax_restart
End Sub

Private Sub oHttp_OnResponseDataAvailable(data() As Byte)
    '本事件可用于传输进度
    
End Sub

Private Sub oHttp_OnResponseFinished()
'    D "http status: " & oHttp.Status
    If isAsync = optAsync Then
        RaiseEvent HttpRep(repData())
    End If
    Call LongAjax_restart
End Sub

Private Function LongAjax_restart() As Boolean
    If isLongAjax Then
        isLongAjax = False
        LongAjax_restart = LongAjax_start()
    Else
        LongAjax_restart = False
    End If
End Function
Public Sub d(ByVal S$)
    Debug.Print Now() & vbTab & S
End Sub
