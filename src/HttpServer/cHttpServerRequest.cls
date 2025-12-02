VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EnumRequestMethod
    ReqPost = 1
    ReqGet = 2
    ReqPut = 3
    ReqDelete = 4
    ReqOptions = 5
End Enum
Public Enum EnumRequestContentType
    JsonString = 1
    FormEncode = 2
    FormMultiPart = 3
    TextPlain = 4
    TextHtml = 5
    Custom = 6
End Enum

Public Method As EnumRequestMethod
Public MethodName As String
Public HttpVersion As String
Public url As String
Public PathInfo As String
Public PathInfoController As String
Public PathInfoAction As String
Public QueryInfo As String
Public QueryString As New Scripting.Dictionary
Public Form As New Scripting.Dictionary
Public Json As New cJson
Public Header As New Scripting.Dictionary
Public HeaderRaw  As String
Public HeaderUserAgent As String
Public HeaderAccept As String
Public HeaderHost As String
Public HeaderHostDomain As String
Public HeaderHostPort As Long
Public HeaderContentType As String
Public HeaderContentLength As Long
Dim pvRawBodyText As String
Dim pvRawBodyBin() As Byte

Public LastError As String

Public Property Get RawBodyText() As String
    RawBodyText = pvRawBodyText
End Property
Public Property Get RawBodyBin() As Byte()
    RawBodyBin = pvRawBodyBin
End Property
Public Property Get Item(ByVal key) As Variant
Attribute Item.VB_Description = "自动从请求参数，表单参数，json参数获取值"
Attribute Item.VB_UserMemId = 0
    '
End Property

Private Function ParseUri(sRequest As String) As Boolean
    Dim oneLine: oneLine = Split(sRequest, vbCrLf, 2)
    Dim a: a = Split(oneLine(0), " ")
    MethodName = UCase(a(0))
    Method = ToolsHttp.MapMethod(MethodName)
    HttpVersion = a(2)
    If a(1) = "/" Then PathInfo = "/": GoTo QH
    url = ToolsHttp.UrlDecodeUtf8(a(1))
    Dim PQ: PQ = Split(url, "?", 2)
    PathInfo = PQ(0)
    If UBound(PQ) > 0 Then QueryInfo = PQ(1)
    If PathInfo <> "/" Then
        Dim p: p = Split(PathInfo & "///", "/", 4)
        PathInfoController = p(1)
        PathInfoAction = p(2)
    End If
    If QueryInfo <> "" Then Call ToolsHttp.ParseContent(QueryInfo, QueryString)
QH:
    ParseUri = True
    Exit Function
End Function

Friend Function ParseHeader(sRequest As String) As Boolean
    On Error GoTo EH
    HeaderRaw = sRequest
    Dim UH: UH = Split(sRequest, vbCrLf, 2)
    Dim UI As String: UI = Trim(UH(0))
    If ParseUri(UI) = False Then ERR.Raise 500, "ParseUri", LastError
    Dim HI As String: HI = Trim(UH(1))
    Call ToolsHttp.ParseKeyValue(HI, Header)
    If Header.Exists("User-Agent") = True Then HeaderUserAgent = Header("User-Agent")
    If Header.Exists("Accept") = True Then HeaderAccept = Header("Accept")
    If Header.Exists("Content-Type") = True Then HeaderContentType = Header("Content-Type")
    If Header.Exists("Content-Length") = True Then HeaderContentLength = Header("Content-Length")
    If Header.Exists("Host") = True Then
        HeaderHost = Header("Host")
        Dim a: a = Split(HeaderHost & ":", ":")
        HeaderHostDomain = a(0)
        If a(1) = "" Then
            HeaderHostPort = 80
        Else
            HeaderHostPort = CLng(a(1))
        End If
    End If
QH:
    ParseHeader = True
    Exit Function
EH:
    LastError = "[PARSE_HEADER ERROR]" & ERR.Number & vbCrLf & sRequest & vbCrLf & ERR.Description
End Function

Friend Function ParseBody(aReqBody() As Byte) As Boolean
    On Error GoTo EH
    Dim b As String
    '先进行编码转换
    b = ToolsUtf8.Decode(aReqBody)
    '解析表单
    If ToolsStr.HasStr(Consts.TOOLS_HTTP_CONTENT_TYPE_FORM_URLENCODED, HeaderContentType) > 0 Then
        '        b = ToolsHttp.UrlDecodeUtf8(b)
        If b = "" Then GoTo QH
        Call ToolsHttp.ParseContent(b, Form): GoTo QH
    End If
    '解析json
    If ToolsStr.HasStr(Consts.TOOLS_HTTP_CONTENT_TYPE_JSON, HeaderContentType) > 0 Then
        Call Json.Decode(b)
    End If
    '解析
QH:
    '赋值给原始值属性组
    pvRawBodyText = b
    pvRawBodyBin = aReqBody
    ParseBody = True
    Exit Function
EH:
    LastError = "[PARSE_BODY ERROR]" & ERR.Number & vbCrLf & ERR.Description
End Function

Private Sub Class_Initialize()
    Header.CompareMode = TextCompare
    QueryString.CompareMode = TextCompare
    Form.CompareMode = TextCompare
End Sub
