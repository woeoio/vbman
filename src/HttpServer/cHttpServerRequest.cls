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
    ReqOption = 5
End Enum

Public Method As EnumRequestMethod
Public MethodName As String
Public HttpVersion As String
Public RawContent As String
Public Url As String
Public PathInfo As String
Public PathInfoController As String
Public PathInfoAction As String
Public QueryInfo As String
Public QueryString As New Scripting.Dictionary
Public Form As New Scripting.Dictionary
Public Json As New Scripting.Dictionary
Public Header As New Scripting.Dictionary
Public HeaderUserAgent As String
Public HeaderAccept As String
Public HeaderHost As String
Public HeaderHostDomain As String
Public HeaderHostPort As Long
Public HeaderContentType As String
Public HeaderContentLength As Long


Friend Function ParseHeaderAndBody(sRequest As String) As Boolean
    Dim HB: HB = Split(sRequest, vbCrLf & vbCrLf, 2)
    Dim UH: UH = Split(HB(0), vbCrLf, 2)
    Dim H As String: H = Trim(UH(1))
    Call ToolsHttp.ParseKeyValue(H, Header)
    If Header.Exists("User-Agent") = True Then HeaderUserAgent = Header("User-Agent")
    If Header.Exists("Accept") = True Then HeaderAccept = Header("Accept")
    If Header.Exists("Content-Type") = True Then HeaderContentType = Header("Content-Type")
    If Header.Exists("Content-Length") = True Then HeaderContentLength = Header("Content-Length")
    If Header.Exists("Host") = True Then
        HeaderHost = Header("Host")
        Dim A: A = Split(HeaderHost & ":", ":")
        HeaderHostDomain = A(0)
        If A(1) = "" Then
            HeaderHostPort = 80
        Else
            HeaderHostPort = CLng(A(1))
        End If
    End If
    '
    Dim b As String: b = Trim(HB(1))
    If b = "" Then GoTo QH
    '先进行编码转换
    b = ToolsHttp.UrlDecodeUtf8(b)
    '解析表单
    If ToolsStr.HasStr(ToolsHttp.C_CONTENT_TYPE_FORM_URLENCODED, HeaderContentType) > 0 Then
        Call ToolsHttp.ParseContent(b, Form): GoTo QH
    End If
    '解析json
    If ToolsStr.HasStr(ToolsHttp.C_CONTENT_TYPE_JSON, HeaderContentType) > 0 Then
        
    End If
    '解析
QH:
    ParseHeaderAndBody = True
    Exit Function
End Function

Friend Function ParseUri(sRequest As String) As Boolean
    RawContent = sRequest
    Dim oneLine: oneLine = Split(sRequest, vbCrLf, 2)
    Dim A: A = Split(oneLine(0), " ")
    MethodName = UCase(A(0))
    Method = ToolsHttp.MapMethod(MethodName)
    HttpVersion = A(2)
    If A(1) = "/" Then PathInfo = "/": GoTo QH
    Url = ToolsHttp.UrlDecodeUtf8(A(1))
    Dim PQ: PQ = Split(Url, "?", 2)
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


