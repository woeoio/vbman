VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim Client As cClientCallback
Dim Sck As cTlsReMaster
Dim CrossDomain As cHttpCrossDomain

Dim JsonInst As New cJson

Const ServerName As String = "VBMAN"

Public JsonPackFieldNameCode As String
Public JsonPackFieldNameMsg As String
Public JsonPackFieldNameData As String
Public JsonPackFieldNameCount As String

Public Header As New Scripting.Dictionary

'兼容asp对象
Public Buffer As Boolean
Public CacheControl  As String
Public CharSet As String
Public CodePage  As Long
Public ContentType As String
Public Cookies As Scripting.Dictionary
Public Expires  As Variant
Public ExpiresAbsolute  As Variant
Public LCID As Long
Public Status  As String

Friend Property Set fClient(o As Object)
    Set Client = o
    Set Sck = Client.Socket
    Set CrossDomain = Client.Parent.CrossDomain
End Property


'Public Sub Write_()
'    '
'End Sub

Public Sub Text(Data As String)
    Call SendBodyText(Data, "text/plain")
End Sub
Public Sub Html(Data As String)
    Call SendBodyText(Data, "text/html")
End Sub
Public Sub View(TplPath As String)
    '
End Sub

'返回 json
Public Sub Json(Data As Variant, Optional Code As Long = -1, Optional msg As String, Optional Count As Long, Optional Whitespace As Variant)
    '判断是否 Recordset 对象，是就先转换为集合
    If TypeOf Data Is ADODB.Recordset Then
        Set Data = ToolsList.RsToCollection(Data)
    End If
    If Code = -1 Then
        Call SendBodyText(JsonInst.Encode(Data, Whitespace), "application/json")
        Exit Sub
    End If
    Dim Dic As New Scripting.Dictionary
    With Dic
        .Add JsonPackFieldNameCode, Code
        .Add JsonPackFieldNameMsg, msg
        .Add JsonPackFieldNameData, Data
        .Add JsonPackFieldNameCount, Count
    End With
    Call SendBodyText(JsonInst.Encode(Dic, Whitespace), "application/json")
End Sub

'返回 文件
Public Sub File(Path As String)
    '自动判断是否ajax请求，决定返回页面还是json
    Dim sPath As String: sPath = ToolsFso.AutoCompleteFullPath(Client.Context.Server.WebRoot & Path, True)
    If ToolsFso.Inst.FileExists(sPath) = False Then Call State404: Exit Sub
    Dim Data() As Byte
    If ToolsStream.LoadFileAsBinary(sPath, Data) = False Then Call State404: Exit Sub
    SendBodyByte Data, ToolsHttpMimeType.MapMimeType(sPath)
End Sub

Public Sub State(Code As Long, Text As String, Optional Say As String = "未知状态")
    '自动判断是否ajax请求，决定返回页面还是json
    SendBodyText Say, "text/html", Code, Text
End Sub
Public Sub State400(Optional Say As String)
    '自动判断是否ajax请求，决定返回页面还是json
    State 400, "BadRequest", Say                                                '"没有找到页面" & vbCrLf & Say
End Sub
Public Sub State401(Optional Say As String)
    '自动判断是否ajax请求，决定返回页面还是json
    State 401, "Unauthorized", Say                                              '"没有找到页面" & vbCrLf & Say
End Sub
Public Sub State403(Optional Say As String)
    '自动判断是否ajax请求，决定返回页面还是json
    State 403, "Forbidden", Say                                                 '"没有找到页面" & vbCrLf & Say
End Sub
Public Sub State404(Optional Say As String)
    '自动判断是否ajax请求，决定返回页面还是json
    State 404, "NotFound", Say                                                  '"没有找到页面" & vbCrLf & Say
End Sub
Public Sub State502(Optional Say As String)
    '自动判断是否ajax请求，决定返回页面还是json
    State 502, "GatewayError", Say                                              '"网关错误" & vbCrLf & Say
End Sub
Public Sub State500(Optional Say As String)
    '自动判断是否ajax请求，决定返回页面还是json
    State 500, "InternalError", Say                                             '"内部错误" & vbCrLf & Say
End Sub

Private Sub Class_Initialize()
    JsonPackFieldNameCode = "code"
    JsonPackFieldNameData = "data"
    JsonPackFieldNameMsg = "msg"
    JsonPackFieldNameCount = "count"
    '默认 utf8 ，可根据前端请求标头自动设置编码输出
    CharSet = "utf-8"
    CodePage = 65001
End Sub

'Private Sub SendBodyTextAsUtf8(Data As String, DataType As String)
'    Call SendBodyByte(ToolsUtf8.Encode(Data), DataType)
'End Sub
Private Sub SendBodyText(Data As String, DataType As String, Optional Status As Long = 200, Optional StatusText As String = "OK")
    Dim d() As Byte
    If Data = "" Then
        Header("Content-Type") = "text/plain; charset=" & CharSet
        Header("Content-Length") = 0
        SendHeader Status, StatusText
        Header.RemoveAll
        Exit Sub
    End If
    If CodePage = 65001 Then
        d = ToolsUtf8.Encode(Data)
    Else
        d = StrConv(Data, vbFromUnicode)
    End If
    Call SendBodyByte(d, DataType, Status, StatusText)
End Sub
Private Sub SendBodyByte(Data() As Byte, DataType As String, Optional Status As Long = 200, Optional StatusText As String = "OK")
    Header("Content-Type") = DataType & "; charset=" & CharSet
    Header("Content-Length") = UBound(Data) + 1                                 '注意空数组
    SendHeader Status, StatusText
    Sck.SendData Data
    Header.RemoveAll
End Sub

Private Sub SendHeader(Optional Status As Long = 200, Optional StatusText As String = "OK", Optional HttpVersioin As String = "1.1")
    Dim Head As String: Head = "HTTP/" & HttpVersioin & " " & Status & " " & StatusText & vbCrLf
    '
    Dim UserVer As String: UserVer = "FREE"
    Dim Ver As String: Ver = "[" & UserVer & "] "
    Header("Server") = Ver & ServerName & " " & Common.Version()
    If UserVer = "FREE" Then
        Header("PowerBy") = "www.vb6.pro"
    End If
    If CrossDomain.Enable = True Then
        CrossDomain.AddTo Header
    End If
    Dim x, i As Long
    For Each x In Header.Keys()
        Head = Head & x & ": " & Header.Item(x) & vbCrLf
    Next
    Sck.SendData Head & vbCrLf
End Sub


Public Sub test()
    Dim sRequest As String: sRequest = Client.Context.Request.HeaderRaw
    Dim Index As Long: Index = Client.Index
    Dim vSplit              As Variant
    Dim sBody               As String
    vSplit = Split(sRequest, vbCrLf)
    If UBound(vSplit) >= 0 Then
        '        Debug.Print vSplit(0)
        sBody = Join(vSplit, vbCrLf & Index & ": ") & vbCrLf & _
        vbCrLf & Index & ": Current time is " & Now & _
        vbCrLf & Index & ": RemoteHostIP is " & Sck.RemoteHostIP & _
        vbCrLf & Index & ": RemotePort is " & Sck.RemotePort
        Sck.SendData "HTTP/1.1 200 OK" & vbCrLf & _
"Content-Type: text/html" & vbCrLf & _
"Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
        sBody
    End If
    '    Debug.Print "OnDataArrival, Index=" & Index & ", done", Timer
End Sub
