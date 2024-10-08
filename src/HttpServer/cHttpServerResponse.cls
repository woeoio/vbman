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
End Property


'Public Sub Write_()
'    '
'End Sub

Public Sub Text()
    '
End Sub
Public Sub Html()
    '
End Sub
Public Sub View()
    '
End Sub

'自定义对象

Public Sub Json(Data As Variant, Optional Code As Long, Optional Msg As String, Optional Count As Long)
    '
End Sub

Public Sub File(Path As String)
    '自动判断是否ajax请求，决定返回页面还是json
    Dim sPath As String: sPath = ToolsFso.AutoCompleteFullPath(Client.Context.Server.WebRoot & Path, True)
    If ToolsFso.Inst.FileExists(sPath) = False Then Call State404: Exit Sub
    Dim Data() As Byte
    If ToolsStream.LoadFileAsBinary(sPath, Data) = False Then Call State404: Exit Sub
    SendBodyByte Data, ToolsHttpMimeType.MapMimeType(sPath)
End Sub

Public Sub State(Code As Long, Optional Say As String = "没有找到页面")
    '自动判断是否ajax请求，决定返回页面还是json
    Sck.SendData ""
End Sub
Public Sub State404(Optional Say As String = "没有找到页面")
    '自动判断是否ajax请求，决定返回页面还是json
    Sck.SendData ""
End Sub
Public Sub State502(Optional Say As String = "没有找到页面")
    '自动判断是否ajax请求，决定返回页面还是json
    Sck.SendData ""
End Sub
Public Sub State500(Optional Say As String = "没有找到页面")
    '自动判断是否ajax请求，决定返回页面还是json
    SendHeader
    SendBodyText Say, "text/html"
End Sub

Private Sub Class_Initialize()
    JsonPackFieldNameCode = "code"
    JsonPackFieldNameData = "data"
    JsonPackFieldNameMsg = "msg"
    JsonPackFieldNameCount = "count"
End Sub


Private Sub SendBodyText(Data As String, DataType As String)
    '
End Sub
Private Sub SendBodyByte(Data() As Byte, DataType As String)
    Header("Content-Type") = DataType
    Header("Content-Length") = UBound(Data) + 1
    SendHeader
    Sck.SendData Data
    Header.RemoveAll
End Sub

Private Sub SendHeader(Optional Status As Long = 200, Optional StatusText As String = "OK", Optional HttpVersioin As String = "1.1")
    Dim Head As String: Head = "HTTP/" & HttpVersioin & " " & Status & " " & StatusText & vbCrLf
    '
    Header("Server") = ServerName & " " & Client.Parent.Version
    Dim x, i As Long
    For Each x In Header.Keys()
        Head = Head & x & ": " & Header.Item(x) & vbCrLf
    Next
    Sck.SendData Head & vbCrLf
End Sub


Public Sub Test()
    Dim sRequest As String: sRequest = Client.Context.Request.RawContent
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
