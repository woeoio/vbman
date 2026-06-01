VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebDavClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub test()
    
    Dim xmlhttp As Object
    Dim url As String
    Dim response As String
    Dim username As String
    Dim password As String
    Dim authHeader As String
    
    url = "http://192.168.168.4:7070/webDAV/admin"
    username = "admin"
    password = "00000000"
    
    ' 将用户名和密码组合为 "username:password"
    authHeader = "Basic YWRtaW46MDAwMDAwMDA="                                   '& EncodeBase64(username & ":" & password)
    
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "PROPFIND", url, False
    xmlhttp.SetRequestHeader "Depth", "1"                                       ' 请求深度，通常是 1 或 0
    xmlhttp.SetRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    xmlhttp.SetRequestHeader "Authorization", authHeader                        ' 添加认证头
    xmlhttp.Send "<D:propfind xmlns:D='DAV:'><D:allprop/></D:propfind>"
    
    response = xmlhttp.ResponseText
    '    Form1.Text1.Text = response
    
    Set xmlhttp = Nothing
End Sub

