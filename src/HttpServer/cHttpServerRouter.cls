VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'路由才是开发者和框架之间的桥梁

Private Const C_METHOD_ANY As String = "ANY"
Private Const C_METHOD_GET As String = "GET"
Private Const C_METHOD_POST As String = "POST"
Private Const C_METHOD_PUT As String = "PUT"
Private Const C_METHOD_DELETE As String = "DELTET"


Dim Root As New Scripting.Dictionary




Friend Function Entry(Client As cClientCallback) As Boolean
    '    Debug.Print "OnDataArrival, Index=" & Index & ", bytesTotal=" & oCallback.dataTotal, Timer
    Dim sRequest            As String
    Client.Socket.GetData sRequest
    
    Call test(Client, sRequest)
    Entry = True
End Function





Private Sub test(oCallback As cClientCallback, sRequest As String)
    Dim Index As Long: Index = oCallback.Index
    Dim vSplit              As Variant
    Dim sBody               As String
    vSplit = Split(sRequest, vbCrLf)
    If UBound(vSplit) >= 0 Then
        '        Debug.Print vSplit(0)
        sBody = Join(vSplit, vbCrLf & Index & ": ") & vbCrLf & _
        vbCrLf & Index & ": Current time is " & Now & _
        vbCrLf & Index & ": RemoteHostIP is " & oCallback.Socket.RemoteHostIP & _
        vbCrLf & Index & ": RemotePort is " & oCallback.Socket.RemotePort
        oCallback.Socket.SendData "HTTP/1.1 200 OK" & vbCrLf & _
"Content-Type: text/html" & vbCrLf & _
"Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
        sBody
    End If
    '    Debug.Print "OnDataArrival, Index=" & Index & ", done", Timer
End Sub


