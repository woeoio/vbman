VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dHttpSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub NianBao(ctx As cHttpServerContext)
    ctx.Response.Text ctx.Request.QueryString("id")
End Sub
' 正常请求测试
Public Sub Index(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Normal request received"
End Sub

' 粘包测试接口 - 处理 /test0, /test1 等
Public Sub Test(ctx As cHttpServerContext)
    Dim requestId As String
    requestId = ctx.Request.QueryString("")
    ctx.Response.Text "OK: Sticky packet test " & requestId
End Sub

' 拆包/半包测试接口
Public Sub Partial(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Partial packet test received"
End Sub

' 延迟发送测试接口
Public Sub Delayed(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Delayed send test received"
End Sub

' 畸形请求测试接口
Public Sub Malformed(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Malformed request test received"
End Sub

Public Sub Double2(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Double CRLF test received"
End Sub

Public Sub NoHeaders(ctx As cHttpServerContext)
    ctx.Response.Text "OK: No headers test received"
End Sub

Public Sub Huge(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Huge header test received"
End Sub

Public Sub Incomplete(ctx As cHttpServerContext)
    ctx.Response.Text "OK: Incomplete request test received"
End Sub

' HTTP流水线测试接口 - 处理 /pipeline0, /pipeline1 等
Public Sub Pipeline(ctx As cHttpServerContext)
    Dim pipelineId As String
    pipelineId = ctx.Request.QueryString("")
    ctx.Response.Text "OK: Pipeline test " & pipelineId
End Sub

' 压力测试接口 - 处理 /stress/{worker_id}/{request_id}
Public Sub Stress(ctx As cHttpServerContext)
    Dim pathParts() As String
    Dim workerId As String
    Dim requestId As String
    
    pathParts = Split(ctx.Request.PathInfo, "/")
    If UBound(pathParts) >= 3 Then
        workerId = pathParts(2)
        requestId = pathParts(3)
        ctx.Response.Text "OK: Stress test - Worker " & workerId & " Request " & requestId
    Else
        ctx.Response.State400 "Bad Request: Invalid stress test path"
    End If
End Sub
