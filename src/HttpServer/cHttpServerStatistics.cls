VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' HttpServer 性能统计对象
' 全部使用 Public Long 字段，递增操作零开销（无 Dictionary / 属性调用开销）
' 外部使用: Server.Statistics.TotalRequests 或 ctx.Statistics.TotalRequests

' ============================================================
' 请求统计
' ============================================================

' 累计请求总数
Public TotalRequests As Long

' 按请求方法分类
Public GetRequests As Long
Public PostRequests As Long
Public PutRequests As Long
Public DeleteRequests As Long
Public OptionsRequests As Long
Public HeadRequests As Long
Public PatchRequests As Long
Public OtherRequests As Long

' 按响应状态码分类（1xx~5xx）
Public Status1xx As Long
Public Status2xx As Long
Public Status3xx As Long
Public Status4xx As Long
Public Status5xx As Long

' ============================================================
' 连接统计
' ============================================================

' 累计接受的连接总数
Public TotalConnectionsAccepted As Long
' 因 MaxConnections 被拒绝的连接数
Public RejectedConnections As Long
' 峰值并发连接数
Public PeakConnections As Long
' 因空闲超时被清理的连接数
Public IdleConnectionsCleaned As Long

' ============================================================
' 错误统计
' ============================================================

' 请求处理异常次数（ProcessHttpRequest 中 EH 路径）
Public RequestErrors As Long
' 413 Payload Too Large 次数
Public RequestSizeRejected As Long
' SSE Entry 失败次数
Public SSEEntryErrors As Long

' ============================================================
' 流量统计（字节）
' ============================================================

' 累计接收字节数
Public TotalBytesReceived As Long
' 累计发送字节数
Public TotalBytesSent As Long

' ============================================================
' SSE 统计
' ============================================================

' SSE 连接总数
Public SSEConnectionsAccepted As Long

' ============================================================
' Session 统计
' ============================================================

' 累计创建的 Session 总数（仅 HasID 的有效 Session）
Public TotalSessionsCreated As Long

' ============================================================
' 时间统计
' ============================================================

' 服务器启动时间
Public StartTime As Date

' ============================================================
' 辅助方法
' ============================================================

' 获取服务器运行时长（秒）
Public Property Get UptimeSeconds() As Long
    If StartTime = 0 Then
        UptimeSeconds = 0
    Else
        UptimeSeconds = CLng(DateDiff("s", StartTime, Now))
    End If
End Property

' 获取每秒平均请求数（QPS）
Public Property Get AverageQPS() As Double
    Dim sec As Long
    sec = UptimeSeconds
    If sec > 0 Then
        AverageQPS = TotalRequests / CDbl(sec)
    Else
        AverageQPS = 0
    End If
End Property

' 按方法名递增对应计数器（内部使用，比 Select Case 更高效的方式留给调用方）
Friend Sub IncrementByMethod(ByVal MethodValue As Long)
    Select Case MethodValue
        Case 1: PostRequests = PostRequests + 1    ' ReqPost = 1
        Case 2: GetRequests = GetRequests + 1      ' ReqGet = 2
        Case 3: PutRequests = PutRequests + 1      ' ReqPut = 3
        Case 4: DeleteRequests = DeleteRequests + 1 ' ReqDelete = 4
        Case 5: OptionsRequests = OptionsRequests + 1 ' ReqOptions = 5
        Case 6: HeadRequests = HeadRequests + 1
        Case 7: PatchRequests = PatchRequests + 1
        Case Else: OtherRequests = OtherRequests + 1
    End Select
End Sub

' 按状态码区间递增（传入实际状态码如 200、404、500）
Friend Sub IncrementByStatusCode(ByVal StatusCode As Long)
    Select Case StatusCode \ 100
        Case 1: Status1xx = Status1xx + 1
        Case 2: Status2xx = Status2xx + 1
        Case 3: Status3xx = Status3xx + 1
        Case 4: Status4xx = Status4xx + 1
        Case 5: Status5xx = Status5xx + 1
    End Select
End Sub

' 更新峰值连接数
Friend Sub UpdatePeakConnections(ByVal CurrentCount As Long)
    If CurrentCount > PeakConnections Then
        PeakConnections = CurrentCount
    End If
End Sub

' 重置所有统计
Public Sub Reset()
    TotalRequests = 0
    GetRequests = 0
    PostRequests = 0
    PutRequests = 0
    DeleteRequests = 0
    OptionsRequests = 0
    HeadRequests = 0
    PatchRequests = 0
    OtherRequests = 0
    
    Status1xx = 0
    Status2xx = 0
    Status3xx = 0
    Status4xx = 0
    Status5xx = 0
    
    TotalConnectionsAccepted = 0
    RejectedConnections = 0
    PeakConnections = 0
    IdleConnectionsCleaned = 0
    
    RequestErrors = 0
    RequestSizeRejected = 0
    SSEEntryErrors = 0
    
    TotalBytesReceived = 0
    TotalBytesSent = 0
    
    SSEConnectionsAccepted = 0
    TotalSessionsCreated = 0
    
    StartTime = Now
End Sub

Private Sub Class_Initialize()
    StartTime = Now
End Sub
