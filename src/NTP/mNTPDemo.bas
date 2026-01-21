Attribute VB_Name = "mNTPDemo"
'=========================================================================
'
' mNTPDemo - NTP 模块演示代码
'
' Purpose: 展示 NTP 客户端和服务端的使用方法
'
'=========================================================================

Option Explicit

'=========================================================================
' NTP 客户端演示
'=========================================================================

Public Sub NTPClientDemo_Sync()
    ' 同步获取服务器时间
    Dim oClient As cNTPClient
    Dim dtServerTime As Date
    Dim dOffset As Double
    
    Set oClient = New cNTPClient
    
    ' 方法1: 直接获取服务器时间
    dtServerTime = oClient.GetServerTime("pool.ntp.org", 123)
    Debug.Print "Server Time: " & Format(dtServerTime, "yyyy-mm-dd hh:mm:ss")
    
    ' 方法2: 获取时间偏移量
    dOffset = oClient.GetTimeOffset("pool.ntp.org", 123)
    Debug.Print "Time Offset: " & dOffset & " seconds"
    If dOffset > 0 Then
        Debug.Print "Local time is " & dOffset & " seconds ahead of server"
    ElseIf dOffset < 0 Then
        Debug.Print "Local time is " & Abs(dOffset) & " seconds behind server"
    Else
        Debug.Print "Local time is synchronized with server"
    End If
End Sub

Public Sub NTPClientDemo_Async()
    ' 异步获取服务器时间
    Dim WithEvents oClient As cNTPClient
    Dim dtServerTime As Date
    Dim dOffset As Double
    
    Set oClient = New cNTPClient
    oClient.ServerHost = "pool.ntp.org"
    oClient.ServerPort = 123
    
    ' 启动异步请求
    Debug.Print "Sending async NTP request..."
    oClient.RequestTimeAsync
    
    ' 事件处理代码
    ' (在实际应用中，这应该在类模块中实现 WithEvents)
    ' Private Sub oClient_TimeSyncComplete(ByVal ServerHost As String, ...
End Sub

Public Sub NTPClientDemo_MultiServer()
    ' 从多个服务器获取时间并取平均值
    Dim aServers(2) As String
    Dim aTimes(2) As Date
    Dim aOffsets(2) As Double
    Dim i As Long
    Dim dtAvgTime As Date
    Dim dAvgOffset As Double
    Dim dTotalSecs As Double
    
    aServers(0) = "pool.ntp.org"
    aServers(1) = "time.windows.com"
    aServers(2) = "time.google.com"
    
    For i = 0 To 2
        Dim oClient As cNTPClient
        Set oClient = New cNTPClient
        Debug.Print "Querying " & aServers(i) & "..."
        aTimes(i) = oClient.GetServerTime(aServers(i), 123)
        aOffsets(i) = oClient.GetTimeOffset(aServers(i), 123)
        Debug.Print "  Time: " & Format(aTimes(i), "hh:mm:ss") & ", Offset: " & aOffsets(i) & "s"
    Next i
    
    ' 计算平均时间
    For i = 0 To 2
        dTotalSecs = dTotalSecs + DateDiff("s", #1/1/1970#, aTimes(i))
    Next i
    dtAvgTime = DateAdd("s", dTotalSecs / 3, #1/1/1970#)
    
    ' 计算平均偏移
    dAvgOffset = (aOffsets(0) + aOffsets(1) + aOffsets(2)) / 3
    
    Debug.Print "Average Time: " & Format(dtAvgTime, "yyyy-mm-dd hh:mm:ss")
    Debug.Print "Average Offset: " & dAvgOffset & " seconds"
End Sub

'=========================================================================
' NTP 服务端演示
'=========================================================================

Public Sub NTPServerDemo_Basic()
    ' 基本 NTP 服务器
    Dim WithEvents oServer As cNTPServer
    
    Set oServer = New cNTPServer
    
    ' 配置服务器
    oServer.Stratum = 2
    oServer.ReferenceID = "LOCL"
    
    ' 启动服务器
    Debug.Print "Starting NTP server on port 123..."
    oServer.Start 123
    
    Debug.Print "Server is running. Press Ctrl+Break to stop."
    
    ' 服务器会在后台运行，响应 NTP 请求
    ' (在实际应用中，可以通过事件监听来处理)
End Sub

' (在实际的类模块中实现 WithEvents)
' Private Sub oServer_ServerStarted(ByVal Port As Long)
'     Debug.Print "Server started on port " & Port
' End Sub
'
' Private Sub oServer_RequestReceived(ByVal RemoteIP As String, ByVal RemotePort As Long)
'     Debug.Print "Request from: " & RemoteIP & ":" & RemotePort
' End Sub
'
' Private Sub oServer_ResponseSent(ByVal RemoteIP As String, ByVal RemotePort As Long)
'     Debug.Print "Response sent to: " & RemoteIP & ":" & RemotePort
' End Sub

Public Sub NTPServerDemo_CustomTime()
    ' 使用自定义时间源的 NTP 服务器
    Dim oServer As cNTPServer
    Dim oProvider As cNTPTimeProvider
    
    Set oServer = New cNTPServer
    Set oProvider = New cNTPTimeProvider
    
    ' 配置时间提供者
    oProvider.TimeSource = TimeSource_Custom
    oProvider.UseCustomTime = True
    oProvider.CustomTime = DateSerial(2024, 1, 1) + TimeSerial(12, 0, 0)
    oProvider.CustomOffset = 3600  ' 偏移1小时
    
    ' 设置服务器的提供者
    Set oServer.TimeProvider = oProvider
    
    ' 启动服务器
    Debug.Print "Starting NTP server with custom time source..."
    oServer.Start 1230  ' 使用非标准端口避免权限问题
    
    Debug.Print "Server is running on port 1230"
    Debug.Print "Providing time: " & Format(oProvider.GetCurrentTime(), "yyyy-mm-dd hh:mm:ss")
End Sub

Public Sub NTPServerDemo_ExternalTime()
    ' 使用外部时间源的 NTP 服务器
    Dim oServer As cNTPServer
    Dim oProvider As cNTPTimeProvider
    
    Set oServer = New cNTPServer
    Set oProvider = New cNTPTimeProvider
    
    ' 配置时间提供者为外部服务器
    oProvider.TimeSource = TimeSource_External
    oProvider.ExternalServer = "pool.ntp.org"
    oProvider.ExternalPort = 123
    
    ' 同步外部时间
    If oProvider.SyncExternalSource() Then
        Debug.Print "Synced with external time source"
    Else
        Debug.Print "Failed to sync with external time source, using system time"
        oProvider.TimeSource = TimeSource_System
    End If
    
    ' 设置服务器
    Set oServer.TimeProvider = oProvider
    oServer.Stratum = 3  ' 三级服务器
    
    ' 启动服务器
    Debug.Print "Starting NTP server on port 1230..."
    oServer.Start 1230
    
    Debug.Print "Server is running"
End Sub

'=========================================================================
' NTP 数据包演示
'=========================================================================

Public Sub NTPPacketDemo_Create()
    ' 创建 NTP 客户端请求包
    Dim oPacket As cNTPPacket
    
    Set oPacket = New cNTPPacket
    Set oPacket = oPacket.CreateClientRequest()
    
    ' 查看包内容
    Debug.Print "=== Client Request Packet ==="
    Debug.Print "LI: " & oPacket.LI
    Debug.Print "VN: " & oPacket.VN
    Debug.Print "Mode: " & oPacket.Mode
    Debug.Print "Stratum: " & oPacket.Stratum
    Debug.Print "Poll: " & oPacket.Poll
    Debug.Print "Precision: " & oPacket.Precision
    
    ' 编码为字节数组
    Dim baData() As Byte
    baData = oPacket.ToBytes()
    Debug.Print "Packet Size: " & UBound(baData) + 1 & " bytes"
    
    ' 创建服务器响应包
    Dim oResponse As cNTPPacket
    Set oResponse = New cNTPPacket
    Set oResponse = oResponse.CreateServerResponse(oPacket)
    
    Debug.Print ""
    Debug.Print "=== Server Response Packet ==="
    Debug.Print "LI: " & oResponse.LI
    Debug.Print "VN: " & oResponse.VN
    Debug.Print "Mode: " & oResponse.Mode
    Debug.Print "Stratum: " & oResponse.Stratum
    Debug.Print "ReferenceID: " & oResponse.ReferenceID
    Debug.Print "Transmit Time: " & Format(oResponse.TransmitTimestamp, "yyyy-mm-dd hh:mm:ss")
End Sub

Public Sub NTPPacketDemo_Parse()
    ' 解析 NTP 数据包
    Dim oPacket As New cNTPPacket
    
    ' 创建并编码一个包
    Dim oRequest As cNTPPacket
    Set oRequest = New cNTPPacket
    Set oRequest = oRequest.CreateClientRequest()
    
    ' 编码
    Dim baData() As Byte
    baData = oRequest.ToBytes()
    
    ' 解析
    If oPacket.FromBytes(baData) Then
        Debug.Print "Packet parsed successfully"
        Debug.Print "Version: " & oPacket.VN
        Debug.Print "Mode: " & oPacket.Mode
    Else
        Debug.Print "Failed to parse packet"
    End If
End Sub

'=========================================================================
' 综合演示
'=========================================================================

Public Sub NTPDemo_Complete()
    ' 完整的 NTP 客户端-服务端演示
    Debug.Print "=== NTP Complete Demo ===" & vbCrLf
    
    ' 1. 启动本地 NTP 服务器
    Dim oServer As cNTPServer
    Set oServer = New cNTPServer
    oServer.Stratum = 2
    oServer.ReferenceID = "TEST"
    
    Debug.Print "1. Starting local NTP server on port 1123..."
    oServer.Start 1123
    Debug.Print "   Server started" & vbCrLf
    
    ' 2. 从本地服务器查询时间
    Dim oClient As cNTPClient
    Set oClient = New cNTPClient
    
    Debug.Print "2. Querying local server..."
    Dim dtTime As Date
    dtTime = oClient.GetServerTime("127.0.0.1", 1123)
    Debug.Print "   Server time: " & Format(dtTime, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    
    ' 3. 查询公共 NTP 服务器
    Debug.Print "3. Querying public NTP server..."
    Dim dtPublicTime As Date
    dtPublicTime = oClient.GetServerTime("pool.ntp.org", 123)
    Debug.Print "   Public time: " & Format(dtPublicTime, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    
    ' 4. 计算本地服务器与公共服务器的时间差
    Dim dDiff As Double
    dDiff = DateDiff("s", dtTime, dtPublicTime)
    Debug.Print "4. Time difference: " & dDiff & " seconds" & vbCrLf
    
    ' 5. 停止服务器
    Debug.Print "5. Stopping server..."
    oServer.Stop
    Debug.Print "   Server stopped"
    
    Debug.Print vbCrLf & "=== Demo Complete ==="
End Sub

'=========================================================================
' 实用函数
'=========================================================================

' 批量查询多个 NTP 服务器
Public Function QueryMultipleServers(ByRef ServerList() As String) As Collection
    Dim oResults As New Collection
    Dim i As Long
    Dim oClient As cNTPClient
    Dim dtTime As Date
    Dim dOffset As Double
    Dim sResult As String
    
    For i = LBound(ServerList) To UBound(ServerList)
        On Error Resume Next
        Set oClient = New cNTPClient
        dtTime = oClient.GetServerTime(ServerList(i), 123)
        dOffset = oClient.GetTimeOffset(ServerList(i), 123)
        
        If Err.Number = 0 Then
            sResult = ServerList(i) & "|" & Format(dtTime, "yyyy-mm-dd hh:mm:ss") & "|" & dOffset
            oResults.Add sResult
            Debug.Print "? " & ServerList(i) & ": " & Format(dtTime, "hh:mm:ss") & " (offset: " & dOffset & "s)"
        Else
            Debug.Print "? " & ServerList(i) & ": Error - " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    Set QueryMultipleServers = oResults
End Function

' 查找最佳 NTP 服务器 (最小延迟)
Public Function FindBestServer(ByRef ServerList() As String) As String
    Dim oClient As cNTPClient
    Dim i As Long
    Dim dMinOffset As Double
    Dim dOffset As Double
    Dim sBestServer As String
    Dim lStartTime As Double
    Dim dLatency As Double
    Dim dMinLatency As Double
    
    dMinLatency = 9999
    sBestServer = ""
    
    For i = LBound(ServerList) To UBound(ServerList)
        On Error Resume Next
        Set oClient = New cNTPClient
        oClient.ServerHost = ServerList(i)
        oClient.ServerPort = 123
        oClient.Timeout = 3
        
        lStartTime = Timer
        dOffset = oClient.GetTimeOffset(ServerList(i), 123)
        dLatency = Timer - lStartTime
        
        If Err.Number = 0 Then
            If dLatency < dMinLatency Then
                dMinLatency = dLatency
                sBestServer = ServerList(i)
            End If
            Debug.Print ServerList(i) & ": latency=" & Format(dLatency * 1000, "0") & "ms, offset=" & dOffset & "s"
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    If LenB(sBestServer) <> 0 Then
        Debug.Print vbCrLf & "Best server: " & sBestServer & " (latency: " & Format(dMinLatency * 1000, "0") & "ms)"
    End If
    
    FindBestServer = sBestServer
End Function

'=========================================================================
' 主测试入口
'=========================================================================

Public Sub RunNTPDemos()
    Debug.Print String(50, "=")
    Debug.Print "NTP Module Demo"
    Debug.Print String(50, "=")
    Debug.Print
    
    ' 选择要运行的演示
    ' 1. NTPPacketDemo_Create
    ' 2. NTPClientDemo_Sync
    ' 3. NTPServerDemo_Basic
    ' 4. NTPDemo_Complete
    
    ' 运行完整演示
    NTPDemo_Complete
End Sub
