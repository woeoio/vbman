Attribute VB_Name = "mSSEDemo"
Option Explicit
'SSE 客户端使用示例
'作者：邓伟，QQ：215879458

Dim WithEvents SSE As cSSEClient
Dim JsonParser As cJson

'基础示例
Public Sub Demo_Basic()
    Set SSE = New cSSEClient
    
    '配置 SSE 客户端
    SSE.AutoReconnect = True          '启用自动重连
    SSE.ReconnectInterval = 3000      '3秒重连间隔
    SSE.MaxReconnectAttempts = 10     '最多重连10次
    
    '连接到 SSE 服务端
    '替换为实际的 SSE 服务端地址
    SSE.Connect "http://localhost:3000/events"
    
    Debug.Print "正在连接 SSE 服务端..."
End Sub

'带认证的 SSE 连接示例
Public Sub Demo_WithAuth()
    Set SSE = New cSSEClient
    
    '配置 SSE 客户端
    SSE.AutoReconnect = True
    SSE.ReconnectInterval = 5000
    
    '连接前可以设置 HttpClient 的请求头（需要在连接前手动设置）
    '注意：这里需要在 Connect 之前通过 SSE.HttpClient 访问
    '但由于 HttpClient 是私有变量，需要在 SSE 类中暴露它
    
    '连接到需要认证的 SSE 服务端
    SSE.Connect "http://localhost:3000/protected-events"
    
    Debug.Print "正在连接需要认证的 SSE 服务端..."
End Sub

'断开连接示例
Public Sub Demo_Disconnect()
    If Not SSE Is Nothing Then
        SSE.DisableAutoReconnect  '关闭自动重连
        SSE.Disconnect           '断开连接
        Debug.Print "SSE 连接已断开"
    End If
End Sub

'重置重连计数器
Public Sub Demo_ResetReconnect()
    If Not SSE Is Nothing Then
        SSE.ResetReconnectAttempts
        Debug.Print "重连计数器已重置"
    End If
End Sub

'获取连接状态
Public Sub Demo_GetStatus()
    If SSE Is Nothing Then
        Debug.Print "SSE 客户端未初始化"
        Exit Sub
    End If
    
    Debug.Print "连接状态: " & IIf(SSE.Connected, "已连接", "未连接")
    Debug.Print "正在连接: " & IIf(SSE.Connecting, "是", "否")
    Debug.Print "最后事件ID: " & SSE.LastReceivedEventId
    Debug.Print "服务器重连间隔: " & SSE.ServerReconnectInterval & " 毫秒"
    Debug.Print "当前重连延迟: " & SSE.CurrentReconnectDelay & " 毫秒"
End Sub

'====== SSE 事件处理程序 ======

'连接建立事件
Private Sub SSE_OnOpen()
    Debug.Print "[SSE] 连接已建立"
End Sub

'接收消息事件
Private Sub SSE_OnMessage(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    Debug.Print "[SSE] 收到消息 - 事件: " & EventName & ", ID: " & Id
    Debug.Print "[SSE] 数据: " & Data
    
    '如果是 JSON 数据，解析处理
    If EventName = "message" Or EventName = "update" Then
        If JsonParser Is Nothing Then Set JsonParser = New cJson
        Call JsonParser.Decode(Data)
        
        '在这里处理业务逻辑
        '例如：更新UI、存储数据等
    End If
    
    '处理特定事件类型
    Select Case EventName
        Case "notification"
            Debug.Print "[SSE] 收到通知: " & Data
        Case "alert"
            Debug.Print "[SSE] 收到告警: " & Data
        Case "heartbeat"
            '心跳消息，可以忽略或用于保活
            Debug.Print "[SSE] 收到心跳"
        Case Else
            Debug.Print "[SSE] 未知事件类型: " & EventName
    End Select
End Sub

'错误事件
Private Sub SSE_OnError(ByVal Description As String, ByVal ErrorNumber As Long)
    Debug.Print "[SSE] 错误 - " & Description & " (代码: " & ErrorNumber & ")"
    
    '如果达到最大重连次数，可以提示用户
    If InStr(Description, "最大重连次数") > 0 Then
        Debug.Print "[SSE] 已达到最大重连次数，请检查网络或服务端状态"
    End If
End Sub

'连接关闭事件
Private Sub SSE_OnClose()
    Debug.Print "[SSE] 连接已关闭"
End Sub

'====== 高级使用示例 ======

'聊天室实时消息示例
Public Sub Demo_ChatRoom()
    Set SSE = New cSSEClient
    
    '连接到聊天室 SSE 端点
    SSE.AutoReconnect = True
    SSE.Connect "http://localhost:3000/chat/room1"
    
    Debug.Print "正在加入聊天室..."
End Sub

Private Sub SSE_OnMessage_ForChat(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    '处理聊天消息
    If EventName = "chat" Then
        '解析 JSON 消息
        Set JsonParser = New cJson
        JsonParser.Decode Data
        
        Dim UserName As String
        Dim Message As String
        Dim TimeStamp As String
        
        '假设消息格式: {"user": "张三", "message": "你好", "time": "2025-01-19 10:30:00"}
        UserName = JsonParser.Item("user")
        Message = JsonParser.Item("message")
        TimeStamp = JsonParser.Item("time")
        
        Debug.Print "[聊天室] " & UserName & " (" & TimeStamp & "): " & Message
        
        '可以在这里更新聊天界面
    ElseIf EventName = "system" Then
        '系统消息
        Debug.Print "[系统消息] " & Data
    End If
End Sub

'股票价格推送示例
Public Sub Demo_StockPrice()
    Set SSE = New cSSEClient
    
    '连接到股票价格推送端点
    SSE.AutoReconnect = True
    SSE.ReconnectInterval = 2000
    SSE.Connect "http://localhost:3000/stock/price"
    
    Debug.Print "正在订阅股票价格推送..."
End Sub

Private Sub SSE_OnMessage_ForStock(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    '处理股票价格更新
    If EventName = "price" Then
        Set JsonParser = New cJson
        JsonParser.Decode Data
        
        Dim Symbol As String
        Dim Price As Double
        Dim Change As Double
        Dim PercentChange As Double
        
        Symbol = JsonParser.Item("symbol")
        Price = CDbl(JsonParser.Item("price"))
        Change = CDbl(JsonParser.Item("change"))
        PercentChange = CDbl(JsonParser.Item("percent"))
        
        Debug.Print "[股票] " & Symbol & " 价格: " & Price & _
                   " 涨跌: " & Change & " (" & PercentChange & "%)"
    End If
End Sub

'服务器日志推送示例
Public Sub Demo_ServerLog()
    Set SSE = New cSSEClient
    
    '连接到服务器日志推送端点
    SSE.AutoReconnect = True
    SSE.MaxReconnectAttempts = 20
    SSE.Connect "http://localhost:3000/logs/stream"
    
    Debug.Print "正在接收服务器日志推送..."
End Sub

Private Sub SSE_OnMessage_ForLog(ByVal EventName As String, ByVal Data As String, ByVal Id As String)
    '处理日志推送
    If EventName = "log" Then
        '日志格式: {"level": "INFO", "message": "...", "timestamp": "..."}
        Set JsonParser = New cJson
        JsonParser.Decode Data
        
        Dim Level As String
        Dim Message As String
        Dim TimeStamp As String
        
        Level = JsonParser.Item("level")
        Message = JsonParser.Item("message")
        TimeStamp = JsonParser.Item("timestamp")
        
        '根据日志级别显示不同颜色（在支持的环境下）
        Select Case Level
            Case "ERROR"
                Debug.Print "[ERROR] " & TimeStamp & " - " & Message
            Case "WARN"
                Debug.Print "[WARN]  " & TimeStamp & " - " & Message
            Case "INFO"
                Debug.Print "[INFO]  " & TimeStamp & " - " & Message
            Case "DEBUG"
                Debug.Print "[DEBUG] " & TimeStamp & " - " & Message
        End Select
    End If
End Sub
