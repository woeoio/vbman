Attribute VB_Name = "Test_HttpServerStickPackage"
' ===========================================================================
' 模块名称: Test_HttpServerStickPackage.bas
' 功能说明: HttpServer TCP粘包处理单元测试模块
' 作者: 邓伟
' 邮箱: 215879458@qq.com
' 创建日期: 2026-04-03
'
' 测试目标:
'   - ContainsCompleteRequest: 完整请求检测逻辑
'   - ExtractOneRequest: 请求提取逻辑
'   - RemoveProcessedData: 缓冲区清理逻辑
'   - 完整工作流: 多请求粘包处理
'
' 使用说明:
'   在 Immediate 窗口中执行: RunAllTests
'   或在代码中调用: Test_HttpServerStickPackage.RunAllTests
' ===========================================================================

Option Explicit

' 测试结果数据结构
Private Type TestResult
    passed As Boolean
    testName As String
    message As String
    duration As Double
End Type

' 测试结果集合
Private testResults As Collection
Private totalTests As Long
Private passedTests As Long
Private failedTests As Long

' Windows API - 内存复制
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' ===========================================================================
' 公共接口：运行所有测试
' ===========================================================================
Public Sub RunAllTests()
    InitializeTestEnvironment
    
    ' 测试 ContainsCompleteRequest 函数
    Test_ContainsCompleteRequest_GetRequestOnlyHeaders
    Test_ContainsCompleteRequest_PostRequestWithBody
    Test_ContainsCompleteRequest_IncompleteHeaders
    Test_ContainsCompleteRequest_IncompleteBody
    Test_ContainsCompleteRequest_LargeBodyFragmented
    Test_ContainsCompleteRequest_MultipleRequestsInBuffer
    
    ' 测试 ExtractOneRequest 函数
    Test_ExtractOneRequest_GetRequest
    Test_ExtractOneRequest_PostRequestWithBody
    Test_ExtractOneRequest_MultipleRequests
    
    ' 测试 RemoveProcessedData 函数
    Test_RemoveProcessedData_SingleRequest
    Test_RemoveProcessedData_MultipleRequests
    Test_RemoveProcessedData_EmptyBufferAfterRemove
    
    ' 测试完整工作流
    Test_FullWorkflow_FragmentedHeaders
    Test_FullWorkflow_StickyPackages
    Test_FullWorkflow_LargeBodyFragmented
    Test_FullWorkflow_RealWorldScenario
    
    ' 输出测试报告
    PrintTestReport
    
    CleanupTestEnvironment
End Sub

' ===========================================================================
' 初始化测试环境
' ===========================================================================
Private Sub InitializeTestEnvironment()
    Set testResults = New Collection
    totalTests = 0
    passedTests = 0
    failedTests = 0
    
    Debug.Print String(80, "=")
    Debug.Print "HttpServer TCP粘包处理单元测试"
    Debug.Print String(80, "=")
    Debug.Print "开始时间: " & Now
    Debug.Print String(80, "-")
End Sub

' ===========================================================================
' 清理测试环境
' ===========================================================================
Private Sub CleanupTestEnvironment()
    Set testResults = Nothing
    Debug.Print String(80, "=")
    Debug.Print "测试完成时间: " & Now
    Debug.Print String(80, "=")
End Sub

' ===========================================================================
' 记录测试结果
' ===========================================================================
Private Sub RecordResult(testName As String, passed As Boolean, message As String, duration As Double)
    Dim result As TestResult
    
    totalTests = totalTests + 1
    
    With result
        .testName = testName
        .passed = passed
        .message = message
        .duration = duration
    End With
    
    testResults.Add result
    
    If passed Then
        passedTests = passedTests + 1
        Debug.Print "? " & message & " [" & Format(duration * 1000, "0.00") & "ms]"
    Else
        failedTests = failedTests + 1
        Debug.Print "? " & message & " [" & Format(duration * 1000, "0.00") & "ms]"
    End If
End Sub

' ===========================================================================
' 测试: ContainsCompleteRequest - GET请求（只有头部）
' ===========================================================================
Private Sub Test_ContainsCompleteRequest_GetRequestOnlyHeaders()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET / HTTP/1.1" & vbCrLf & "Host: localhost" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Boolean
    result = server.ContainsCompleteRequest(callback)
    
    RecordResult "ContainsCompleteRequest_Get", result, "GET请求（仅头部）应识别为完整请求", Timer - startTime
End Sub

' ===========================================================================
' 测试: ContainsCompleteRequest - POST请求（有请求体）
' ===========================================================================
Private Sub Test_ContainsCompleteRequest_PostRequestWithBody()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("POST /api HTTP/1.1" & vbCrLf & _
"Content-Length: 5" & vbCrLf & vbCrLf & _
    "12345", vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Boolean
    result = server.ContainsCompleteRequest(callback)
    
    RecordResult "ContainsCompleteRequest_PostWithBody", result, "POST请求（含正确长度body）应识别为完整请求", Timer - startTime
End Sub

' ===========================================================================
' 测试: ContainsCompleteRequest - 不完整的请求头
' ===========================================================================
Private Sub Test_ContainsCompleteRequest_IncompleteHeaders()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET / HTTP/1.1" & vbCrLf & "Host: local", vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Boolean
    result = server.ContainsCompleteRequest(callback)
    
    RecordResult "ContainsCompleteRequest_IncompleteHeaders", Not result, _
                 "不完整的请求头应识别为不完整", Timer - startTime
End Sub

' ===========================================================================
' 测试: ContainsCompleteRequest - 不完整的请求体
' ===========================================================================
Private Sub Test_ContainsCompleteRequest_IncompleteBody()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("POST /api HTTP/1.1" & vbCrLf & _
                   "Content-Length: 10" & vbCrLf & vbCrLf & _
                   "12345", vbFromUnicode)  ' 只发送了5字节，还需要5字节
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Boolean
    result = server.ContainsCompleteRequest(callback)
    
    RecordResult "ContainsCompleteRequest_IncompleteBody", Not result, _
                 "不完整的请求体应识别为不完整", Timer - startTime
End Sub

' ===========================================================================
' 测试: ContainsCompleteRequest - 大请求体分包
' ===========================================================================
Private Sub Test_ContainsCompleteRequest_LargeBodyFragmented()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim bodyPart1 As String
    bodyPart1 = String(3000, "A")
    
    Dim data() As Byte
    data = StrConv("POST /upload HTTP/1.1" & vbCrLf & _
"Content-Length: 10000" & vbCrLf & vbCrLf & _
    bodyPart1, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Boolean
    result = server.ContainsCompleteRequest(callback)
    
    RecordResult "ContainsCompleteRequest_LargeBodyFragmented", Not result, "大请求体分包（部分到达）应识别为不完整", Timer - startTime
End Sub

' ===========================================================================
' 测试: ContainsCompleteRequest - 缓冲区中有多个请求
' ===========================================================================
Private Sub Test_ContainsCompleteRequest_MultipleRequestsInBuffer()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET /req1 HTTP/1.1" & vbCrLf & vbCrLf & _
                   "GET /req2 HTTP/1.1" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Boolean
    result = server.ContainsCompleteRequest(callback)
    
    RecordResult "ContainsCompleteRequest_MultipleRequests", result, _
                 "缓冲区中有多个请求时，应识别第一个为完整", Timer - startTime
End Sub

' ===========================================================================
' 测试: ExtractOneRequest - 提取GET请求
' ===========================================================================
Private Sub Test_ExtractOneRequest_GetRequest()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET /test HTTP/1.1" & vbCrLf & "Host: localhost" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Variant
    result = server.ExtractOneRequest(callback)
    
    Dim headerBytes() As Byte
    headerBytes = result(0)
    Dim bodyBytes() As Byte
    bodyBytes = result(1)
    Dim totalLen As Long
    totalLen = result(2)
    
    Dim headerStr As String
    headerStr = StrConv(headerBytes, vbUnicode)
    
    Dim passed As Boolean
    passed = (InStr(headerStr, "GET /test") > 0 And totalLen = 38 And UBound(bodyBytes) = -1)
    
    RecordResult "ExtractOneRequest_Get", passed, _
                 "GET请求提取应返回正确头部和长度", Timer - startTime
End Sub

' ===========================================================================
' 测试: ExtractOneRequest - 提取POST请求（含body）
' ===========================================================================
Private Sub Test_ExtractOneRequest_PostRequestWithBody()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("POST /api HTTP/1.1" & vbCrLf & _
"Content-Length: 5" & vbCrLf & vbCrLf & _
    "12345", vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Variant
    result = server.ExtractOneRequest(callback)
    
    Dim headerBytes() As Byte
    headerBytes = result(0)
    Dim bodyBytes() As Byte
    bodyBytes = result(1)
    Dim totalLen As Long
    totalLen = result(2)
    
    Dim bodyStr As String
    bodyStr = StrConv(bodyBytes, vbUnicode)
    
    Dim passed As Boolean
    passed = (bodyStr = "12345" And totalLen = 40)
    
    RecordResult "ExtractOneRequest_PostWithBody", passed, "POST请求提取应返回正确头部、Body和长度", Timer - startTime
End Sub

' ===========================================================================
' 测试: ExtractOneRequest - 从多个请求中提取第一个
' ===========================================================================
Private Sub Test_ExtractOneRequest_MultipleRequests()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET /first HTTP/1.1" & vbCrLf & vbCrLf & _
    "GET /second HTTP/1.1" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim result As Variant
    result = server.ExtractOneRequest(callback)
    
    Dim headerBytes() As Byte
    headerBytes = result(0)
    Dim totalLen As Long
    totalLen = result(2)
    
    Dim headerStr As String
    headerStr = StrConv(headerBytes, vbUnicode)
    
    Dim passed As Boolean
    passed = (InStr(headerStr, "GET /first") > 0 And totalLen = 26)
    
    RecordResult "ExtractOneRequest_Multiple", passed, "从多个请求中提取应只返回第一个请求", Timer - startTime
End Sub

' ===========================================================================
' 测试: RemoveProcessedData - 清理单个请求后的缓冲区
' ===========================================================================
Private Sub Test_RemoveProcessedData_SingleRequest()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET /first HTTP/1.1" & vbCrLf & vbCrLf & _
                   "GET /second HTTP/1.1" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    server.RemoveProcessedData callback, 26  ' 移除第一个请求
    
    Dim remainingStr As String
    remainingStr = Left(StrConv(callback.RecvBuffer, vbUnicode), callback.recvBufferLen)
    
    Dim passed As Boolean
    passed = (InStr(remainingStr, "GET /second") > 0 And callback.recvBufferLen = 27)
    
    RecordResult "RemoveProcessedData_Single", passed, _
                 "清理单个请求后应保留剩余数据", Timer - startTime
End Sub

' ===========================================================================
' 测试: RemoveProcessedData - 清理后缓冲区中仍有多个请求
' ===========================================================================
Private Sub Test_RemoveProcessedData_MultipleRequests()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET /req1 HTTP/1.1" & vbCrLf & vbCrLf & _
                   "GET /req2 HTTP/1.1" & vbCrLf & vbCrLf & _
                   "GET /req3 HTTP/1.1" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    server.RemoveProcessedData callback, 25  ' 移除第一个请求
    
    Dim remainingStr As String
    remainingStr = Left(StrConv(callback.RecvBuffer, vbUnicode), callback.recvBufferLen)
    
    Dim passed As Boolean
    passed = (InStr(remainingStr, "GET /req2") > 0 And InStr(remainingStr, "GET /req3") > 0)
    
    RecordResult "RemoveProcessedData_Multiple", passed, _
                 "清理后应保留后续多个请求", Timer - startTime
End Sub

' ===========================================================================
' 测试: RemoveProcessedData - 清理所有数据后缓冲区为空
' ===========================================================================
Private Sub Test_RemoveProcessedData_EmptyBufferAfterRemove()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim data() As Byte
    data = StrConv("GET / HTTP/1.1" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    server.RemoveProcessedData callback, callback.recvBufferLen
    
    Dim passed As Boolean
    passed = (callback.recvBufferLen = 0)
    
    RecordResult "RemoveProcessedData_Empty", passed, _
                 "清理所有数据后缓冲区长度应为0", Timer - startTime
End Sub

' ===========================================================================
' 测试: 完整工作流 - 请求头分包场景
' ===========================================================================
Private Sub Test_FullWorkflow_FragmentedHeaders()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    ' 第一次接收：不完整的头部
    Dim part1() As Byte
    part1 = StrConv("GET / HTTP/1.1" & vbCrLf & "Hos", vbFromUnicode)
    callback.RecvBuffer = part1
    callback.recvBufferLen = UBound(part1) + 1
    
    Dim complete1 As Boolean
    complete1 = server.ContainsCompleteRequest(callback)
    
    ' 第二次接收：补全头部
    Dim part2() As Byte
    part2 = StrConv("t: localhost" & vbCrLf & vbCrLf, vbFromUnicode)
    
    ReDim Preserve callback.RecvBuffer(0 To callback.recvBufferLen + UBound(part2))
    CopyMemory callback.RecvBuffer(callback.recvBufferLen), part2(0), UBound(part2) + 1
    callback.recvBufferLen = callback.recvBufferLen + UBound(part2) + 1
    
    Dim complete2 As Boolean
    complete2 = server.ContainsCompleteRequest(callback)
    
    Dim passed As Boolean
    passed = (Not complete1 And complete2)
    
    RecordResult "Workflow_FragmentedHeaders", passed, _
                 "请求头分包场景：第一次不完整，第二次完整", Timer - startTime
End Sub

' ===========================================================================
' 测试: 完整工作流 - 请求粘包场景
' ===========================================================================
Private Sub Test_FullWorkflow_StickyPackages()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    ' 一次发送两个请求
    Dim data() As Byte
    data = StrConv("GET /req1 HTTP/1.1" & vbCrLf & "Host: test" & vbCrLf & vbCrLf & _
    "GET /req2 HTTP/1.1" & vbCrLf & "Host: test" & vbCrLf & vbCrLf, vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    Dim processedCount As Long
    processedCount = 0
    
    ' 模拟处理循环
    Do While server.ContainsCompleteRequest(callback)
        Dim req As Variant
        req = server.ExtractOneRequest(callback)
        server.RemoveProcessedData callback, req(2)
        processedCount = processedCount + 1
    Loop
    
    RecordResult "Workflow_StickyPackages", (processedCount = 2), "粘包场景应正确处理两个请求", Timer - startTime
End Sub

' ===========================================================================
' 测试: 完整工作流 - 大请求体分包场景
' ===========================================================================
Private Sub Test_FullWorkflow_LargeBodyFragmented()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    ' 第一次接收：头部+部分body
    Dim part1() As Byte
    part1 = StrConv("POST /upload HTTP/1.1" & vbCrLf & _
"Content-Length: 10" & vbCrLf & vbCrLf & _
    "12345", vbFromUnicode)
    callback.RecvBuffer = part1
    callback.recvBufferLen = UBound(part1) + 1
    
    Dim complete1 As Boolean
    complete1 = server.ContainsCompleteRequest(callback)
    
    ' 第二次接收：剩余body
    Dim part2() As Byte
    part2 = StrConv("67890", vbFromUnicode)
    
    ReDim Preserve callback.RecvBuffer(0 To callback.recvBufferLen + UBound(part2))
    CopyMemory callback.RecvBuffer(callback.recvBufferLen), part2(0), UBound(part2) + 1
    callback.recvBufferLen = callback.recvBufferLen + UBound(part2) + 1
    
    Dim complete2 As Boolean
    complete2 = server.ContainsCompleteRequest(callback)
    
    Dim req As Variant
    req = server.ExtractOneRequest(callback)
    Dim bodyBytes() As Byte
    bodyBytes = req(1)
    Dim bodyStr As String
    bodyStr = StrConv(bodyBytes, vbUnicode)
    
    Dim passed As Boolean
    passed = (Not complete1 And complete2 And bodyStr = "1234567890")
    
    RecordResult "Workflow_LargeBodyFragmented", passed, "大请求体分包场景：两次接收后应完整提取", Timer - startTime
End Sub

' ===========================================================================
' 测试: 完整工作流 - 真实场景（混合情况）
' ===========================================================================
Private Sub Test_FullWorkflow_RealWorldScenario()
    Dim startTime As Double
    startTime = Timer
    
    Dim callback As cClientCallback
    Set callback = New cClientCallback
    
    Dim server As cHttpServer
    Set server = New cHttpServer
    
    ' 模拟真实网络：多个请求粘包 + 分包
    Dim data() As Byte
    data = StrConv("GET /api/user HTTP/1.1" & vbCrLf & "Host: api.example.com" & vbCrLf & vbCrLf & _
                   "POST /api/data HTTP/1.1" & vbCrLf & "Host: api.example.com" & vbCrLf & _
                   "Content-Length: 15" & vbCrLf & vbCrLf & _
                   "Hello, World!!!", vbFromUnicode)
    callback.RecvBuffer = data
    callback.recvBufferLen = UBound(data) + 1
    
    Dim processedCount As Long
    processedCount = 0
    Dim firstRequestMethod As String
    Dim secondRequestBody As String
    
    ' 处理第一个请求（GET）
    If server.ContainsCompleteRequest(callback) Then
        Dim req1 As Variant
        req1 = server.ExtractOneRequest(callback)
        Dim header1() As Byte
        header1 = req1(0)
        firstRequestMethod = StrConv(header1, vbUnicode)
        server.RemoveProcessedData callback, req1(2)
        processedCount = processedCount + 1
    End If
    
    ' 处理第二个请求（POST）
    If server.ContainsCompleteRequest(callback) Then
        Dim req2 As Variant
        req2 = server.ExtractOneRequest(callback)
        Dim body2() As Byte
        body2 = req2(1)
        secondRequestBody = StrConv(body2, vbUnicode)
        server.RemoveProcessedData callback, req2(2)
        processedCount = processedCount + 1
    End If
    
    Dim passed As Boolean
    passed = (processedCount = 2 And _
              InStr(firstRequestMethod, "GET /api/user") > 0 And _
              secondRequestBody = "Hello, World!!!")
    
    RecordResult "Workflow_RealWorldScenario", passed, _
                 "真实场景：应正确处理GET和POST粘包请求", Timer - startTime
End Sub

' ===========================================================================
' 输出测试报告
' ===========================================================================
Private Sub PrintTestReport()
    Dim i As Long
    Dim result As TestResult
    Dim totalDuration As Double
    
    Debug.Print String(80, "=")
    Debug.Print "测试报告"
    Debug.Print String(80, "=")
    
    ' 统计信息
    For i = 1 To testResults.Count
        totalDuration = totalDuration + testResults(i).duration
    Next i
    
    Debug.Print "总测试数: " & totalTests
    Debug.Print "通过数: " & passedTests & " (" & Format(passedTests / totalTests * 100, "0.00") & "%)"
    Debug.Print "失败数: " & failedTests
    Debug.Print "总耗时: " & Format(totalDuration * 1000, "0.00") & "ms"
    Debug.Print String(80, "-")
    
    ' 详细结果
    Debug.Print "详细结果:"
    For i = 1 To testResults.Count
        Set result = testResults(i)
        Debug.Print IIf(result.passed, "[PASS]", "[FAIL]") & " " & _
                    result.testName & " - " & result.message
    Next i
    
    Debug.Print String(80, "=")
    
    If failedTests = 0 Then
        Debug.Print "? 所有测试通过！粘包处理逻辑正常工作。"""
    Else
        Debug.Print "? 测试失败 " & failedTests & " 项，请检查实现代码。"""
    End If
    
    Debug.Print String(80, "=")
End Sub
