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

Dim JsonInst As cJson

Const ServerName As String = "VBMAN"

Public JsonPackFieldNameCode As String
Public JsonPackFieldNameMsg As String
Public JsonPackFieldNameData As String
Public JsonPackFieldNameCount As String

Public Header As Dictionary

'兼容asp对象
Public Buffer As Boolean
Public CacheControl  As String
Public CharSet As String
Public CodePage  As Long
Public ContentType As String
Public Cookies As Dictionary
Public Expires  As Variant
Public ExpiresAbsolute  As Variant
Public LCID As Long
Public Status  As String

' 统计用：记录本次响应的状态码（在 SendHeader 中赋值）
Public StatusCode As Long

' 标记是否已经发送过响应，防止重复发送
Private m_HasSent As Boolean

Friend Property Set fClient(o As Object)
    If o Is Nothing Then
        Set Client = Nothing
    Else
        Set Client = o
    End If
End Property

Friend Property Get fHasSent() As Boolean
    fHasSent = m_HasSent
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

' 302 重定向
Public Sub Redirect(Url As String)
    Header("Location") = Url
    SendBodyText "", "text/html", 302, "Found"
    ' SendBodyText 内部已调用 Header.RemoveAll，此处无需重复
End Sub
Public Sub View(TplPath As String)
    '
End Sub

'发送 json
Public Sub Json(Data As Variant, Optional Code As Long = -1, Optional msg As String, Optional Count As Long, Optional Whitespace As Variant)
    '判断是否 Recordset 对象，如果是就转为数组
    '    If TypeOf Data Is ADODB.Recordset Then
    '        Set Data = ToolsList.RsToCollection(Data)
    '    End If
    If Code = -1 Then
        Call SendBodyText(JsonInst.Encode(Data, Whitespace), "application/json")
        Exit Sub
    End If
    Dim Dic As Dictionary
    Set Dic = New Dictionary
    With Dic
        .Add JsonPackFieldNameCode, Code
        .Add JsonPackFieldNameMsg, msg
        .Add JsonPackFieldNameData, Data
        .Add JsonPackFieldNameCount, Count
    End With
    Call SendBodyText(JsonInst.Encode(Dic, Whitespace), "application/json")
End Sub

'发送 文件
Public Sub File(Path As String)
    On Error GoTo EH
    
    ' 防御性检查：确保 Client -> Context -> Server 引用链有效
    If Client Is Nothing Then
        Call State500("File: Client is Nothing")
        Exit Sub
    End If
    If Client.Context Is Nothing Then
        Call State500("File: Context is Nothing")
        Exit Sub
    End If
    If Client.Context.Server Is Nothing Then
        Call State500("File: Server is Nothing")
        Exit Sub
    End If
    
    ' 路径穿越防护：拒绝包含 .. 的路径
    If InStr(1, Path, "..", vbTextCompare) > 0 Then
        Call State403("Forbidden")
        Exit Sub
    End If
    
    Dim sPath As String
    sPath = ToolsFso.AutoCompleteFullPath(Client.Context.Server.WebRoot & Path, True)
    
    ' 如果路径对应目录，则尝试默认文档
    If ToolsFso.Inst.FolderExists(sPath) = True Then
        Dim defaultDocs As Collection
        Set defaultDocs = Client.Context.Server.GetDefaultDocuments()
        Dim doc As Variant
        For Each doc In defaultDocs
            Dim tryPath As String
            tryPath = sPath & "\" & doc
            If ToolsFso.Inst.FileExists(tryPath) = True Then
                sPath = tryPath
                GoTo SendFile
            End If
        Next
        ' 目录存在但无默认文档，返回 403
        State403 "Directory listing not allowed"
        Exit Sub
    End If
    
SendFile:
    If ToolsFso.Inst.FileExists(sPath) = False Then Call State404: Exit Sub
    
    ' === ETag + 304 + 懒缓存 ===
    
    ' 1. 获取文件信息，计算 ETag（仅 FSO 元数据，不读文件内容）
    Dim fileInfo As Object  ' Scripting.File
    Set fileInfo = ToolsFso.Inst.GetFile(sPath)
    Dim currentETag As String
    ' 修复：fileInfo.Size 是 Variant 类型，超过 2GB 时隐式转 Long 会溢出崩溃
    ' 改用 Variant 接收，ComputeETag 也改为接受 Variant 参数
    Dim fileSize As Variant
    fileSize = fileInfo.Size
    currentETag = ToolsDateTime.ComputeETag(fileInfo.DateLastModified, fileSize)
    Dim lastModTime As Date
    lastModTime = fileInfo.DateLastModified
    Set fileInfo = Nothing
    
    ' 2. 设置 ETag + Last-Modified 响应头
    Header("ETag") = currentETag
    Header("Last-Modified") = ToolsDateTime.FormatHttpDate(lastModTime)
    
    ' 3. 304 Not Modified 协商缓存（比对请求的 If-None-Match）
    '    防御性检查：Request 和 Header 可能在重入场景下被释放
    If Not Client.Context Is Nothing Then
        If Not Client.Context.Request Is Nothing Then
            If Not Client.Context.Request.Header Is Nothing Then
                Dim reqHeader As Dictionary
                Set reqHeader = Client.Context.Request.Header
                If reqHeader.Exists("If-None-Match") Then
                    If reqHeader("If-None-Match") = currentETag Then
                        Send304 currentETag, lastModTime
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    ' 4. 查找内容缓存
    Dim cachedEntry As cHttpServerFileCacheEntry
    Set cachedEntry = Client.Context.Server.GetCachedFile(sPath)
    
    If Not cachedEntry Is Nothing Then
        ' 缓存命中：检查 TTL 过期 + ETag 变化
        Dim ttlExpired As Boolean
        ttlExpired = False
        If Client.Context.Server.CacheTTLSeconds > 0 Then
            If DateDiff("s", cachedEntry.CachedAt, Now) > Client.Context.Server.CacheTTLSeconds Then
                ttlExpired = True
            End If
        End If
        
        If cachedEntry.ETag <> currentETag Then
            ' 文件已更新：移除旧缓存，重新读取
            Client.Context.Server.InvalidateFileCache sPath
            Set cachedEntry = Nothing
        ElseIf ttlExpired Then
            ' TTL 过期但文件未变：刷新缓存时间，继续使用
            cachedEntry.CachedAt = Now
        End If
    End If
    
    ' 5. 发送内容
    Dim mimeType As String
    mimeType = ToolsHttpMimeType.MapMimeType(sPath)
    
    If Not cachedEntry Is Nothing Then
        ' 从缓存内存输出（跳过磁盘 I/O）
        ' 先读取到局部变量，避免 Property Get 返回值直接传入 ByRef 参数的潜在问题
        Dim cachedData() As Byte
        cachedData = cachedEntry.Content
        ' 修复：缓存数据数组安全检查，防止未初始化数组传入 SendBodyByte 导致 UBound 崩溃
        If (Not cachedData) = -1 Then
            SendBodyText "", mimeType
            Exit Sub
        End If
        SendBodyByte cachedData, mimeType
    Else
        ' 从磁盘读取
        Dim Data() As Byte
        If ToolsStream.LoadFileAsBinary(sPath, Data) = False Then Call State404: Exit Sub
        
        ' 修复：LoadFileAsBinary 成功但 Data 仍可能为未初始化数组的防护
        ' (Not Data) = -1 是 VB6 检查动态数组是否未初始化的惯用写法
        If (Not Data) = -1 Then
            SendBodyText "", mimeType
            Exit Sub
        End If
        
        ' 懒缓存：首次访问且未超出大小上限时加入缓存
        If Client.Context.Server.MaxCacheFileSize > 0 Then
            If UBound(Data) + 1 <= Client.Context.Server.MaxCacheFileSize Then
                Client.Context.Server.AddFileToCache sPath, Data, currentETag, lastModTime
            End If
        End If
        
        SendBodyByte Data, mimeType
    End If
    
    Exit Sub
EH:
    ' 修复：先标记 m_HasSent 防止 State500 递归调用 SendBodyByte 再次写入已损坏的 Socket
    On Error Resume Next  ' 防止 State500 内 SendData 再出错导致递归崩溃
    m_HasSent = True
    Dim ErrInfo As String: ErrInfo = ERR.Number & "#" & ERR.Description
    ' 直接写原始错误响应，避免走完整 SendHeader→SendBodyByte 流程的递归风险
    If Not Client Is Nothing Then
        If Not Client.Socket Is Nothing Then
            Dim errMsg As String
            errMsg = "HTTP/1.1 500 Internal Error" & vbCrLf & _
                     "Content-Type: text/plain" & vbCrLf & _
                     "Content-Length: " & Len(ErrInfo) & vbCrLf & _
                     "Connection: close" & vbCrLf & vbCrLf & ErrInfo
            Client.Socket.SendData errMsg
        End If
    End If
End Sub

Public Sub State(Code As Long, Text As String, Optional Say As String = "未知状态")
    On Error Resume Next  ' 防止 EH 块调用 State* 时 SendData 再次出错导致递归崩溃
    SendBodyText Say, "text/html", Code, Text
End Sub
Public Sub State400(Optional Say As String)
    State 400, "BadRequest", Say
End Sub
Public Sub State401(Optional Say As String)
    State 401, "Unauthorized", Say
End Sub
Public Sub State403(Optional Say As String)
    State 403, "Forbidden", Say
End Sub
Public Sub State404(Optional Say As String)
    State 404, "NotFound", Say
End Sub
' BUG-3 修复：添加 413 请求体过大状态码
Public Sub State413(Optional Say As String = "Request Entity Too Large")
    State 413, "PayloadTooLarge", Say
End Sub
Public Sub State502(Optional Say As String)
    State 502, "GatewayError", Say
End Sub
Public Sub State500(Optional Say As String)
    On Error Resume Next  ' 防止错误处理器内的二次错误导致崩溃
    State 500, "InternalError", Say
End Sub

' 发送 304 Not Modified（协商缓存命中，不传输文件内容）
Private Sub Send304(ETagValue As String, LastModifiedTime As Date)
    If m_HasSent Then Exit Sub
    Header("ETag") = ETagValue
    Header("Last-Modified") = ToolsDateTime.FormatHttpDate(LastModifiedTime)
    Header("Content-Length") = 0
    Header("Cache-Control") = "no-cache": SendHeader 304, "Not Modified"
    Header.RemoveAll
    m_HasSent = True
End Sub

Private Sub Class_Initialize()
    JsonPackFieldNameCode = "code"
    JsonPackFieldNameData = "data"
    JsonPackFieldNameMsg = "msg"
    JsonPackFieldNameCount = "count"
    '默认 utf8 并生成相应的响应头会自动设置编码
    CharSet = "utf-8"
    CodePage = 65001
    m_HasSent = False
    StatusCode = 0
    ' 显式创建子对象（替代 As New 隐式创建）
    Set Header = New Dictionary
    Set JsonInst = New cJson
End Sub

Private Sub Class_Terminate()
    ' 显式断开循环引用并销毁内部对象
    Set Client = Nothing
    If Not Header Is Nothing Then
        Header.RemoveAll
        Set Header = Nothing
    End If
    Set JsonInst = Nothing
    If Not Cookies Is Nothing Then Set Cookies = Nothing
End Sub

'Private Sub SendBodyTextAsUtf8(Data As String, DataType As String)
'    Call SendBodyByte(ToolsUtf8.Encode(Data), DataType)
'End Sub
Private Sub SendBodyText(Data As String, DataType As String, Optional Status As Long = 200, Optional StatusText As String = "OK")
    ' BUG-8 修复：防止重复发送响应
    If m_HasSent Then Exit Sub
    
    Dim D() As Byte
    If Data = "" Then
        ' 修复：空字符串时使用传入的 DataType 而非硬编码 text/plain
        If ShouldIncludeCharset(DataType) Then
            Header("Content-Type") = DataType & "; charset=" & CharSet
        Else
            Header("Content-Type") = DataType
        End If
        Header("Content-Length") = 0
        SendHeader Status, StatusText
        Header.RemoveAll
        m_HasSent = True
        Exit Sub
    End If
    If CodePage = 65001 Then
        D = ToolsUtf8.Encode(Data)
    Else
        D = StrConv(Data, vbFromUnicode)
    End If
    Call SendBodyByte(D, DataType, Status, StatusText)
End Sub
Private Sub SendBodyByte(Data() As Byte, DataType As String, Optional Status As Long = 200, Optional StatusText As String = "OK")
    ' BUG-8 修复：防止重复发送响应
    If m_HasSent Then Exit Sub
    
    On Error GoTo EH
    
    ' 数组安全检查：确保 Data 已初始化
    ' (Not Data) = -1 是 VB6 检查数组是否未初始化的惯用写法
    Dim dataLen As Long
    If (Not Data) = -1 Then
        ' 数组未初始化，发送空响应
        If ShouldIncludeCharset(DataType) Then
            Header("Content-Type") = DataType & "; charset=" & CharSet
        Else
            Header("Content-Type") = DataType
        End If
        Header("Content-Length") = 0
        SendHeader Status, StatusText
        Header.RemoveAll
        m_HasSent = True
        Exit Sub
    End If
    dataLen = UBound(Data) + 1
    
    ' 智能 Content-Type：仅对文本类型附加 charset
    If ShouldIncludeCharset(DataType) Then
        Header("Content-Type") = DataType & "; charset=" & CharSet
    Else
        Header("Content-Type") = DataType
    End If
    Header("Content-Length") = dataLen
    SendHeader Status, StatusText
    
    ' 修复：只在有数据时才发送 body，避免 0 长度 Data 传给 SendData 触发 Winsock 内部异常
    If dataLen > 0 Then
        If Not Client Is Nothing Then
            If Not Client.Socket Is Nothing Then
                On Error Resume Next
                Client.Socket.SendData Data
                ' 统计：记录Body发送字节数
                If Not Client.Parent Is Nothing Then
                    Client.Parent.Statistics.TotalBytesSent = Client.Parent.Statistics.TotalBytesSent + dataLen
                End If
                On Error GoTo 0
            End If
        End If
    End If
    
    Header.RemoveAll
    m_HasSent = True
    Exit Sub
EH:
    ' SendBodyByte 内部异常保护：标记已发送防止递归
    m_HasSent = True
End Sub

' 判断 MIME 类型是否需要附加 charset
' 文本类型需要 charset（如 text/html, text/css, application/json）
' 二进制类型不需要（如 image/png, video/mp4, application/octet-stream）
Private Function ShouldIncludeCharset(ByVal mimeType As String) As Boolean
    Dim lowerType As String
    lowerType = LCase(mimeType)
    If Left(lowerType, 5) = "text/" Then
        ShouldIncludeCharset = True
    ElseIf Left(lowerType, 12) = "application/" Then
        If InStr(lowerType, "json") > 0 Or _
           InStr(lowerType, "javascript") > 0 Or _
           InStr(lowerType, "xml") > 0 Or _
           InStr(lowerType, "yaml") > 0 Then
            ShouldIncludeCharset = True
        End If
    End If
End Function

Private Sub SendHeader(Optional Status As Long = 200, Optional StatusText As String = "OK", Optional HttpVersioin As String = "1.1")
    ' 记录状态码供统计使用
    StatusCode = Status
    Dim Head As String: Head = "HTTP/" & HttpVersioin & " " & Status & " " & StatusText & vbCrLf
    '
    Dim UserVer As String: UserVer = "FREE"
    Dim Ver As String: Ver = "[" & UserVer & "] "
    Header("Server") = Ver & ServerName & " " & Common.Version()
    If UserVer = "FREE" Then
        Header("PowerBy") = "www.vb6.pro"
    End If
    
    ' BUG 修复：原代码在 If Not Client Is Nothing 之前访问 Client.Parent.CrossDomain
    ' 当 Client 为 Nothing 时会导致崩溃。现将所有 Client 访问统一在防护块内
    If Not Client Is Nothing Then
        ' 跨域处理
        If Not Client.Parent Is Nothing Then
            If Not Client.Parent.CrossDomain Is Nothing Then
                If Client.Parent.CrossDomain.Enable = True Then
                    Client.Parent.CrossDomain.AddTo Header
                End If
            End If
        End If
        
        ' 处理响应 Cookie（从 Context 中的 Cookies 获取）
        If Not Client.Context Is Nothing Then
            If Not Client.Context.Cookies Is Nothing Then
                Dim cookieLines As String
                Dim cookieLine As Variant
                cookieLines = Client.Context.Cookies.Encode
                If cookieLines <> "" Then
                    ' 每个 Cookie 需要单独的 Set-Cookie 头
                    Dim cookieArray() As String
                    cookieArray = Split(cookieLines, vbCrLf)
                    Dim iCookie As Long
                    For iCookie = 0 To UBound(cookieArray)
                        If cookieArray(iCookie) <> "" Then
                            Head = Head & "Set-Cookie: " & cookieArray(iCookie) & vbCrLf
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    Dim x, i As Long
    For Each x In Header.Keys()
        Head = Head & x & ": " & Header.Item(x) & vbCrLf
    Next
    
    ' BUG-8 修复：SendData 添加错误处理
    If Not Client Is Nothing Then
        If Not Client.Socket Is Nothing Then
            On Error Resume Next
            Client.socket.SendData Head & vbCrLf
            ' 统计：记录Header发送字节数
            If Not Client.Parent Is Nothing Then
                Client.Parent.Statistics.TotalBytesSent = Client.Parent.Statistics.TotalBytesSent + Len(Head) + 2
            End If
            On Error GoTo 0
        End If
    End If
End Sub


Public Sub test()
    Dim sRequest As String: sRequest = Client.Context.Request.HeaderRaw
    Dim hSocket As Long: hSocket = Client.hSocket
    Dim vSplit              As Variant
    Dim sBody               As String
    vSplit = Split(sRequest, vbCrLf)
    If UBound(vSplit) >= 0 Then
        sBody = Join(vSplit, vbCrLf & hSocket & ": ") & vbCrLf & _
        vbCrLf & hSocket & ": Current time is " & Now & _
        vbCrLf & hSocket & ": RemoteHostIP is " & Client.Socket.RemoteHostIP & _
        vbCrLf & hSocket & ": RemotePort is " & Client.Socket.RemotePort
        Client.Socket.SendData "HTTP/1.1 200 OK" & vbCrLf & _
"Content-Type: text/html" & vbCrLf & _
"Content-Length: " & Len(sBody) & vbCrLf & vbCrLf & _
        sBody
    End If
End Sub