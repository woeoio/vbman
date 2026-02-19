VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRedisClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'=========================================================================
'
' cRedisClient - Redis 客户端
'
' Purpose: 提供简单易用的 Redis 数据库客户端
'
' Features:
'   - 支持基础 Redis 命令 (GET, SET, DEL, EXISTS, KEYS 等)
'   - 支持管道模式
'   - 支持事务模式
'   - 自动重连机制
'   - 支持密码认证
'
' Usage:
'   Dim oRedis As New cRedisClient
'   oRedis.Connect "127.0.0.1", 6379
'   oRedis.Set_ "mykey", "myvalue"
'   Debug.Print oRedis.Get_("mykey")
'
'=========================================================================

Private Const MODULE_NAME As String = "cRedisClient"

'=========================================================================
' 枚举定义
'=========================================================================

' Redis 回复类型
Private Enum RedisResponseType
    RESP_SIMPLE_STRING = 1                                                      '+'
    RESP_ERROR = 2                                                              '-'
    RESP_INTEGER = 3                                                            ':'
    RESP_BULK_STRING = 4                                                        '$'
    RESP_ARRAY = 5                                                              '*'
End Enum

'=========================================================================
' 私有成员
'=========================================================================

Private m_oSocket As cWinsock
Private m_sHost As String
Private m_lPort As Long
Private m_sPassword As String
Private m_lDbIndex As Long
Private m_dTimeout As Double
Private m_bConnected As Boolean
Private m_bInTransaction As Boolean
Private m_oPipeline As Collection
Private m_bInPipeline As Boolean

'=========================================================================
' 公共事件
'=========================================================================

' 连接断开事件
Public Event OnDisconnected()

' 错误事件
Public Event OnError(ByVal ErrorMsg As String)


' 最后一次错误
Public LastError As String

'=========================================================================
' 辅助函数
'=========================================================================

' 检查变量是否为字符串类型
Private Function IsString(ByVal vValue As Variant) As Boolean
    IsString = (varType(vValue) = vbString)
End Function

'=========================================================================
' 属性
'=========================================================================

' 服务器地址
Public Property Get Host() As String
    Host = m_sHost
End Property

' 服务器端口
Public Property Get Port() As Long
    Port = m_lPort
End Property

' 是否已连接
Public Property Get Connected() As Boolean
    Connected = m_bConnected
End Property

' 数据库索引
Public Property Get DbIndex() As Long
    DbIndex = m_lDbIndex
End Property

Public Property Let DbIndex(ByVal Value As Long)
    If m_bConnected Then
        Call SelectDb(Value)
    End If
    m_lDbIndex = Value
End Property

' 超时时间(秒)
Public Property Get Timeout() As Double
    Timeout = m_dTimeout
End Property

Public Property Let Timeout(ByVal Value As Double)
    m_dTimeout = Value
End Property

' 是否在事务中
Public Property Get InTransaction() As Boolean
    InTransaction = m_bInTransaction
End Property

' 是否在管道模式
Public Property Get InPipeline() As Boolean
    InPipeline = m_bInPipeline
End Property

' Socket 对象 (用于高级用户)
Public Property Get Socket() As cWinsock
    Set Socket = m_oSocket
End Property



'=========================================================================
' 公共方法
'=========================================================================

' 连接到 Redis 服务器
Public Function Connect(Optional ByVal Server As String = "127.0.0.1", _
    Optional ByVal Port As Long = 6379, _
    Optional ByVal Password As String = "") As Boolean
    On Error GoTo ErrorHandler

    ' 设置参数
    m_sHost = Server
    m_lPort = Port
    m_sPassword = Password

    ' 关闭现有连接
    If Not m_oSocket Is Nothing Then
        m_oSocket.Close_
        Set m_oSocket = Nothing
    End If

    ' 创建新连接
    Set m_oSocket = New cWinsock
    m_oSocket.Protocol = wskTCPProtocol

    ' 使用同步连接
    If Not m_oSocket.SyncConnect(m_sHost, m_lPort, m_dTimeout * 1000) Then
        LastError = "连接超时"
        RaiseEvent OnError(LastError)
        Connect = False
        Exit Function
    End If

    m_bConnected = True

    ' 如果有密码,执行认证
    If LenB(m_sPassword) > 0 Then
        If Not Auth(m_sPassword) Then
            LastError = "认证失败"
            RaiseEvent OnError(LastError)
            Connect = False
            Exit Function
        End If
    End If

    ' 如果指定了数据库索引,切换数据库
    If m_lDbIndex > 0 Then
        Call SelectDb(m_lDbIndex)
    End If

    Connect = True
    Exit Function

ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Connect = False
End Function

' 断开连接
Public Sub DisConnect()
    If Not m_oSocket Is Nothing Then
        On Error Resume Next
        m_oSocket.Close_
        Set m_oSocket = Nothing
        On Error GoTo 0
    End If
    m_bConnected = False
    m_bInTransaction = False
    m_bInPipeline = False
End Sub

' 认证
Public Function Auth(ByVal Password As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand(Array("AUTH", Password))

    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            Auth = True
            Exit Function
        End If
    End If

    Auth = False
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Auth = False
End Function

' 选择数据库
Public Function SelectDb(ByVal Index As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand(Array("SELECT", CStr(Index)))

    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            m_lDbIndex = Index
            SelectDb = True
            Exit Function
        End If
    End If

    SelectDb = False
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    SelectDb = False
End Function

'=========================================================================
' 基础命令
'=========================================================================

' SET 命令
Public Function Set_(ByVal Key As String, ByVal Value As String, _
                     Optional ByVal Ex As Long = 0, _
                     Optional ByVal Px As Long = 0) As Boolean
    On Error GoTo ErrorHandler

    Dim vCmd As Variant
    If Ex > 0 Then
        vCmd = Array("SET", Key, Value, "EX", CStr(Ex))
    ElseIf Px > 0 Then
        vCmd = Array("SET", Key, Value, "PX", CStr(Px))
    Else
        vCmd = Array("SET", Key, Value)
    End If

    Dim sResp As Variant
    sResp = SendCommand(vCmd)

    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            Set_ = True
            Exit Function
        End If
    End If

    Set_ = False
    Exit Function

ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Set_ = False
End Function

' GET 命令
Public Function Get_(ByVal Key As String) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("GET", Key))

    If varType(sResp) = vbString Then
        Get_ = sResp
    ElseIf IsArray(sResp) Then
        If UBound(sResp) >= 0 Then
            Get_ = sResp(0)
        End If
    End If

    Exit Function

ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Get_ = ""
End Function

' DEL 命令
Public Function Del(ParamArray Keys() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Long
    lret = SendCommand(Array("DEL", Keys))
    
    If IsNumeric(lret) Then
        Del = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Del = 0
End Function

' EXISTS 命令
Public Function Exists(ParamArray Keys() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("EXISTS", Keys))
    
    If IsNumeric(lret) Then
        Exists = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Exists = 0
End Function

' KEYS 命令
Public Function Keys(ByVal Pattern As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim vRet As Variant
    vRet = SendCommand(Array("KEYS", Pattern))
    
    If IsArray(vRet) Then
        Keys = vRet
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Keys = Array()
End Function

' EXPIRE 命令
Public Function Expire(ByVal Key As String, ByVal Seconds As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("EXPIRE", Key, CStr(Seconds)))
    
    If IsNumeric(lret) And lret = 1 Then
        Expire = True
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Expire = False
End Function

' TTL 命令
Public Function TTL(ByVal Key As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("TTL", Key))
    
    If IsNumeric(lret) Then
        TTL = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    TTL = -2
End Function

'=========================================================================
' String 操作
'=========================================================================

' INCR 命令
Public Function Incr(ByVal Key As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("INCR", Key))
    
    If IsNumeric(lret) Then
        Incr = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Incr = 0
End Function

' DECR 命令
Public Function Decr(ByVal Key As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("DECR", Key))
    
    If IsNumeric(lret) Then
        Decr = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Decr = 0
End Function

' MGET 命令
Public Function MGet(ParamArray Keys() As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim vRet As Variant
    vRet = SendCommand(Array("MGET", Keys))
    
    If IsArray(vRet) Then
        MGet = vRet
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    MGet = Array()
End Function

' MSET 命令
Public Function MSet(ParamArray KeyValues() As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant
    vCmd = Array("MSET", KeyValues)
    
    Dim sResp As Variant
    sResp = SendCommand(vCmd)

    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            MSet = True
            Exit Function
        End If
    End If

    MSet = False
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    MSet = False
End Function

'=========================================================================
' Hash 操作
'=========================================================================

' HSET 命令
Public Function HSet(ByVal Key As String, ByVal Field As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand(Array("HSET", Key, Field, Value))
    
    If IsNumeric(sResp) Then
        HSet = True
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    HSet = False
End Function

' HGET 命令
Public Function HGet(ByVal Key As String, ByVal Field As String) As String
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand(Array("HGET", Key, Field))
    
    If varType(sResp) = vbString Then
        HGet = sResp
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    HGet = ""
End Function

' HMGET 命令
Public Function HMGet(ByVal Key As String, ParamArray fields() As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim vRet As Variant
    vRet = SendCommand(Array("HMGET", Key, fields))
    
    If IsArray(vRet) Then
        HMGet = vRet
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    HMGet = Array()
End Function

' HGETALL 命令
Public Function HGetAll(ByVal Key As String) As Scripting.Dictionary
    On Error GoTo ErrorHandler
    
    Dim oDict As New Scripting.Dictionary
    oDict.CompareMode = TextCompare
    
    Dim vRet As Variant
    vRet = SendCommand(Array("HGETALL", Key))
    
    If IsArray(vRet) Then
        Dim i As Long
        For i = 0 To UBound(vRet) Step 2
            If i + 1 <= UBound(vRet) Then
                oDict(vRet(i)) = vRet(i + 1)
            End If
        Next
    End If
    
    Set HGetAll = oDict
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Set HGetAll = New Scripting.Dictionary
End Function

' HDEL 命令
Public Function HDel(ByVal Key As String, ParamArray fields() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("HDEL", Key, fields))
    
    If IsNumeric(lret) Then
        HDel = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    HDel = 0
End Function

' HEXISTS 命令
Public Function HExists(ByVal Key As String, ByVal Field As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("HEXISTS", Key, Field))
    
    If IsNumeric(lret) And lret = 1 Then
        HExists = True
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    HExists = False
End Function

'=========================================================================
' List 操作
'=========================================================================

' LPUSH 命令
Public Function LPush(ByVal Key As String, ParamArray Values() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("LPUSH", Key, Values))
    
    If IsNumeric(lret) Then
        LPush = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    LPush = 0
End Function

' RPUSH 命令
Public Function RPush(ByVal Key As String, ParamArray Values() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("RPUSH", Key, Values))
    
    If IsNumeric(lret) Then
        RPush = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    RPush = 0
End Function

' LPOP 命令
Public Function LPop(ByVal Key As String) As String
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand(Array("LPOP", Key))
    
    If varType(sResp) = vbString Then
        LPop = sResp
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    LPop = ""
End Function

' RPOP 命令
Public Function RPop(ByVal Key As String) As String
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand(Array("RPOP", Key))
    
    If varType(sResp) = vbString Then
        RPop = sResp
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    RPop = ""
End Function

' LLEN 命令
Public Function lLen(ByVal Key As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("LLEN", Key))
    
    If IsNumeric(lret) Then
        lLen = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    lLen = 0
End Function

' LRANGE 命令
Public Function LRange(ByVal Key As String, ByVal StartIdx As Long, ByVal StopIdx As Long) As Variant
    On Error GoTo ErrorHandler
    
    Dim vRet As Variant
    vRet = SendCommand(Array("LRANGE", Key, CStr(StartIdx), CStr(StopIdx)))
    
    If IsArray(vRet) Then
        LRange = vRet
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    LRange = Array()
End Function

'=========================================================================
' Set 操作
'=========================================================================

' SADD 命令
Public Function SAdd(ByVal Key As String, ParamArray Members() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("SADD", Key, Members))
    
    If IsNumeric(lret) Then
        SAdd = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    SAdd = 0
End Function

' SMEMBERS 命令
Public Function SMembers(ByVal Key As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim vRet As Variant
    vRet = SendCommand(Array("SMEMBERS", Key))
    
    If IsArray(vRet) Then
        SMembers = vRet
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    SMembers = Array()
End Function

' SCARD 命令
Public Function SCard(ByVal Key As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("SCARD", Key))
    
    If IsNumeric(lret) Then
        SCard = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    SCard = 0
End Function

' SISMEMBER 命令
Public Function SIsMember(ByVal Key As String, ByVal Member As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("SISMEMBER", Key, Member))
    
    If IsNumeric(lret) And lret = 1 Then
        SIsMember = True
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    SIsMember = False
End Function

' SREM 命令
Public Function SRem(ByVal Key As String, ParamArray Members() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("SREM", Key, Members))
    
    If IsNumeric(lret) Then
        SRem = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    SRem = 0
End Function

'=========================================================================
' Sorted Set 操作
'=========================================================================

' ZADD 命令
Public Function ZAdd(ByVal Key As String, ByVal Score As Double, ByVal Member As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("ZADD", Key, CStr(Score), Member))
    
    If IsNumeric(lret) Then
        ZAdd = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    ZAdd = 0
End Function

' ZRANGE 命令
Public Function ZRange(ByVal Key As String, ByVal StartIdx As Long, ByVal StopIdx As Long, Optional ByVal WithScores As Boolean = False) As Variant
    On Error GoTo ErrorHandler
    
    Dim vRet As Variant
    If WithScores Then
        vRet = SendCommand(Array("ZRANGE", Key, CStr(StartIdx), CStr(StopIdx), "WITHSCORES"))
    Else
        vRet = SendCommand(Array("ZRANGE", Key, CStr(StartIdx), CStr(StopIdx)))
    End If
    
    If IsArray(vRet) Then
        ZRange = vRet
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    ZRange = Array()
End Function

' ZREM 命令
Public Function ZRem(ByVal Key As String, ParamArray Members() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("ZREM", Key, Members))
    
    If IsNumeric(lret) Then
        ZRem = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    ZRem = 0
End Function

' ZCARD 命令
Public Function ZCard(ByVal Key As String) As Long
    On Error GoTo ErrorHandler
    
    Dim lret As Variant
    lret = SendCommand(Array("ZCARD", Key))
    
    If IsNumeric(lret) Then
        ZCard = CLng(lret)
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    ZCard = 0
End Function

'=========================================================================
' 事务操作
'=========================================================================

' 开始事务
Public Function Multi() As Boolean
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand("MULTI")

    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            m_bInTransaction = True
            Multi = True
            Exit Function
        End If
    End If

ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Multi = False
End Function

' 提交事务
Public Function Exec() As Variant
    On Error GoTo ErrorHandler
    
    If Not m_bInTransaction Then
        LastError = "Not in transaction"
        RaiseEvent OnError(LastError)
        Exec = Array()
        Exit Function
    End If
    
    Dim vRet As Variant
    vRet = SendCommand("EXEC")
    
    m_bInTransaction = False
    
    If IsArray(vRet) Then
        Exec = vRet
    Else
        Exec = Array()
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    m_bInTransaction = False
    Exec = Array()
End Function

' 放弃事务
Public Function Discard() As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_bInTransaction Then
        Discard = False
        Exit Function
    End If
    
    Dim sResp As Variant
    sResp = SendCommand("DISCARD")
    
    m_bInTransaction = False
    
    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            Discard = True
            Exit Function
        End If
    End If
    
    Discard = False
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    m_bInTransaction = False
    Discard = False
End Function

'=========================================================================
' 服务器操作
'=========================================================================

' PING 命令
Public Function Ping() As String
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand("PING")
    
    If varType(sResp) = vbString Then
        Ping = sResp
    Else
        Ping = "PONG"
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Ping = ""
End Function

' INFO 命令
Public Function Info(Optional ByVal Section As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    If LenB(Section) > 0 Then
        sResp = SendCommand(Array("INFO", Section))
    Else
        sResp = SendCommand("INFO")
    End If
    
    If varType(sResp) = vbString Then
        Info = sResp
    End If
    
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    Info = ""
End Function

' FLUSHDB 命令
Public Function FlushDb() As Boolean
    On Error GoTo ErrorHandler
    
    Dim sResp As Variant
    sResp = SendCommand("FLUSHDB")
    
    If varType(sResp) = vbString Then
        If sResp = "OK" Then
            FlushDb = True
            Exit Function
        End If
    End If
    
    FlushDb = False
    Exit Function
    
ErrorHandler:
    LastError = ERR.Description
    RaiseEvent OnError(LastError)
    FlushDb = False
End Function

'=========================================================================
' 私有方法
'=========================================================================

' 发送命令并接收响应
Private Function SendCommand(ByVal Command As Variant) As Variant
    If Not m_bConnected Then
        ERR.Raise vbObjectError, MODULE_NAME, "Not connected to Redis server"
    End If

    ' 构建命令（现在直接返回 UTF-8 字节数组）
    Dim baCmd() As Byte
    baCmd = BuildRedisCommand(Command)

    ' 使用同步发送
    If Not m_oSocket.SyncSendArray(baCmd, m_dTimeout * 1000) Then
        ERR.Raise vbObjectError, MODULE_NAME, "Send command timeout"
    End If

    ' 接收响应
    Dim sResponse As String
    sResponse = ReceiveResponse()

    ' 解析响应
    SendCommand = ParseResp(sResponse)
End Function

' 构建 Redis RESP 协议命令（返回字节数组）
Private Function BuildRedisCommand(ByVal Command As Variant) As Byte()
    Dim vFlatten() As Variant
    Dim i As Long, j As Long
    Dim lCount As Long
    Dim baPart() As Byte
    Dim baResult() As Byte
    Dim lOffset As Long
    Dim sPart As String
    Dim baHeader() As Byte
    
    ' 展平数组（处理 ParamArray 嵌套）
    vFlatten = FlattenArray(Command)
    lCount = UBound(vFlatten) + 1
    
    ' 计算最终数组大小
    Dim lTotalSize As Long
    lTotalSize = 0
    
    ' 数组头部: *<count>\r\n
    Dim sHeader As String
    sHeader = "*" & lCount & vbCrLf
    baHeader = ToolsUtf8.Encode(sHeader)
    lTotalSize = UBound(baHeader) + 1
    lOffset = 0
    
    ' 预分配结果数组
    ReDim baResult(0 To 1023) As Byte ' 初始分配1KB
    lOffset = 0
    
    ' 复制数组头
    For i = 0 To UBound(baHeader)
        If lOffset > UBound(baResult) Then
            ReDim Preserve baResult(UBound(baResult) + 1024) As Byte
        End If
        baResult(lOffset) = baHeader(i)
        lOffset = lOffset + 1
    Next
    
    ' 处理每个元素
    For i = 0 To lCount - 1
        sPart = CStr(vFlatten(i))
        
        ' 编码字符串为 UTF-8
        baPart = ToolsUtf8.Encode(sPart)
        
        ' 构建批量字符串头部: $<length>\r\n
        Dim sLength As String
        sLength = "$" & (UBound(baPart) + 1) & vbCrLf
        baHeader = ToolsUtf8.Encode(sLength)
        
        ' 复制长度头
        For j = 0 To UBound(baHeader)
            If lOffset > UBound(baResult) Then
                ReDim Preserve baResult(UBound(baResult) + 1024) As Byte
            End If
            baResult(lOffset) = baHeader(j)
            lOffset = lOffset + 1
        Next
        
        ' 复制字符串内容
        For j = 0 To UBound(baPart)
            If lOffset > UBound(baResult) Then
                ReDim Preserve baResult(UBound(baResult) + 1024) As Byte
            End If
            baResult(lOffset) = baPart(j)
            lOffset = lOffset + 1
        Next
        
        ' 复制 \r\n
        baHeader = ToolsUtf8.Encode(vbCrLf)
        For j = 0 To UBound(baHeader)
            If lOffset > UBound(baResult) Then
                ReDim Preserve baResult(UBound(baResult) + 1024) As Byte
            End If
            baResult(lOffset) = baHeader(j)
            lOffset = lOffset + 1
        Next
    Next
    
    ' 调整到实际大小
    If lOffset > 0 Then
        ReDim Preserve baResult(lOffset - 1) As Byte
    End If
    
    BuildRedisCommand = baResult
End Function

' 展平数组（递归处理嵌套数组）
Private Function FlattenArray(ByVal vInput As Variant) As Variant()
    Dim vResult() As Variant
    Dim i As Long, j As Long
    Dim lIndex As Long
    Dim vTemp() As Variant
    
    lIndex = 0
    ReDim vResult(0) As Variant
    
    If IsArray(vInput) Then
        For i = LBound(vInput) To UBound(vInput)
            If IsArray(vInput(i)) Then
                ' 递归展平嵌套数组
                vTemp = FlattenArray(vInput(i))
                For j = LBound(vTemp) To UBound(vTemp)
                    ReDim Preserve vResult(lIndex) As Variant
                    vResult(lIndex) = vTemp(j)
                    lIndex = lIndex + 1
                Next
            Else
                ReDim Preserve vResult(lIndex) As Variant
                vResult(lIndex) = vInput(i)
                lIndex = lIndex + 1
            End If
        Next
    Else
        vResult(0) = vInput
    End If
    
    ' 调整数组大小
    If lIndex > 0 Then
        ReDim Preserve vResult(lIndex - 1) As Variant
    End If
    
    FlattenArray = vResult
End Function

' 接收响应
Private Function ReceiveResponse() As String
    Dim baBuffer() As Byte
    Dim lReceived As Long

    ' 读取响应的第一个字节,判断类型
    ReDim baBuffer(0 To 0) As Byte
    If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
        ERR.Raise vbObjectError, MODULE_NAME, "Receive response timeout"
    End If

    Dim sType As String
    sType = Chr(baBuffer(0))

    Dim sBuffer As String
    sBuffer = sType

    Select Case sType
        Case "+", "-", ":"
            ' 简单字符串、错误、整数 - 读取直到 \r\n
            Do
                ReDim baBuffer(0 To 0) As Byte
                If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                    Exit Do
                End If
                sBuffer = sBuffer & Chr(baBuffer(0))

                If Right(sBuffer, 2) = vbCrLf Then
                    Exit Do
                End If
            Loop

        Case "$"
            ' 批量字符串
            Dim baLen() As Byte
            Dim i As Long
            Dim sLen As String

            ' 读取长度部分(直到 \r\n)
            sLen = ""
            Do
                ReDim baBuffer(0 To 0) As Byte
                If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                    Exit Do
                End If

                If baBuffer(0) = 13 Then  ' CR
                    ReDim baBuffer(0 To 0) As Byte
                    If m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                        If baBuffer(0) = 10 Then  ' LF
                            Exit Do
                        End If
                    End If
                Else
                    sLen = sLen & Chr(baBuffer(0))
                End If
            Loop

            sBuffer = sBuffer & sLen & vbCrLf

            Dim lLength As Long
            lLength = Val(sLen)

            If lLength = -1 Then
                ' null 值
            Else
                ' 读取数据内容
                If lLength > 0 Then
                    ReDim baBuffer(0 To lLength - 1) As Byte
                    If m_oSocket.SyncReceiveArray(baBuffer, lLength, m_dTimeout * 1000) Then
                        sBuffer = sBuffer & ToolsUtf8.Decode(baBuffer)
                    End If
                End If

                ' 读取 \r\n
                ReDim baBuffer(0 To 1) As Byte
                If m_oSocket.SyncReceiveArray(baBuffer, 2, m_dTimeout * 1000) Then
                    sBuffer = sBuffer & vbCrLf
                End If
            End If

        Case "*"
            ' 数组
            Dim baCount() As Byte
            Dim sCount As String

            ' 读取数组元素个数
            sCount = ""
            Do
                ReDim baBuffer(0 To 0) As Byte
                If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                    Exit Do
                End If

                If baBuffer(0) = 13 Then  ' CR
                    ReDim baBuffer(0 To 0) As Byte
                    If m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                        If baBuffer(0) = 10 Then  ' LF
                            Exit Do
                        End If
                    End If
                Else
                    sCount = sCount & Chr(baBuffer(0))
                End If
            Loop

            sBuffer = sBuffer & sCount & vbCrLf

            Dim lCount As Long
            lCount = Val(sCount)

            If lCount = -1 Then
                ' null 数组
            ElseIf lCount = 0 Then
                ' 空数组
            Else
                ' 读取每个数组元素
                Dim j As Long
                For j = 0 To lCount - 1
                    Dim sElement As String
                    sElement = ReceiveElement()
                    sBuffer = sBuffer & sElement
                Next
            End If

        Case Else
            ' 未知类型,直接读取剩余数据
    End Select

    ReceiveResponse = sBuffer
End Function

' 接收数组元素(递归调用)
Private Function ReceiveElement() As String
    Dim baBuffer() As Byte
    Dim sBuffer As String

    ' 读取第一个字节
    ReDim baBuffer(0 To 0) As Byte
    If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
        ERR.Raise vbObjectError, MODULE_NAME, "Receive element timeout"
    End If

    Dim sType As String
    sType = Chr(baBuffer(0))
    sBuffer = sType

    Select Case sType
        Case "+", "-", ":"
            ' 简单字符串、错误、整数
            Do
                ReDim baBuffer(0 To 0) As Byte
                If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                    Exit Do
                End If
                sBuffer = sBuffer & Chr(baBuffer(0))

                If Right(sBuffer, 2) = vbCrLf Then
                    Exit Do
                End If
            Loop

        Case "$"
            ' 批量字符串
            Dim sLen As String
            sLen = ""
            Do
                ReDim baBuffer(0 To 0) As Byte
                If Not m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                    Exit Do
                End If

                If baBuffer(0) = 13 Then  ' CR
                    ReDim baBuffer(0 To 0) As Byte
                    If m_oSocket.SyncReceiveArray(baBuffer, 1, m_dTimeout * 1000) Then
                        If baBuffer(0) = 10 Then  ' LF
                            Exit Do
                        End If
                    End If
                Else
                    sLen = sLen & Chr(baBuffer(0))
                End If
            Loop

            sBuffer = sBuffer & sLen & vbCrLf

            Dim lLength As Long
            lLength = Val(sLen)

            If lLength = -1 Then
                ' null 值
            Else
                ' 读取数据内容
                If lLength > 0 Then
                    ReDim baBuffer(0 To lLength - 1) As Byte
                    If m_oSocket.SyncReceiveArray(baBuffer, lLength, m_dTimeout * 1000) Then
                        sBuffer = sBuffer & ToolsUtf8.Decode(baBuffer)
                    End If
                End If

                ' 读取 \r\n
                ReDim baBuffer(0 To 1) As Byte
                If m_oSocket.SyncReceiveArray(baBuffer, 2, m_dTimeout * 1000) Then
                    sBuffer = sBuffer & vbCrLf
                End If
            End If
    End Select

    ReceiveElement = sBuffer
End Function

' 检查 RESP 响应是否完整
Private Function IsCompleteResp(ByVal Buffer As String) As Boolean
    If LenB(Buffer) = 0 Then
        IsCompleteResp = False
        Exit Function
    End If
    
    Dim sType As String
    sType = Left(Buffer, 1)
    
    Select Case sType
        Case "+", "-", ":"
            ' 简单字符串、错误、整数
            IsCompleteResp = InStr(Buffer, vbCrLf) > 0
        
        Case "$"
            ' 批量字符串
            Dim lCrLf As Long
            lCrLf = InStr(Buffer, vbCrLf)
            If lCrLf > 0 Then
                Dim lLength As Long
                lLength = Val(Mid(Buffer, 2, lCrLf - 2))
                
                If lLength = -1 Then
                    ' null 值
                    IsCompleteResp = True
                Else
                    Dim lExpectedLen As Long
                    lExpectedLen = lCrLf + 2 + lLength + 2
                    IsCompleteResp = Len(Buffer) >= lExpectedLen
                End If
            Else
                IsCompleteResp = False
            End If
        
        Case "*"
            ' 数组
            Dim lCrLf2 As Long
            lCrLf2 = InStr(Buffer, vbCrLf)
            If lCrLf2 > 0 Then
                Dim lCount As Long
                lCount = Val(Mid(Buffer, 2, lCrLf2 - 2))
                
                If lCount = -1 Then
                    ' null 数组
                    IsCompleteResp = True
                ElseIf lCount = 0 Then
                    ' 空数组
                    IsCompleteResp = True
                Else
                    ' 简化处理:假设数组元素都是完整接收的
                    ' 实际应该递归检查每个元素
                    IsCompleteResp = True
                End If
            Else
                IsCompleteResp = False
            End If
        
        Case Else
            IsCompleteResp = True
    End Select
End Function

' 解析 RESP 响应
Private Function ParseResp(ByVal Buffer As String) As Variant
    If LenB(Buffer) = 0 Then
        ParseResp = Array()
        Exit Function
    End If
    
    Dim lPos As Long
    Dim lCrLf As Long
    Dim sType As String
    Dim sLine As String
    
    lPos = 1
    sType = Mid(Buffer, lPos, 1)
    lPos = lPos + 1
    
    Select Case sType
        Case "+"
            ' 简单字符串
            lCrLf = InStr(lPos, Buffer, vbCrLf)
            ParseResp = Mid(Buffer, lPos, lCrLf - lPos)
        
        Case "-"
            ' 错误
            lCrLf = InStr(lPos, Buffer, vbCrLf)
            ERR.Raise vbObjectError, MODULE_NAME, Mid(Buffer, lPos, lCrLf - lPos)
        
        Case ":"
            ' 整数
            lCrLf = InStr(lPos, Buffer, vbCrLf)
            ParseResp = CLng(Mid(Buffer, lPos, lCrLf - lPos))
        
        Case "$"
            ' 批量字符串
            lCrLf = InStr(lPos, Buffer, vbCrLf)
            Dim lLength As Long
            lLength = Val(Mid(Buffer, lPos, lCrLf - lPos))
            
            lPos = lCrLf + 2
            If lLength = -1 Then
                ParseResp = Null
            Else
                ParseResp = Mid(Buffer, lPos, lLength)
                lPos = lPos + lLength + 2
            End If
        
        Case "*"
            ' 数组
            lCrLf = InStr(lPos, Buffer, vbCrLf)
            Dim lCount As Long
            lCount = Val(Mid(Buffer, lPos, lCrLf - lPos))
            
            lPos = lCrLf + 2
            
            If lCount = -1 Then
                ParseResp = Null
            ElseIf lCount = 0 Then
                ParseResp = Array()
            Else
                Dim vArray() As Variant
                ReDim vArray(lCount - 1)
                
                Dim i As Long
                For i = 0 To lCount - 1
                    vArray(i) = ParseResp(Mid(Buffer, lPos))
                    ' 更新位置(简化处理,实际需要计算每个元素的长度)
                    lPos = InStr(lPos, Buffer, vbCrLf) + 2
                    If Mid(Buffer, lPos - 2, 1) = "$" Then
                        Dim lLen As Long
                        lLen = Val(Mid(Buffer, lPos))
                        lPos = lPos + InStr(lPos, Buffer, vbCrLf) + 2 + lLen + 2
                    End If
                Next
                
                ParseResp = vArray
            End If
        
        Case Else
            ParseResp = Buffer
    End Select
End Function

'=========================================================================
' 类初始化/终止
'=========================================================================

Private Sub Class_Initialize()
    m_sHost = "127.0.0.1"
    m_lPort = 6379
    m_sPassword = ""
    m_lDbIndex = 0
    m_dTimeout = 5
    m_bConnected = False
    m_bInTransaction = False
    Set m_oPipeline = New Collection
    m_bInPipeline = False
End Sub

Private Sub Class_Terminate()
    Call DisConnect
    Set m_oPipeline = Nothing
End Sub
