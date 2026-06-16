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
'   - 支持多种数据结构：String、Hash、List、Set、Sorted Set
'   - 支持管道模式（Pipeline）
'   - 支持事务模式（MULTI/EXEC/WATCH）
'   - 支持自动重连
'   - 支持 Redis 密码认证
'   - 批量接收 + 二进制安全解析
'
'=========================================================================

Private Const MODULE_NAME As String = "cRedisClient"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' 枚举定义
'=========================================================================

' SET 命令条件选项
Public Enum RedisSetCondition
    RSC_NONE = 0          ' 无条件
    RSC_NX = 1            ' SET NX (仅在 key 不存在时设置)
    RSC_XX = 2            ' SET XX (仅在 key 存在时设置)
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
Private m_bInPipeline As Boolean
Private m_lPipelineCount As Long
Private m_bAutoReconnect As Boolean
Private m_sLastError As String

'=========================================================================
' 事件
'=========================================================================

Public Event OnDisconnected()
Public Event OnError(ByVal ErrorMsg As String)

'=========================================================================
' 属性
'=========================================================================

Public Property Get Host() As String
    Host = m_sHost
End Property

Public Property Get Port() As Long
    Port = m_lPort
End Property

Public Property Get Connected() As Boolean
    Connected = m_bConnected
End Property

Public Property Get DbIndex() As Long
    DbIndex = m_lDbIndex
End Property

Public Property Let DbIndex(ByVal Value As Long)
    If m_bConnected Then
        Call SelectDb(Value)
    End If
    m_lDbIndex = Value
End Property

Public Property Get Timeout() As Double
    Timeout = m_dTimeout
End Property

Public Property Let Timeout(ByVal Value As Double)
    m_dTimeout = Value
End Property

Public Property Get InTransaction() As Boolean
    InTransaction = m_bInTransaction
End Property

Public Property Get InPipeline() As Boolean
    InPipeline = m_bInPipeline
End Property

Public Property Get AutoReconnect() As Boolean
    AutoReconnect = m_bAutoReconnect
End Property

Public Property Let AutoReconnect(ByVal Value As Boolean)
    m_bAutoReconnect = Value
End Property

Public Property Get LastError() As String
    LastError = m_sLastError
End Property

Public Property Get Socket() As cWinsock
    Set Socket = m_oSocket
End Property

'=========================================================================
' 连接管理
'=========================================================================

' 连接到 Redis 服务器
Public Function Connect(Optional ByVal Server As String = "127.0.0.1", _
    Optional ByVal Port As Long = 6379, _
    Optional ByVal Password As String = "") As Boolean
    On Error GoTo ErrorHandler

    m_sHost = Server
    m_lPort = Port
    m_sPassword = Password

    ' 关闭已有连接
    If Not m_oSocket Is Nothing Then
        m_oSocket.Close_
        Set m_oSocket = Nothing
    End If

    ' 创建新连接
    Set m_oSocket = New cWinsock
    m_oSocket.Protocol = wskTCPProtocol

    If Not m_oSocket.SyncConnect(m_sHost, m_lPort, m_dTimeout * 1000) Then
        m_sLastError = "连接超时"
        RaiseEvent OnError(m_sLastError)
        Connect = False
        Exit Function
    End If

    m_bConnected = True

    ' 密码认证
    If LenB(m_sPassword) > 0 Then
        If Not Auth(m_sPassword) Then
            m_sLastError = "认证失败"
            RaiseEvent OnError(m_sLastError)
            Connect = False
            Exit Function
        End If
    End If

    ' 切换数据库
    If m_lDbIndex > 0 Then
        Call SelectDb(m_lDbIndex)
    End If

    Connect = True
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    m_bConnected = False
    Connect = False
End Function

' 断开连接
Public Sub Disconnect()
    If Not m_oSocket Is Nothing Then
        On Error Resume Next
        m_oSocket.Close_
        Set m_oSocket = Nothing
        On Error GoTo 0
    End If
    m_bConnected = False
    m_bInTransaction = False
    m_bInPipeline = False
    m_lPipelineCount = 0
End Sub

' 认证
Public Function Auth(ByVal Password As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("AUTH", Password))
    Auth = (sResp = "OK")
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Auth = False
End Function

' 选择数据库
Public Function SelectDb(ByVal Index As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("SELECT", CStr(Index)))
    If sResp = "OK" Then
        m_lDbIndex = Index
        SelectDb = True
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SelectDb = False
End Function

' 确保连接可用（自动重连）
Private Function EnsureConnected() As Boolean
    If m_bConnected Then
        EnsureConnected = True
    ElseIf m_bAutoReconnect Then
        EnsureConnected = Connect(m_sHost, m_lPort, m_sPassword)
    End If
End Function

'=========================================================================
' 基础命令
'=========================================================================

' SET 命令（支持 EX/PX/NX/XX）
Public Function Set_(ByVal Key As String, ByVal Value As String, _
                     Optional ByVal Ex As Long = 0, _
                     Optional ByVal Px As Long = 0, _
                     Optional ByVal Condition As RedisSetCondition = RSC_NONE) As Boolean
    On Error GoTo ErrorHandler

    Dim vParts() As Variant
    Dim lCount As Long
    ReDim vParts(0 To 7) As Variant

    lCount = 0
    vParts(lCount) = "SET": lCount = lCount + 1
    vParts(lCount) = Key: lCount = lCount + 1
    vParts(lCount) = Value: lCount = lCount + 1

    If Ex > 0 Then
        vParts(lCount) = "EX": lCount = lCount + 1
        vParts(lCount) = CStr(Ex): lCount = lCount + 1
    ElseIf Px > 0 Then
        vParts(lCount) = "PX": lCount = lCount + 1
        vParts(lCount) = CStr(Px): lCount = lCount + 1
    End If

    If Condition = RSC_NX Then
        vParts(lCount) = "NX": lCount = lCount + 1
    ElseIf Condition = RSC_XX Then
        vParts(lCount) = "XX": lCount = lCount + 1
    End If

    ReDim Preserve vParts(0 To lCount - 1) As Variant

    Dim sResp As Variant
    sResp = SendCommand(vParts)

    ' SET NX/XX 返回 Null 表示未设置
    If IsNull(sResp) Then
        Set_ = False
    Else
        Set_ = (sResp = "OK")
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Set_ = False
End Function

' SETNX 命令（仅在 key 不存在时设置）
Public Function SetNx(ByVal Key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("SETNX", Key, Value))

    If IsNumeric(lRet) Then
        SetNx = (CLng(lRet) = 1)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SetNx = False
End Function

' GET 命令
Public Function Get_(ByVal Key As String) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("GET", Key))

    If IsNull(sResp) Then
        Get_ = ""
    ElseIf varType(sResp) = vbString Then
        Get_ = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Get_ = ""
End Function

' DEL 命令
Public Function Del(ParamArray Keys() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Keys
    vCmd = MergeParamArray("DEL", Keys2)
    
    Dim lRet As Variant
    lRet = SendCommand(vCmd)
    
    If IsNumeric(lRet) Then
        Del = CLng(lRet)
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Del = 0
End Function

' EXISTS 命令
Public Function Exists(ParamArray Keys() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Keys
    vCmd = MergeParamArray("EXISTS", Keys2)
    
    Dim lRet As Variant
    lRet = SendCommand(vCmd)
    
    If IsNumeric(lRet) Then
        Exists = CLng(lRet)
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Exists = 0
End Function

' KEYS 命令（生产环境建议使用 Scan 替代）
Public Function Keys(ByVal Pattern As String) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    vRet = SendCommand(Array("KEYS", Pattern))

    If IsArray(vRet) Then
        Keys = vRet
    Else
        Keys = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Keys = Array()
End Function

' SCAN 命令（生产安全的遍历方式）
' 返回 Variant 数组: vResult(0) = 下一轮游标, vResult(1) = 键数组
Public Function Scan(ByVal Cursor As Long, Optional ByVal Pattern As String = "", Optional ByVal count As Long = 0) As Variant
    On Error GoTo ErrorHandler

    Dim vParts() As Variant
    Dim lCount As Long
    ReDim vParts(0 To 5) As Variant

    lCount = 0
    vParts(lCount) = "SCAN": lCount = lCount + 1
    vParts(lCount) = CStr(Cursor): lCount = lCount + 1

    If Len(Pattern) > 0 Then
        vParts(lCount) = "MATCH": lCount = lCount + 1
        vParts(lCount) = Pattern: lCount = lCount + 1
    End If

    If count > 0 Then
        vParts(lCount) = "COUNT": lCount = lCount + 1
        vParts(lCount) = CStr(count): lCount = lCount + 1
    End If

    ReDim Preserve vParts(0 To lCount - 1) As Variant

    Scan = SendCommand(vParts)
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Scan = Array(CStr(0), Array())
End Function

' EXPIRE 命令
Public Function Expire(ByVal Key As String, ByVal Seconds As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("EXPIRE", Key, CStr(Seconds)))

    If IsNumeric(lRet) Then
        Expire = (CLng(lRet) = 1)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Expire = False
End Function

' TTL 命令
Public Function TTL(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("TTL", Key))

    If IsNumeric(lRet) Then
        TTL = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    TTL = -2
End Function

' TYPE 命令
Public Function Type_(ByVal Key As String) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("TYPE", Key))

    If varType(sResp) = vbString Then
        Type_ = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Type_ = "none"
End Function

'=========================================================================
' String 命令
'=========================================================================

' INCR 命令
Public Function Incr(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("INCR", Key))

    If IsNumeric(lRet) Then
        Incr = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Incr = 0
End Function

' DECR 命令
Public Function Decr(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("DECR", Key))

    If IsNumeric(lRet) Then
        Decr = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Decr = 0
End Function

' INCRBY 命令
Public Function IncrBy(ByVal Key As String, ByVal Increment As Long) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("INCRBY", Key, CStr(Increment)))

    If IsNumeric(lRet) Then
        IncrBy = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    IncrBy = 0
End Function

' DECRBY 命令
Public Function DecrBy(ByVal Key As String, ByVal Decrement As Long) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("DECRBY", Key, CStr(Decrement)))

    If IsNumeric(lRet) Then
        DecrBy = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    DecrBy = 0
End Function

' MGET 命令
Public Function MGet(ParamArray Keys() As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Keys
    vCmd = MergeParamArray("MGET", Keys2)
    
    Dim vRet As Variant
    vRet = SendCommand(vCmd)
    
    If IsArray(vRet) Then
        MGet = vRet
    Else
        MGet = Array()
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    MGet = Array()
End Function

' MSET 命令
Public Function MSet(ParamArray KeyValues() As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = KeyValues
    vCmd = MergeParamArray("MSET", Keys2)
    
    Dim sResp As Variant
    sResp = SendCommand(vCmd)
    MSet = (sResp = "OK")
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    MSet = False
End Function

' APPEND 命令
Public Function Append_(ByVal Key As String, ByVal Value As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("APPEND", Key, Value))

    If IsNumeric(lRet) Then
        Append_ = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Append_ = 0
End Function

' STRLEN 命令
Public Function StrLen(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("STRLEN", Key))

    If IsNumeric(lRet) Then
        StrLen = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    StrLen = 0
End Function

'=========================================================================
' Hash 命令
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
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HSet = False
End Function

' HMSET 命令（批量设置 Hash 字段）
Public Function HmSet(ByVal Key As String, ParamArray FieldValues() As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = FieldValues
    vCmd = MergeParamArray2("HMSET", Key, Keys2)
    
    Dim sResp As Variant
    sResp = SendCommand(vCmd)
    HmSet = (sResp = "OK")
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HmSet = False
End Function

' HGET 命令
Public Function HGet(ByVal Key As String, ByVal Field As String) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("HGET", Key, Field))

    If IsNull(sResp) Then
        HGet = ""
    ElseIf varType(sResp) = vbString Then
        HGet = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HGet = ""
End Function

' HMGET 命令
Public Function HMGet(ByVal Key As String, ParamArray Fields() As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Fields
    vCmd = MergeParamArray2("HMGET", Key, Keys2)
    
    Dim vRet As Variant
    vRet = SendCommand(vCmd)
    
    If IsArray(vRet) Then
        HMGet = vRet
    Else
        HMGet = Array()
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HMGet = Array()
End Function

' HGETALL 命令（后期绑定 Dictionary，无需工程引用）
Public Function HGetAll(ByVal Key As String) As Object
    On Error GoTo ErrorHandler

    Dim oDict As Object
    Set oDict = CreateObject("Scripting.Dictionary")
    oDict.CompareMode = vbTextCompare

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
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Set HGetAll = CreateObject("Scripting.Dictionary")
End Function

' HDEL 命令
Public Function HDel(ByVal Key As String, ParamArray Fields() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Fields
    vCmd = MergeParamArray2("HDEL", Key, Keys2)
    
    Dim lRet As Variant
    lRet = SendCommand(vCmd)
    
    If IsNumeric(lRet) Then
        HDel = CLng(lRet)
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HDel = 0
End Function

' HEXISTS 命令
Public Function HExists(ByVal Key As String, ByVal Field As String) As Boolean
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("HEXISTS", Key, Field))

    If IsNumeric(lRet) Then
        HExists = (CLng(lRet) = 1)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HExists = False
End Function

' HINCRBY 命令
Public Function HIncrBy(ByVal Key As String, ByVal Field As String, ByVal Increment As Long) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("HINCRBY", Key, Field, CStr(Increment)))

    If IsNumeric(lRet) Then
        HIncrBy = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HIncrBy = 0
End Function

' HLEN 命令
Public Function HLen(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("HLEN", Key))

    If IsNumeric(lRet) Then
        HLen = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HLen = 0
End Function

' HKEYS 命令
Public Function HKeys(ByVal Key As String) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    vRet = SendCommand(Array("HKEYS", Key))

    If IsArray(vRet) Then
        HKeys = vRet
    Else
        HKeys = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HKeys = Array()
End Function

' HVALS 命令
Public Function HVals(ByVal Key As String) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    vRet = SendCommand(Array("HVALS", Key))

    If IsArray(vRet) Then
        HVals = vRet
    Else
        HVals = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    HVals = Array()
End Function

'=========================================================================
' List 命令
'=========================================================================

' LPUSH 命令
Public Function LPush(ByVal Key As String, ParamArray Values() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Values
    vCmd = MergeParamArray2("LPUSH", Key, Keys2)
    
    Dim lRet As Variant
    lRet = SendCommand(vCmd)
    
    If IsNumeric(lRet) Then
        LPush = CLng(lRet)
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    LPush = 0
End Function

' RPUSH 命令
Public Function RPush(ByVal Key As String, ParamArray Values() As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Values
    vCmd = MergeParamArray2("RPUSH", Key, Keys2)
    
    Dim lRet As Variant
    lRet = SendCommand(vCmd)
    
    If IsNumeric(lRet) Then
        RPush = CLng(lRet)
    End If
    Exit Function
    
ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    RPush = 0
End Function

' LPOP 命令
Public Function LPop(ByVal Key As String) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("LPOP", Key))

    If IsNull(sResp) Then
        LPop = ""
    ElseIf varType(sResp) = vbString Then
        LPop = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    LPop = ""
End Function

' RPOP 命令
Public Function RPop(ByVal Key As String) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("RPOP", Key))

    If IsNull(sResp) Then
        RPop = ""
    ElseIf varType(sResp) = vbString Then
        RPop = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    RPop = ""
End Function

' LLEN 命令
Public Function lLen(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("LLEN", Key))

    If IsNumeric(lRet) Then
        lLen = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    lLen = 0
End Function

' LRANGE 命令
Public Function LRange(ByVal Key As String, ByVal startIdx As Long, ByVal StopIdx As Long) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    vRet = SendCommand(Array("LRANGE", Key, CStr(startIdx), CStr(StopIdx)))

    If IsArray(vRet) Then
        LRange = vRet
    Else
        LRange = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    LRange = Array()
End Function

' LINDEX 命令
Public Function lIndex(ByVal Key As String, ByVal Index As Long) As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("LINDEX", Key, CStr(Index)))

    If IsNull(sResp) Then
        lIndex = ""
    ElseIf varType(sResp) = vbString Then
        lIndex = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    lIndex = ""
End Function

' LREM 命令
Public Function LRem(ByVal Key As String, ByVal count As Long, ByVal Value As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("LREM", Key, CStr(count), Value))

    If IsNumeric(lRet) Then
        LRem = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    LRem = 0
End Function

' LTRIM 命令
Public Function LTrim(ByVal Key As String, ByVal startIdx As Long, ByVal StopIdx As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("LTRIM", Key, CStr(startIdx), CStr(StopIdx)))
    LTrim = (sResp = "OK")
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    LTrim = False
End Function

'=========================================================================
' Set 命令
'=========================================================================

' SADD 命令
Public Function SAdd(ByVal Key As String, ParamArray Members() As Variant) As Long
    On Error GoTo ErrorHandler

    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Members
    vCmd = MergeParamArray2("SADD", Key, Keys2)

    Dim lRet As Variant
    lRet = SendCommand(vCmd)

    If IsNumeric(lRet) Then
        SAdd = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SAdd = 0
End Function

' SMEMBERS 命令
Public Function SMembers(ByVal Key As String) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    vRet = SendCommand(Array("SMEMBERS", Key))

    If IsArray(vRet) Then
        SMembers = vRet
    Else
        SMembers = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SMembers = Array()
End Function

' SCARD 命令
Public Function SCard(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("SCARD", Key))

    If IsNumeric(lRet) Then
        SCard = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SCard = 0
End Function

' SISMEMBER 命令
Public Function SIsMember(ByVal Key As String, ByVal Member As String) As Boolean
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("SISMEMBER", Key, Member))

    If IsNumeric(lRet) Then
        SIsMember = (CLng(lRet) = 1)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SIsMember = False
End Function

' SREM 命令
Public Function SRem(ByVal Key As String, ParamArray Members() As Variant) As Long
    On Error GoTo ErrorHandler

    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Members
    vCmd = MergeParamArray2("SREM", Key, Keys2)

    Dim lRet As Variant
    lRet = SendCommand(vCmd)

    If IsNumeric(lRet) Then
        SRem = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    SRem = 0
End Function

'=========================================================================
' Sorted Set 命令
'=========================================================================

' ZADD 命令
Public Function ZAdd(ByVal Key As String, ByVal Score As Double, ByVal Member As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("ZADD", Key, CStr(Score), Member))

    If IsNumeric(lRet) Then
        ZAdd = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZAdd = 0
End Function

' ZRANGE 命令
Public Function ZRange(ByVal Key As String, ByVal startIdx As Long, ByVal StopIdx As Long, Optional ByVal WithScores As Boolean = False) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    If WithScores Then
        vRet = SendCommand(Array("ZRANGE", Key, CStr(startIdx), CStr(StopIdx), "WITHSCORES"))
    Else
        vRet = SendCommand(Array("ZRANGE", Key, CStr(startIdx), CStr(StopIdx)))
    End If

    If IsArray(vRet) Then
        ZRange = vRet
    Else
        ZRange = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZRange = Array()
End Function

' ZREVRANGE 命令
Public Function ZRevRange(ByVal Key As String, ByVal startIdx As Long, ByVal StopIdx As Long, Optional ByVal WithScores As Boolean = False) As Variant
    On Error GoTo ErrorHandler

    Dim vRet As Variant
    If WithScores Then
        vRet = SendCommand(Array("ZREVRANGE", Key, CStr(startIdx), CStr(StopIdx), "WITHSCORES"))
    Else
        vRet = SendCommand(Array("ZREVRANGE", Key, CStr(startIdx), CStr(StopIdx)))
    End If

    If IsArray(vRet) Then
        ZRevRange = vRet
    Else
        ZRevRange = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZRevRange = Array()
End Function

' ZSCORE 命令
Public Function ZScore(ByVal Key As String, ByVal Member As String) As Variant
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("ZSCORE", Key, Member))

    If IsNull(sResp) Then
        ZScore = Null
    ElseIf IsNumeric(sResp) Then
        ZScore = CDbl(sResp)
    Else
        ZScore = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZScore = Null
End Function

' ZRANK 命令
Public Function ZRank(ByVal Key As String, ByVal Member As String) As Variant
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("ZRANK", Key, Member))

    If IsNull(lRet) Then
        ZRank = Null
    ElseIf IsNumeric(lRet) Then
        ZRank = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZRank = Null
End Function

' ZREM 命令
Public Function ZRem(ByVal Key As String, ParamArray Members() As Variant) As Long
    On Error GoTo ErrorHandler

    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Members
    vCmd = MergeParamArray2("ZREM", Key, Keys2)

    Dim lRet As Variant
    lRet = SendCommand(vCmd)

    If IsNumeric(lRet) Then
        ZRem = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZRem = 0
End Function

' ZCARD 命令
Public Function ZCard(ByVal Key As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("ZCARD", Key))

    If IsNumeric(lRet) Then
        ZCard = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    ZCard = 0
End Function

'=========================================================================
' 事务命令
'=========================================================================

' WATCH 命令（乐观锁）
Public Function Watch(ParamArray Keys() As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim vCmd As Variant, Keys2() As Variant
    Keys2 = Keys
    vCmd = MergeParamArray("WATCH", Keys2)

    Dim sResp As Variant
    sResp = SendCommand(vCmd)
    Watch = (sResp = "OK")
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Watch = False
End Function

' UNWATCH 命令
Public Function Unwatch() As Boolean
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("UNWATCH"))
    Unwatch = (sResp = "OK")
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Unwatch = False
End Function

' MULTI 命令
Public Function Multi() As Boolean
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("MULTI"))

    If sResp = "OK" Then
        m_bInTransaction = True
        Multi = True
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Multi = False
End Function

' EXEC 命令
Public Function Exec() As Variant
    On Error GoTo ErrorHandler

    If Not m_bInTransaction Then
        m_sLastError = "Not in transaction"
        RaiseEvent OnError(m_sLastError)
        Exec = Array()
        Exit Function
    End If

    Dim vRet As Variant
    vRet = SendCommand(Array("EXEC"))

    m_bInTransaction = False

    If IsArray(vRet) Then
        Exec = vRet
    Else
        Exec = Array()
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    m_bInTransaction = False
    Exec = Array()
End Function

' DISCARD 命令
Public Function Discard() As Boolean
    On Error GoTo ErrorHandler

    If Not m_bInTransaction Then
        Discard = False
        Exit Function
    End If

    Dim sResp As Variant
    sResp = SendCommand(Array("DISCARD"))

    m_bInTransaction = False

    If sResp = "OK" Then
        Discard = True
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    m_bInTransaction = False
    Discard = False
End Function

'=========================================================================
' Pipeline 命令
'=========================================================================

' 开始管道模式
Public Function PipelineStart() As Boolean
    If m_bInPipeline Then
        m_sLastError = "Already in pipeline mode"
        PipelineStart = False
        Exit Function
    End If
    If m_bInTransaction Then
        m_sLastError = "Cannot start pipeline in transaction"
        PipelineStart = False
        Exit Function
    End If

    m_bInPipeline = True
    m_lPipelineCount = 0
    PipelineStart = True
End Function

' 刷新管道，读取所有响应
Public Function PipelineFlush() As Variant
    On Error GoTo ErrorHandler

    If Not m_bInPipeline Then
        m_sLastError = "Not in pipeline mode"
        PipelineFlush = Array()
        Exit Function
    End If

    If m_lPipelineCount = 0 Then
        m_bInPipeline = False
        PipelineFlush = Array()
        Exit Function
    End If

    Dim vResults() As Variant
    ReDim vResults(m_lPipelineCount - 1) As Variant

    Dim i As Long
    For i = 0 To m_lPipelineCount - 1
        vResults(i) = SafeReadResp()
    Next

    m_bInPipeline = False
    m_lPipelineCount = 0
    PipelineFlush = vResults
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    m_bInPipeline = False
    m_lPipelineCount = 0
    PipelineFlush = Array()
End Function

' 取消管道（安全断开，避免响应错位）
Public Sub PipelineDiscard()
    If m_bInPipeline Then
        m_bInPipeline = False
        m_lPipelineCount = 0
        Call Disconnect
    End If
End Sub

'=========================================================================
' 服务器命令
'=========================================================================

' PING 命令
Public Function Ping() As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("PING"))

    If varType(sResp) = vbString Then
        Ping = sResp
    Else
        Ping = "PONG"
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Ping = ""
End Function

' INFO 命令
Public Function Info(Optional ByVal Section As String = "") As String
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    If LenB(Section) > 0 Then
        sResp = SendCommand(Array("INFO", Section))
    Else
        sResp = SendCommand(Array("INFO"))
    End If

    If varType(sResp) = vbString Then
        Info = sResp
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Info = ""
End Function

' FLUSHDB 命令
Public Function FlushDb() As Boolean
    On Error GoTo ErrorHandler

    Dim sResp As Variant
    sResp = SendCommand(Array("FLUSHDB"))
    FlushDb = (sResp = "OK")
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    FlushDb = False
End Function

' DBSIZE 命令
Public Function DBSize() As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("DBSIZE"))

    If IsNumeric(lRet) Then
        DBSize = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    DBSize = 0
End Function

' PUBLISH 命令
Public Function Publish(ByVal Channel As String, ByVal message As String) As Long
    On Error GoTo ErrorHandler

    Dim lRet As Variant
    lRet = SendCommand(Array("PUBLISH", Channel, message))

    If IsNumeric(lRet) Then
        Publish = CLng(lRet)
    End If
    Exit Function

ErrorHandler:
    m_sLastError = ERR.Description
    RaiseEvent OnError(m_sLastError)
    Publish = 0
End Function

'=========================================================================
' 私有方法 - 命令发送
'=========================================================================

' 发送命令并返回响应
Private Function SendCommand(ByVal Command As Variant) As Variant
    If Not EnsureConnected() Then
        ERR.Raise vbObjectError, MODULE_NAME, "Not connected to Redis server"
    End If

    ' 构建命令（生成 UTF-8 字节数组）
    Dim baCmd() As Byte
    baCmd = BuildRedisCommand(Command)

    ' 同步发送
    If Not m_oSocket.SyncSendArray(baCmd, m_dTimeout * 1000) Then
        m_bConnected = False
        ERR.Raise vbObjectError, MODULE_NAME, "Send command timeout"
    End If

    If m_bInPipeline Then
        ' 管道模式：只发送不接收
        m_lPipelineCount = m_lPipelineCount + 1
        SendCommand = Empty
    Else
        ' 正常模式：发送并接收响应
        SendCommand = ReadResp()
    End If
End Function

'=========================================================================
' 私有方法 - RESP 接收与解析
'=========================================================================

' 从 Socket 读取一行（以 CRLF 结尾），返回不含 CRLF 的字符串
' 优化：使用字节数组累积代替字符串拼接，避免 O(N^2) 的 BSTR 重分配
Private Function ReadLine() As String
    Dim baBuf() As Byte
    ReDim baBuf(0 To 255) As Byte
    Dim lLen As Long

    Do
        Dim baOne(0 To 0) As Byte
        If Not m_oSocket.SyncReceiveArray(baOne, 1, m_dTimeout * 1000) Then
            ERR.Raise vbObjectError, MODULE_NAME, "Read line timeout"
        End If

        ' 翻倍扩容策略
        If lLen > UBound(baBuf) Then
            ReDim Preserve baBuf(0 To UBound(baBuf) * 2) As Byte
        End If

        baBuf(lLen) = baOne(0)
        lLen = lLen + 1

        ' 检测 CRLF
        If lLen >= 2 Then
            If baBuf(lLen - 2) = 13 And baBuf(lLen - 1) = 10 Then
                Exit Do
            End If
        End If
    Loop

    ' UTF-8 解码（去除末尾 CRLF）
    If lLen > 2 Then
        ReDim Preserve baBuf(0 To lLen - 3) As Byte
        ReadLine = ToolsUtf8.Decode(baBuf)
    Else
        ReadLine = ""
    End If
End Function

' 从 Socket 读取指定字节数，返回原始字节数组
' 用于 Bulk String 的二进制安全读取
Private Function ReadBytes(ByVal nBytes As Long) As Byte()
    If nBytes <= 0 Then
        Dim baEmpty() As Byte
        ReadBytes = baEmpty
        Exit Function
    End If

    Dim baResult() As Byte
    ReDim baResult(0 To nBytes - 1) As Byte
    If Not m_oSocket.SyncReceiveArray(baResult, nBytes, m_dTimeout * 1000) Then
        ERR.Raise vbObjectError, MODULE_NAME, "Read bytes timeout"
    End If
    ReadBytes = baResult
End Function

' 读取一个完整的 RESP 值（递归解析，二进制安全）
' 替代原有的 ReceiveResponse + ReceiveElement + ParseResp 三合一
Private Function ReadResp() As Variant
    Dim sLine As String
    sLine = ReadLine()

    If Len(sLine) = 0 Then
        ERR.Raise vbObjectError, MODULE_NAME, "Empty RESP response"
    End If

    Dim sType As String
    sType = Left$(sLine, 1)
    Dim sPayload As String
    sPayload = Mid$(sLine, 2)

    Select Case sType
        Case "+"
            ' Simple String
            ReadResp = sPayload

        Case "-"
            ' Error
            ERR.Raise vbObjectError, MODULE_NAME, sPayload

        Case ":"
            ' Integer
            ReadResp = CLng(sPayload)

        Case "$"
            ' Bulk String —— 按字节长度精准读取，确保二进制安全
            Dim lLen As Long
            lLen = CLng(sPayload)
            If lLen = -1 Then
                ReadResp = Null
            ElseIf lLen = 0 Then
                Call ReadLine    ' 消耗尾随 CRLF
                ReadResp = ""
            Else
                Dim baData() As Byte
                baData = ReadBytes(lLen)
                ReadResp = ToolsUtf8.Decode(baData)
                Call ReadLine    ' 消耗尾随 CRLF
            End If

        Case "*"
            ' Array —— 递归解析每个元素
            Dim lCount As Long
            lCount = CLng(sPayload)
            If lCount = -1 Then
                ReadResp = Null
            ElseIf lCount = 0 Then
                ReadResp = Array()
            Else
                Dim vArr() As Variant
                ReDim vArr(lCount - 1) As Variant
                Dim i As Long
                For i = 0 To lCount - 1
                    vArr(i) = ReadResp()
                Next
                ReadResp = vArr
            End If

        Case Else
            ERR.Raise vbObjectError, MODULE_NAME, "Unknown RESP type: " & sType
    End Select
End Function

' 安全读取 RESP 响应（用于 Pipeline，错误不抛异常而是返回错误字符串）
Private Function SafeReadResp() As Variant
    On Error GoTo SafeHandler
    SafeReadResp = ReadResp()
    Exit Function
SafeHandler:
    SafeReadResp = "ERROR: " & ERR.Description
End Function

'=========================================================================
' 私有方法 - 命令构建
'=========================================================================

' 构建 Redis RESP 协议命令（返回 UTF-8 字节数组）
' 优化：翻倍扩容 + CopyMemory 批量拷贝代替逐字节赋值
Private Function BuildRedisCommand(ByVal Command As Variant) As Byte()
    Dim vFlat() As Variant
    vFlat = FlattenArray(Command)
    Dim lCount As Long
    lCount = UBound(vFlat) + 1

    ' 构建命令头: *<count>\r\n
    Dim sHeader As String
    sHeader = "*" & lCount & vbCrLf
    Dim baHeader() As Byte
    baHeader = ToolsUtf8.Encode(sHeader)

    ' 预分配结果缓冲区（翻倍策略）
    Dim lCapacity As Long
    lCapacity = 2048
    Dim baResult() As Byte
    ReDim baResult(0 To lCapacity - 1) As Byte
    Dim lOffset As Long

    ' 写入命令头
    Call AppendBytes(baResult, lOffset, lCapacity, baHeader)

    ' 写入每个元素
    Dim i As Long
    For i = 0 To lCount - 1
        Dim sPart As String
        sPart = CStr(vFlat(i))

        ' 编码为 UTF-8
        Dim baPart() As Byte
        baPart = ToolsUtf8.Encode(sPart)

        ' 元素头: $<length>\r\n
        Dim sBulkHeader As String
        sBulkHeader = "$" & (UBound(baPart) + 1) & vbCrLf
        Dim baBulkHeader() As Byte
        baBulkHeader = ToolsUtf8.Encode(sBulkHeader)

        Call AppendBytes(baResult, lOffset, lCapacity, baBulkHeader)
        Call AppendBytes(baResult, lOffset, lCapacity, baPart)
        Call AppendBytes(baResult, lOffset, lCapacity, ToolsUtf8.Encode(vbCrLf))
    Next

    ' 截断到实际大小
    If lOffset > 0 Then
        ReDim Preserve baResult(0 To lOffset - 1) As Byte
    End If

    BuildRedisCommand = baResult
End Function

' 追加字节数组到结果（自动翻倍扩容 + CopyMemory 高效拷贝）
Private Sub AppendBytes(ByRef baResult() As Byte, ByRef lOffset As Long, ByRef lCapacity As Long, ByRef baSource() As Byte)
    Dim lSrcLen As Long
    lSrcLen = UBound(baSource) + 1
    If lSrcLen <= 0 Then Exit Sub

    ' 确保容量足够（翻倍扩容）
    Do While lOffset + lSrcLen > lCapacity
        lCapacity = lCapacity * 2
        ReDim Preserve baResult(0 To lCapacity - 1) As Byte
    Loop

    ' 内存拷贝，比逐字节赋值快一个数量级
    CopyMemory baResult(lOffset), baSource(0), lSrcLen
    lOffset = lOffset + lSrcLen
End Sub

'=========================================================================
' 私有方法 - 数组工具
'=========================================================================

' 展平嵌套数组
' 优化：先递归计数，一次性分配，再递归填充，避免逐元素 ReDim Preserve
Private Function FlattenArray(ByVal vInput As Variant) As Variant()
    Dim lCount As Long
    lCount = CountElements(vInput)

    If lCount = 0 Then
        Dim vEmpty(0) As Variant
        FlattenArray = vEmpty
        Exit Function
    End If

    Dim vResult() As Variant
    ReDim vResult(0 To lCount - 1) As Variant
    Dim lIndex As Long
    lIndex = 0
    Call FillElements(vInput, vResult, lIndex)

    FlattenArray = vResult
End Function

' 递归计数数组元素
Private Function CountElements(ByVal vInput As Variant) As Long
    If IsArray(vInput) Then
        Dim i As Long
        For i = LBound(vInput) To UBound(vInput)
            CountElements = CountElements + CountElements(vInput(i))
        Next
    Else
        CountElements = 1
    End If
End Function

' 递归填充数组元素
Private Sub FillElements(ByVal vInput As Variant, ByRef vResult() As Variant, ByRef lIndex As Long)
    If IsArray(vInput) Then
        Dim i As Long
        For i = LBound(vInput) To UBound(vInput)
            FillElements vInput(i), vResult, lIndex
        Next
    Else
        vResult(lIndex) = vInput
        lIndex = lIndex + 1
    End If
End Sub

' 合并命令名 + ParamArray 为一维数组
' 用于 Del/Exists/MGet/MSet/Keys 等只有 ParamArray 的命令
' 替代原来的 Array("CMD", Params) 嵌套方式，避免 FlattenArray 深层递归
Private Function MergeParamArray(ByVal CmdName As String, ByRef params() As Variant) As Variant
    Dim lCount As Long
    On Error Resume Next
    lCount = UBound(params) - LBound(params) + 1
    If ERR.Number <> 0 Then lCount = 0
    On Error GoTo 0
    
    Dim vCmd() As Variant
    ReDim vCmd(0 To lCount) As Variant
    vCmd(0) = CmdName
    
    Dim i As Long
    For i = 0 To lCount - 1
        vCmd(i + 1) = params(LBound(params) + i)
    Next
    
    MergeParamArray = vCmd
End Function

' 合并命令名 + Key + ParamArray 为一维数组
' 用于 LPush/RPush/SAdd/HDel 等有 Key + ParamArray 的命令
Private Function MergeParamArray2(ByVal CmdName As String, ByVal Key As String, ByRef params() As Variant) As Variant
    Dim lCount As Long
    On Error Resume Next
    lCount = UBound(params) - LBound(params) + 1
    If ERR.Number <> 0 Then lCount = 0
    On Error GoTo 0

    Dim vCmd() As Variant
    ReDim vCmd(0 To lCount + 1) As Variant
    vCmd(0) = CmdName
    vCmd(1) = Key

    Dim i As Long
    For i = 0 To lCount - 1
        vCmd(i + 2) = params(LBound(params) + i)
    Next

    MergeParamArray2 = vCmd
End Function

'=========================================================================
' 初始化 / 终止
'=========================================================================

Private Sub Class_Initialize()
    m_sHost = "127.0.0.1"
    m_lPort = 6379
    m_sPassword = ""
    m_lDbIndex = 0
    m_dTimeout = 5
    m_bConnected = False
    m_bInTransaction = False
    m_bInPipeline = False
    m_lPipelineCount = 0
    m_bAutoReconnect = False
    m_sLastError = ""
End Sub

Private Sub Class_Terminate()
    Call Disconnect
End Sub
