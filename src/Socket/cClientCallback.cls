VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClientCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public hSocket              As Long
Public dataTotal            As Long
Public ClientInfo           As cHttpServerClientInfo
Public Context              As cHttpServerContext
Public Parent               As cHttpServer
Public WithEvents Socket    As cTlsReMaster
Attribute Socket.VB_VarHelpID = -1
Public IsSSE                As Boolean
' 最后活动时间，用于空闲连接超时检测
Public LastActivity         As Date

' 接收缓冲区（用于解决 TCP 粘包问题）
Private recvBuffer()        As Byte
Private recvBufferLen       As Long
' 缓冲区缩容阈值：当缓冲区容量超过此值且使用率低于 1/4 时缩容
Private Const RECVBUFFER_SHRINK_CAP As Long = 65536

' 缓冲区访问方法
Friend Property Get GetRecvBuffer() As Byte()
    GetRecvBuffer = recvBuffer
End Property

Friend Property Let GetRecvBuffer(data() As Byte)
    recvBuffer = data
End Property

Friend Property Get GetRecvBufferLen() As Long
    GetRecvBufferLen = recvBufferLen
End Property

Friend Property Let GetRecvBufferLen(ByVal value As Long)
    recvBufferLen = value
End Property

' 追加数据到缓冲区（解决粘包）
Friend Sub AppendToBuffer(data() As Byte)
    If recvBufferLen = 0 Then
        recvBuffer = data
        recvBufferLen = UBound(data) + 1
    Else
        Dim totalLen As Long
        totalLen = recvBufferLen + UBound(data) + 1
        
        If UBound(recvBuffer) + 1 < totalLen Then
            ReDim Preserve recvBuffer(0 To totalLen - 1)
        End If
        
        CopyMemory recvBuffer(recvBufferLen), data(0), UBound(data) + 1
        recvBufferLen = totalLen
    End If
End Sub

' 从缓冲区移除已处理的数据
Friend Sub RemoveFromBuffer(processedLen As Long)
    If processedLen >= recvBufferLen Then
        recvBufferLen = 0
        ' BUG-4 修复：缓冲区已空，释放大 buffer 防止内存膨胀
        If (Not recvBuffer) = -1 Then Exit Sub  ' 未初始化
        If UBound(recvBuffer) + 1 > RECVBUFFER_SHRINK_CAP Then
            Erase recvBuffer
        End If
    Else
        Dim remainingLen As Long
        remainingLen = recvBufferLen - processedLen
        CopyMemory recvBuffer(0), recvBuffer(processedLen), remainingLen
        recvBufferLen = remainingLen
        ' BUG-4 修复：如果缓冲区容量过大但使用率低，缩容
        Dim cap As Long
        If (Not recvBuffer) = -1 Then Exit Sub
        cap = UBound(recvBuffer) + 1
        If cap > RECVBUFFER_SHRINK_CAP And remainingLen < cap \ 4 Then
            ReDim Preserve recvBuffer(0 To remainingLen - 1)
        End If
    End If
End Sub

' BUG-1 修复：用 hSocket 替代 Index 作为连接池查找键
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    dataTotal = bytesTotal
    LastActivity = Now  ' BUG-2 修复：更新活动时间
    If Not Parent Is Nothing Then Parent.OnDataArrival hSocket
End Sub

Private Sub Socket_CloseSck()
    If Not Parent Is Nothing Then Parent.OnCloseSck hSocket
End Sub

' 显式释放所有引用，断开循环引用链
' 释放顺序：Response.fClient → Context.Release → Parent → Socket
Friend Sub Release()
    ' 1. 断开 Response 对 Client 的循环引用
    If Not Context Is Nothing Then
        If Not Context.Response Is Nothing Then
            Set Context.Response.fClient = Nothing
        End If
        ' 2. 释放 Context 内部所有对象引用
        Context.Release
        Set Context = Nothing
    End If
    ' 3. 断开 Client 对 cHttpServer 的引用，解开 cClientCallback → cHttpServer 循环引用
    Set Parent = Nothing
    ' 4. 释放 ClientInfo
    Set ClientInfo = Nothing
    ' 5. 关闭并释放 Socket
    If Not Socket Is Nothing Then
        On Error Resume Next
        Socket.CloseSck
        On Error GoTo 0
        Set Socket = Nothing
    End If
    ' 6. 彻底清空接收缓冲区 (BUG-4 修复)
    Erase recvBuffer
    recvBufferLen = 0
    IsSSE = False
    hSocket = 0
End Sub
