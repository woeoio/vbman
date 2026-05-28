VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMSMQMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===========================================================================
' 类名:    cMSMQMessage
' 描述:    MSMQ 消息封装类
' 作者:    [作者]
' 日期:    [日期]
' 版本:    1.0.0
' 说明:    封装 MSMQMessage 对象，提供友好的属性访问
'===========================================================================

'----------------------------------
' 成员变量
'----------------------------------
Private m_Body As Variant
Private m_Label As String
Private m_Priority As Long
Private m_Delivery As Long
Private m_SentTime As Date
Private m_ArrivedTime As Date
Private m_MessageId As String
Private m_CorrelationId As String
Private m_AppSpecific As Long

'----------------------------------
' 属性: 消息正文
'----------------------------------
Public Property Get Body() As Variant
    Body = m_Body
End Property

Public Property Let Body(ByVal Value As Variant)
    m_Body = Value
End Property

'----------------------------------
' 属性: 消息标签
'----------------------------------
Public Property Get Label() As String
    Label = m_Label
End Property

Public Property Let Label(ByVal Value As String)
    m_Label = Value
End Property

'----------------------------------
' 属性: 优先级 (0-7)
'----------------------------------
Public Property Get Priority() As Long
    Priority = m_Priority
End Property

Public Property Let Priority(ByVal Value As Long)
    If Value < 0 Then Value = 0
    If Value > 7 Then Value = 7
    m_Priority = Value
End Property

'----------------------------------
' 属性: 发送时间
'----------------------------------
Public Property Get SentTime() As Date
    SentTime = m_SentTime
End Property

'----------------------------------
' 属性: 到达时间
'----------------------------------
Public Property Get ArrivedTime() As Date
    ArrivedTime = m_ArrivedTime
End Property

'----------------------------------
' 属性: 消息ID
'----------------------------------
Public Property Get MessageId() As String
    MessageId = m_MessageId
End Property

'----------------------------------
' 属性: 关联ID
'----------------------------------
Public Property Get CorrelationId() As String
    CorrelationId = m_CorrelationId
End Property

Public Property Let CorrelationId(ByVal Value As String)
    m_CorrelationId = Value
End Property

'----------------------------------
' 属性: 应用程序特定值
'----------------------------------
Public Property Get AppSpecific() As Long
    AppSpecific = m_AppSpecific
End Property

Public Property Let AppSpecific(ByVal Value As Long)
    m_AppSpecific = Value
End Property

'----------------------------------
' 方法: 从 MSMQMessage 对象加载数据
'----------------------------------
Friend Sub LoadFromMSMQMessage(ByRef SourceMsg As Object)
    On Error Resume Next
    
    m_Body = SourceMsg.Body
    m_Label = SourceMsg.Label
    m_Priority = SourceMsg.Priority
    m_Delivery = SourceMsg.Delivery
    m_AppSpecific = SourceMsg.AppSpecific
    
    ' 时间戳处理（MSMQ 返回的是 UTC 时间）
    m_SentTime = SourceMsg.SentTime
    m_ArrivedTime = SourceMsg.ArrivedTime
    
    ' ID 转换为字符串
    m_MessageId = IdToString(SourceMsg.Id)
    m_CorrelationId = IdToString(SourceMsg.CorrelationId)
    
    On Error GoTo 0
End Sub

'----------------------------------
' 私有方法: 将字节数组ID转换为字符串
'----------------------------------
Private Function IdToString(ByRef Id As Variant) As String
    On Error Resume Next
    
    Dim bytes() As Byte
    Dim i As Long
    Dim result As String
    
    ' MSMQ ID 是字节数组
    bytes = Id
    
    For i = LBound(bytes) To UBound(bytes)
        result = result & Right$("0" & Hex$(bytes(i)), 2)
    Next i
    
    IdToString = result
    On Error GoTo 0
End Function

'----------------------------------
' 方法: 获取JSON表示（用于调试）
'----------------------------------
Public Function ToJSON() As String
    Dim json As String
    
    json = "{"
    json = json & """Label"": """ & m_Label & ""","
    json = json & """Priority"": " & m_Priority & ","
    json = json & """AppSpecific"": " & m_AppSpecific & ","
    json = json & """SentTime"": """ & m_SentTime & ""","
    json = json & """MessageId"": """ & m_MessageId & ""","
    json = json & """Body"": """ & CStr(m_Body) & """"
    json = json & "}"
    
    ToJSON = json
End Function

'----------------------------------
' 初始化
'----------------------------------
Private Sub Class_Initialize()
    m_Priority = 3  ' 默认普通优先级
End Sub
