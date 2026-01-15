VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpServerSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 使用 Scripting.Dictionary 存储会话数据
Dim Data As New Scripting.Dictionary

' SessionID
Private pvSessionID As String
' 创建时间
Private pvCreatedAt As Date
' 最后访问时间
Private pvLastAccessedAt As Date
' 过期时间（分钟），0 表示不过期
Private pvTimeout As Long

' 获取或设置会话数据
Public Property Get Item(ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    If Data.Exists(Key) Then
        Item = Data(Key)
    End If
End Property

Public Property Let Item(ByVal Key As String, v As Variant)
    Data(Key) = v
    pvLastAccessedAt = Now
End Property

Public Property Set Item(ByVal Key As String, v As Variant)
    Set Data(Key) = v
    pvLastAccessedAt = Now
End Property

' SessionID
Public Property Get sessionID() As String
    If pvSessionID = "" Then
        ' 生成新的 SessionID
        pvSessionID = GenerateSessionID
        pvCreatedAt = Now
        pvLastAccessedAt = Now
    End If
    sessionID = pvSessionID
End Property

Public Property Let sessionID(v As String)
    pvSessionID = v
End Property

' 创建时间
Public Property Get CreatedAt() As Date
    CreatedAt = pvCreatedAt
End Property

' 最后访问时间
Public Property Get LastAccessedAt() As Date
    LastAccessedAt = pvLastAccessedAt
End Property

' 超时时间（分钟）
Public Property Get TimeOut() As Long
    TimeOut = pvTimeout
End Property

Public Property Let TimeOut(v As Long)
    pvTimeout = v
End Property

' 检查会话是否过期
Public Function IsExpired() As Boolean
    If pvTimeout <= 0 Then
        IsExpired = False
        Exit Function
    End If
    Dim dtExpire As Date
    dtExpire = DateAdd("n", pvTimeout, pvLastAccessedAt)
    IsExpired = (Now > dtExpire)
End Function

' 检查键是否存在
Public Function Exists(ByVal Key As String) As Boolean
    Exists = Data.Exists(Key)
End Function

' 删除指定的键
Public Sub Remove(ByVal Key As String)
    If Data.Exists(Key) Then
        Data.Remove Key
    End If
End Sub

' 清空所有会话数据
Public Sub Clear()
    Data.RemoveAll
    pvSessionID = ""
End Sub

' 放弃会话（清空并重新生成 SessionID）
Public Sub Abandon()
    Data.RemoveAll
    pvSessionID = GenerateSessionID
    pvCreatedAt = Now
    pvLastAccessedAt = Now
End Sub

' 更新最后访问时间
Public Sub Touch()
    pvLastAccessedAt = Now
End Sub

' 生成 SessionID
Private Function GenerateSessionID() As String
    ' 使用 GUID 生成 SessionID
    GenerateSessionID = ToolsStr.GetGUID(False)
End Function

Private Sub Class_Initialize()
    pvTimeout = 20 ' 默认 20 分钟超时
    pvCreatedAt = Now
    pvLastAccessedAt = Now
    Data.CompareMode = TextCompare
End Sub

' ========== Session 持久化相关方法 ==========

' 序列化 Session 数据为 JSON 字符串
Public Function Serialize() As String
    On Error GoTo EH
    Dim Json As New cJson
    Dim SessionData As New Scripting.Dictionary
    
    With SessionData
        .Add "SessionID", pvSessionID
        .Add "CreatedAt", pvCreatedAt
        .Add "LastAccessedAt", pvLastAccessedAt
        .Add "Timeout", pvTimeout
        .Add "Data", Data
    End With
    
    Serialize = Json.Encode(SessionData)
    Exit Function
EH:
    Serialize = ""
End Function

' 从 JSON 字符串反序列化 Session 数据
Public Function Deserialize(ByVal JsonString As String) As Boolean
    On Error GoTo EH
    If JsonString = "" Then
        Deserialize = False
        Exit Function
    End If
    
    Dim Json As New cJson
    Dim SessionData As Scripting.Dictionary
    
    Set SessionData = Json.Decode(JsonString).Root
    
    If SessionData.Exists("SessionID") Then pvSessionID = SessionData("SessionID")
    If SessionData.Exists("CreatedAt") Then pvCreatedAt = SessionData("CreatedAt")
    If SessionData.Exists("LastAccessedAt") Then pvLastAccessedAt = SessionData("LastAccessedAt")
    If SessionData.Exists("Timeout") Then pvTimeout = SessionData("Timeout")
    If SessionData.Exists("Data") Then
        Set Data = SessionData("Data")
    End If
    
    Deserialize = True
    Exit Function
EH:
    Deserialize = False
End Function

' 保存 Session 到文件
Public Function SaveToFile(ByVal FilePath As String) As Boolean
    On Error GoTo EH
    Dim JsonString As String
    JsonString = Serialize()
    If JsonString = "" Then
        SaveToFile = False
        Exit Function
    End If
    
    FilePath = ToolsFso.AutoCompleteFullPath(FilePath, True)
    Call ToolsFso.AutoMakeDir(FilePath, True)
    Call ToolsStream.SaveFileAsText(FilePath, JsonString, "UTF-8")
    
    SaveToFile = True
    Exit Function
EH:
    SaveToFile = False
End Function

' 从文件加载 Session
Public Function LoadFromFile(ByVal FilePath As String) As Boolean
    On Error GoTo EH
    If ToolsFso.Inst.FileExists(FilePath) = False Then
        LoadFromFile = False
        Exit Function
    End If
    
    Dim JsonString As String
    JsonString = ToolsStream.LoadFileAsText(FilePath, "UTF-8")
    
    LoadFromFile = Deserialize(JsonString)
    Exit Function
EH:
    LoadFromFile = False
End Function

' 保存 Session 到数据库
' 需要表结构：SessionID (主键), SessionData (TEXT/CLOB), CreatedAt (DATETIME), LastAccessedAt (DATETIME), ExpiresAt (DATETIME)
Public Function SaveToDatabase(ByVal Db As cDataBase, Optional ByVal TableName As String = "sessions") As Boolean
    On Error GoTo EH
    Dim JsonString As String
    JsonString = Serialize()
    If JsonString = "" Then
        SaveToDatabase = False
        Exit Function
    End If
    
    Dim dtExpires As Date
    If pvTimeout > 0 Then
        dtExpires = DateAdd("n", pvTimeout, pvLastAccessedAt)
    Else
        dtExpires = DateAdd("yyyy", 1, Now)                                     ' 默认1年后过期
    End If
    
    ' 使用 INSERT OR REPLACE 或 UPDATE 语句
    Dim Sql As String
    Sql = "INSERT OR REPLACE INTO " & TableName & " (SessionID, SessionData, CreatedAt, LastAccessedAt, ExpiresAt) VALUES (?, ?, ?, ?, ?)"
    
    ' 如果数据库不支持 INSERT OR REPLACE，先尝试 UPDATE，失败则 INSERT
    If Db.Sql("SELECT SessionID FROM " & TableName & " WHERE SessionID=?") _
        .Param("SessionID", pvSessionID, adVarWChar) _
        .QueryParam Then
        
        If Not Db.Rs.EOF Then
            ' 更新现有记录
            If Not Db.Sql("UPDATE " & TableName & " SET SessionData=?, LastAccessedAt=?, ExpiresAt=? WHERE SessionID=?") _
                .Param("SessionData", JsonString, adVarWChar) _
                .Param("LastAccessedAt", pvLastAccessedAt, adDate) _
                .Param("ExpiresAt", dtExpires, adDate) _
                .Param("SessionID", pvSessionID, adVarWChar) _
                .ExecParam Then
                Db.Rs.Close
                SaveToDatabase = False
                Exit Function
            End If
        Else
            ' 插入新记录
            If Not Db.Sql("INSERT INTO " & TableName & " (SessionID, SessionData, CreatedAt, LastAccessedAt, ExpiresAt) VALUES (?, ?, ?, ?, ?)") _
                .Param("SessionID", pvSessionID, adVarWChar) _
                .Param("SessionData", JsonString, adVarWChar) _
                .Param("CreatedAt", pvCreatedAt, adDate) _
                .Param("LastAccessedAt", pvLastAccessedAt, adDate) _
                .Param("ExpiresAt", dtExpires, adDate) _
                .ExecParam Then
                Db.Rs.Close
                SaveToDatabase = False
                Exit Function
            End If
        End If
        Db.Rs.Close
    Else
        SaveToDatabase = False
        Exit Function
    End If
    
    SaveToDatabase = True
    Exit Function
EH:
    SaveToDatabase = False
End Function

' 从数据库加载 Session
Public Function LoadFromDatabase(ByVal Db As cDataBase, ByVal sessionID As String, Optional ByVal TableName As String = "sessions") As Boolean
    On Error GoTo EH
    If Not Db.Sql("SELECT SessionData, CreatedAt, LastAccessedAt FROM " & TableName & " WHERE SessionID=? AND ExpiresAt>?") _
        .Param("SessionID", sessionID, adVarWChar) _
        .Param("ExpiresAt", Now, adDate) _
        .QueryParam Then
        LoadFromDatabase = False
        Exit Function
    End If
    
    If Db.Rs.EOF Then
        Db.Rs.Close
        LoadFromDatabase = False
        Exit Function
    End If
    
    Dim JsonString As String
    JsonString = Db.Rs("SessionData")
    
    If Deserialize(JsonString) = True Then
        ' 更新最后访问时间
        If Db.Sql("UPDATE " & TableName & " SET LastAccessedAt=? WHERE SessionID=?") _
            .Param("LastAccessedAt", Now, adDate) _
            .Param("SessionID", sessionID, adVarWChar) _
            .ExecParam Then
            Db.Rs.Close
            LoadFromDatabase = True
        Else
            Db.Rs.Close
            LoadFromDatabase = False
        End If
    Else
        Db.Rs.Close
        LoadFromDatabase = False
    End If
    Exit Function
EH:
    LoadFromDatabase = False
End Function
