VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'作者：邓伟，215879458@qq.com
'引用：
'Microsoft ActiveX Data Objects 2.8 Library
'依赖：
'ToolsDic.bas

Public Event AsyncExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)

Public IsConnect As Boolean
'
Public LastErr As String
Public LastErrNumber As Long
Public LastErrDescription As String
'Save rs in class inst.
Public Rs As New ADODB.Recordset
'convert rs to Dictionary object an save in class inst.
Public Rows As New Collection
Public Row As New Scripting.Dictionary
'Connection inst 20241011 今天发现conn居然有异步执行sql的能力, 后期挖掘下, 实现和前端异步交互，还有 stream 模式
Public WithEvents Conn As ADODB.Connection
Attribute Conn.VB_VarHelpID = -1
'cDatabase pool, 多个数据库对象池，主要用于动态根据字符串key使用数据库对象的场景
Dim Connections As New Scripting.Dictionary

'
Private IsPage As Boolean
Private IsTrans As Boolean
Private pvRawSqlString As String
Private pvPageNum As Long
Private pvPageLimit As Long
Private pvPageOffset As Long

Dim IsAsync As Boolean
Dim Cmd As ADODB.Command

Public Enum enumDbType
    Access = 1
    Mysql = 2
    MsSql = 3
    Csv = 4
End Enum
Public DatabaseType As enumDbType

Public Property Get Async() As cDataBase
    IsAsync = True
    Set Async = Me
End Property

Public Function ConnInst(ByVal InstName As String, Optional ByVal IsCloneMasterConnection As Boolean = True) As cDataBase
    If Connections.Exists(InstName) = False Then
        Connections.Add InstName, New cDataBase
        Set ConnInst = Connections(InstName)
        If IsCloneMasterConnection = True Then
            ConnInst.Conn = Me.Conn
            ConnInst.Connect DatabaseType
        End If
    Else
        Set ConnInst = Connections(InstName)
    End If
End Function
Public Sub ConnInstRemove(Optional ByVal InstName As String)
    If InstName = "" Then Connections.RemoveAll: Exit Sub
    If Connections.Exists(InstName) = False Then Exit Sub
    Set Connections(InstName) = Nothing
    Connections.Remove InstName
End Sub

Public Function TransBegin() As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    Conn.BeginTrans
    IsTrans = True
    TransBegin = True
    Exit Function
eee:
    LastErr = "TransBegin Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function
Public Function TransCommit() As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    Conn.CommitTrans
    IsTrans = False
    TransCommit = True
    Exit Function
eee:
    LastErr = "TransCommit Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    '如果提交失败，自动回滚
    If IsTrans Then Call TransRollback
End Function
Public Function TransRollback() As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    Conn.RollbackTrans
    IsTrans = False
    TransRollback = True
    Exit Function
eee:
    LastErr = "TransRollback Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

Public Function Page(Optional num As Long = 1, Optional Limit As Long = 10) As cDataBase
    Set Page = Me
    IsPage = True
    pvPageNum = num
    pvPageLimit = Limit
    If pvPageNum < 1 Then pvPageNum = 1
    If pvPageLimit < 1 Then pvPageLimit = 10
    pvPageOffset = (pvPageNum - 1) * pvPageLimit
End Function

Public Function Connect( _
    ByVal DbType As enumDbType, _
    Optional ByVal DbAddress As String = "127.0.0.1,1433", _
    Optional ByVal username As String = "sa", _
    Optional ByVal password As String = "Sa123456", _
    Optional ByVal DefaultDataBase As String = "master") As Boolean
    
    On Error GoTo eee
    
    '如果连接标记为真，则返回。否则会出错
    If IsConnect = True Then
        '返回成功
        Connect = True
        Exit Function
    End If
    DatabaseType = DbType
    '设置连接字符串ConnectionString属性
    If Conn.ConnectionString = "" Then
        Select Case DbType
        Case Access
            DbAddress = ToolsFso.AutoCompleteFullPath(DbAddress, True)
            Conn.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & DbAddress & ";"
        Case Csv
            DbAddress = ToolsFso.AutoCompleteFullPath(DbAddress, True)
            Conn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" & DbAddress
        Case MsSql
            Conn.ConnectionString = "Driver={SQL Server};Server=" & DbAddress & ";Uid=" & username & ";pwd=" & password & ";Database=" & DefaultDataBase & ";"
        Case Mysql
            Dim Svr: Svr = Split(DbAddress, ",")
            Dim Address$: Address = Svr(0)
            Dim Port$: Port = "3306"
            If UBound(Svr) > 0 Then Port = Svr(1)
            DbAddress = Replace(DbAddress, "，", ",")
            '如果包含逗号，说明有端口号，转换为冒号格式
            If InStr(DbAddress, ",") > 0 Then
                DbAddress = Replace(DbAddress, ",", ":")
            End If
            Conn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" & DbAddress & ";Uid=" & username & ";pwd=" & password & ";Database=" & DefaultDataBase & ";"
        End Select
    End If
    '打开到数据库的连接
    Conn.Open
    '判断连接的状态
    If Conn.State <> adStateOpen Then
        '如果连接不成功，则显示提示信息，退出程序
        Exit Function
    End If
    '设置rst的ActiveConnection属性，指定与其关联的数据库连接
    Set Rs.ActiveConnection = Conn
    '根据链接类型设置一些必要的查询缺省值，减少外部查询代码的累赘
    Select Case DbType
    Case Access
        'todo
    Case Csv
        'todo
    Case MsSql
        'todo
    Case Mysql
        'todo
    End Select
    
    Conn.CursorLocation = adUseClient
    '设置连接标记，表示已经连接到数据库
    IsConnect = True
    Connect = True
    Exit Function
eee:
    LastErr = "Connect Err[" & ERR.Number & "]" & ERR.Description & "{" & DbAddress & "}{" & DbType & "}{" & username & "}{" & DefaultDataBase & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ERR.Raise LastErrNumber, , LastErr
End Function

'断开与数据库的连接
Public Function Disconnect() As Boolean
    On Error GoTo eee
    
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then
        Exit Function
    End If
    '如果正在事务中，先回滚
    If IsTrans Then Call TransRollback
    '关闭连接
    Conn.Close
    If Rs.State <> adStateClosed Then Rs.Close
    '释放conn
    Set Conn = Nothing
    Set Rs = Nothing
    
    '设置连接标记，表示已经断开与数据库的连接
    IsConnect = False
    Disconnect = True
    
    Exit Function
eee:
    LastErr = "Disconnect Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

Public Function Sql(ByVal RawSqlString As String) As cDataBase
    Set Sql = Me
    pvRawSqlString = RawSqlString
    IsPage = False
End Function

'参数化查询 - 防止SQL注入
Public Function Param(ByVal ParamName As String, ByVal ParamValue As Variant, Optional ByVal ParamType As DataTypeEnum = adVarChar) As cDataBase
    Set Param = Me
    If Cmd Is Nothing Then Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = Conn
    Cmd.CommandText = pvRawSqlString
    Cmd.CommandType = adCmdText
    
    Dim p As ADODB.Parameter
    Set p = Cmd.CreateParameter(ParamName, ParamType, adParamInput, , ParamValue)
    Cmd.Parameters.Append p
End Function

'执行参数化查询
Public Function ExecParam(Optional RecordsAffected) As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Cmd Is Nothing Then LastErr = "no parameters set": Exit Function
    
    Set Rs = Cmd.Execute(RecordsAffected)
    ExecParam = True
    
    '清理参数
    Do While Cmd.Parameters.Count > 0
        Cmd.Parameters.Delete 0
    Loop
    Set Cmd = Nothing
    Exit Function
eee:
    LastErr = "ExecParam Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    If Not Cmd Is Nothing Then
        Do While Cmd.Parameters.Count > 0
            Cmd.Parameters.Delete 0
        Loop
        Set Cmd = Nothing
    End If
End Function

'查询参数化
Public Function QueryParam(Optional CurType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockOptimistic) As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Cmd Is Nothing Then LastErr = "no parameters set": Exit Function
    If Rs.State <> adStateClosed Then Rs.Close
    
    Cmd.CommandType = adCmdText
    Set Rs = Cmd.Execute(, , adCmdText)
    QueryParam = True
    
    '清理参数
    Do While Cmd.Parameters.Count > 0
        Cmd.Parameters.Delete 0
    Loop
    Set Cmd = Nothing
    Exit Function
eee:
    LastErr = "QueryParam Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    If Not Cmd Is Nothing Then
        Do While Cmd.Parameters.Count > 0
            Cmd.Parameters.Delete 0
        Loop
        Set Cmd = Nothing
    End If
End Function

'执行数据库操作语句
Public Function Exec(Optional RecordsAffected, Optional Options As Long = -1) As Boolean
    On Error GoTo eee
    IsPage = False
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvRawSqlString = "" Then LastErr = "no have sql string": Exit Function
    
    '异步执行
    If IsAsync Then
        Options = Options Or adExecuteAsync
        IsAsync = False
    End If
    
    '设置要执行的命令文本
    Set Rs = Conn.Execute(pvRawSqlString, RecordsAffected, Options)
    '返回记录集d对象
    Exec = True
    Exit Function
eee:
    LastErr = "Exec Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'执行数据库查询语句
Public Function Fetch(Optional CurType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockOptimistic, Optional Options As Long = -1) As Boolean
    Fetch = Query(CurType, LockType, Options)
    If Fetch = True Then
        Set Rows = ToolsList.RsToCollection(Rs)
        If Rows.Count > 0 Then Set Row = Rows(1)
    End If
End Function
Public Function Query(Optional CurType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockOptimistic, Optional Options As Long = -1) As Boolean
    On Error GoTo eee
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Rs.State <> adStateClosed Then Rs.Close
    
    '处理分页
    Dim SqlString As String
    SqlString = pvRawSqlString
    If IsPage Then
        Select Case DatabaseType
        Case MsSql
            'SQL Server 2012+ 使用 OFFSET FETCH
            If InStr(1, UCase(SqlString), "ORDER BY") = 0 Then
                SqlString = SqlString & " ORDER BY (SELECT NULL)"
            End If
            SqlString = SqlString & " OFFSET " & pvPageOffset & " ROWS FETCH NEXT " & pvPageLimit & " ROWS ONLY"
        Case Mysql
            'MySQL 使用 LIMIT
            SqlString = SqlString & " LIMIT " & pvPageLimit & " OFFSET " & pvPageOffset
        Case Access
            'Access 不支持标准分页，使用TOP和子查询
            '这里简化处理，使用TOP
            SqlString = "SELECT TOP " & (pvPageOffset + pvPageLimit) & " * FROM (" & SqlString & ") AS T"
        Case Else
            '其他数据库类型暂不支持分页
        End Select
        IsPage = False
    End If
    
    '打开记录集
    Rs.Open SqlString, , CurType, LockType, Options
    Query = True
    Exit Function
eee:
    LastErr = "Query Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

Private Sub Class_Initialize()
    Set Conn = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    Call Disconnect
End Sub

'获取最后插入的ID（自增主键）
Public Function LastInsertId() As Variant
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    
    Dim SqlString As String
    Select Case DatabaseType
    Case MsSql
        SqlString = "SELECT SCOPE_IDENTITY() AS LastId"
    Case Mysql
        SqlString = "SELECT LAST_INSERT_ID() AS LastId"
    Case Access
        SqlString = "SELECT @@IDENTITY AS LastId"
    Case Else
        LastErr = "Database type not support LastInsertId"
        Exit Function
    End Select
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        LastInsertId = rsTemp("LastId")
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "LastInsertId Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'统计记录数
Public Function Count(Optional ByVal TableName As String = "") As Long
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    
    Dim SqlString As String
    If TableName <> "" Then
        SqlString = "SELECT COUNT(*) AS cnt FROM " & TableName
    Else
        '从现有SQL中提取表名并统计
        SqlString = "SELECT COUNT(*) AS cnt FROM (" & pvRawSqlString & ") AS T"
    End If
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        Count = rsTemp("cnt")
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "Count Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'判断表是否存在
Public Function TableExists(ByVal TableName As String) As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    
    Dim SqlString As String
    Select Case DatabaseType
    Case MsSql
        SqlString = "SELECT COUNT(*) AS cnt FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '" & TableName & "'"
    Case Mysql
        SqlString = "SELECT COUNT(*) AS cnt FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = '" & TableName & "'"
    Case Access
        SqlString = "SELECT COUNT(*) AS cnt FROM MSysObjects WHERE Name = '" & TableName & "' AND Type = 1"
    Case Else
        LastErr = "Database type not support TableExists"
        Exit Function
    End Select
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        TableExists = (rsTemp("cnt") > 0)
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "TableExists Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'获取表的字段列表
Public Function GetTableFields(ByVal TableName As String) As Collection
    On Error GoTo eee
    Set GetTableFields = New Collection
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    
    Dim SqlString As String
    Select Case DatabaseType
    Case MsSql
        SqlString = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & TableName & "' ORDER BY ORDINAL_POSITION"
    Case Mysql
        SqlString = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = '" & TableName & "' ORDER BY ORDINAL_POSITION"
    Case Access
        SqlString = "SELECT Name FROM MSysObjects WHERE Type = 4 AND ParentId = (SELECT Id FROM MSysObjects WHERE Name = '" & TableName & "' AND Type = 1)"
    Case Else
        LastErr = "Database type not support GetTableFields"
        Exit Function
    End Select
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    Do Until rsTemp.EOF
        GetTableFields.Add rsTemp(0)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "GetTableFields Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'批量插入数据（使用事务）
Public Function BatchInsert(ByVal TableName As String, ByVal Data As Collection) As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Data.Count = 0 Then BatchInsert = True: Exit Function
    
    '开始事务
    Call TransBegin
    
    Dim i As Long
    Dim Row As Scripting.Dictionary
    Dim Fields As String, Values As String
    Dim SqlString As String
    
    For i = 1 To Data.Count
        Set Row = Data(i)
        Fields = ""
        Values = ""
        
        Dim Key As Variant
        For Each Key In Row.Keys
            If Fields <> "" Then Fields = Fields & ", ": Values = Values & ", "
            Fields = Fields & Key
            '处理值，转义单引号
            Dim Val As String
            Val = CStr(Row(Key))
            Val = Replace(Val, "'", "''")
            Values = Values & "'" & Val & "'"
        Next
        
        SqlString = "INSERT INTO " & TableName & " (" & Fields & ") VALUES (" & Values & ")"
        Call Conn.Execute(SqlString)
    Next
    
    '提交事务
    Call TransCommit
    BatchInsert = True
    Exit Function
eee:
    '回滚事务
    If IsTrans Then Call TransRollback
    LastErr = "BatchInsert Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'获取数据库版本信息
Public Function GetVersion() As String
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    GetVersion = Conn.Version
    Exit Function
eee:
    LastErr = "GetVersion Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'获取数据库名称列表
Public Function GetDatabases() As Collection
    On Error GoTo eee
    Set GetDatabases = New Collection
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    
    Dim SqlString As String
    Select Case DatabaseType
    Case MsSql
        SqlString = "SELECT name FROM sys.databases ORDER BY name"
    Case Mysql
        SqlString = "SHOW DATABASES"
    Case Else
        LastErr = "Database type not support GetDatabases"
        Exit Function
    End Select
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    Do Until rsTemp.EOF
        GetDatabases.Add rsTemp(0)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "GetDatabases Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'获取表名列表
Public Function GetTables(Optional ByVal DatabaseName As String = "") As Collection
    On Error GoTo eee
    Set GetTables = New Collection
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    
    Dim SqlString As String
    Select Case DatabaseType
    Case MsSql
        If DatabaseName = "" Then
            SqlString = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
        Else
            SqlString = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG = '" & DatabaseName & "' AND TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
        End If
    Case Mysql
        If DatabaseName = "" Then
            SqlString = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = DATABASE() ORDER BY TABLE_NAME"
        Else
            SqlString = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '" & DatabaseName & "' ORDER BY TABLE_NAME"
        End If
    Case Access
        SqlString = "SELECT Name FROM MSysObjects WHERE Type = 1 AND Flags = 0 ORDER BY Name"
    Case Else
        LastErr = "Database type not support GetTables"
        Exit Function
    End Select
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    Do Until rsTemp.EOF
        GetTables.Add rsTemp(0)
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "GetTables Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'转义SQL字符串（防止SQL注入）
Public Function Escape(ByVal Str As String) As String
    Escape = Replace(Str, "'", "''")
End Function

'检查连接状态并尝试重连
Public Function CheckConnection() As Boolean
    On Error GoTo eee
    CheckConnection = False
    If Conn Is Nothing Then Exit Function
    If Conn.State = adStateOpen Then
        CheckConnection = True
    Else
        '尝试重新连接
        If IsConnect Then
            IsConnect = False
            Conn.Open
            If Conn.State = adStateOpen Then
                IsConnect = True
                CheckConnection = True
            End If
        End If
    End If
    Exit Function
eee:
    LastErr = "CheckConnection Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

Private Sub Conn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    'async sql callback
    RaiseEvent AsyncExecuteComplete(RecordsAffected, pError, adStatus, pCommand, pRecordset, pConnection)
End Sub


