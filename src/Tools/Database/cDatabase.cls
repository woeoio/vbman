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
Public Rows As New ADODB.Recordset
Public Row As New Scripting.Dictionary
'Connection inst 20241011 今天发现conn居然有异步执行sql的能力, 后期挖掘下, 实现和前端异步交互，还有 stream 模式
Public WithEvents Conn As ADODB.Connection
Attribute Conn.VB_VarHelpID = -1
'cDatabase pool, 多个数据库对象池，主要用于动态根据字符串key使用数据库对象的场景
Dim Connections As New Scripting.Dictionary
'
Private isTrans As Boolean
Private pvRawSqlString As String

Public Enum enumDbType
    Access
    mysql
    MsSql
    CSV
End Enum

Public Function ConnInst(ByVal InstName As String) As cDataBase
    If Connections.Exists(InstName) = False Then Connections.Add InstName, New cDataBase
    Set ConnInst = Connections(InstName)
End Function
Public Sub ConnInstRemove(ByVal InstName As String)
    If Connections.Exists(InstName) = False Then Exit Sub
    Set Connections(InstName) = Nothing
    Connections.Remove InstName
End Sub

Public Function Connect( _
    ByVal dbType As enumDbType, _
    Optional ByVal DbAddress As String = "127.0.0.1,1433", _
    Optional ByVal Username As String = "sa", _
    Optional ByVal Password As String = "Sa123456", _
    Optional ByVal DefaultDataBase As String = "master") As Boolean
    
    On Error GoTo eee
    
    '如果连接标记为真，则返回。否则会出错
    If IsConnect = True Then
        '返回成功
        Connect = True
        Exit Function
    End If
    
    '设置连接字符串ConnectionString属性
    Select Case dbType
    Case Access
        Conn.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & DbAddress & ";"
    Case CSV
        Conn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" & DbAddress
    Case MsSql
        Conn.ConnectionString = "Driver={SQL Server};Server=" & DbAddress & ";Uid=" & Username & ";pwd=" & Password & ";Database=" & DefaultDataBase & ";"
    Case mysql
        Dim Svr: Svr = Split(Server, ",")
        Dim Address$: Address = Svr(0)
        Dim Port$: Port = "3306"
        If UBound(Svr) > 0 Then Port = Svr(1)
        Conn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" & DbAddress & ";Port=" & Port & ";Uid=" & Username & ";pwd=" & Password & ";Database=" & DefaultDataBase & ";OPTION=3"
    End Select
    '打开到数据库的连接
    Conn.Open
    '判断连接的状态
    If Conn.State <> adStateOpen Then
        '如果连接不成功，则显示提示信息，退出程序
        Exit Function
    End If
    '设置Exec_Cmd的ActiveConnection属性，指定与其关联的数据库连接
    Set Cmd.ActiveConnection = Conn
    '设置rst的ActiveConnection属性，指定与其关联的数据库连接
    Set Rs.ActiveConnection = Conn
    '根据链接类型设置一些必要的查询缺省值，减少外部查询代码的累赘
    Select Case dbType
    Case Access
        'todo
    Case CSV
        'todo
    Case MsSql
        'todo
    Case mysql
        'todo
    End Select
    
    Conn.CursorLocation = adUseClient
    '设置连接标记，表示已经连接到数据库
    IsConnect = True
    Connect = True
    Exit Function
eee:
    LastErr = "Connect Err[" & Err.Number & "]" & Err.Description & "{" & DbAddress & "}{" & dbType & "}{" & Username & "}{" & DefaultDataBase & "}"
    LastErrNumber = Err.Number
    LastErrDescription = Err.Description
End Function

'断开与数据库的连接
Public Function Disconnect() As Boolean
    On Error GoTo eee
    
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then
        Exit Function
    End If
    '关闭连接
    Conn.Close
    If Rs.State <> adStateClosed Then Rs.Close
    '释放conn
    Set Conn = Nothing
    Set Rs = Nothing
    Set Cmd = Nothing
    
    '设置连接标记，表示已经断开与数据库的连接
    IsConnect = False
    Disconnect = True
    
    Exit Function
eee:
    LastErr = "Disconnect Err[" & Err.Number & "]" & Err.Description
    LastErrNumber = Err.Number
    LastErrDescription = Err.Description
End Function

Public Function Sql(ByVal RawSqlString As String) As cDataBase
    Set Sql = Me
    pvRawSqlString = RawSqlString
End Function

'执行数据库操作语句
Public Function ExecAsync(Optional RecordsAffected, Optional Options As Long = -1) As Boolean
    ExecAsync = Exec(RecordsAffected, adAsyncExecute)
End Function
Public Function Exec(Optional RecordsAffected, Optional Options As Long = -1) As Boolean
    On Error GoTo eee
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If RawSqlString = "" Then LastErr = "no have sql string": Exit Function
    '设置要执行的命令文本
    Set Rs = Conn.Execute(RawSqlString, RecordsAffected, Options)
    '返回记录集d对象
    Exec = True
    Exit Function
eee:
    LastErr = "Exec Err[" & Err.Number & "]" & Err.Description & "{" & RawSqlString & "}"
    LastErrNumber = Err.Number
    LastErrDescription = Err.Description
End Function

'执行数据库查询语句
Public Function Query(Optional CurType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockOptimistic, Optional Options As Long = -1) As Boolean
    On Error GoTo eee
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Rs.State <> adStateClosed Then Rs.Close
    '打开记录集
    Rs.Open RawSqlString, , CurType, LockType, Options
    Query = True
    Exit Function
eee:
    LastErr = "Query Err[" & Err.Number & "]" & Err.Description & "{" & RawSqlString & "}"
    LastErrNumber = Err.Number
    LastErrDescription = Err.Description
End Function

Private Sub Class_Initialize()
    Set Conn = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    Call Disconnect
End Sub

Private Sub Conn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    'async sql callback
    RaiseEvent AsyncExecuteComplete(RecordsAffected, pError, adStatus, pCommand, pRecordset, pConnection)
End Sub


