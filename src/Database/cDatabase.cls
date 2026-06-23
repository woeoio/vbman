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
Public Rows As cCollection
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
Private pvTableName As String
Private pvWhereCondition As String
Private pvWhereParams As New Collection
Private pvFields As New Scripting.Dictionary  ' Field() 链式调用缓存 {fieldName: value}
Private pvPageNum As Long
Private pvPageLimit As Long
Private pvPageOffset As Long
'━━━ 链式查询构建器新增内部状态 ━━━
Private pvSelectFields As String          ' Columns() 设置的查询列列表
Private pvIsDistinct As Boolean           ' Distinct 标记
Private pvOrderByClause As String         ' OrderBy 排序子句
Private pvGroupByClause As String         ' GroupBy 分组子句
Private pvHavingClause As String          ' Having 分组过滤子句
Private pvJoinClauses As String           ' Join 连接子句（可叠加多个）
Private pvLimitValue As Long              ' Limit 限制条数
Private pvOffsetValue As Long             ' Offset 偏移量
Private pvOrWhereCondition As String      ' OrWhere OR条件
Private pvWhereInClauses As String        ' WhereIn 子句片段（可叠加）
Private pvWhereBetweenClauses As String   ' WhereBetween 子句片段
Private pvWhereNullClauses As String      ' WhereNull/WhereNotNull 子句片段
Public IsCloneInstance As Boolean                                               '标记是否为连接池克隆实例
'Public CustomConnectionString As String

Dim IsAsync As Boolean
Dim Cmd As ADODB.Command

Public Enum enumDbType
    Access = 1
    Mysql = 2
    MsSql = 3
    Csv = 4
    Custom = 5
End Enum
Public DatabaseType As enumDbType

'ADO DataTypeEnum 常用枚举，方便外部引用 VBMAN 后直接使用
'如 VBMAN.cDataBase.adVarWChar 或直接 adVarWChar
Public Enum AdoDataType
    adEmpty = 0
    adTinyInt = 16
    adSmallInt = 2
    adInteger = 3
    adBigInt = 20
    adUnsignedTinyInt = 17
    adUnsignedSmallInt = 18
    adUnsignedInt = 19
    adUnsignedBigInt = 21
    adSingle = 4
    adDouble = 5
    adCurrency = 6
    adDecimal = 14
    adBoolean = 11
    adDate = 7
    adDBDate = 133
    adDBTime = 134
    adDBTimeStamp = 135
    adVarChar = 200
    adLongVarChar = 201
    adVarWChar = 202
    adLongVarWChar = 203
    adBinary = 128
    adVarBinary = 204
    adLongVarBinary = 205
    adGUID = 72
End Enum

Public Property Get Async() As cDataBase
    IsAsync = True
    Set Async = Me
End Property

Public Function ConnInst(ByVal InstName As String, Optional ByVal IsCloneMasterConnection As Boolean = True) As cDataBase
    If Connections.Exists(InstName) = False Then
        Connections.Add InstName, New cDataBase
        Set ConnInst = Connections(InstName)
        If IsCloneMasterConnection = True Then
            '直接使用主连接对象，不需要重新连接
            Set ConnInst.Conn = Me.Conn
            ConnInst.IsConnect = Me.IsConnect
            ConnInst.DatabaseType = Me.DatabaseType
            '标记为克隆实例，断开时不关闭共享连接
            ConnInst.IsCloneInstance = True
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

Public Function Page(Optional Num As Long = 1, Optional Limit As Long = 10) As cDataBase
    Set Page = Me
    IsPage = True
    pvPageNum = Num
    pvPageLimit = Limit
    If pvPageNum < 1 Then pvPageNum = 1
    If pvPageLimit < 1 Then pvPageLimit = 10
    pvPageOffset = (pvPageNum - 1) * pvPageLimit
End Function

Public Function Connect( _
    ByVal DbType As enumDbType, _
    Optional ByVal DbAddress As String = "127.0.0.1,1433", _
    Optional ByVal username As String = "sa", _
    Optional ByVal Password As String = "Sa123456", _
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
        Case Custom
            Conn.ConnectionString = DbAddress
        Case Access
            DbAddress = ToolsFso.AutoCompleteFullPath(DbAddress, True)
            Conn.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & DbAddress & ";"
        Case Csv
            DbAddress = ToolsFso.AutoCompleteFullPath(DbAddress, True)
            Conn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" & DbAddress
        Case MsSql
            Conn.ConnectionString = "Driver={SQL Server};Server=" & DbAddress & ";Uid=" & username & ";pwd=" & Password & ";Database=" & DefaultDataBase & ";"
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
            Conn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=" & DbAddress & ";Uid=" & username & ";pwd=" & Password & ";Database=" & DefaultDataBase & ";"
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
Public Function DisConnect() As Boolean
    On Error GoTo eee

    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then
        Exit Function
    End If
    '如果正在事务中，先回滚
    If IsTrans Then Call TransRollback

    '如果是克隆实例，只清空本地引用，不关闭共享连接
    If IsCloneInstance Then
        If Rs.State <> adStateClosed Then Rs.Close
        Set Rs = Nothing
        Set Conn = Nothing
        IsConnect = False
        DisConnect = True
        Exit Function
    End If

    '关闭连接
    Conn.Close
    If Rs.State <> adStateClosed Then Rs.Close
    '释放conn
    Set Conn = Nothing
    Set Rs = Nothing

    '设置连接标记，表示已经断开与数据库的连接
    IsConnect = False
    DisConnect = True

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

'设置当前操作的表名，用于后续 AddRow/RowUpdate 等方法
Public Function Table(ByVal TableName As String) As cDataBase
    Set Table = Me
    pvTableName = TableName
End Function

'设置 WHERE 条件，支持 ? 占位符 + ParamArray 参数值
'用法: db.Where("serial_no=? AND channel=?", serialNo, 1)
Public Function Where(ByVal Condition As String, ParamArray Params() As Variant) As cDataBase
    Set Where = Me
    ' 替换 ? 占位符为参数值
    Dim i As Long
    pvWhereCondition = Condition
    For i = 0 To UBound(Params)
        ' 用 ADO 参数化方式存储参数值
        pvWhereParams.Add CStr(Params(i))
        ' 替换第一个 ? 为值（用单引号转义）
        Dim pos As Long: pos = InStr(pvWhereCondition, "?")
        If pos > 0 Then
            Dim val As String: val = Replace(CStr(Params(i)), "'", "''")
            pvWhereCondition = Left$(pvWhereCondition, pos - 1) & "'" & val & "'" & Mid$(pvWhereCondition, pos + 1)
        End If
    Next
End Function

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'  链式查询构建器 - Columns / Distinct / OrderBy / GroupBy / Having
'  / Join / LeftJoin / RightJoin / OrWhere / WhereIn / WhereNotIn
'  / WhereBetween / WhereLike / WhereNull / WhereNotNull / Limit / Offset
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

'指定查询列（默认 SELECT *，调用后改为指定列）
'用法: db.Table("t").Columns("id,name,age").RowRead
'     db.Table("t").Columns("id", "name", "age").RowRead  (多参数形式)
'注意: 方法名使用 Columns 而非 Select，避免与 VB6 保留字 Select Case 冲突
Public Function Columns(ParamArray Cols() As Variant) As cDataBase
    Set Columns = Me
    If UBound(Cols) < 0 Then Exit Function
    '如果只传一个参数且包含逗号，视为完整列列表字符串
    If UBound(Cols) = 0 And InStr(CStr(Cols(0)), ",") > 0 Then
        pvSelectFields = CStr(Cols(0))
    Else
        Dim i As Long, s As String
        For i = 0 To UBound(Cols)
            If s <> "" Then s = s & ", "
            s = s & CStr(Cols(i))
        Next
        pvSelectFields = s
    End If
End Function

'去重查询 SELECT DISTINCT
'用法: db.Table("t").Distinct.Columns("category").RowRead
Public Function Distinct() As cDataBase
    Set Distinct = Me
    pvIsDistinct = True
End Function

'排序 ASC（可多次调用叠加多个排序字段）
'用法: db.Table("t").OrderBy("age").OrderBy("name").RowRead
Public Function OrderBy(ByVal FieldName As String) As cDataBase
    Set OrderBy = Me
    If pvOrderByClause <> "" Then pvOrderByClause = pvOrderByClause & ", "
    pvOrderByClause = pvOrderByClause & FieldName & " ASC"
End Function

'排序 DESC（可多次调用叠加多个排序字段）
'用法: db.Table("t").OrderByDesc("age").OrderByDesc("name").RowRead
Public Function OrderByDesc(ByVal FieldName As String) As cDataBase
    Set OrderByDesc = Me
    If pvOrderByClause <> "" Then pvOrderByClause = pvOrderByClause & ", "
    pvOrderByClause = pvOrderByClause & FieldName & " DESC"
End Function

'自定义排序（指定字段和方向）
'用法: db.Table("t").OrderByRaw("age DESC, name ASC").RowRead
Public Function OrderByRaw(ByVal RawClause As String) As cDataBase
    Set OrderByRaw = Me
    If pvOrderByClause <> "" Then pvOrderByClause = pvOrderByClause & ", "
    pvOrderByClause = pvOrderByClause & RawClause
End Function

'分组 GROUP BY（可多次调用叠加多个分组字段）
'用法: db.Table("t").GroupBy("category").Having("COUNT(*)>5").RowRead
Public Function GroupBy(ByVal FieldName As String) As cDataBase
    Set GroupBy = Me
    If pvGroupByClause <> "" Then pvGroupByClause = pvGroupByClause & ", "
    pvGroupByClause = pvGroupByClause & FieldName
End Function

'分组过滤 HAVING（配合 GroupBy 使用）
'用法: db.Table("t").GroupBy("category").Having("COUNT(*)>5").RowRead
Public Function Having(ByVal Condition As String) As cDataBase
    Set Having = Me
    pvHavingClause = Condition
End Function

'INNER JOIN 表连接
'用法: db.Table("orders").Join("users", "orders.user_id = users.id").RowRead
Public Function Join(ByVal TableName As String, ByVal OnCondition As String) As cDataBase
    Set Join = Me
    pvJoinClauses = pvJoinClauses & " INNER JOIN " & TableName & " ON " & OnCondition
End Function

'LEFT JOIN 左连接
'用法: db.Table("orders").LeftJoin("users", "orders.user_id = users.id").RowRead
Public Function LeftJoin(ByVal TableName As String, ByVal OnCondition As String) As cDataBase
    Set LeftJoin = Me
    pvJoinClauses = pvJoinClauses & " LEFT JOIN " & TableName & " ON " & OnCondition
End Function

'RIGHT JOIN 右连接
'用法: db.Table("orders").RightJoin("users", "orders.user_id = users.id").RowRead
Public Function RightJoin(ByVal TableName As String, ByVal OnCondition As String) As cDataBase
    Set RightJoin = Me
    pvJoinClauses = pvJoinClauses & " RIGHT JOIN " & TableName & " ON " & OnCondition
End Function

'OR WHERE 条件（与 Where 并列，用 OR 连接）
'用法: db.Table("t").Where("age>18").OrWhere("status='active'").RowRead
Public Function OrWhere(ByVal Condition As String, ParamArray Params() As Variant) As cDataBase
    Set OrWhere = Me
    Dim s As String: s = Condition
    Dim i As Long
    For i = 0 To UBound(Params)
        Dim pos As Long: pos = InStr(s, "?")
        If pos > 0 Then
            Dim val As String: val = Replace(CStr(Params(i)), "'", "''")
            s = Left$(s, pos - 1) & "'" & val & "'" & Mid$(s, pos + 1)
        End If
    Next
    If pvOrWhereCondition <> "" Then pvOrWhereCondition = pvOrWhereCondition & " OR "
    pvOrWhereCondition = pvOrWhereCondition & s
End Function

'WHERE IN 条件（字段值在指定列表中）
'用法: db.Table("t").WhereIn("status", "active,pending,closed").RowRead
'     db.Table("t").WhereIn("id", 1, 2, 3).RowRead  (多参数形式)
Public Function WhereIn(ByVal FieldName As String, ParamArray Values() As Variant) As cDataBase
    Set WhereIn = Me
    Dim s As String
    '如果只传一个参数且包含逗号，视为完整列表字符串
    If UBound(Values) = 0 And InStr(CStr(Values(0)), ",") > 0 Then
        Dim items() As String: items = Split(CStr(Values(0)), ",")
        Dim j As Long
        For j = 0 To UBound(items)
            If s <> "" Then s = s & ", "
            s = s & "'" & Trim(Replace(items(j), "'", "''")) & "'"
        Next
    Else
        Dim i As Long
        For i = 0 To UBound(Values)
            If s <> "" Then s = s & ", "
            s = s & "'" & Replace(CStr(Values(i)), "'", "''") & "'"
        Next
    End If
    If pvWhereInClauses <> "" Then pvWhereInClauses = pvWhereInClauses & " AND "
    pvWhereInClauses = pvWhereInClauses & FieldName & " IN (" & s & ")"
End Function

'WHERE NOT IN 条件（字段值不在指定列表中）
Public Function WhereNotIn(ByVal FieldName As String, ParamArray Values() As Variant) As cDataBase
    Set WhereNotIn = Me
    Dim s As String
    If UBound(Values) = 0 And InStr(CStr(Values(0)), ",") > 0 Then
        Dim items() As String: items = Split(CStr(Values(0)), ",")
        Dim j As Long
        For j = 0 To UBound(items)
            If s <> "" Then s = s & ", "
            s = s & "'" & Trim(Replace(items(j), "'", "''")) & "'"
        Next
    Else
        Dim i As Long
        For i = 0 To UBound(Values)
            If s <> "" Then s = s & ", "
            s = s & "'" & Replace(CStr(Values(i)), "'", "''") & "'"
        Next
    End If
    If pvWhereInClauses <> "" Then pvWhereInClauses = pvWhereInClauses & " AND "
    pvWhereInClauses = pvWhereInClauses & FieldName & " NOT IN (" & s & ")"
End Function

'WHERE BETWEEN 条件（字段值在两个值之间）
'用法: db.Table("t").WhereBetween("age", 18, 30).RowRead
Public Function WhereBetween(ByVal FieldName As String, ByVal Value1 As Variant, ByVal Value2 As Variant) As cDataBase
    Set WhereBetween = Me
    If pvWhereBetweenClauses <> "" Then pvWhereBetweenClauses = pvWhereBetweenClauses & " AND "
    pvWhereBetweenClauses = pvWhereBetweenClauses & FieldName & " BETWEEN '" & Replace(CStr(Value1), "'", "''") & "' AND '" & Replace(CStr(Value2), "'", "''") & "'"
End Function

'WHERE NOT BETWEEN 条件
Public Function WhereNotBetween(ByVal FieldName As String, ByVal Value1 As Variant, ByVal Value2 As Variant) As cDataBase
    Set WhereNotBetween = Me
    If pvWhereBetweenClauses <> "" Then pvWhereBetweenClauses = pvWhereBetweenClauses & " AND "
    pvWhereBetweenClauses = pvWhereBetweenClauses & FieldName & " NOT BETWEEN '" & Replace(CStr(Value1), "'", "''") & "' AND '" & Replace(CStr(Value2), "'", "''") & "'"
End Function

'WHERE LIKE 条件（模糊匹配）
'用法: db.Table("t").WhereLike("name", "%张%").RowRead
Public Function WhereLike(ByVal FieldName As String, ByVal Pattern As String) As cDataBase
    Set WhereLike = Me
    '复用 pvWhereCondition 累加
    If pvWhereCondition <> "" Then pvWhereCondition = pvWhereCondition & " AND "
    pvWhereCondition = pvWhereCondition & FieldName & " LIKE '" & Replace(Pattern, "'", "''") & "'"
End Function

'WHERE NOT LIKE 条件
Public Function WhereNotLike(ByVal FieldName As String, ByVal Pattern As String) As cDataBase
    Set WhereNotLike = Me
    If pvWhereCondition <> "" Then pvWhereCondition = pvWhereCondition & " AND "
    pvWhereCondition = pvWhereCondition & FieldName & " NOT LIKE '" & Replace(Pattern, "'", "''") & "'"
End Function

'WHERE IS NULL 条件
'用法: db.Table("t").WhereNull("deleted_at").RowRead
Public Function WhereNull(ByVal FieldName As String) As cDataBase
    Set WhereNull = Me
    If pvWhereNullClauses <> "" Then pvWhereNullClauses = pvWhereNullClauses & " AND "
    pvWhereNullClauses = pvWhereNullClauses & FieldName & " IS NULL"
End Function

'WHERE IS NOT NULL 条件
'用法: db.Table("t").WhereNotNull("email").RowRead
Public Function WhereNotNull(ByVal FieldName As String) As cDataBase
    Set WhereNotNull = Me
    If pvWhereNullClauses <> "" Then pvWhereNullClauses = pvWhereNullClauses & " AND "
    pvWhereNullClauses = pvWhereNullClauses & FieldName & " IS NOT NULL"
End Function

'限制返回条数 LIMIT（非分页场景，简单截断）
'用法: db.Table("t").OrderByDesc("id").Limit(10).RowRead
Public Function Limit(ByVal N As Long) As cDataBase
    Set Limit = Me
    pvLimitValue = N
End Function

'偏移量 OFFSET（配合 Limit 使用，跳过前 N 条）
'用法: db.Table("t").OrderByDesc("id").Limit(10).Offset(20).RowRead
Public Function Offset(ByVal N As Long) As cDataBase
    Set Offset = Me
    pvOffsetValue = N
End Function

'重置链式查询构建器所有内部状态（每次终结操作后自动调用）
Private Sub ResetBuilder()
    pvTableName = ""
    pvWhereCondition = ""
    pvOrWhereCondition = ""
    pvWhereInClauses = ""
    pvWhereBetweenClauses = ""
    pvWhereNullClauses = ""
    pvSelectFields = ""
    pvIsDistinct = False
    pvOrderByClause = ""
    pvGroupByClause = ""
    pvHavingClause = ""
    pvJoinClauses = ""
    pvLimitValue = 0
    pvOffsetValue = 0
    pvRawSqlString = ""
    IsPage = False
    pvPageNum = 0
    pvPageLimit = 0
    pvPageOffset = 0
End Sub

'拼接 WHERE 子句（合并 Where/WhereIn/WhereBetween/WhereNull/OrWhere）
'由 RowUpdate/RowDelete 等需要 WHERE 的终结方法内部调用
Private Function BuildWhereClause() As String
    Dim whereParts As String
    If pvWhereCondition <> "" Then whereParts = pvWhereCondition
    If pvWhereInClauses <> "" Then
        If whereParts <> "" Then whereParts = whereParts & " AND "
        whereParts = whereParts & pvWhereInClauses
    End If
    If pvWhereBetweenClauses <> "" Then
        If whereParts <> "" Then whereParts = whereParts & " AND "
        whereParts = whereParts & pvWhereBetweenClauses
    End If
    If pvWhereNullClauses <> "" Then
        If whereParts <> "" Then whereParts = whereParts & " AND "
        whereParts = whereParts & pvWhereNullClauses
    End If
    If pvOrWhereCondition <> "" Then
        If whereParts <> "" Then
            whereParts = whereParts & " OR (" & pvOrWhereCondition & ")"
        Else
            whereParts = pvOrWhereCondition
        End If
    End If
    BuildWhereClause = whereParts
End Function

'拼接完整 SELECT SQL（由 RowRead/Fetch 等终结方法内部调用）
'构建逻辑: SELECT [DISTINCT] {fields} FROM {table} {joins} WHERE {conditions} GROUP BY {group} HAVING {having} ORDER BY {order} LIMIT/OFFSET
Private Function BuildSelectSql() As String
    Dim SqlString As String
    
    'SELECT 部分
    If pvIsDistinct Then
        SqlString = "SELECT DISTINCT "
    Else
        SqlString = "SELECT "
    End If
    If pvSelectFields <> "" Then
        SqlString = SqlString & pvSelectFields
    Else
        SqlString = SqlString & "*"
    End If
    
    'FROM 部分
    SqlString = SqlString & " FROM " & pvTableName
    
    'JOIN 部分
    If pvJoinClauses <> "" Then
        SqlString = SqlString & pvJoinClauses
    End If
    
    'WHERE 部分 - 合并所有条件子句
    Dim whereParts As String
    If pvWhereCondition <> "" Then whereParts = pvWhereCondition
    If pvWhereInClauses <> "" Then
        If whereParts <> "" Then whereParts = whereParts & " AND "
        whereParts = whereParts & pvWhereInClauses
    End If
    If pvWhereBetweenClauses <> "" Then
        If whereParts <> "" Then whereParts = whereParts & " AND "
        whereParts = whereParts & pvWhereBetweenClauses
    End If
    If pvWhereNullClauses <> "" Then
        If whereParts <> "" Then whereParts = whereParts & " AND "
        whereParts = whereParts & pvWhereNullClauses
    End If
    If pvOrWhereCondition <> "" Then
        If whereParts <> "" Then
            whereParts = whereParts & " OR (" & pvOrWhereCondition & ")"
        Else
            whereParts = pvOrWhereCondition
        End If
    End If
    If whereParts <> "" Then
        SqlString = SqlString & " WHERE " & whereParts
    End If
    
    'GROUP BY 部分
    If pvGroupByClause <> "" Then
        SqlString = SqlString & " GROUP BY " & pvGroupByClause
    End If
    
    'HAVING 部分
    If pvHavingClause <> "" Then
        SqlString = SqlString & " HAVING " & pvHavingClause
    End If
    
    'ORDER BY 部分
    If pvOrderByClause <> "" Then
        SqlString = SqlString & " ORDER BY " & pvOrderByClause
    End If
    
    'LIMIT / OFFSET 部分（根据数据库类型适配语法）
    If pvLimitValue > 0 Or pvOffsetValue > 0 Then
        Select Case DatabaseType
        Case Mysql
            If pvOffsetValue > 0 Then
                SqlString = SqlString & " LIMIT " & pvOffsetValue & ", " & pvLimitValue
            ElseIf pvLimitValue > 0 Then
                SqlString = SqlString & " LIMIT " & pvLimitValue
            End If
        Case MsSql
            'SQL Server: OFFSET {n} ROWS FETCH NEXT {m} ROWS ONLY (需 ORDER BY)
            If pvOrderByClause = "" Then
                '没有 ORDER BY 时用 id 排序兜底（SQL Server OFFSET 必须有 ORDER BY）
                SqlString = SqlString & " ORDER BY (SELECT 0)"
            End If
            If pvOffsetValue > 0 Then
                SqlString = SqlString & " OFFSET " & pvOffsetValue & " ROWS FETCH NEXT " & pvLimitValue & " ROWS ONLY"
            ElseIf pvLimitValue > 0 Then
                SqlString = SqlString & " OFFSET 0 ROWS FETCH NEXT " & pvLimitValue & " ROWS ONLY"
            End If
        Case Access
            'Access 不支持 LIMIT/OFFSET，使用 ADO 分页机制替代
            If pvLimitValue > 0 Then
                pvPageLimit = pvLimitValue
                pvPageNum = (pvOffsetValue \ pvLimitValue) + 1
                IsPage = True
            End If
        Case Else
            '其他数据库类型，使用 LIMIT 语法
            If pvLimitValue > 0 Then
                SqlString = SqlString & " LIMIT " & pvLimitValue
            End If
        End Select
    End If
    
    BuildSelectSql = SqlString
End Function

'参数化查询 - 防止SQL注入
Public Function Param(ByVal ParamName As String, ByVal ParamValue As Variant, Optional ByVal ParamType As DataTypeEnum = adVarChar) As cDataBase
    Set Param = Me
    If Cmd Is Nothing Then Set Cmd = New ADODB.Command
    Cmd.ActiveConnection = Conn
    Cmd.CommandText = pvRawSqlString
    Cmd.CommandType = adCmdText
    
    Dim P As ADODB.Parameter
    Dim lSize As Long
    '根据参数类型设置大小
    Select Case ParamType
    Case adVarChar, adVarWChar, adChar, adWChar
        lSize = 255
    Case adInteger, adSmallInt, adTinyInt
        lSize = 4
    Case adBigInt
        lSize = 8
    Case adDouble, adSingle, adCurrency
        lSize = 8
    Case adDate, adDBTimeStamp
        lSize = 8
    Case adBoolean
        lSize = 1
    Case Else
        lSize = Len(CStr(ParamValue))
        If lSize = 0 Then lSize = 255
    End Select
    
    Set P = Cmd.CreateParameter(ParamName, ParamType, adParamInput, lSize, ParamValue)
    Cmd.Parameters.Append P
End Function

'清理参数化查询的参数和Command对象
Public Sub ClearParam()
    If Not Cmd Is Nothing Then
        Do While Cmd.Parameters.count > 0
            Cmd.Parameters.Delete 0
        Loop
        Set Cmd = Nothing
    End If
End Sub

'执行参数化查询
Public Function ExecParam(Optional RecordsAffected) As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Cmd Is Nothing Then LastErr = "no parameters set": Exit Function
    
    Set Rs = Cmd.Execute(RecordsAffected)
    ExecParam = True
    
    '清理参数
    ClearParam
    Exit Function
eee:
    LastErr = "ExecParam Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ClearParam
End Function

Public Function FetchParam(Optional CurType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockOptimistic) As Boolean
    FetchParam = QueryParam(CurType, LockType)
    If FetchParam = True Then
        Set Rows = ToolsList.RsToCollection(Rs)
        If Rows.count > 0 Then
            Set Row = Rows(1)
        Else
            Set Row = New Scripting.Dictionary
        End If
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
    ClearParam
    Exit Function
eee:
    LastErr = "QueryParam Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ClearParam
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
    
    '设置要执行的命令文本，使用 adExecuteNoRecords 标志避免创建记录集
    Dim executeOptions As Long
    executeOptions = Options Or adExecuteNoRecords
    
    Dim lRecordsAffected As Long
    Conn.Execute pvRawSqlString, lRecordsAffected, executeOptions
    If Not IsMissing(RecordsAffected) Then RecordsAffected = lRecordsAffected
    
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
        If Rows.count > 0 Then
            Set Row = Rows(1)
        Else
            Set Row = New Scripting.Dictionary
        End If
    End If
End Function
Public Function Query(Optional CurType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockOptimistic, Optional Options As Long = -1) As Boolean
    On Error GoTo eee
    '如果连接标记为假，标明已经断开连接，则直接返回
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Rs.State <> adStateClosed Then Rs.Close
    
    '打开记录集
    Rs.Open pvRawSqlString, Conn, CurType, LockType, Options
    
    '处理分页（使用 ADO 原生分页属性，支持所有数据库类型）
    If IsPage Then
        Rs.PageSize = pvPageLimit
        Rs.AbsolutePage = pvPageNum
        IsPage = False
    End If
    
    Query = True
    Exit Function
eee:
    LastErr = "Query Err[" & ERR.Number & "]" & ERR.Description & "{" & pvRawSqlString & "}"
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

Private Sub Class_Initialize()
    Set Conn = New ADODB.Connection
    Connections.CompareMode = TextCompare
    Row.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Call DisConnect
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
Public Function count(Optional ByVal TableName As String = "") As Long
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
        count = rsTemp("cnt")
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
        GetTableFields.Add rsTemp(0).Value
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    Set GetTableFields = Nothing
    LastErr = "GetTableFields Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'批量插入数据（使用事务）
Public Function BatchInsert(ByVal TableName As String, ByVal Data As Collection) As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If Data.count = 0 Then BatchInsert = True: Exit Function
    
    '开始事务
    Call TransBegin
    
    Dim i As Long
    Dim bRow As Scripting.Dictionary
    Dim fields As String, Values As String
    Dim SqlString As String
    
    For i = 1 To Data.count
        Set bRow = Data(i)
        fields = ""
        Values = ""
        
        Dim key As Variant
        For Each key In bRow.Keys
            If fields <> "" Then fields = fields & ", ": Values = Values & ", "
            fields = fields & key
            '处理值，转义单引号
            Dim Val As String
            Val = CStr(bRow(key))
            Val = Replace(Val, "'", "''")
            Values = Values & "'" & Val & "'"
        Next
        
        SqlString = "INSERT INTO " & TableName & " (" & fields & ") VALUES (" & Values & ")"
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
        GetDatabases.Add rsTemp(0).Value
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    Set GetDatabases = Nothing
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
        GetTables.Add rsTemp(0).Value
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    Set GetTables = Nothing
    LastErr = "GetTables Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
End Function

'设置字段值（链式），存入内部缓存，最终由 RowCreate/RowUpdate 自动写入 Recordset
'用法: .Field "serial_no", serialNo  或  .Field "is_new", True
Public Function Field(ByVal FieldName As String, ByVal FieldValue As Variant) As cDataBase
    Set Field = Me
    If pvFields.Exists(FieldName) Then
        pvFields(FieldName) = FieldValue
    Else
        pvFields.Add FieldName, FieldValue
    End If
End Function

'清除已缓存的字段值（RowCreate/RowUpdate 执行后自动调用）
Private Sub ClearFields()
    pvFields.RemoveAll
End Sub

'将缓存字段写入当前 Recordset（RowCreate/RowUpdate 内部调用）
Private Sub FlushFields()
    Dim key As Variant
    For Each key In pvFields.Keys
        Rs.Fields(CStr(key)).Value = pvFields(key)
    Next
End Sub

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'  ORM 链式操作: RowCreate / RowRead / RowUpdate / RowDelete
'  用法:
'    创建: db.Table("t").Field("k",v).RowCreate
'    读取: db.Table("t").Where("id=?",1).RowRead : Debug.Print .rs!name
'    更新: db.Table("t").Where("id=?",1).Field("k",v).RowUpdate
'    删除: db.Table("t").Where("id=?",1).RowDelete
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

'插入新行 - 需先调用 Table() 设置表名
'支持两种用法:
'  1. 链式 Field: With db.Table("devices") : .Field "serial_no", "xxx" : .RowCreate : End With
'  2. 传统 Rs:   With db.Table("devices").RowCreate : .rs!serial_no = "xxx" : .rs.Update : End With
Public Function RowCreate() As cDataBase
    On Error GoTo eee
    Set RowCreate = Me
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowCreate: no table set, call Table() first": Exit Function
    
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.Open "SELECT * FROM " & pvTableName & " WHERE 1=0", Conn, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    
    '自动填充 Field() 链式设置的字段值
    If pvFields.count > 0 Then
        FlushFields
        Rs.Update
        ClearFields
    End If
    '注意: RowCreate 后不自动 ResetBuilder，因为用户可能继续用 Rs 编辑字段
    '但 pvTableName 保留以便后续操作，pvWhere 等查询状态无需重置
    Exit Function
eee:
    LastErr = "RowCreate Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ClearFields
    ResetBuilder
End Function

'查询行 - 链式构建器终结方法，拼接完整 SQL 并打开 Recordset
'用法: db.Table("devices").Where("id=?", id).RowRead
'     db.Table("t").Columns("id,name").OrderBy("age").GroupBy("dept").RowRead
Public Function RowRead() As cDataBase
    On Error GoTo eee
    Set RowRead = Me
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowRead: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = BuildSelectSql()
    ResetBuilder  ' 构建完毕，重置状态
    
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.Open SqlString, Conn, adOpenKeyset, adLockOptimistic
    
    '处理分页
    If IsPage Then
        Rs.PageSize = pvPageLimit
        Rs.AbsolutePage = pvPageNum
        IsPage = False
    End If
    Exit Function
eee:
    LastErr = "RowRead Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'更新行 - 需先调用 Table() + Where() 定位目标行
'支持两种用法:
'  1. 链式 Field: With db.Table("devices").Where("serial_no=?", sn) : .Field "device_json", json : .RowUpdate : End With
'  2. 传统 Rs:   With db.Table("devices").Where("serial_no=?", sn).RowUpdate : .rs!device_json = json : .rs.Update : End With
Public Function RowUpdate() As cDataBase
    On Error GoTo eee
    Set RowUpdate = Me
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowUpdate: no table set, call Table() first": Exit Function
    
    '拼接 WHERE 条件（合并所有链式条件子句）
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql = "" Then LastErr = "RowUpdate: no condition set, call Where() first": ResetBuilder: Exit Function
    
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.Open "SELECT * FROM " & pvTableName & " WHERE " & whereSql, Conn, adOpenKeyset, adLockOptimistic
    If Not Rs.EOF Then
        '定位到第一行（通常只有一行），准备编辑
    Else
        LastErr = "RowUpdate: no row found matching condition"
        ResetBuilder
        Exit Function
    End If
    
    '自动填充 Field() 链式设置的字段值
    If pvFields.count > 0 Then
        FlushFields
        Rs.Update
        ClearFields
    End If
    ResetBuilder  ' 用完重置
    Exit Function
eee:
    LastErr = "RowUpdate Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
    ClearFields
End Function

'删除行 - 需先调用 Table() + Where() 定位目标行
'用法: db.Table("devices").Where("id=?", id).RowDelete
'     db.Table("t").WhereIn("status", "expired,deleted").RowDelete
Public Function RowDelete() As cDataBase
    On Error GoTo eee
    Set RowDelete = Me
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowDelete: no table set, call Table() first": Exit Function
    
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql = "" Then LastErr = "RowDelete: no condition set, call Where() first": ResetBuilder: Exit Function
    
    Dim SqlString As String
    SqlString = "DELETE FROM " & pvTableName & " WHERE " & whereSql
    ResetBuilder  ' 用完重置
    
    Conn.Execute SqlString, , adExecuteNoRecords
    Exit Function
eee:
    LastErr = "RowDelete Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'  链式构建器终结方法 - RowCount / RowExists / RowPluck / RowMax / RowMin / RowSum / RowAvg
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

'链式统计行数 - 返回符合条件的记录总数
'用法: n = db.Table("t").Where("age>18").RowCount
Public Function RowCount() As Long
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowCount: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = "SELECT COUNT(*) AS cnt FROM " & pvTableName
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql <> "" Then SqlString = SqlString & " WHERE " & whereSql
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then RowCount = rsTemp("cnt")
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "RowCount Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'链式判断是否存在 - 返回布尔值
'用法: If db.Table("t").Where("serial_no=?", sn).RowExists Then ...
Public Function RowExists() As Boolean
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowExists: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = "SELECT 1 AS exists_flag FROM " & pvTableName
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql <> "" Then SqlString = SqlString & " WHERE " & whereSql
    '添加 LIMIT 1 提升性能
    Select Case DatabaseType
    Case Mysql: SqlString = SqlString & " LIMIT 1"
    Case MsSql: SqlString = SqlString & " OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY"  '需有ORDER BY，此处简化
    End Select
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    RowExists = Not rsTemp.EOF
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "RowExists Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'链式提取单列值列表 - 返回 Collection
'用法: Dim names As Collection: Set names = db.Table("t").Where("age>18").RowPluck("name")
Public Function RowPluck(ByVal FieldName As String) As Collection
    On Error GoTo eee
    Set RowPluck = New Collection
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowPluck: no table set, call Table() first": Exit Function
    
    '强制修改 Columns 为单字段
    pvSelectFields = FieldName
    Dim SqlString As String
    SqlString = BuildSelectSql()
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    Do Until rsTemp.EOF
        RowPluck.Add rsTemp(0).Value
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    Set RowPluck = New Collection
    LastErr = "RowPluck Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'链式聚合函数 - MAX
'用法: maxAge = db.Table("t").Where("dept='IT'").RowMax("age")
Public Function RowMax(ByVal FieldName As String) As Variant
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowMax: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = "SELECT MAX(" & FieldName & ") AS agg_val FROM " & pvTableName
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql <> "" Then SqlString = SqlString & " WHERE " & whereSql
    If pvGroupByClause <> "" Then SqlString = SqlString & " GROUP BY " & pvGroupByClause
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then RowMax = rsTemp("agg_val")
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "RowMax Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'链式聚合函数 - MIN
Public Function RowMin(ByVal FieldName As String) As Variant
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowMin: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = "SELECT MIN(" & FieldName & ") AS agg_val FROM " & pvTableName
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql <> "" Then SqlString = SqlString & " WHERE " & whereSql
    If pvGroupByClause <> "" Then SqlString = SqlString & " GROUP BY " & pvGroupByClause
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then RowMin = rsTemp("agg_val")
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "RowMin Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'链式聚合函数 - SUM
Public Function RowSum(ByVal FieldName As String) As Variant
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowSum: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = "SELECT SUM(" & FieldName & ") AS agg_val FROM " & pvTableName
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql <> "" Then SqlString = SqlString & " WHERE " & whereSql
    If pvGroupByClause <> "" Then SqlString = SqlString & " GROUP BY " & pvGroupByClause
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then RowSum = rsTemp("agg_val")
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "RowSum Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'链式聚合函数 - AVG
Public Function RowAvg(ByVal FieldName As String) As Variant
    On Error GoTo eee
    If IsConnect = False Then LastErr = "Database not connect": Exit Function
    If pvTableName = "" Then LastErr = "RowAvg: no table set, call Table() first": Exit Function
    
    Dim SqlString As String
    SqlString = "SELECT AVG(" & FieldName & ") AS agg_val FROM " & pvTableName
    Dim whereSql As String
    whereSql = BuildWhereClause()
    If whereSql <> "" Then SqlString = SqlString & " WHERE " & whereSql
    If pvGroupByClause <> "" Then SqlString = SqlString & " GROUP BY " & pvGroupByClause
    ResetBuilder
    
    Dim rsTemp As New ADODB.Recordset
    rsTemp.Open SqlString, Conn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then RowAvg = rsTemp("agg_val")
    rsTemp.Close
    Set rsTemp = Nothing
    Exit Function
eee:
    LastErr = "RowAvg Err[" & ERR.Number & "]" & ERR.Description
    LastErrNumber = ERR.Number
    LastErrDescription = ERR.Description
    ResetBuilder
End Function

'转义SQL字符串（防止SQL注入）
Public Function Escape(ByVal str As String) As String
    Escape = Replace(str, "'", "''")
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


