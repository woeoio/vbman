VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'操作类型枚举
Public Enum CsvOperationType
    OperationNone = 0
    OperationLoadFrom = 1
    OperationLoadFromRecordSet = 2
    OperationSaveTo = 3
    OperationSaveFromRecordSet = 4
End Enum

'事件声明
Public Event Progress(ByVal Operation As CsvOperationType, ByVal CurrentLine As Long, ByVal TotalLines As Long, ByVal Percent As String, LineDataArray As Variant)
Public Event Complete(ByVal Operation As CsvOperationType, ByVal TotalLines As Long, ByVal Success As Boolean)

'todo 需要改为文件流的处理方式，解决自动识别文件编码问题，以及字段内回车问题
Public Data As New Scripting.Dictionary
Public CurrentLineDic As New Scripting.Dictionary
Public CurrentLineNumber As Long
Public Delimiter As String
Public LastError As String
'用于列表控件渲染
Public HasHeadTitle As Boolean
Public HeadTitle As New Scripting.Dictionary
Public HeadTitleWidth As New Scripting.Dictionary
Public HeadTitleWidthDefault As Long                                            '默认宽度1500
Public HeadTitleWidthMax As Long                                                '最大宽度8000
Public HeadTitleWidthAuto As Boolean                                            '是否自动宽度
Public HeadTitleWidthWord As Long                                               '单个双字节字符宽度
Public IsFirstLineToHead As Boolean                                             '是否首行作为标题

'循环中断标志
'只有外部引用了 Progress 事件，并在事件中有执行 DoEvents 才能响应 StopLoop 函数
'或者先调用 CanStop() 开启循环中的 DoEvents ，再调用 load /save 函数
Private StopLoopFlag As Boolean
Private EnableDoEvents As Boolean
'tmp
Dim x, y, z, i As Long, ii As Long

Public Sub Reset(Optional OnlyData As Boolean)
    Data.RemoveAll
    HeadTitle.RemoveAll
    CurrentLineDic.RemoveAll
    HeadTitleWidth.RemoveAll
    CurrentLineNumber = 0
    EnableDoEvents = False
    If OnlyData = False Then Call Class_Initialize
End Sub

Public Sub StopLoop()
    '设置中断标志，用于在循环中中断操作
    StopLoopFlag = True
End Sub

Public Function CanStop() As cCsv
    '启用循环中的 DoEvents，使 StopLoop 能够响应
    EnableDoEvents = True
    Set CanStop = Me
End Function

Public Function ShowTo(Lv As Object, Optional DefaultWidth As Long = 800) As Boolean
    On Error GoTo EH
    ' 初始化ListView控件
    With Lv
        .ColumnHeaders.Clear
        .ListItems.Clear
        .View = 3                                                               ' lvwReport                                                       ' 设置为报告视图
        .FullRowSelect = True                                                   ' 允许选择整行
        .GridLines = True                                                       ' 显示网格线
        .LabelEdit = 1                                                          ' lvwManual                                                  '禁止编辑行号
        Dim Row As Scripting.Dictionary
        'Dim Col As Scripting.Dictionary
        Dim itm As Object
        'Dim itm As ListItem
        
        For i = 1 To Data.Count
            Set Row = Data(i)
            If i > 1 Then Set itm = .ListItems.Add(, , i - 1)
            For ii = 1 To Row.Count
                If i = 1 Then
                    ' 添加列标题
                    Dim hh As String: hh = ii & "#" & Row(ii)
                    If ii = 1 Then .ColumnHeaders.Add , , "行号", DefaultWidth
                    .ColumnHeaders.Add , , hh, HeadTitleWidth(ii)
                Else
                    ' 添加数据到ListView
                    itm.SubItems(ii) = Row(ii)
                End If
            Next
        Next
    End With
    ShowTo = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
End Function

Public Function SetHeadTitleWidthByDic(ByVal Dic As Scripting.Dictionary) As Boolean
    For Each x In Dic.Keys()
        If HeadTitleWidth.Exists(x) = True Then HeadTitleWidth(x) = Dic(x)
    Next
End Function
Public Function SetHeadTitleWidth(ByVal ColumnNumber As Long, ByVal ColumnWidth As Long) As Boolean
    If HeadTitleWidth.Exists(ColumnNumber) = False Then LastError = "列号不存在": Exit Function
    If ColumnWidth > HeadTitleWidthMax Then ColumnWidth = HeadTitleWidthMax
    HeadTitleWidth(ColumnNumber) = ColumnWidth
    SetHeadTitleWidth = True
End Function
Public Function DeleteColumn(Optional ColumnNumber As Long) As Boolean
    If GetLine(1) = False Then Exit Function
    If ColumnNumber = 0 Then ColumnNumber = CurrentLineDic.Count
    '20240828 待完善，目前只删除最后一列
    For Each x In Data.Keys()
        Data(x).Remove ColumnNumber
        If HeadTitle.Exists(ColumnNumber) = True Then HeadTitle.Remove ColumnNumber
        If HeadTitleWidth.Exists(ColumnNumber) = True Then HeadTitleWidth.Remove ColumnNumber
    Next
    If CurrentLineDic.Exists(ColumnNumber) = True Then CurrentLineDic.Remove ColumnNumber
    DeleteColumn = True
End Function
Public Function NewColumn(ByVal Title As String, Optional Width As Long, Optional Content As String) As Boolean
    Dim CN As Long, lN As Long, L, C As String
    Dim F As Scripting.Dictionary
    If IsMissing(Content) = True Or IsNull(Content) Or IsEmpty(Content) Then Content = ""
    For Each L In Data.Keys()
        lN = lN + 1
        Set F = Data(L)
        CN = F.Count + 1
        GoSub AddColumn
    Next
    Exit Function
AddColumn:
    If IsFirstLineToHead = True And lN = 1 Then
        C = Title
    Else
        C = Content
    End If
    If Width = 0 Then Width = HeadTitleWidthDefault
    F.Add CN, C
    'add width and title
    HeadTitle(CN) = Title
    HeadTitleWidth(CN) = Width
    Return
End Function
Public Function NewLine(ParamArray FieldsData() As Variant) As Boolean
    Dim lN As Long: lN = Data.Count + 1
    Dim F As New Scripting.Dictionary, D() As Variant: D = FieldsData
    If UBound(D) = -1 Then ReDim D(0)
    If TypeOf D(0) Is Scripting.Dictionary Then
        Dim ff As Scripting.Dictionary: Set ff = D(0)
        Set F = ToolsDic.DeepCopy(ff)
    Else
        Dim CN As Long, x
        For Each x In D
            CN = CN + 1
            If IsMissing(x) = True Or IsNull(x) Or IsEmpty(x) Then x = ""
            F.Add CN, x
        Next
    End If
    SyncColumnLength F
    Data.Add lN, F
    CurrentLineNumber = lN
    Set CurrentLineDic = F
    Call SyncColumnWidth
    Set F = Nothing
    NewLine = True
End Function
Private Sub SyncColumnLength(LineDic As Scripting.Dictionary)
    Dim dn As Long: dn = Data.Count
    If dn > 0 Then
        Dim d0n As Long: d0n = Data(1).Count
        With LineDic
            Do Until .Count = d0n
                If .Count > d0n Then
                    .Remove .Count
                Else
                    .Add .Count + 1, ""
                End If
            Loop
        End With
    End If
End Sub
Private Sub SyncColumnWidth()
    '同步宽度
    Dim dn As Long: dn = Data.Count
    If dn > 0 Then
        Dim d0n As Long: d0n = Data(1).Count
        With HeadTitleWidth
            Do Until .Count = d0n
                If .Count > d0n Then
                    .Remove .Count
                Else
                    .Add .Count + 1, HeadTitleWidthDefault
                End If
            Loop
        End With
        '同步标题
        With HeadTitle
            Do Until .Count = d0n
                If .Count > d0n Then
                    .Remove .Count
                Else
                    .Add .Count + 1, "(无标题)"
                End If
            Loop
        End With
    End If
End Sub
Public Function GetLine(ByVal LineNumber As Long, Optional ByRef LineDic As Scripting.Dictionary) As Boolean
    If Data.Exists(LineNumber) = False Then LastError = "行号不存在": Exit Function
    Set CurrentLineDic = Data(LineNumber)
    CurrentLineNumber = LineNumber
    If IsMissing(LineDic) = False Then Set LineDic = CurrentLineDic
    GetLine = True
End Function
Public Function SetLine(ByVal LineNumber As Long, ByRef LineDic As Scripting.Dictionary) As Boolean
    If Data.Exists(LineNumber) = False Then LastError = "行号不存在": Exit Function
    Set CurrentLineDic = LineDic
    SyncColumnLength CurrentLineDic
    Set Data(LineNumber) = CurrentLineDic
    Set CurrentLineDic = Nothing
    SetLine = True
End Function
Public Property Get CountLine() As Long
    CountLine = Data.Count
End Property
Public Property Get CountColumn() As Long
    If Data.Count > 0 Then CountColumn = Data(1).Count
End Property
Public Property Get Value(ByVal LineNumber As Long, ByVal ColumnNumber As Long) As Variant
    Value = vbNull
    If Data.Exists(LineNumber) = False Then LastError = "行号不存在": Exit Property
    If Data(LineNumber).Exists(ColumnNumber) = False Then LastError = "列号不存在": Exit Property
    Value = Data(LineNumber)(ColumnNumber)
End Property
Public Property Set Value(ByVal LineNumber As Long, ByVal ColumnNumber As Long, ByVal Dat As Variant)
    If Data.Exists(LineNumber) = False Then LastError = "行号不存在": Exit Property
    If Data(LineNumber).Exists(ColumnNumber) = False Then LastError = "列号不存在": Exit Property
    Set Data(LineNumber)(ColumnNumber) = Dat
End Property
Public Property Let Value(ByVal LineNumber As Long, ByVal ColumnNumber As Long, ByVal Dat As Variant)
    If Data.Exists(LineNumber) = False Then LastError = "行号不存在": Exit Property
    If Data(LineNumber).Exists(ColumnNumber) = False Then LastError = "列号不存在": Exit Property
    Data(LineNumber)(ColumnNumber) = Dat
End Property
Public Function SetCurrentLineData(ByVal ColumnNumber As Long, ByVal Dat As Variant, Optional ByRef LineDic As Scripting.Dictionary) As Boolean
    If IsMissing(LineDic) = True Then Set LineDic = CurrentLineDic
    If CurrentLineDic.Exists(ColumnNumber) = False Then LastError = "列号不存在": Exit Function
    CurrentLineDic(ColumnNumber) = Dat
    SetCurrentLineData = True
End Function

Public Function LoadFrom(ByVal FilePath As String, Optional IsFirstHead As Boolean = True, Optional CharSet As String = "UTF-8") As cCsv
    '20240828 可能需要同步处理 headtitle 和 headwidth'
    On Error GoTo EH
    Set LoadFrom = Me
    Dim Dict As New Dictionary
    
    Dim fo As New cToolsStream
    fo.UseLineMode(True).LoadFileAsText FilePath, CharSet                       ' 1: ForReading
    
    Set Data = Dict
    Dim lineText As String
    Dim TotalLines As Long
    TotalLines = fo.LineData.Count
    
    StopLoopFlag = False
    Dim i As Long
    For i = 1 To TotalLines
        If EnableDoEvents Then DoEvents
        If StopLoopFlag Then
            StopLoopFlag = False
            LastError = "用户中断"
            RaiseEvent Complete(OperationLoadFrom, i - 1, False)
            Exit Function
        End If
        Dim LineDic As New Scripting.Dictionary
        lineText = fo.ReadLine
        If lineText = "" Then
            Dim Percent1 As Long: Percent1 = (i / TotalLines) * 100
            Dim EmptyData() As Variant
            RaiseEvent Progress(OperationLoadFrom, i, TotalLines, Percent1 & "%", EmptyData)
            GoTo CT
        End If
        If ParseCsvLine(lineText, LineDic) = True Then
            'todo 此处也应该做首行对齐，避免用户的csv文件不规范
            SyncColumnLength LineDic
            Dict.Add fo.LineNumber - 1, LineDic
            ' RaiseEvent RowLoaded(fo.LineNumber - 1, LineDic)
        End If
        Dim Percent2 As Long: Percent2 = (i / TotalLines) * 100
        Dim LineData1() As Variant: LineData1 = LineDic.Items()
        RaiseEvent Progress(OperationLoadFrom, i, TotalLines, Percent2 & "%", LineData1)
        Set LineDic = Nothing
CT:
    Next
    
    fo.Inst.Close
    
    Call SyncColumnWidth
    RaiseEvent Complete(OperationLoadFrom, TotalLines, True)
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description & "：" & FilePath
    RaiseEvent Complete(OperationLoadFrom, 0, False)
End Function

Public Function LoadFromRecordSet(ByVal Rs As Object, Optional IsFirstHead As Boolean = True) As cCsv
    '从数据库记录集载入数据
'rs:ADODB.Recordset 对象
'IsFirstHead: 是否将字段名作为标题行
    On Error GoTo EH
    Set LoadFromRecordSet = Me
    
    If Rs Is Nothing Then LastError = "Recordset对象为空": Exit Function
    If Rs.State <> 1 Then LastError = "Recordset状态不正确": Exit Function      ' adStateOpen = 1
    
    Dim Dict As New Dictionary
    Set Data = Dict
    
    Dim FieldCount As Long
    FieldCount = Rs.fields.Count
    
    Dim RowCount As Long
    RowCount = 0
    
    '计算总行数用于进度事件
    Dim TotalRows As Long
    If Not (Rs.EOF And Rs.BOF) Then
        Rs.MoveLast
        TotalRows = Rs.RecordCount
        If IsFirstHead Then TotalRows = TotalRows + 1
        Rs.MoveFirst
    End If
    
    '如果IsFirstHead=True，先添加标题行
    If IsFirstHead = True And FieldCount > 0 Then
        Dim TitleDic As New Scripting.Dictionary
        Dim i As Long
        For i = 0 To FieldCount - 1
            TitleDic.Add i + 1, Rs.fields(i).Name
        Next
        SyncColumnLength TitleDic
        Dict.Add 1, TitleDic
        RowCount = 1
        ' RaiseEvent RowLoaded(1, TitleDic)
        Dim Percent3 As Long: Percent3 = (RowCount / TotalRows) * 100
        Dim LineData3() As Variant: LineData3 = TitleDic.Items()
        RaiseEvent Progress(OperationLoadFromRecordSet, RowCount, TotalRows, Percent3 & "%", LineData3)
    End If
    
    '遍历记录集
    If Not (Rs.EOF And Rs.BOF) Then
        StopLoopFlag = False
        Do While Not Rs.EOF
            If EnableDoEvents Then DoEvents
            If StopLoopFlag Then
                StopLoopFlag = False
                LastError = "用户中断"
                RaiseEvent Complete(OperationLoadFromRecordSet, RowCount - 1, False)
                Exit Function
            End If
            RowCount = RowCount + 1
            Dim LineDic As New Scripting.Dictionary
            Dim J As Long
            For J = 0 To FieldCount - 1
                Dim FieldValue As Variant
                FieldValue = Rs.fields(J).Value
                If IsNull(FieldValue) Then FieldValue = ""
                LineDic.Add J + 1, FieldValue
            Next
            SyncColumnLength LineDic.Items()
            Dict.Add RowCount, LineDic
            ' RaiseEvent RowLoaded(RowCount, LineDic)
            Dim Percent4 As Long: Percent4 = (RowCount / TotalRows) * 100
            Dim LineData4() As Variant: LineData4 = LineDic.Items()
            RaiseEvent Progress(OperationLoadFromRecordSet, RowCount, TotalRows, Percent4 & "%", LineData4)
            Set LineDic = Nothing
            Rs.MoveNext
        Loop
    End If
    
    Call SyncColumnWidth
    RaiseEvent Complete(OperationLoadFromRecordSet, RowCount, True)
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
    RaiseEvent Complete(OperationLoadFromRecordSet, 0, False)
End Function

Private Function ParseCsvLine(ByVal Content As String, ByRef LineDic As Scripting.Dictionary) As Boolean
    If Trim(Content) = "" Then LastError = "内容为空": Exit Function
    LineDic.RemoveAll
    Dim inQuotes As Boolean
    inQuotes = False
    
    Dim i As Long
    Dim currentField As String
    currentField = ""
    
    Dim fields() As String
    ReDim fields(0)
    Dim fieldIndex As Long
    fieldIndex = 0
    
    For i = 1 To Len(Content)
        Dim currentChar As String
        currentChar = Mid(Content, i, 1)
        
        Select Case currentChar
        Case """"
            If inQuotes And Mid(Content, i + 1, 1) = """" Then
                ' 双引号内的双引号，视为转义双引号，加入字段
                currentField = currentField & """"
                i = i + 1                                                       ' 跳过下一个引号
            Else
                inQuotes = Not inQuotes                                         ' 切换引号状态
            End If
        Case Delimiter
            If inQuotes Then
                ' 引号内的逗号，加入字段
                currentField = currentField & currentChar
            Else
                ' 引号外的逗号，作为字段分隔符
                fields(fieldIndex) = Trim(currentField)
                fieldIndex = fieldIndex + 1
                ReDim Preserve fields(fieldIndex)
                currentField = ""
            End If
        Case Else
            currentField = currentField & currentChar
        End Select
    Next i
    
    ' 最后一个字段
    fields(fieldIndex) = Trim(currentField)
    
    ' 转存到字典
    For i = 0 To fieldIndex
        LineDic.Add i + 1, fields(i)
    Next
    
    ParseCsvLine = True
End Function

'----

Public Function SaveTo(ByVal FilePath As String, Optional CharSet As String = "UTF-8") As Boolean
    On Error GoTo EH
    Dim Fso As New cToolsStream
    Fso.UseLine = True
    Dim Key As Variant, i As Long, LineFields() As String
    Dim LineDic As Scripting.Dictionary
    Dim TotalRows As Long
    Dim CurrentRow As Long
    
    TotalRows = Data.Count
    
    StopLoopFlag = False
    For Each Key In Data.Keys
        If EnableDoEvents Then DoEvents
        If StopLoopFlag Then
            StopLoopFlag = False
            LastError = "用户中断"
            RaiseEvent Complete(OperationSaveTo, CurrentRow - 1, False)
            Exit Function
        End If
        CurrentRow = CurrentRow + 1
        Set LineDic = Data(Key)
        ReDim LineFields(LineDic.Count - 1)
        For i = 1 To LineDic.Count
            LineFields(i - 1) = QuoteCsvField(LineDic(i))
        Next
        Fso.WriteLine Join(LineFields, Delimiter)
        ' RaiseEvent RowSaved(CurrentRow, LineDic)
        Dim Percent5 As Long: Percent5 = (CurrentRow / TotalRows) * 100
        Dim LineData5() As Variant: LineData5 = LineDic.Items()
        RaiseEvent Progress(OperationSaveTo, CurrentRow, TotalRows, Percent5 & "%", LineData5)
    Next
    Fso.SaveFileAsText FilePath, , CharSet
    Fso.Inst.Close
    RaiseEvent Complete(OperationSaveTo, TotalRows, True)
    SaveTo = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
    RaiseEvent Complete(OperationSaveTo, 0, False)
End Function

Public Function SaveFromRecordSet(ByVal Rs As Object, ByVal SaveToFilePath As String, Optional IncludeFieldNames As Boolean = True) As Boolean
    '从ADODB.RecordSet记录集直接保存到CSV文件（高效版）
'rs:ADODB.Recordset对象
'SaveToFilePath: 保存到的文件路径
'IncludeFieldNames: 是否包含字段名作为第一行
    On Error GoTo EH
    
    If Rs Is Nothing Then LastError = "Recordset对象为空": Exit Function
    If Rs.State <> 1 Then LastError = "Recordset状态不正确": Exit Function      ' adStateOpen = 1
    If SaveToFilePath = "" Then LastError = "保存路径为空": Exit Function
    
    Dim FieldCount As Long
    FieldCount = Rs.fields.Count
    If FieldCount = 0 Then LastError = "Recordset字段数为0": Exit Function
    
    '创建文件流对象
    Dim Fso As New cToolsStream
    Fso.UseLine = True
    
    Dim i As Long
    Dim LineFields() As String
    Dim CurrentRow As Long
    Dim TotalRows As Long
    
    '计算总行数用于进度事件
    If Not (Rs.EOF And Rs.BOF) Then
        Rs.MoveLast
        TotalRows = Rs.RecordCount
        If IncludeFieldNames Then TotalRows = TotalRows + 1
        Rs.MoveFirst
    End If
    
    '写入字段名行（如果需要）
    If IncludeFieldNames = True Then
        ReDim LineFields(FieldCount - 1)
        For i = 0 To FieldCount - 1
            LineFields(i) = QuoteCsvField(Rs.fields(i).Name)
        Next
        Fso.WriteLine Join(LineFields, Delimiter)
        CurrentRow = 1
        Dim Percent6 As Long: Percent6 = (CurrentRow / TotalRows) * 100
        RaiseEvent Progress(OperationSaveFromRecordSet, CurrentRow, TotalRows, Percent6 & "%", LineFields)
    End If
    
    '遍历记录集，直接写入文件
    If Not (Rs.EOF And Rs.BOF) Then
        StopLoopFlag = False
        Do While Not Rs.EOF
            If EnableDoEvents Then DoEvents
            If StopLoopFlag Then
                StopLoopFlag = False
                LastError = "用户中断"
                RaiseEvent Complete(OperationSaveFromRecordSet, CurrentRow - 1, False)
                Exit Function
            End If
            CurrentRow = CurrentRow + 1
            ReDim LineFields(FieldCount - 1)
            
            For i = 0 To FieldCount - 1
                Dim FieldValue As Variant
                FieldValue = Rs.fields(i).Value
                If IsNull(FieldValue) Then
                    LineFields(i) = ""
                Else
                    LineFields(i) = QuoteCsvField(CStr(FieldValue))
                End If
            Next
            
            Fso.WriteLine Join(LineFields, Delimiter)
            Dim Percent7 As Long: Percent7 = (CurrentRow / TotalRows) * 100
            RaiseEvent Progress(OperationSaveFromRecordSet, CurrentRow, TotalRows, Percent7 & "%", LineFields)
            Rs.MoveNext
        Loop
    End If
    
    '保存文件
    Fso.SaveFileAsText SaveToFilePath, , "UTF-8"
    Fso.Inst.Close
    
    RaiseEvent Complete(OperationSaveFromRecordSet, CurrentRow, True)
    SaveFromRecordSet = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
    RaiseEvent Complete(OperationSaveFromRecordSet, 0, False)
End Function

Private Function QuoteCsvField(ByVal Field As String) As String
    ' 检查字段是否包含逗号或双引号
    If InStr(Field, Delimiter) > 0 Or InStr(Field, """") > 0 Then
        ' 双引号转义，将单个双引号变成两个双引号
        Field = Replace(Field, """", """""")
        ' 用双引号包裹字段
        QuoteCsvField = """" & Field & """"
    Else
        QuoteCsvField = Field
    End If
End Function

Private Sub Class_Initialize()
    Delimiter = ","
    HasHeadTitle = True
    HeadTitleWidthDefault = 1500
    HeadTitleWidthMax = 8000
    IsFirstLineToHead = True
End Sub


'Private Sub Command1_Click()
'    Dim f As String: f = App.Path & "\..\1.csv"
'    Dim cc As New cCsv
'    If cc.LoadFrom(f) = True Then
'        MsgBox cc.Data.Count, , cc.GetDataByPostion(5, 6)
'        cc.SetDataByPostion 4, 1, Now()
'        'cc.Data(55)(1) = "444"
'        cc.SaveTo f
'    Else
'        MsgBox cc.LastError
'    End If
'End Sub
'
'Private Sub Command2_Click()
'    Dim f1 As String: f1 = App.Path & "\..\1.csv"
'    Dim c As New cCsv
'    c.LoadFrom f1
'    If c.GetLine(2) = False Then MsgBox c.LastError
'    c.SetCurrentLineData 3, 999
'    Dim cc As New cCsv
'    cc.NewLine "12345", "45,""9999""45", "ffgygy", "48546", "1234", 123415
'    cc.NewLine "84565846", "56564,"
'    cc.NewLine , , "ffgygy", "455562", "gfhfghfh", "45,""9999""45"
'    If cc.NewLine(c.CurrentLineDic) = False Then MsgBox cc.LastError
'    cc.NewColumn "888", , "999"
'    cc.NewColumn
'    cc.NewColumn
'    cc.NewLine
'    cc.NewLine
'    Dim f11 As String: f11 = App.Path & "\..\11.csv"
'    If cc.SaveTo(f11) = False Then MsgBox cc.LastError
'End Sub
