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
'tmp
Dim x, Y, z, i As Long, ii As Long

Public Sub Reset(Optional OnlyData As Boolean)
    Data.RemoveAll
    HeadTitle.RemoveAll
    CurrentLineDic.RemoveAll
    HeadTitleWidth.RemoveAll
    CurrentLineNumber = 0
    If OnlyData = False Then Call Class_Initialize
End Sub

'Public Function ShowTo(Lv As Object) As Boolean
'    On Error GoTo Eh
'    ' 初始化ListView控件
'    With Lv
'        .ColumnHeaders.Clear
'        .ListItems.Clear
'        .View = lvwReport                                                       ' 设置为报告视图
'        .FullRowSelect = True                                                   ' 允许选择整行
'        .GridLines = True                                                       ' 显示网格线
'        .LabelEdit = lvwManual                                                  '禁止编辑行号
'        Dim Row As Scripting.Dictionary
'        'Dim Col As Scripting.Dictionary
'        Dim itm As Object
'        'Dim itm As ListItem
'
'        For i = 1 To Data.Count
'            Set Row = Data(i)
'            If i > 1 Then Set itm = .ListItems.Add(, , i - 1)
'            For ii = 1 To Row.Count
'                If i = 1 Then
'                    ' 添加列标题
'                    Dim hh As String: hh = ii & "#" & Row(ii)
'                    If ii = 1 Then .ColumnHeaders.Add , , "行号", 800
'                    .ColumnHeaders.Add , , hh, HeadTitleWidth(ii)
'                Else
'                    ' 添加数据到ListView
'                    itm.SubItems(ii) = Row(ii)
'                End If
'            Next
'        Next
'    End With
'    ShowTo = True
'    Exit Function
'Eh:
'    LastError = ERR.Number & "#" & ERR.Description
'End Function

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
    Dim CN As Long, ln As Long, l, c As String
    Dim f As Scripting.Dictionary
    If IsMissing(Content) = True Or IsNull(Content) Or IsEmpty(Content) Then Content = ""
    For Each l In Data.Keys()
        ln = ln + 1
        Set f = Data(l)
        CN = f.Count + 1
        GoSub AddColumn
    Next
    Exit Function
AddColumn:
    If IsFirstLineToHead = True And ln = 1 Then
        c = Title
    Else
        c = Content
    End If
    If Width = 0 Then Width = HeadTitleWidthDefault
    f.Add CN, c
    'add width and title
    HeadTitle(CN) = Title
    HeadTitleWidth(CN) = Width
    Return
End Function
Public Function NewLine(ParamArray FieldsData() As Variant) As Boolean
    Dim ln As Long: ln = Data.Count + 1
    Dim f As New Scripting.Dictionary, d() As Variant: d = FieldsData
    If UBound(d) = -1 Then ReDim d(0)
    If TypeOf d(0) Is Scripting.Dictionary Then
        Dim ff As Scripting.Dictionary: Set ff = d(0)
        Set f = ToolsDic.DeepCopy(ff)
    Else
        Dim CN As Long, x
        For Each x In d
            CN = CN + 1
            If IsMissing(x) = True Or IsNull(x) Or IsEmpty(x) Then x = ""
            f.Add CN, x
        Next
    End If
    SyncColumnLength f
    Data.Add ln, f
    CurrentLineNumber = ln
    Set CurrentLineDic = f
    Call SyncColumnWidth
    Set f = Nothing
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
Public Property Get Value(ByVal LineNumber As Long, ByVal ColumnNumber As Long) As Variant
    Value = vbNull
    If Data.Exists(LineNumber) = False Then LastError = "行号不存在": Exit Property
    If Data(LineNumber).Exists(ColumnNumber) = False Then LastError = "列号不存在": Exit Property
    Value = Data(LineNumber)(ColumnNumber)
End Property
Public Property Set Value(ByVal LineNumber As Long, ByVal ColumnNumber As Long, ByVal Dat As Variant)
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

Public Function LoadFrom(ByVal FilePath As String, Optional IsFirstHead As Boolean = True) As cCsv
    '20240828 可能需要同步处理 headtitle 和 headwidth'
    On Error GoTo EH
    Set LoadFrom = Me
    Dim Dict As New Dictionary
    
    Dim Fso As New Scripting.FileSystemObject
    
    Dim fo As Scripting.TextStream
    Set fo = Fso.OpenTextFile(FilePath, 1)                                      ' 1: ForReading
    Set Data = Dict
    Dim LineText As String, LineNumber As Long
    
    Do While Not fo.AtEndOfStream
        Dim LineDic As New Scripting.Dictionary
        LineNumber = LineNumber + 1
        LineText = fo.ReadLine
        If ParseCsvLine(LineText, LineDic) = True Then
            'todo 此处也应该做首行对齐，避免用户的csv文件不规范
            SyncColumnLength LineDic
            Dict.Add LineNumber, LineDic
        End If
        Set LineDic = Nothing
    Loop
    
    fo.Close
    
    Call SyncColumnWidth
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description & "：" & FilePath
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

Public Function SaveTo(ByVal FilePath As String) As Boolean
    On Error GoTo EH
    Dim Fso As New Scripting.FileSystemObject
    
    Dim File As Scripting.TextStream
    Set File = Fso.CreateTextFile(FilePath, True)                               ' True: Overwrite if exists
    
    Dim Key As Variant, i As Long, LineFields() As String
    Dim LineDic As Scripting.Dictionary
    For Each Key In Data.Keys
        Set LineDic = Data(Key)
        ReDim LineFields(LineDic.Count - 1)
        For i = 1 To LineDic.Count
            LineFields(i - 1) = QuoteCsvField(LineDic(i))
        Next
        File.WriteLine Join(LineFields, Delimiter)
    Next
    File.Close
    SaveTo = True
    Exit Function
EH:
    LastError = ERR.Number & "#" & ERR.Description
End Function

Private Function QuoteCsvField(ByVal field As String) As String
    ' 检查字段是否包含逗号或双引号
    If InStr(field, Delimiter) > 0 Or InStr(field, """") > 0 Then
        ' 双引号转义，将单个双引号变成两个双引号
        field = Replace(field, """", """""")
        ' 用双引号包裹字段
        QuoteCsvField = """" & field & """"
    Else
        QuoteCsvField = field
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
