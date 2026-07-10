VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function RsToCollection(Obj As Variant) As cCollection
    '把记录集转为字典数组
    Dim Rs As New ADODB.Recordset
    Dim fd As ADODB.Field
    Dim i As Long, ii As Long
    Dim RsArr As New cCollection
    Dim RsRow As Dictionary
    Set Rs = Obj.Clone
    '克隆的对象记录集指针居然是初始位置
    Dim p As Long: p = IIf(Obj.AbsolutePage < 1, 1, Obj.AbsolutePage)
    Rs.Filter = Obj.Filter                                                      '过滤器要咋欧阳分页器
    '
    If Rs.EOF = True Then GoTo QH
    Rs.AbsolutePage = p
    '    rs.RecordCount = cnt
    '    rs.AbsolutePosition = Obj.AbsolutePosition
    Do Until Rs.EOF = True
        '注意分页，不知道为啥分页参数不能控制rs记录集数量了，有空研究下
        If Rs.AbsolutePage <> p Then Exit Do
        Set RsRow = New Dictionary
        RsRow.CompareMode = TextCompare
        For i = 0 To Rs.fields.Count - 1
            Set fd = Rs.fields(i)
            RsRow.Add fd.Name, fd.Value
        Next
        RsArr.Add RsRow
        Set RsRow = Nothing
        '
        Rs.MoveNext
    Loop
QH:
    Set RsToCollection = RsArr
End Function

