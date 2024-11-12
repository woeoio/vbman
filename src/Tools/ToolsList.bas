Attribute VB_Name = "ToolsList"
Option Explicit


Public Function RsToCollection(Obj As Variant) As Collection
    '把记录集转为字典数组
    Dim Rs As New ADODB.Recordset
    Dim fd As ADODB.field
    Dim i As Long, ii As Long
    Dim RsArr As New VBA.Collection
    Dim RsRow As Scripting.Dictionary
    Set Rs = Obj.Clone
    '克隆的对象记录集指针居然是初始位置
    Dim p As Long: p = IIf(Obj.AbsolutePage < 1, 1, Obj.AbsolutePage)
    Rs.Filter = Obj.Filter                                                      '过滤器要咋欧阳分页器
    Rs.AbsolutePage = p
    '    rs.RecordCount = cnt
    '    rs.AbsolutePosition = Obj.AbsolutePosition
    Do Until Rs.EOF = True
        '注意分页，不知道为啥分页参数不能控制rs记录集数量了，有空研究下
        If Rs.AbsolutePage <> p Then Exit Do
        Set RsRow = New Scripting.Dictionary
        For i = 0 To Rs.fields.Count - 1
            Set fd = Rs.fields(i)
            RsRow.Add fd.Name, fd.Value
        Next
        RsArr.Add RsRow
        Set RsRow = Nothing
        '
        Rs.MoveNext
    Loop
    Set RsToCollection = RsArr
End Function
