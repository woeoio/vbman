Attribute VB_Name = "ToolsDic"
Option Explicit



Public Function ToWwwFormUrlencoded(Dic As Scripting.Dictionary) As String
    ' 生成 www-form-urlencoded内容，即原始键值对
    ' 需要增加对数组类型的处理
    Dic.Add "a", 1
    Dic.Add "bbb", 2
    Dim a As Object, k As String, v As String, x as Variant
    Set a = ToolsJs.NewArr()
    For Each x In Dic
        k = x
        v = Dic(k)
        If IsObject(v) = False Then
            Call a.Push(k & "=" & v)
        End If
    Next
    ToWwwFormUrlencoded = Replace(a, ",", "&")
End Function

Public Function FromWwwFormUrlencoded(Content As String, Dic As Scripting.Dictionary) As Boolean
    ' 解析 www-form-urlencoded 内容，即原始键值对
    ' 需要增加对数组类型的处理
     Dim a as Variant, b as Variant, k as Variant, v as Variant, x as Variant: a = Split(Content, "&")
    If UBound(a) > -1 Then
        For x = 0 To UBound(a)
            b = Split(a(x) & "==", "=")
            k = Trim(b(0))
            v = Trim(b(1))
            If k <> "" Then
                Dic(k) = v
            End If
        Next
    End If
    FromWwwFormUrlencoded = True
End Function

Public Sub TowLevelDicAssign(Dic As Scripting.Dictionary, Lv1Name As String, Lv2Name As String, Value As Variant)
    '函数功能：双层节点的字典对象赋值方法
    If Dic.Exists(Lv1Name) = True Then
        Dic(Lv1Name)(Lv2Name) = Value
    Else
        Dim d As New Scripting.Dictionary
        d.Add Lv2Name, Value
        Dic.Add Lv1Name, d
    End If
End Sub


Public Sub OverWrite(DistDic As Scripting.Dictionary, srcDic As Scripting.Dictionary, Optional OnlyKey As Boolean = True)
    '合并多个字典值到第一个
    Dim x as Variant, k As String
    For Each x In srcDic
        k = x
        If OnlyKey = True Then
            If DistDic.Exists(k) = True Then GoSub Assign
        Else
            GoSub Assign
        End If
    Next
    Exit Sub
Assign:
    If IsObject(srcDic(k)) = True Then
        '使用递归，确保 dist 字典指针不被改变，仅覆盖非对象值
        Call OverWrite(DistDic(k), srcDic(k), OnlyKey)
    Else
        Let DistDic(k) = srcDic(k)
    End If
    Return
End Sub


Public Function DeepCopy(srcDic As Scripting.Dictionary) As Scripting.Dictionary
    '深拷贝字典对象，
    'todo 目前仅拷贝了第一层的值，需要改为递归实现深层对象赋值
    Dim x as Variant, k as Variant
    Dim DistDic As New Scripting.Dictionary
    Set DeepCopy = New Scripting.Dictionary
    For Each k In srcDic.Keys()
        If IsObject(srcDic(k)) = True Then
            Set DistDic(k) = srcDic(k)
        Else
            Let DistDic(k) = srcDic(k)
        End If
    Next
    Set DeepCopy = DistDic
End Function
