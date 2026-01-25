VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolsDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function ToWwwFormUrlencoded(Dic As Scripting.Dictionary) As String
    ' 生成 www-form-urlencoded内容，即原始键值对
    ' 需要增加对数组类型的处理
    Dic.Add "a", 1
    Dic.Add "bbb", 2
    Dim A As Object, K As String, v As String, x As Variant
    Set A = ToolsJs.NewArr()
    For Each x In Dic
        K = x
        v = Dic(K)
        If IsObject(v) = False Then
            Call A.Push(K & "=" & v)
        End If
    Next
    ToWwwFormUrlencoded = Replace(A, ",", "&")
End Function

Public Function FromWwwFormUrlencoded(Content As String, Dic As Scripting.Dictionary) As Boolean
    ' 解析 www-form-urlencoded 内容，即原始键值对
    ' 需要增加对数组类型的处理
     Dim A As Variant, b As Variant, K As Variant, v As Variant, x As Variant: A = Split(Content, "&")
    If UBound(A) > -1 Then
        For x = 0 To UBound(A)
            b = Split(A(x) & "==", "=")
            K = Trim(b(0))
            v = Trim(b(1))
            If K <> "" Then
                Dic(K) = v
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
        Dim D As New Scripting.Dictionary
        D.Add Lv2Name, Value
        Dic.Add Lv1Name, D
    End If
End Sub


Public Sub OverWrite(DistDic As Scripting.Dictionary, srcDic As Scripting.Dictionary, Optional OnlyKey As Boolean = True)
    '合并多个字典值到第一个
    Dim x As Variant, K As String
    For Each x In srcDic
        K = x
        If OnlyKey = True Then
            If DistDic.Exists(K) = True Then GoSub Assign
        Else
            GoSub Assign
        End If
    Next
    Exit Sub
Assign:
    If IsObject(srcDic(K)) = True Then
        '使用递归，确保 dist 字典指针不被改变，仅覆盖非对象值
        Call OverWrite(DistDic(K), srcDic(K), OnlyKey)
    Else
        Let DistDic(K) = srcDic(K)
    End If
    Return
End Sub


Public Function DeepCopy(srcDic As Scripting.Dictionary) As Scripting.Dictionary
    '深拷贝字典对象，
    'todo 目前仅拷贝了第一层的值，需要改为递归实现深层对象赋值
    Dim x As Variant, K As Variant
    Dim DistDic As New Scripting.Dictionary
    Set DeepCopy = New Scripting.Dictionary
    For Each K In srcDic.Keys()
        If IsObject(srcDic(K)) = True Then
            Set DistDic(K) = srcDic(K)
        Else
            Let DistDic(K) = srcDic(K)
        End If
    Next
    Set DeepCopy = DistDic
End Function

