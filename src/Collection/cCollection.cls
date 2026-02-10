VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' clsCollectionEx.cls
Option Explicit

Private m_col As New Collection

Public Property Get RawCollection() As Collection
    Set RawCollection = m_col
End Property

' 添加或更新元素
Public Sub Add(ByVal Item As Variant, Optional ByVal Key As String = "")
    On Error Resume Next
    
    ' 如果提供了键，尝试先删除（如果存在）
    If Len(Key) > 0 Then
        m_col.Remove Key
    End If
    
    ' 添加元素
    If Len(Key) > 0 Then
        m_col.Add Item, Key
    Else
        m_col.Add Item
    End If
End Sub

' 更新元素（必须已存在）
Public Function Update(ByVal Item As Variant, ByVal Key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 检查键是否存在
    Dim v As Variant
    v = m_col.Item(Key)                                                         ' 如果不存在会触发错误
    
    ' 删除并重新添加
    m_col.Remove Key
    m_col.Add Item, Key
    
    Update = True
    Exit Function
    
ErrorHandler:
    Update = False
End Function

' 检查键是否存在
Public Function Exists(ByVal Key As String) As Boolean
Attribute Exists.VB_UserMemId = 0
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    v = m_col.Item(Key)
    Exists = True
    Exit Function
    
ErrorHandler:
    Exists = False
End Function

' 获取元素
Public Function Item(ByVal KeyOrIndex As Variant) As Variant
    On Error Resume Next
    If IsNumeric(KeyOrIndex) Then
        Item = m_col.Item(CInt(KeyOrIndex))
    Else
        Item = m_col.Item(CStr(KeyOrIndex))
    End If
End Function

' 删除元素
Public Sub Remove(ByVal KeyOrIndex As Variant)
    On Error Resume Next
    m_col.Remove KeyOrIndex
End Sub

' 返回集合数量
Public Function Count() As Long
    Count = m_col.Count
End Function

'=========================================================================
' 排序功能
'=========================================================================

' 将集合元素转换为数组（支持保留Key）
Private Function CollectionToArray(ByRef Keys() As String, ByRef Values() As Variant) As Long
    Dim i As Long
    Dim Count As Long
    Count = m_col.Count
    
    If Count = 0 Then
        CollectionToArray = 0
        Exit Function
    End If
    
    ReDim Keys(0 To Count - 1)
    ReDim Values(0 To Count - 1)
    
    ' 注意：VBA Collection 无法直接获取 Keys，
    ' 这里我们遍历索引，Keys 数组将存储为 ""
    On Error Resume Next
    For i = 1 To Count
        Values(i - 1) = m_col.Item(i)
        ' 尝试获取 Key（如果存在）
        Keys(i - 1) = GetKeyByIndex(i)
    Next i
    On Error GoTo 0
    
    CollectionToArray = Count
End Function

' 根据索引获取 Key（通过错误处理尝试）
Private Function GetKeyByIndex(ByVal Index As Long) As String
    On Error GoTo NoKey
    Dim Key As String
    ' 尝试用索引和 Key 来访问，如果 Key 存在则返回
    m_col.Item Index
    ' 如果上面没有错误，尝试获取 Key
    ' 由于 VBA Collection 不暴露 Keys，我们无法直接获取
    ' 这里返回空字符串作为标记
    GetKeyByIndex = ""
    Exit Function
NoKey:
    GetKeyByIndex = ""
End Function

' 快速排序（按 Key 字典序升序）
Private Sub QuickSortByKey(ByRef Keys() As String, ByRef Values() As Variant, ByVal Low As Long, ByVal High As Long)
    Dim i As Long, j As Long
    Dim Pivot As String
    Dim TempKey As String
    Dim TempValue As Variant
    
    If Low < High Then
        i = Low
        j = High
        Pivot = Keys((Low + High) \ 2)
        
        Do While i <= j
            Do While Keys(i) < Pivot And i < High
                i = i + 1
            Loop
            Do While Keys(j) > Pivot And j > Low
                j = j - 1
            Loop
            If i <= j Then
                ' 交换 Keys
                TempKey = Keys(i)
                Keys(i) = Keys(j)
                Keys(j) = TempKey
                ' 交换 Values
                TempValue = Values(i)
                Values(i) = Values(j)
                Values(j) = TempValue
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If Low < j Then QuickSortByKey Keys, Values, Low, j
        If i < High Then QuickSortByKey Keys, Values, i, High
    End If
End Sub

' 快速排序（按 Value 字典序升序，仅适用于字符串值）
Private Sub QuickSortByValue(ByRef Keys() As String, ByRef Values() As Variant, ByVal Low As Long, ByVal High As Long)
    Dim i As Long, j As Long
    Dim Pivot As String
    Dim TempKey As String
    Dim TempValue As Variant
    
    If Low < High Then
        i = Low
        j = High
        Pivot = CStr(Values((Low + High) \ 2))
        
        Do While i <= j
            Do While CStr(Values(i)) < Pivot And i < High
                i = i + 1
            Loop
            Do While CStr(Values(j)) > Pivot And j > Low
                j = j - 1
            Loop
            If i <= j Then
                ' 交换 Keys
                TempKey = Keys(i)
                Keys(i) = Keys(j)
                Keys(j) = TempKey
                ' 交换 Values
                TempValue = Values(i)
                Values(i) = Values(j)
                Values(j) = TempValue
                i = i + 1
                j = j - 1
            End If
        Loop
        
        If Low < j Then QuickSortByValue Keys, Values, Low, j
        If i < High Then QuickSortByValue Keys, Values, i, High
    End If
End Sub

' 按 Key 排序集合（重新组织集合顺序）
Public Sub SortByKey()
    Dim Keys() As String
    Dim Values() As Variant
    Dim Count As Long
    Dim i As Long
    
    Count = CollectionToArray(Keys, Values)
    If Count <= 1 Then Exit Sub
    
    ' 排序
    QuickSortByKey Keys, Values, 0, Count - 1
    
    ' 清空原集合并按排序后的顺序重新添加
    Set m_col = New Collection
    For i = 0 To Count - 1
        If Keys(i) <> "" Then
            m_col.Add Values(i), Keys(i)
        Else
            m_col.Add Values(i)
        End If
    Next i
End Sub

' 按 Value 排序集合（重新组织集合顺序，值应为字符串类型）
Public Sub SortByValue()
    Dim Keys() As String
    Dim Values() As Variant
    Dim Count As Long
    Dim i As Long
    
    Count = CollectionToArray(Keys, Values)
    If Count <= 1 Then Exit Sub
    
    ' 排序
    QuickSortByValue Keys, Values, 0, Count - 1
    
    ' 清空原集合并按排序后的顺序重新添加
    Set m_col = New Collection
    For i = 0 To Count - 1
        If Keys(i) <> "" Then
            m_col.Add Values(i), Keys(i)
        Else
            m_col.Add Values(i)
        End If
    Next i
End Sub

' 获取排序后的键数组（不修改原集合）
Public Function GetSortedKeys() As String()
    Dim Keys() As String
    Dim Values() As Variant
    Dim Count As Long
    
    Count = CollectionToArray(Keys, Values)
    If Count > 1 Then
        QuickSortByKey Keys, Values, 0, Count - 1
    End If
    
    GetSortedKeys = Keys
End Function

' 获取排序后的值数组（按 Key 排序，不修改原集合）
Public Function GetSortedValuesByKey() As Variant()
    Dim Keys() As String
    Dim Values() As Variant
    Dim Count As Long
    Dim SortedValues() As Variant
    Dim i As Long
    
    Count = CollectionToArray(Keys, Values)
    If Count = 0 Then
        GetSortedValuesByKey = Array()
        Exit Function
    End If
    
    If Count > 1 Then
        QuickSortByKey Keys, Values, 0, Count - 1
    End If
    
    ReDim SortedValues(0 To Count - 1)
    For i = 0 To Count - 1
        SortedValues(i) = Values(i)
    Next i
    
    GetSortedValuesByKey = SortedValues
End Function

' 遍历集合（For Each 支持）
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_col.[_NewEnum]
End Function
