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
Private m_Keys As New Collection  ' 用于存储 Key，保持与 m_col 同步

Public Property Get RawCollection() As Collection
    Set RawCollection = m_col
End Property

' 添加或更新元素
Public Sub Add(ByVal Item As Variant, Optional ByVal Key As String = "")
    Dim RemoveIndex As Long
    
    On Error Resume Next
    
    ' 如果提供了键，尝试先删除（如果存在）
    If Len(Key) > 0 Then
        ' 先找到并删除旧的 Key-Value 对
        RemoveIndex = FindKeyIndex(Key)
        If RemoveIndex > 0 Then
            m_col.Remove RemoveIndex
            m_Keys.Remove RemoveIndex
        End If
        
        ' 添加新元素
        m_col.Add Item, Key
        m_Keys.Add Key
    Else
        m_col.Add Item
        m_Keys.Add ""                                                           ' 空字符串表示无 Key
    End If
End Sub

' 根据 Key 查找索引
Private Function FindKeyIndex(ByVal Key As String) As Long
    Dim i As Long
    FindKeyIndex = 0
    
    On Error Resume Next
    For i = 1 To m_Keys.Count
        If m_Keys.Item(i) = Key Then
            FindKeyIndex = i
            Exit Function
        End If
    Next i
End Function

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
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    v = m_col.Item(Key)
    Exists = True
    Exit Function
    
ErrorHandler:
    Exists = False
End Function

' 获取或设置元素（Property 实现，兼容 Dictionary 用法）
Public Property Get Item(ByVal KeyOrIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    
    '    Dim v As Variant
    '    If IsNumeric(KeyOrIndex) Then
    '        v = m_col.Item(CInt(KeyOrIndex))
    '    Else
    '        v = m_col.Item(CStr(KeyOrIndex))
    '    End If
    
    ' 检查是否是对象，如果是则使用 Set 返回对象引用
    If IsObject(m_col(KeyOrIndex)) = True Then
        Set Item = m_col(KeyOrIndex)
    Else
        Item = m_col(KeyOrIndex)
    End If
End Property

Public Property Let Item(ByVal KeyOrIndex As Variant, ByVal Value As Variant)
    On Error Resume Next
    If IsNumeric(KeyOrIndex) Then
        ' 通过索引更新（需要知道对应的 Key）
        Dim idx As Long
        idx = CInt(KeyOrIndex)
        If idx >= 1 And idx <= m_col.Count Then
            Dim Key As String
            Key = m_Keys.Item(idx)
            m_col.Remove idx
            m_Keys.Remove idx
            If Key <> "" Then
                m_col.Add Value, Key
            Else
                m_col.Add Value
            End If
            m_Keys.Add Key
        End If
    Else
        ' 通过 Key 更新
        Dim sKey As String
        sKey = CStr(KeyOrIndex)
        idx = FindKeyIndex(sKey)
        If idx > 0 Then
            m_col.Remove idx
            m_Keys.Remove idx
            m_col.Add Value, sKey
            m_Keys.Add sKey
        Else
            ' Key 不存在则添加
            m_col.Add Value, sKey
            m_Keys.Add sKey
        End If
    End If
End Property

Public Property Set Item(ByVal KeyOrIndex As Variant, ByVal Value As Object)
    On Error Resume Next
    If IsNumeric(KeyOrIndex) Then
        Dim idx As Long
        idx = CInt(KeyOrIndex)
        If idx >= 1 And idx <= m_col.Count Then
            Dim Key As String
            Key = m_Keys.Item(idx)
            m_col.Remove idx
            m_Keys.Remove idx
            If Key <> "" Then
                m_col.Add Value, Key
            Else
                m_col.Add Value
            End If
            m_Keys.Add Key
        End If
    Else
        Dim sKey As String
        sKey = CStr(KeyOrIndex)
        idx = FindKeyIndex(sKey)
        If idx > 0 Then
            m_col.Remove idx
            m_Keys.Remove idx
            m_col.Add Value, sKey
            m_Keys.Add sKey
        Else
            m_col.Add Value, sKey
            m_Keys.Add sKey
        End If
    End If
End Property

' 删除元素
Public Sub Remove(ByVal KeyOrIndex As Variant)
    On Error Resume Next
    Dim idx As Long
    
    If IsNumeric(KeyOrIndex) Then
        idx = CInt(KeyOrIndex)
    Else
        ' 通过 Key 找到索引
        idx = FindKeyIndex(CStr(KeyOrIndex))
    End If
    
    If idx > 0 Then
        m_col.Remove idx
        m_Keys.Remove idx
    End If
End Sub

' 返回集合数量
Public Function Count() As Long
    Count = m_col.Count
End Function

' 通过索引获取 Key（用于遍历）
' @param Index 索引（从1开始）
' @return Key 字符串，如果索引无效则返回空字符串
Public Function KeyByIndex(ByVal Index As Long) As String
    On Error GoTo ErrorHandler
    
    If Index >= 1 And Index <= m_Keys.Count Then
        KeyByIndex = m_Keys.Item(Index)
    Else
        KeyByIndex = ""
    End If
    
    Exit Function
    
ErrorHandler:
    KeyByIndex = ""
End Function

'=========================================================================
' 标准字典成员（Scripting.Dictionary 兼容）
'=========================================================================

' 返回所有 Key 的数组
' @return 包含所有 Key 的 String 数组
Public Function Keys() As String()
    Dim KeyArray() As String
    Dim i As Long
    Dim N As Long
    
    N = m_Keys.Count
    If N = 0 Then
        Keys = Array()
        Exit Function
    End If
    
    ReDim KeyArray(0 To N - 1)
    For i = 1 To N
        KeyArray(i - 1) = m_Keys.Item(i)
    Next i
    
    Keys = KeyArray
End Function

' 返回所有 Value 的数组
' @return 包含所有 Value 的 Variant 数组
Public Function Items() As Variant()
    Dim ItemArray() As Variant
    Dim i As Long
    Dim N As Long
    
    N = m_col.Count
    If N = 0 Then
        Items = Array()
        Exit Function
    End If
    
    ReDim ItemArray(0 To N - 1)
    For i = 1 To N
        ItemArray(i - 1) = m_col.Item(i)
    Next i
    
    Items = ItemArray
End Function

' 清空集合（删除所有元素）
Public Sub RemoveAll()
    Set m_col = New Collection
    Set m_Keys = New Collection
End Sub

' 修改指定 Key 的名称（Dictionary 兼容方法）
' @param OldKey 原 Key 名称
' @param NewKey 新 Key 名称
' @return 是否修改成功
Public Function RenameKey(ByVal OldKey As String, ByVal NewKey As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 检查原 Key 是否存在
    Dim idx As Long
    idx = FindKeyIndex(OldKey)
    If idx = 0 Then GoTo ErrorHandler
    
    ' 检查新 Key 是否已存在
    If FindKeyIndex(NewKey) > 0 Then GoTo ErrorHandler
    
    ' 获取 Value
    Dim Value As Variant
    Value = m_col.Item(OldKey)
    
    ' 删除旧的
    m_col.Remove idx
    m_Keys.Remove idx
    
    ' 添加新的
    m_col.Add Value, NewKey
    m_Keys.Add NewKey
    
    RenameKey = True
    Exit Function
    
ErrorHandler:
    RenameKey = False
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
    
    ' 使用 m_Keys 集合获取 Keys
    For i = 1 To Count
        Values(i - 1) = m_col.Item(i)
        Keys(i - 1) = m_Keys.Item(i)
    Next i
    
    CollectionToArray = Count
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
    Set m_Keys = New Collection
    For i = 0 To Count - 1
        If Keys(i) <> "" Then
            m_col.Add Values(i), Keys(i)
        Else
            m_col.Add Values(i)
        End If
        m_Keys.Add Keys(i)
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
    Set m_Keys = New Collection
    For i = 0 To Count - 1
        If Keys(i) <> "" Then
            m_col.Add Values(i), Keys(i)
        Else
            m_col.Add Values(i)
        End If
        m_Keys.Add Keys(i)
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
