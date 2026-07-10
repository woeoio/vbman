VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Dictionary.cls
' 完全兼容 Dictionary 的字典类
' 内部使用 cCollection 实现（Collection 本身不支持按 Key 区分大小写查找，
' 这里通过 CompareMode 实现二进制/文本/数据库三种比较模式）
'
' 兼容的 Dictionary 公共 API:
'   - Add(Key, Item)              参数顺序为 (Key, Item)
'   - Exists(Key)                 按 CompareMode 区分大小写
'   - Item(Key) [default member]  支持 Get/Let/Set
'   - Remove(Key)
'   - RemoveAll()
'   - Keys()                      返回 String 数组
'   - Items()                     返回 Variant 数组
'   - Count                       属性
'   - CompareMode                 属性 (vbBinaryCompare/vbTextCompare/vbDatabaseCompare)
'   - Key(Key) = NewKey           重命名 Key
'   - NewEnum                     For Each 支持
'
' 扩展 API（cCollection 兼容，非 Dictionary 标准）:
'   - Update(Item, Key)
'   - Exists(Key, CompareCase)
'   - Count()                     方法形式（属性和方法可同时使用）
'   - KeyByIndex(Index)
'   - RenameKey(OldKey, NewKey)
'   - SortByKey / SortByValue
'   - GetSortedKeys / GetSortedValuesByKey
'   - RawCollection / InnerCollection

Public m_col As New cCollection
Private m_CompareMode As CompareMethod       ' 缺省 vbBinaryCompare

Private Sub Class_Initialize()
    m_CompareMode = vbBinaryCompare
End Sub

'=========================================================================
' Dictionary 标准 API
'=========================================================================

' 设置/获取比较模式：vbBinaryCompare=0, vbTextCompare=1, vbDatabaseCompare=2
Public Property Get CompareMode() As CompareMethod
    CompareMode = m_CompareMode
End Property

Public Property Let CompareMode(ByVal Value As CompareMethod)
    m_CompareMode = Value
End Property

' 添加或更新元素（参数顺序：Key 在前，Item 在后）
Public Sub Add(ByVal Key As String, ByVal Item As Variant)
    m_col.Add Item, Key
End Sub

' 检查键是否存在（按当前 CompareMode 区分大小写）
Public Function Exists(ByVal Key As String) As Boolean
    Dim bCaseSensitive As Boolean
    bCaseSensitive = (m_CompareMode = vbBinaryCompare)
    Exists = m_col.Exists(Key, bCaseSensitive)
End Function

' 默认成员：Item(Key)
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If IsObject(m_col.Item(Key)) Then
        Set Item = m_col.Item(Key)
    Else
        Item = m_col.Item(Key)
    End If
End Property

Public Property Let Item(ByVal Key As Variant, ByVal Value As Variant)
    m_col.Item(Key) = Value
End Property

Public Property Set Item(ByVal Key As Variant, ByVal Value As Object)
    Set m_col.Item(Key) = Value
End Property

' 删除元素
Public Sub Remove(ByVal Key As Variant)
    m_col.Remove Key
End Sub

' 清空集合
Public Sub RemoveAll()
    m_col.RemoveAll
End Sub

' 返回所有 Key（0 基数组）
Public Function Keys() As String()
    Keys = m_col.Keys
End Function

' 返回所有 Item（0 基数组）
Public Function Items() As Variant()
    Items = m_col.Items
End Function

' 元素数量（Dictionary 中是属性）
Public Property Get Count() As Long
    Count = m_col.Count
End Property

' 重命名 Key（Dictionary 的 Key 语法是 Property Let Key(OldKey) = NewKey）
Public Property Let Key(ByVal OldKey As String, ByVal NewKey As String)
    On Error Resume Next
    m_col.RenameKey OldKey, NewKey
End Property

' For Each 支持
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_col.NewEnum
End Function

'=========================================================================
' 扩展 API（与 cCollection 兼容）
'=========================================================================

' 返回内部原始 Collection 对象
Public Property Get RawCollection() As Collection
    Set RawCollection = m_col.RawCollection
End Property

' 返回内部 cCollection 对象
Public Property Get InnerCollection() As cCollection
    Set InnerCollection = m_col
End Property

' 通过索引获取 Key（用于按索引遍历）
Public Function KeyByIndex(ByVal Index As Long) As String
    KeyByIndex = m_col.KeyByIndex(Index)
End Function

' 更新元素（必须已存在）
Public Function Update(ByVal Item As Variant, ByVal Key As String) As Boolean
    Update = m_col.Update(Item, Key)
End Function

' 检查键是否存在（显式指定大小写敏感）
Public Function ExistsEx(ByVal Key As String, Optional ByVal CompareCase As Boolean = False) As Boolean
    ExistsEx = m_col.Exists(Key, CompareCase)
End Function

' 修改指定 Key 的名称
Public Function RenameKey(ByVal OldKey As String, ByVal NewKey As String) As Boolean
    RenameKey = m_col.RenameKey(OldKey, NewKey)
End Function

' 按 Key 排序
Public Sub SortByKey()
    m_col.SortByKey
End Sub

' 按 Value 排序
Public Sub SortByValue()
    m_col.SortByValue
End Sub

' 获取排序后的 Key 数组
Public Function GetSortedKeys() As String()
    GetSortedKeys = m_col.GetSortedKeys
End Function

' 获取按 Key 排序后的 Value 数组
Public Function GetSortedValuesByKey() As Variant()
    GetSortedValuesByKey = m_col.GetSortedValuesByKey
End Function
