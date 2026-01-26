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
