VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSTransactionObject
END
Attribute VB_Name = "cHttpServerFileCacheEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 静态文件缓存条目（懒缓存机制）
' 默认不缓存任何文件，首次访问且未超出缓存上限时才加入

' 文件内容（Byte数组）
' VB6限制：Public动态数组不能作为类公开成员，必须用Private+Property暴露
Private m_Content() As Byte

' ETag 值（基于文件修改时间+大小计算）
Public ETag As String
' 文件最后修改时间（用于 Last-Modified 头）
Public LastModified As Date
' 文件大小（字节）
Public FileSize As Long
' 缓存时间（用于 Expires 头）/TTL 过期判断
Public CachedAt As Date

' 文件内容属性（只读：通过 Let 写入，Get 读取数组引用）
Public Property Get Content() As Byte()
    Content = m_Content
End Property

Public Property Let Content(RHS() As Byte)
    m_Content = RHS
End Property
