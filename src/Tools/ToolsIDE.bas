Attribute VB_Name = "ToolsIDE"
Option Explicit

Public Property Get IsIDE() As Boolean
    IsIDE = App.LogMode = 0
End Property
