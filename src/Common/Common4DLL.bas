Attribute VB_Name = "Common"
Option Explicit



Public Function Version(Optional HostApp As Object) As String
    If HostApp Is Nothing Then Set HostApp = App
    Version = "v" & HostApp.Major & "." & HostApp.Minor & "." & HostApp.Revision
End Function

