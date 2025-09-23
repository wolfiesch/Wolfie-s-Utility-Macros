' Ribbon Customization for Version Control
' Implements custom ribbon tab with version control commands

Option Explicit

' Ribbon callback procedures
Public Sub OnCreateSnapshot(control As IRibbonControl)
    Call VersionControlAddin_Simple.CreateVersionSnapshot
End Sub

Public Sub OnCompareVersions(control As IRibbonControl)
    Call VersionControlAddin_Simple.CompareToVersion
End Sub

Public Sub OnListVersions(control As IRibbonControl)
    Call VersionControlAddin_Simple.ListVersions
End Sub

Public Sub OnRollback(control As IRibbonControl)
    Call VersionControlAddin_Simple.RollbackToVersion
End Sub

Public Sub OnShowStats(control As IRibbonControl)
    Call VersionControlAddin_Simple.ShowProjectStats
End Sub

Public Sub OnTestConnection(control As IRibbonControl)
    If VersionControlAddin_Simple.TestPythonConnection() Then
        MsgBox "Python connection test successful!", vbInformation, "Version Control"
    Else
        MsgBox "Python connection test failed. Please check your Python installation and script paths.", _
               vbExclamation, "Version Control"
    End If
End Sub