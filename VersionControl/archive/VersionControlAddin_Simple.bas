' Simplified Version Control Excel Add-in
' Main module for the Excel Version Control System
' Uses simple Shell commands instead of complex API calls

Option Explicit

' Constants
Private Const PYTHON_BRIDGE_PATH As String = "C:\Users\wschoenberger\FuzzySum\VersionControl\vba_python_bridge.py"
Private Const TEMP_DIR As String = "C:\temp\VersionControl\"

' Global variables
Public g_VersionControlEnabled As Boolean
Public g_CurrentWorkbookPath As String

' Main entry points for ribbon/menu integration
Public Sub CreateVersionSnapshot()
    On Error GoTo ErrorHandler

    ' Get current workbook path
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    If Not ActiveWorkbook.Saved Then
        Dim response As VbMsgBoxResult
        response = MsgBox("The workbook has unsaved changes. Save before creating snapshot?", _
                         vbYesNoCancel + vbQuestion, "Version Control")

        Select Case response
            Case vbYes
                ActiveWorkbook.Save
            Case vbCancel
                Exit Sub
            Case vbNo
                ' Continue without saving
        End Select
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    ' Show snapshot creation dialog
    Call ShowCreateSnapshotDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error creating version snapshot: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Public Sub CompareToVersion()
    On Error GoTo ErrorHandler

    ' Validate active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    ' Show version comparison dialog
    Call ShowCompareVersionDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error comparing versions: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Public Sub ListVersions()
    On Error GoTo ErrorHandler

    ' Validate active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    ' Show versions list dialog
    Call ShowVersionsListDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error listing versions: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Public Sub RollbackToVersion()
    On Error GoTo ErrorHandler

    ' Validate active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    ' Show rollback dialog with warning
    Dim response As VbMsgBoxResult
    response = MsgBox("Rolling back will replace the current workbook with a previous version. " & _
                     "This action cannot be undone. Continue?", _
                     vbYesNo + vbExclamation, "Version Control - Rollback Warning")

    If response = vbYes Then
        Call ShowRollbackDialog
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error during rollback: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Public Sub ShowProjectStats()
    On Error GoTo ErrorHandler

    ' Validate active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    ' Get project statistics from Python backend
    Dim statsJson As String
    statsJson = ExecutePythonCommandSimple("stats", "")

    If Len(statsJson) > 0 Then
        Call ShowProjectStatsDialog(statsJson)
    Else
        MsgBox "Failed to retrieve project statistics.", vbExclamation, "Version Control"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error retrieving project stats: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' Simplified Python command execution using Shell
Public Function ExecutePythonCommandSimple(action As String, parameters As String) As String
    On Error GoTo ErrorHandler

    ' Ensure temp directory exists
    Call EnsureTempDirectory

    ' Create unique temporary file for output
    Dim outputFile As String
    outputFile = TEMP_DIR & "output_" & Format(Now, "yyyymmdd_hhmmss_") & Int(Rnd * 1000) & ".json"

    ' Construct command line
    Dim cmd As String
    cmd = "python """ & PYTHON_BRIDGE_PATH & """ --workbook """ & g_CurrentWorkbookPath & """ --action " & action & _
          " --output """ & outputFile & """"

    If Len(parameters) > 0 Then
        cmd = cmd & " " & parameters
    End If

    ' Execute command using Shell
    Dim taskId As Long
    taskId = Shell("cmd /c " & cmd, vbHide)

    ' Wait for completion (simple approach)
    Application.Wait Now + TimeValue("00:00:03")

    ' Read result
    Dim fileContent As String
    If Dir(outputFile) <> "" Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Open outputFile For Input As fileNum
        If LOF(fileNum) > 0 Then
            fileContent = Input(LOF(fileNum), fileNum)
        End If
        Close fileNum

        ' Clean up temp file
        Kill outputFile
    End If

    ExecutePythonCommandSimple = fileContent
    Exit Function

ErrorHandler:
    ExecutePythonCommandSimple = ""
    If Dir(outputFile) <> "" Then Kill outputFile
End Function

' Ensure temporary directory exists
Private Sub EnsureTempDirectory()
    On Error Resume Next
    If Dir(TEMP_DIR, vbDirectory) = "" Then
        MkDir Left(TEMP_DIR, Len(TEMP_DIR) - 1) ' Remove trailing backslash
    End If
    On Error GoTo 0
End Sub

' Dialog helper functions (simplified versions)
Private Sub ShowCreateSnapshotDialog()
    Dim notes As String
    notes = InputBox("Enter notes for this version snapshot (optional):", _
                    "Create Version Snapshot", "")

    If notes <> "" Or MsgBox("Create snapshot without notes?", vbYesNo + vbQuestion) = vbYes Then
        Dim parameters As String
        If notes <> "" Then
            parameters = "--notes """ & notes & """"
        End If

        ' Show progress indicator
        Application.StatusBar = "Creating version snapshot..."
        Application.ScreenUpdating = False

        Dim result As String
        result = ExecutePythonCommandSimple("create_snapshot", parameters)

        Application.ScreenUpdating = True
        Application.StatusBar = False

        ' Parse and display result
        Call ProcessSnapshotResult(result)
    End If
End Sub

Private Sub ShowCompareVersionDialog()
    ' Get list of available versions first
    Dim versionsJson As String
    versionsJson = ExecutePythonCommandSimple("list_versions", "")

    If Len(versionsJson) > 0 And InStr(versionsJson, "success") > 0 Then
        ' Simple version selection using InputBox
        Dim selectedVersion As String
        selectedVersion = InputBox("Enter version name to compare with (e.g., v001):", _
                                  "Compare Versions", "v001")

        If selectedVersion <> "" Then
            ' Execute comparison
            Application.StatusBar = "Comparing workbook versions..."
            Application.ScreenUpdating = False

            Dim result As String
            result = ExecutePythonCommandSimple("compare", "--version """ & selectedVersion & """")

            Application.ScreenUpdating = True
            Application.StatusBar = False

            Call ProcessComparisonResult(result)
        End If
    Else
        MsgBox "No versions found for comparison or failed to retrieve versions list.", vbInformation, "Version Control"
    End If
End Sub

Private Sub ShowVersionsListDialog()
    ' Get versions list and display in message box
    Dim versionsJson As String
    versionsJson = ExecutePythonCommandSimple("list_versions", "")

    If Len(versionsJson) > 0 Then
        ' Simple display - could be enhanced with a proper form
        MsgBox "Version list retrieved. Check the Python output for details.", vbInformation, "Versions List"
    Else
        MsgBox "Failed to retrieve versions list.", vbExclamation, "Version Control"
    End If
End Sub

Private Sub ShowRollbackDialog()
    ' Simple rollback with version input
    Dim selectedVersion As String
    selectedVersion = InputBox("Enter version name to rollback to (e.g., v001):", _
                              "Rollback to Version", "v001")

    If selectedVersion <> "" Then
        ' Final confirmation
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to rollback to version " & selectedVersion & "?", _
                         vbYesNo + vbCritical, "Confirm Rollback")

        If response = vbYes Then
            ' Execute rollback
            Application.StatusBar = "Rolling back to version " & selectedVersion & "..."
            Application.ScreenUpdating = False

            Dim result As String
            result = ExecutePythonCommandSimple("rollback", "--version """ & selectedVersion & """")

            Application.ScreenUpdating = True
            Application.StatusBar = False

            Call ProcessRollbackResult(result)
        End If
    End If
End Sub

' Result processing functions (simplified)
Private Sub ProcessSnapshotResult(jsonResult As String)
    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Version snapshot created successfully!", vbInformation, "Version Control"
    Else
        MsgBox "Failed to create snapshot. Check the output for details.", vbCritical, "Version Control Error"
    End If
End Sub

Private Sub ProcessComparisonResult(jsonResult As String)
    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Comparison completed successfully! Check the Reports folder for detailed comparison.", vbInformation, "Version Control"
    Else
        MsgBox "Comparison failed. Check the output for details.", vbCritical, "Version Control Error"
    End If
End Sub

Private Sub ProcessRollbackResult(jsonResult As String)
    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Rollback completed successfully! You may need to reopen the workbook.", vbInformation, "Version Control"
    Else
        MsgBox "Rollback failed. Check the output for details.", vbCritical, "Version Control Error"
    End If
End Sub

Private Sub ShowProjectStatsDialog(jsonText As String)
    ' Simple stats display
    If InStr(jsonText, """success"": true") > 0 Then
        MsgBox "Project statistics retrieved successfully. Check Python output for details.", vbInformation, "Project Statistics"
    Else
        MsgBox "Failed to retrieve project statistics.", vbExclamation, "Version Control"
    End If
End Sub

' Test function to verify Python connection
Public Function TestPythonConnection() As Boolean
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        TestPythonConnection = False
        Exit Function
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName
    Dim testResult As String
    testResult = ExecutePythonCommandSimple("stats", "")

    TestPythonConnection = (Len(testResult) > 0 And InStr(testResult, "success") > 0)
    Exit Function

ErrorHandler:
    TestPythonConnection = False
End Function

' Add-in event handlers
Private Sub Workbook_AddinInstall()
    g_VersionControlEnabled = True
    MsgBox "Version Control add-in installed successfully!", vbInformation, "Version Control"
End Sub

Private Sub Workbook_AddinUninstall()
    g_VersionControlEnabled = False
    MsgBox "Version Control add-in uninstalled.", vbInformation, "Version Control"
End Sub