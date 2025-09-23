' Version Control Excel Add-in - Windows Script Host Version
' Uses WSH instead of Shell to bypass corporate security restrictions

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

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName
    Call ShowCompareVersionDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error comparing versions: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Public Sub ListVersions()
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName
    Call ShowVersionsListDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error listing versions: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Public Sub RollbackToVersion()
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

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

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    Dim statsJson As String
    statsJson = ExecutePythonCommandWSH("stats", "")

    If Len(statsJson) > 0 Then
        Call ShowProjectStatsDialog(statsJson)
    Else
        MsgBox "Failed to retrieve project statistics.", vbExclamation, "Version Control"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error retrieving project stats: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' WSH-based Python command execution to bypass Shell restrictions
Public Function ExecutePythonCommandWSH(action As String, parameters As String) As String
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

    ' Use Windows Script Host instead of Shell
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    ' Execute with different methods to avoid blocking
    Dim execResult As Object
    On Error Resume Next

    ' Method 1: Try PowerShell execution (often less restricted)
    Dim psCmd As String
    psCmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Command """ & cmd & """"
    Set execResult = wsh.Exec(psCmd)

    If Err.Number <> 0 Then
        Err.Clear
        ' Method 2: Try direct Python execution
        Set execResult = wsh.Exec(cmd)

        If Err.Number <> 0 Then
            Err.Clear
            ' Method 3: Try cmd.exe with different approach
            wsh.Run "cmd /c " & cmd, 0, True  ' 0=hidden, True=wait
        End If
    End If

    On Error GoTo ErrorHandler

    ' Wait for completion if using Exec method
    If Not execResult Is Nothing Then
        Do While execResult.Status = 0  ' 0 = still running
            Application.Wait Now + TimeValue("00:00:01")
            DoEvents
        Loop
    Else
        ' If using Run method, add extra wait time
        Application.Wait Now + TimeValue("00:00:03")
    End If

    ' Read result file
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

    ExecutePythonCommandWSH = fileContent
    Set wsh = Nothing
    Exit Function

ErrorHandler:
    ExecutePythonCommandWSH = ""
    If Dir(outputFile) <> "" Then
        On Error Resume Next
        Kill outputFile
        On Error GoTo 0
    End If
    Set wsh = Nothing
End Function

' Alternative execution method using batch file (often less restricted)
Public Function ExecutePythonCommandBatch(action As String, parameters As String) As String
    On Error GoTo ErrorHandler

    Call EnsureTempDirectory

    ' Create batch file
    Dim batchFile As String
    Dim outputFile As String
    batchFile = TEMP_DIR & "vc_cmd_" & Format(Now, "yyyymmdd_hhmmss") & ".bat"
    outputFile = TEMP_DIR & "output_" & Format(Now, "yyyymmdd_hhmmss_") & Int(Rnd * 1000) & ".json"

    ' Write batch file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open batchFile For Output As fileNum
    Print #fileNum, "@echo off"
    Print #fileNum, "cd /d """ & Left(PYTHON_BRIDGE_PATH, InStrRev(PYTHON_BRIDGE_PATH, "\") - 1) & """"

    Dim cmd As String
    cmd = "python """ & PYTHON_BRIDGE_PATH & """ --workbook """ & g_CurrentWorkbookPath & """ --action " & action & _
          " --output """ & outputFile & """"
    If Len(parameters) > 0 Then
        cmd = cmd & " " & parameters
    End If

    Print #fileNum, cmd
    Close fileNum

    ' Execute batch file using WSH
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & batchFile & """", 0, True  ' Hidden and wait for completion

    ' Read result
    Dim fileContent As String
    If Dir(outputFile) <> "" Then
        fileNum = FreeFile
        Open outputFile For Input As fileNum
        If LOF(fileNum) > 0 Then
            fileContent = Input(LOF(fileNum), fileNum)
        End If
        Close fileNum
        Kill outputFile
    End If

    ' Clean up batch file
    Kill batchFile

    ExecutePythonCommandBatch = fileContent
    Set wsh = Nothing
    Exit Function

ErrorHandler:
    ExecutePythonCommandBatch = ""
    On Error Resume Next
    If Dir(outputFile) <> "" Then Kill outputFile
    If Dir(batchFile) <> "" Then Kill batchFile
    On Error GoTo 0
    Set wsh = Nothing
End Function

' Ensure temporary directory exists
Private Sub EnsureTempDirectory()
    On Error Resume Next
    If Dir(TEMP_DIR, vbDirectory) = "" Then
        ' Create directory hierarchy
        Dim parts As Variant
        parts = Split(TEMP_DIR, "\")
        Dim currentPath As String
        Dim i As Integer

        For i = 0 To UBound(parts) - 1  ' -1 because last element is empty due to trailing \
            If parts(i) <> "" Then
                currentPath = currentPath & parts(i) & "\"
                If Dir(currentPath, vbDirectory) = "" Then
                    MkDir currentPath
                End If
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

' Dialog helper functions
Private Sub ShowCreateSnapshotDialog()
    Dim notes As String
    notes = InputBox("Enter notes for this version snapshot (optional):", _
                    "Create Version Snapshot", "")

    If notes <> "" Or MsgBox("Create snapshot without notes?", vbYesNo + vbQuestion) = vbYes Then
        Dim parameters As String
        If notes <> "" Then
            parameters = "--notes """ & notes & """"
        End If

        Application.StatusBar = "Creating version snapshot..."
        Application.ScreenUpdating = False

        ' Try WSH method first, then batch method if it fails
        Dim result As String
        result = ExecutePythonCommandWSH("create_snapshot", parameters)

        If Len(result) = 0 Or InStr(result, "error") > 0 Then
            ' Try batch method as fallback
            result = ExecutePythonCommandBatch("create_snapshot", parameters)
        End If

        Application.ScreenUpdating = True
        Application.StatusBar = False

        Call ProcessSnapshotResult(result)
    End If
End Sub

Private Sub ShowCompareVersionDialog()
    Dim selectedVersion As String
    selectedVersion = InputBox("Enter version name to compare with (e.g., v001):", _
                              "Compare Versions", "v001")

    If selectedVersion <> "" Then
        Application.StatusBar = "Comparing workbook versions..."
        Application.ScreenUpdating = False

        Dim result As String
        result = ExecutePythonCommandWSH("compare", "--version """ & selectedVersion & """")

        If Len(result) = 0 Or InStr(result, "error") > 0 Then
            result = ExecutePythonCommandBatch("compare", "--version """ & selectedVersion & """")
        End If

        Application.ScreenUpdating = True
        Application.StatusBar = False

        Call ProcessComparisonResult(result)
    End If
End Sub

Private Sub ShowVersionsListDialog()
    Dim result As String
    result = ExecutePythonCommandWSH("list_versions", "")

    If Len(result) = 0 Or InStr(result, "error") > 0 Then
        result = ExecutePythonCommandBatch("list_versions", "")
    End If

    If Len(result) > 0 Then
        MsgBox "Version list retrieved. Check the Python output for details.", vbInformation, "Versions List"
    Else
        MsgBox "Failed to retrieve versions list.", vbExclamation, "Version Control"
    End If
End Sub

Private Sub ShowRollbackDialog()
    Dim selectedVersion As String
    selectedVersion = InputBox("Enter version name to rollback to (e.g., v001):", _
                              "Rollback to Version", "v001")

    If selectedVersion <> "" Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to rollback to version " & selectedVersion & "?", _
                         vbYesNo + vbCritical, "Confirm Rollback")

        If response = vbYes Then
            Application.StatusBar = "Rolling back to version " & selectedVersion & "..."
            Application.ScreenUpdating = False

            Dim result As String
            result = ExecutePythonCommandWSH("rollback", "--version """ & selectedVersion & """")

            If Len(result) = 0 Or InStr(result, "error") > 0 Then
                result = ExecutePythonCommandBatch("rollback", "--version """ & selectedVersion & """")
            End If

            Application.ScreenUpdating = True
            Application.StatusBar = False

            Call ProcessRollbackResult(result)
        End If
    End If
End Sub

' Result processing functions
Private Sub ProcessSnapshotResult(jsonResult As String)
    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Version snapshot created successfully!", vbInformation, "Version Control"
    ElseIf Len(jsonResult) = 0 Then
        MsgBox "No response from Python backend. Please check:" & vbCrLf & _
               "1. Python is installed and in PATH" & vbCrLf & _
               "2. Required Python packages are installed" & vbCrLf & _
               "3. Script paths are correct", vbExclamation, "Version Control"
    Else
        MsgBox "Failed to create snapshot. Check the output for details.", vbCritical, "Version Control Error"
    End If
End Sub

Private Sub ProcessComparisonResult(jsonResult As String)
    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Comparison completed successfully! Check the Reports folder for detailed comparison.", vbInformation, "Version Control"
    ElseIf Len(jsonResult) = 0 Then
        MsgBox "No response from Python backend.", vbExclamation, "Version Control"
    Else
        MsgBox "Comparison failed. Check the output for details.", vbCritical, "Version Control Error"
    End If
End Sub

Private Sub ProcessRollbackResult(jsonResult As String)
    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Rollback completed successfully! You may need to reopen the workbook.", vbInformation, "Version Control"
    ElseIf Len(jsonResult) = 0 Then
        MsgBox "No response from Python backend.", vbExclamation, "Version Control"
    Else
        MsgBox "Rollback failed. Check the output for details.", vbCritical, "Version Control Error"
    End If
End Sub

Private Sub ShowProjectStatsDialog(jsonText As String)
    If InStr(jsonText, """success"": true") > 0 Then
        MsgBox "Project statistics retrieved successfully. Check Python output for details.", vbInformation, "Project Statistics"
    ElseIf Len(jsonText) = 0 Then
        MsgBox "No response from Python backend.", vbExclamation, "Version Control"
    Else
        MsgBox "Failed to retrieve project statistics.", vbExclamation, "Version Control"
    End If
End Sub

' Test function to verify connection
Public Function TestPythonConnection() As Boolean
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        TestPythonConnection = False
        Exit Function
    End If

    g_CurrentWorkbookPath = ActiveWorkbook.FullName

    ' Try both methods
    Dim testResult As String
    testResult = ExecutePythonCommandWSH("stats", "")

    If Len(testResult) = 0 Or InStr(testResult, "error") > 0 Then
        testResult = ExecutePythonCommandBatch("stats", "")
    End If

    TestPythonConnection = (Len(testResult) > 0 And InStr(testResult, "success") > 0)
    Exit Function

ErrorHandler:
    TestPythonConnection = False
End Function

' Add-in event handlers
Private Sub Workbook_AddinInstall()
    g_VersionControlEnabled = True
    MsgBox "Version Control add-in (WSH version) installed successfully!", vbInformation, "Version Control"
End Sub

Private Sub Workbook_AddinUninstall()
    g_VersionControlEnabled = False
    MsgBox "Version Control add-in uninstalled.", vbInformation, "Version Control"
End Sub