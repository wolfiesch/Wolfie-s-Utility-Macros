' Version Control Excel Add-in
' Main module for the Excel Version Control System
' Provides UI integration with Python backend

Option Explicit

' API declarations for running Python scripts
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
#End If

' Constants
Private Const PYTHON_SCRIPT_PATH As String = "C:\Users\wschoenberger\FuzzySum\VersionControl\version_control.py"
Private Const SW_HIDE As Long = 0
Private Const INFINITE As Long = &HFFFFFFFF

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
    statsJson = ExecutePythonCommand("stats", "")

    If Len(statsJson) > 0 Then
        Call ShowProjectStatsDialog(statsJson)
    Else
        MsgBox "Failed to retrieve project statistics.", vbExclamation, "Version Control"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error retrieving project stats: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' Helper function to execute Python commands
Public Function ExecutePythonCommand(action As String, parameters As String) As String
    On Error GoTo ErrorHandler

    ' Construct command line
    Dim cmd As String
    cmd = "python """ & PYTHON_SCRIPT_PATH & """ --action " & action & _
          " --workbook """ & g_CurrentWorkbookPath & """"

    If Len(parameters) > 0 Then
        cmd = cmd & " " & parameters
    End If

    ' Create temporary file for output
    Dim tempFile As String
    tempFile = Environ("TEMP") & "\vc_output_" & Format(Now, "yyyymmdd_hhmmss") & ".json"

    ' Redirect output to temp file
    cmd = cmd & " > """ & tempFile & """ 2>&1"

    ' Execute command
    Dim result As Long
    result = Shell("cmd /c " & cmd, vbHide)

    ' Wait for completion (simple approach - could be improved)
    Application.Wait Now + TimeValue("00:00:02")

    ' Read result
    Dim fileContent As String
    If Dir(tempFile) <> "" Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Open tempFile For Input As fileNum
        fileContent = Input(LOF(fileNum), fileNum)
        Close fileNum

        ' Clean up temp file
        Kill tempFile
    End If

    ExecutePythonCommand = fileContent
    Exit Function

ErrorHandler:
    ExecutePythonCommand = ""
    If Dir(tempFile) <> "" Then Kill tempFile
End Function

' Dialog helper functions
Private Sub ShowCreateSnapshotDialog()
    ' This would typically load a UserForm
    ' For now, use simple InputBox
    Dim notes As String
    notes = InputBox("Enter notes for this version snapshot (optional):", _
                    "Create Version Snapshot", "")

    If notes <> "" Then
        Dim parameters As String
        parameters = "--notes """ & notes & """"

        ' Show progress indicator
        Application.StatusBar = "Creating version snapshot..."
        Application.ScreenUpdating = False

        Dim result As String
        result = ExecutePythonCommand("create_snapshot", parameters)

        Application.ScreenUpdating = True
        Application.StatusBar = False

        ' Parse and display result
        Call ProcessSnapshotResult(result)
    End If
End Sub

Private Sub ShowCompareVersionDialog()
    ' Get list of available versions first
    Dim versionsJson As String
    versionsJson = ExecutePythonCommand("list_versions", "")

    If Len(versionsJson) > 0 Then
        ' Parse versions and show selection dialog
        Dim versions As Collection
        Set versions = ParseVersionsList(versionsJson)

        If versions.Count > 0 Then
            Dim selectedVersion As String
            selectedVersion = ShowVersionSelector(versions, "Select version to compare with current workbook:")

            If selectedVersion <> "" Then
                ' Execute comparison
                Application.StatusBar = "Comparing workbook versions..."
                Application.ScreenUpdating = False

                Dim result As String
                result = ExecutePythonCommand("compare", "--version """ & selectedVersion & """")

                Application.ScreenUpdating = True
                Application.StatusBar = False

                Call ProcessComparisonResult(result)
            End If
        Else
            MsgBox "No versions found for comparison.", vbInformation, "Version Control"
        End If
    Else
        MsgBox "Failed to retrieve versions list.", vbExclamation, "Version Control"
    End If
End Sub

Private Sub ShowVersionsListDialog()
    ' Get versions list
    Dim versionsJson As String
    versionsJson = ExecutePythonCommand("list_versions", "")

    If Len(versionsJson) > 0 Then
        Call DisplayVersionsList(versionsJson)
    Else
        MsgBox "Failed to retrieve versions list.", vbExclamation, "Version Control"
    End If
End Sub

Private Sub ShowRollbackDialog()
    ' Get list of available versions
    Dim versionsJson As String
    versionsJson = ExecutePythonCommand("list_versions", "")

    If Len(versionsJson) > 0 Then
        Dim versions As Collection
        Set versions = ParseVersionsList(versionsJson)

        If versions.Count > 0 Then
            Dim selectedVersion As String
            selectedVersion = ShowVersionSelector(versions, "Select version to rollback to:")

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
                    result = ExecutePythonCommand("rollback", "--version """ & selectedVersion & """")

                    Application.ScreenUpdating = True
                    Application.StatusBar = False

                    Call ProcessRollbackResult(result)
                End If
            End If
        Else
            MsgBox "No versions available for rollback.", vbInformation, "Version Control"
        End If
    Else
        MsgBox "Failed to retrieve versions list.", vbExclamation, "Version Control"
    End If
End Sub

' Result processing functions
Private Sub ProcessSnapshotResult(jsonResult As String)
    On Error GoTo ErrorHandler

    If InStr(jsonResult, """success"": true") > 0 Then
        ' Extract version name
        Dim versionName As String
        versionName = ExtractJsonValue(jsonResult, "version")

        MsgBox "Version snapshot created successfully!" & vbCrLf & _
               "Version: " & versionName, vbInformation, "Version Control"
    Else
        Dim errorMsg As String
        errorMsg = ExtractJsonValue(jsonResult, "error")
        If errorMsg = "" Then errorMsg = "Unknown error occurred"

        MsgBox "Failed to create snapshot: " & errorMsg, vbCritical, "Version Control Error"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error processing snapshot result: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Private Sub ProcessComparisonResult(jsonResult As String)
    On Error GoTo ErrorHandler

    If InStr(jsonResult, """success"": true") > 0 Then
        ' Extract report path
        Dim reportPath As String
        reportPath = ExtractJsonValue(jsonResult, "report_path")

        If reportPath <> "" And Dir(reportPath) <> "" Then
            Dim response As VbMsgBoxResult
            response = MsgBox("Comparison completed successfully!" & vbCrLf & _
                             "Report saved to: " & reportPath & vbCrLf & vbCrLf & _
                             "Would you like to open the comparison report?", _
                             vbYesNo + vbInformation, "Version Control")

            If response = vbYes Then
                Workbooks.Open reportPath
            End If
        Else
            MsgBox "Comparison completed successfully!", vbInformation, "Version Control"
        End If
    Else
        Dim errorMsg As String
        errorMsg = ExtractJsonValue(jsonResult, "error")
        If errorMsg = "" Then errorMsg = "Unknown error occurred"

        MsgBox "Comparison failed: " & errorMsg, vbCritical, "Version Control Error"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error processing comparison result: " & Err.Description, vbCritical, "Version Control Error"
End Sub

Private Sub ProcessRollbackResult(jsonResult As String)
    On Error GoTo ErrorHandler

    If InStr(jsonResult, """success"": true") > 0 Then
        MsgBox "Rollback completed successfully!" & vbCrLf & _
               "The workbook has been restored to the selected version." & vbCrLf & _
               "You may need to reopen the workbook to see the changes.", _
               vbInformation, "Version Control"
    Else
        Dim errorMsg As String
        errorMsg = ExtractJsonValue(jsonResult, "error")
        If errorMsg = "" Then errorMsg = "Unknown error occurred"

        MsgBox "Rollback failed: " & errorMsg, vbCritical, "Version Control Error"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error processing rollback result: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' Utility functions
Private Function ExtractJsonValue(jsonText As String, key As String) As String
    On Error GoTo ErrorHandler

    Dim startPos As Long
    Dim endPos As Long
    Dim searchPattern As String

    searchPattern = """" & key & """: """
    startPos = InStr(jsonText, searchPattern)

    If startPos > 0 Then
        startPos = startPos + Len(searchPattern)
        endPos = InStr(startPos, jsonText, """")

        If endPos > startPos Then
            ExtractJsonValue = Mid(jsonText, startPos, endPos - startPos)
        End If
    End If

    Exit Function

ErrorHandler:
    ExtractJsonValue = ""
End Function

Private Function ParseVersionsList(jsonText As String) As Collection
    On Error GoTo ErrorHandler

    Dim versions As New Collection

    ' Simple JSON parsing for version list
    ' In production, would use proper JSON parser
    Dim lines() As String
    lines = Split(jsonText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), """value"":") > 0 Then
            Dim versionValue As String
            versionValue = ExtractJsonValue(lines(i), "value")
            If versionValue <> "" Then
                versions.Add versionValue
            End If
        End If
    Next i

    Set ParseVersionsList = versions
    Exit Function

ErrorHandler:
    Set ParseVersionsList = New Collection
End Function

Private Function ShowVersionSelector(versions As Collection, prompt As String) As String
    On Error GoTo ErrorHandler

    ' Create a simple selection dialog
    Dim versionArray() As String
    ReDim versionArray(1 To versions.Count)

    Dim i As Long
    For i = 1 To versions.Count
        versionArray(i) = versions(i)
    Next i

    ' Use InputBox with list (simple approach)
    Dim versionList As String
    For i = 1 To versions.Count
        versionList = versionList & i & ". " & versionArray(i) & vbCrLf
    Next i

    Dim selectedIndex As String
    selectedIndex = InputBox(prompt & vbCrLf & vbCrLf & versionList & vbCrLf & _
                           "Enter the number of the version to select:", _
                           "Select Version")

    If IsNumeric(selectedIndex) Then
        Dim index As Long
        index = CLng(selectedIndex)
        If index >= 1 And index <= versions.Count Then
            ShowVersionSelector = versionArray(index)
        End If
    End If

    Exit Function

ErrorHandler:
    ShowVersionSelector = ""
End Function

Private Sub DisplayVersionsList(jsonText As String)
    ' Create a new worksheet to display versions
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "Version_History_" & Format(Now, "mmdd_hhmm")

    ' Add headers
    ws.Range("A1").Value = "Version"
    ws.Range("B1").Value = "Date"
    ws.Range("C1").Value = "Size (MB)"
    ws.Range("D1").Value = "Notes"

    ' Format headers
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With

    ' Parse and populate data (simplified)
    Dim row As Long
    row = 2

    Dim lines() As String
    lines = Split(jsonText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), """value"":") > 0 Then
            ws.Cells(row, 1).Value = ExtractJsonValue(lines(i), "value")
            row = row + 1
        End If
    Next i

    ' Auto-fit columns
    ws.Columns("A:D").AutoFit

    MsgBox "Version history displayed in new worksheet: " & ws.Name, vbInformation, "Version Control"
End Sub

Private Sub ShowProjectStatsDialog(jsonText As String)
    ' Simple stats display
    Dim statsMsg As String
    statsMsg = "Project Statistics:" & vbCrLf & vbCrLf

    ' Extract key stats
    Dim totalVersions As String
    Dim totalSize As String
    Dim latestVersion As String

    totalVersions = ExtractJsonValue(jsonText, "total_versions")
    totalSize = ExtractJsonValue(jsonText, "total_size_mb")
    latestVersion = ExtractJsonValue(jsonText, "latest_version")

    If totalVersions <> "" Then statsMsg = statsMsg & "Total Versions: " & totalVersions & vbCrLf
    If totalSize <> "" Then statsMsg = statsMsg & "Total Size: " & totalSize & " MB" & vbCrLf
    If latestVersion <> "" Then statsMsg = statsMsg & "Latest Version: " & latestVersion & vbCrLf

    MsgBox statsMsg, vbInformation, "Project Statistics"
End Sub

' Add-in event handlers
Private Sub Workbook_AddinInstall()
    g_VersionControlEnabled = True
    MsgBox "Version Control add-in installed successfully!", vbInformation, "Version Control"
End Sub

Private Sub Workbook_AddinUninstall()
    g_VersionControlEnabled = False
    MsgBox "Version Control add-in uninstalled.", vbInformation, "Version Control"
End Sub