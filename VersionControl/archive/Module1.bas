Attribute VB_Name = "Module1"
' Version Control Excel Add-in - VBA Only Version
' Pure VBA implementation for maximum compatibility in restricted environments
' No external dependencies or process execution required

Option Explicit

' Constants
Private Const VERSION_FOLDER As String = "C:\Users\wschoenberger\FuzzySum\VersionControl\Versions\"
Private Const METADATA_FOLDER As String = "C:\Users\wschoenberger\FuzzySum\VersionControl\Versions\Metadata\"

' Global variables
Public g_VersionControlEnabled As Boolean
Public g_CurrentWorkbookPath As String
Public g_NextVersionNumber As Long

' Main entry points for ribbon/menu integration
Public Sub CreateVersionSnapshot()
    On Error GoTo ErrorHandler

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
    Call ShowCreateSnapshotDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error creating version snapshot: " & Err.Description, vbCritical, "Version Control Error"
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
    Call ShowProjectStatsDialog

    Exit Sub

ErrorHandler:
    MsgBox "Error retrieving project stats: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' Core VBA-only functions
Private Function CreateSnapshotVBA(notes As String) As Boolean
    On Error GoTo ErrorHandler

    ' Ensure directories exist
    Call EnsureDirectoriesExist

    ' Get next version number
    Dim versionNumber As Long
    versionNumber = GetNextVersionNumber()

    ' Create version name and paths
    Dim versionName As String
    Dim timestamp As String
    timestamp = Format(Now, "yyyymmdd_hhmmss")
    versionName = "v" & Format(versionNumber, "000")

    Dim snapshotPath As String
    snapshotPath = VERSION_FOLDER & versionName & "_" & timestamp & ".xlsx"

    ' Create snapshot by copying the workbook
    Application.StatusBar = "Creating version snapshot..."
    Application.DisplayAlerts = False

    ' Save current workbook first
    ActiveWorkbook.Save

    ' Create a copy
    ActiveWorkbook.SaveCopyAs snapshotPath

    Application.DisplayAlerts = True

    ' Create metadata
    Call CreateVersionMetadata(versionName, timestamp, snapshotPath, notes)

    ' Update next version number
    Call SaveNextVersionNumber(versionNumber + 1)

    CreateSnapshotVBA = True
    Application.StatusBar = "Snapshot created: " & versionName

    Exit Function

ErrorHandler:
    Application.DisplayAlerts = True
    Application.StatusBar = False
    CreateSnapshotVBA = False
    MsgBox "Error creating snapshot: " & Err.Description, vbCritical, "Version Control Error"
End Function

Private Sub CreateVersionMetadata(versionName As String, timestamp As String, filePath As String, notes As String)
    On Error GoTo ErrorHandler

    Dim metadataPath As String
    metadataPath = METADATA_FOLDER & versionName & ".txt"

    Dim fileNum As Integer
    fileNum = FreeFile

    Open metadataPath For Output As fileNum
    Print #fileNum, "Version: " & versionName
    Print #fileNum, "Created: " & Now
    Print #fileNum, "Timestamp: " & timestamp
    Print #fileNum, "File: " & filePath
    Print #fileNum, "Original: " & g_CurrentWorkbookPath
    Print #fileNum, "Size: " & FileLen(filePath)
    Print #fileNum, "Notes: " & notes
    Print #fileNum, "User: " & Environ("USERNAME")
    Print #fileNum, "Computer: " & Environ("COMPUTERNAME")
    Close fileNum

    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close fileNum
End Sub

Private Function GetVersionsList() As Collection
    On Error GoTo ErrorHandler

    Dim versions As New Collection
    Dim searchPath As String
    searchPath = METADATA_FOLDER & "*.txt"

    Dim fileName As String
    fileName = Dir(searchPath)

    Do While fileName <> ""
        If Left(fileName, 1) = "v" Then  ' Version files start with 'v'
            Dim versionInfo As Dictionary
            Set versionInfo = ParseMetadataFile(METADATA_FOLDER & fileName)
            If Not versionInfo Is Nothing Then
                versions.Add versionInfo
            End If
        End If
        fileName = Dir
    Loop

    Set GetVersionsList = versions
    Exit Function

ErrorHandler:
    Set GetVersionsList = New Collection
End Function

Private Function ParseMetadataFile(filePath As String) As Dictionary
    On Error GoTo ErrorHandler

    Dim versionInfo As New Scripting.Dictionary

    Dim fileNum As Integer
    fileNum = FreeFile

    Open filePath For Input As fileNum

    Do While Not EOF(fileNum)
        Dim line As String
        Line Input #fileNum, line

        If InStr(line, ":") > 0 Then
            Dim parts As Variant
            parts = Split(line, ":", 2)
            If UBound(parts) >= 1 Then
                versionInfo(Trim(parts(0))) = Trim(parts(1))
            End If
        End If
    Loop

    Close fileNum

    Set ParseMetadataFile = versionInfo
    Exit Function

ErrorHandler:
    If fileNum > 0 Then Close fileNum
    Set ParseMetadataFile = Nothing
End Function

Private Function GetNextVersionNumber() As Long
    On Error GoTo ErrorHandler

    Dim versionFile As String
    versionFile = METADATA_FOLDER & "next_version.txt"

    If Dir(versionFile) <> "" Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Open versionFile For Input As fileNum
        Input #fileNum, GetNextVersionNumber
        Close fileNum
    Else
        GetNextVersionNumber = 1
    End If

    Exit Function

ErrorHandler:
    GetNextVersionNumber = 1
End Function

Private Sub SaveNextVersionNumber(nextNumber As Long)
    On Error Resume Next

    Dim versionFile As String
    versionFile = METADATA_FOLDER & "next_version.txt"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open versionFile For Output As fileNum
    Print #fileNum, nextNumber
    Close fileNum
End Sub

Private Sub EnsureDirectoriesExist()
    On Error Resume Next

    ' Create version folder
    If Dir(VERSION_FOLDER, vbDirectory) = "" Then
        Call CreateDirectoryRecursive(VERSION_FOLDER)
    End If

    ' Create metadata folder
    If Dir(METADATA_FOLDER, vbDirectory) = "" Then
        Call CreateDirectoryRecursive(METADATA_FOLDER)
    End If

    On Error GoTo 0
End Sub

Private Sub CreateDirectoryRecursive(dirPath As String)
    On Error Resume Next

    Dim parts As Variant
    parts = Split(dirPath, "\")
    Dim currentPath As String
    Dim i As Integer

    For i = 0 To UBound(parts)
        If parts(i) <> "" Then
            currentPath = currentPath & parts(i) & "\"
            If Dir(currentPath, vbDirectory) = "" Then
                MkDir currentPath
            End If
        End If
    Next i

    On Error GoTo 0
End Sub

' Dialog functions
Private Sub ShowCreateSnapshotDialog()
    Dim notes As String
    notes = InputBox("Enter notes for this version snapshot (optional):", _
                    "Create Version Snapshot", "")

    If notes <> "" Or MsgBox("Create snapshot without notes?", vbYesNo + vbQuestion) = vbYes Then
        If CreateSnapshotVBA(notes) Then
            MsgBox "Version snapshot created successfully!", vbInformation, "Version Control"
        End If
    End If
End Sub

Private Sub ShowVersionsListDialog()
    Dim versions As Collection
    Set versions = GetVersionsList()

    If versions.Count = 0 Then
        MsgBox "No versions found.", vbInformation, "Version Control"
        Exit Sub
    End If

    ' Create simple list display
    Dim versionList As String
    versionList = "Available Versions:" & vbCrLf & vbCrLf

    Dim i As Integer
    For i = 1 To versions.Count
        Dim version As Dictionary
        Set version = versions(i)
        versionList = versionList & version("Version") & " - " & version("Created")
        If version("Notes") <> "" Then
            versionList = versionList & " - " & version("Notes")
        End If
        versionList = versionList & vbCrLf
    Next i

    MsgBox versionList, vbInformation, "Version Control - Available Versions"
End Sub

Private Sub ShowCompareVersionDialog()
    Dim versions As Collection
    Set versions = GetVersionsList()

    If versions.Count = 0 Then
        MsgBox "No versions available for comparison.", vbInformation, "Version Control"
        Exit Sub
    End If

    ' Simple version selection
    Dim selectedVersion As String
    selectedVersion = InputBox("Enter version name to compare with (e.g., v001):", _
                              "Compare Versions", "v001")

    If selectedVersion <> "" Then
        Call PerformBasicComparison(selectedVersion)
    End If
End Sub

Private Sub ShowRollbackDialog()
    Dim versions As Collection
    Set versions = GetVersionsList()

    If versions.Count = 0 Then
        MsgBox "No versions available for rollback.", vbInformation, "Version Control"
        Exit Sub
    End If

    Dim selectedVersion As String
    selectedVersion = InputBox("Enter version name to rollback to (e.g., v001):", _
                              "Rollback to Version", "v001")

    If selectedVersion <> "" Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to rollback to version " & selectedVersion & "?", _
                         vbYesNo + vbCritical, "Confirm Rollback")

        If response = vbYes Then
            Call PerformRollback(selectedVersion)
        End If
    End If
End Sub

Private Sub ShowProjectStatsDialog()
    Dim versions As Collection
    Set versions = GetVersionsList()

    Dim statsMsg As String
    statsMsg = "Project Statistics:" & vbCrLf & vbCrLf
    statsMsg = statsMsg & "Total Versions: " & versions.Count & vbCrLf

    If versions.Count > 0 Then
        ' Calculate total size
        Dim totalSize As Long
        totalSize = 0

        Dim i As Integer
        For i = 1 To versions.Count
            Dim version As Dictionary
            Set version = versions(i)
            If version.Exists("Size") Then
                totalSize = totalSize + CLng(version("Size"))
            End If
        Next i

        statsMsg = statsMsg & "Total Size: " & Format(totalSize / 1024 / 1024, "0.0") & " MB" & vbCrLf

        ' Latest version
        Dim latestVersion As Dictionary
        Set latestVersion = versions(versions.Count)
        statsMsg = statsMsg & "Latest Version: " & latestVersion("Version") & vbCrLf
        statsMsg = statsMsg & "Created: " & latestVersion("Created") & vbCrLf
    End If

    MsgBox statsMsg, vbInformation, "Project Statistics"
End Sub

' Basic comparison function
Private Sub PerformBasicComparison(versionName As String)
    On Error GoTo ErrorHandler

    ' Find version file
    Dim versionInfo As Dictionary
    Set versionInfo = ParseMetadataFile(METADATA_FOLDER & versionName & ".txt")

    If versionInfo Is Nothing Then
        MsgBox "Version " & versionName & " not found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    Dim versionFilePath As String
    versionFilePath = versionInfo("File")

    If Dir(versionFilePath) = "" Then
        MsgBox "Version file not found: " & versionFilePath, vbExclamation, "Version Control"
        Exit Sub
    End If

    ' Simple comparison message
    Dim compareMsg As String
    compareMsg = "Basic Comparison:" & vbCrLf & vbCrLf
    compareMsg = compareMsg & "Current File: " & g_CurrentWorkbookPath & vbCrLf
    compareMsg = compareMsg & "Current Size: " & Format(FileLen(g_CurrentWorkbookPath) / 1024, "0.0") & " KB" & vbCrLf & vbCrLf
    compareMsg = compareMsg & "Version " & versionName & ": " & versionFilePath & vbCrLf
    compareMsg = compareMsg & "Version Size: " & Format(FileLen(versionFilePath) / 1024, "0.0") & " KB" & vbCrLf
    compareMsg = compareMsg & "Created: " & versionInfo("Created") & vbCrLf & vbCrLf

    If FileLen(g_CurrentWorkbookPath) = FileLen(versionFilePath) Then
        compareMsg = compareMsg & "Files are the same size."
    Else
        compareMsg = compareMsg & "Files are different sizes."
    End If

    MsgBox compareMsg, vbInformation, "Version Comparison"

    ' Offer to open version for manual comparison
    Dim response As VbMsgBoxResult
    response = MsgBox("Would you like to open the version file for manual comparison?", _
                     vbYesNo + vbQuestion, "Version Control")

    If response = vbYes Then
        Workbooks.Open versionFilePath
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error during comparison: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' Rollback function
Private Sub PerformRollback(versionName As String)
    On Error GoTo ErrorHandler

    ' Find version file
    Dim versionInfo As Dictionary
    Set versionInfo = ParseMetadataFile(METADATA_FOLDER & versionName & ".txt")

    If versionInfo Is Nothing Then
        MsgBox "Version " & versionName & " not found.", vbExclamation, "Version Control"
        Exit Sub
    End If

    Dim versionFilePath As String
    versionFilePath = versionInfo("File")

    If Dir(versionFilePath) = "" Then
        MsgBox "Version file not found: " & versionFilePath, vbExclamation, "Version Control"
        Exit Sub
    End If

    ' Create backup of current file
    Application.StatusBar = "Creating backup of current file..."
    Application.DisplayAlerts = False

    Dim backupPath As String
    backupPath = g_CurrentWorkbookPath & ".backup_" & Format(Now, "yyyymmdd_hhmmss")
    ActiveWorkbook.SaveCopyAs backupPath

    ' Close current workbook
    ActiveWorkbook.Close SaveChanges:=False

    ' Copy version file to current location
    FileCopy versionFilePath, g_CurrentWorkbookPath

    ' Reopen the restored file
    Workbooks.Open g_CurrentWorkbookPath

    Application.DisplayAlerts = True
    Application.StatusBar = False

    MsgBox "Rollback completed successfully!" & vbCrLf & _
           "Backup of original file saved as:" & vbCrLf & backupPath, _
           vbInformation, "Version Control"

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.StatusBar = False
    MsgBox "Error during rollback: " & Err.Description, vbCritical, "Version Control Error"
End Sub

' Test function
Public Function TestVBAOnlySystem() As Boolean
    On Error GoTo ErrorHandler

    ' Test directory creation
    Call EnsureDirectoriesExist

    ' Test version numbering
    Dim nextVersion As Long
    nextVersion = GetNextVersionNumber()

    TestVBAOnlySystem = True
    MsgBox "VBA-only version control system test passed!" & vbCrLf & _
           "Next version number: " & nextVersion, vbInformation, "Version Control Test"

    Exit Function

ErrorHandler:
    TestVBAOnlySystem = False
    MsgBox "VBA-only system test failed: " & Err.Description, vbExclamation, "Version Control Test"
End Function

' Add-in event handlers
Private Sub Workbook_AddinInstall()
    g_VersionControlEnabled = True
    MsgBox "Version Control add-in (VBA-only version) installed successfully!" & vbCrLf & _
           "This version works entirely within Excel with no external dependencies.", _
           vbInformation, "Version Control"
End Sub

Private Sub Workbook_AddinUninstall()
    g_VersionControlEnabled = False
    MsgBox "Version Control add-in uninstalled.", vbInformation, "Version Control"
End Sub
