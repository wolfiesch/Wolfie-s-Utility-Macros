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
Public g_DebugMode As Boolean  ' Enable debug logging

' Simple diagnostic function to test Directory and Dictionary functionality
Public Sub DiagnoseVersionControl()
    Debug.Print "=== DIAGNOSING VERSION CONTROL ISSUES ==="
    Debug.Print "Time: " & Now
    
    ' Test 1: Check metadata folder
    Debug.Print "1. Testing metadata folder access..."
    Debug.Print "   Metadata folder: " & METADATA_FOLDER
    If Dir(METADATA_FOLDER, vbDirectory) = "" Then
        Debug.Print "   ERROR: Metadata folder does not exist!"
        Exit Sub
    Else
        Debug.Print "   OK: Metadata folder exists"
    End If
    
    ' Test 2: Test Dir() function
    Debug.Print "2. Testing Dir() function..."
    Dim searchPath As String
    searchPath = METADATA_FOLDER & "*.txt"
    Debug.Print "   Search path: " & searchPath
    
    Dim fileName As String
    fileName = Dir(searchPath)
    Dim fileCount As Integer
    fileCount = 0
    
    Do While fileName <> ""
        fileCount = fileCount + 1
        Debug.Print "   Found file " & fileCount & ": " & fileName
        
        ' Test if it's a version file
        If Left(fileName, 1) = "v" And Right(fileName, 4) = ".txt" Then
            Debug.Print "     -> This is a version file"
            
            ' Test 3: Try to create a Dictionary object using late binding
            On Error Resume Next
            Dim testDict As Object
            Set testDict = CreateObject("Scripting.Dictionary")
            If Err.Number <> 0 Then
                Debug.Print "     -> ERROR: Cannot create Scripting.Dictionary: " & Err.Description
                Err.Clear
            Else
                Debug.Print "     -> OK: Scripting.Dictionary created successfully"
                Set testDict = Nothing
            End If
            On Error GoTo 0
        Else
            Debug.Print "     -> Not a version file (doesn't match pattern)"
        End If
        
        fileName = Dir
    Loop
    
    Debug.Print "3. Summary:"
    Debug.Print "   Total files found: " & fileCount
    Debug.Print "=== DIAGNOSIS COMPLETE ==="
End Sub

' Quick test function for Dictionary issue
Public Sub TestDictionaryCreation()
    On Error Resume Next
    
    Debug.Print "Testing Dictionary creation..."
    
    ' Test late binding (should work)
    Dim dictLate As Object
    Set dictLate = CreateObject("Scripting.Dictionary")
    If Err.Number = 0 Then
        Debug.Print "Late binding Dictionary: SUCCESS"
        dictLate("test") = "value"
        Debug.Print "Dictionary assignment: " & dictLate("test")
        Set dictLate = Nothing
    Else
        Debug.Print "Late binding Dictionary: FAILED - " & Err.Description
        Err.Clear
    End If
    
    Debug.Print "Dictionary test complete."
End Sub

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
    
    If g_DebugMode Then Debug.Print "GetVersionsList: Searching path: " & searchPath
    
    ' Check if metadata folder exists
    If Dir(METADATA_FOLDER, vbDirectory) = "" Then
        If g_DebugMode Then Debug.Print "GetVersionsList: Metadata folder does not exist: " & METADATA_FOLDER
        GoTo ErrorHandler
    End If

    Dim fileName As String
    fileName = Dir(searchPath)
    Dim fileCount As Integer
    fileCount = 0

    Do While fileName <> ""
        fileCount = fileCount + 1
        If g_DebugMode Then Debug.Print "GetVersionsList: Found file: " & fileName
        
        If Left(fileName, 1) = "v" And Right(fileName, 4) = ".txt" Then  ' Version files start with 'v' and end with '.txt'
            Dim versionInfo As Object
            Set versionInfo = ParseMetadataFile(METADATA_FOLDER & fileName)
            If Not versionInfo Is Nothing Then
                versions.Add versionInfo
                If g_DebugMode Then Debug.Print "GetVersionsList: Successfully parsed " & fileName & " (Version: " & versionInfo("Version") & ")"
            Else
                If g_DebugMode Then Debug.Print "GetVersionsList: Failed to parse " & fileName
            End If
        Else
            If g_DebugMode Then Debug.Print "GetVersionsList: Skipping non-version file: " & fileName
        End If
        fileName = Dir
    Loop
    
    If g_DebugMode Then Debug.Print "GetVersionsList: Total files found: " & fileCount & ", Valid versions: " & versions.Count

    Set GetVersionsList = versions
    Exit Function

ErrorHandler:
    If g_DebugMode Then Debug.Print "GetVersionsList: Error occurred: " & Err.Description & " (" & Err.Number & ")"
    Set GetVersionsList = New Collection
End Function

Private Function ParseMetadataFile(filePath As String) As Object
    On Error GoTo ErrorHandler

    If g_DebugMode Then Debug.Print "ParseMetadataFile: Parsing " & filePath
    
    ' Check if file exists
    If Dir(filePath) = "" Then
        If g_DebugMode Then Debug.Print "ParseMetadataFile: File does not exist: " & filePath
        GoTo ErrorHandler
    End If

    Dim versionInfo As Object
    Set versionInfo = CreateObject("Scripting.Dictionary")
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Dim lineCount As Integer
    lineCount = 0

    Open filePath For Input As fileNum

    Do While Not EOF(fileNum)
        Dim line As String
        Line Input #fileNum, line
        lineCount = lineCount + 1
        
        If g_DebugMode And lineCount <= 3 Then Debug.Print "ParseMetadataFile: Line " & lineCount & ": " & line

        If InStr(line, ":") > 0 Then
            Dim parts As Variant
            parts = Split(line, ":", 2)
            If UBound(parts) >= 1 Then
                Dim key As String, value As String
                key = Trim(parts(0))
                value = Trim(parts(1))
                versionInfo(key) = value
                
                If g_DebugMode And (key = "Version" Or key = "File") Then
                    Debug.Print "ParseMetadataFile: Added key '" & key & "' = '" & value & "'"
                End If
            End If
        End If
    Loop

    Close fileNum
    
    ' Validate required fields
    If Not versionInfo.Exists("Version") Then
        If g_DebugMode Then Debug.Print "ParseMetadataFile: Missing 'Version' field in " & filePath
        GoTo ErrorHandler
    End If
    
    If Not versionInfo.Exists("File") Then
        If g_DebugMode Then Debug.Print "ParseMetadataFile: Missing 'File' field in " & filePath
        GoTo ErrorHandler
    End If
    
    If g_DebugMode Then Debug.Print "ParseMetadataFile: Successfully parsed " & versionInfo.Count & " fields from " & filePath

    Set ParseMetadataFile = versionInfo
    Exit Function

ErrorHandler:
    If fileNum > 0 Then Close fileNum
    If g_DebugMode Then Debug.Print "ParseMetadataFile: Error parsing " & filePath & ": " & Err.Description & " (" & Err.Number & ")"
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
        Dim version As Object
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
    If g_DebugMode Then Debug.Print "ShowCompareVersionDialog: Starting version comparison dialog"
    
    Dim versions As Collection
    Set versions = GetVersionsList()
    
    If g_DebugMode Then Debug.Print "ShowCompareVersionDialog: Found " & versions.Count & " versions"

    If versions.Count = 0 Then
        Dim errorMsg As String
        errorMsg = "No versions available for comparison." & vbCrLf & vbCrLf
        errorMsg = errorMsg & "Debug Info:" & vbCrLf
        errorMsg = errorMsg & "Metadata Folder: " & METADATA_FOLDER & vbCrLf
        
        If Dir(METADATA_FOLDER, vbDirectory) = "" Then
            errorMsg = errorMsg & "Status: Metadata folder does not exist" & vbCrLf
        Else
            errorMsg = errorMsg & "Status: Metadata folder exists but no valid version files found" & vbCrLf
        End If
        
        errorMsg = errorMsg & vbCrLf & "Enable debug mode (g_DebugMode = True) and check Immediate window for details."
        
        MsgBox errorMsg, vbExclamation, "Version Control - No Versions Found"
        Exit Sub
    End If

    ' Create list of available versions for user reference
    Dim versionList As String
    versionList = "Available versions for comparison:" & vbCrLf & vbCrLf
    
    Dim i As Integer
    For i = 1 To versions.Count
        Dim version As Object
        Set version = versions(i)
        versionList = versionList & version("Version") & " - " & version("Created")
        If version.Exists("Notes") And version("Notes") <> "" Then
            versionList = versionList & " (" & version("Notes") & ")"
        End If
        versionList = versionList & vbCrLf
    Next i
    
    versionList = versionList & vbCrLf & "Enter the version name you want to compare with:"
    
    ' Show available versions and get user selection
    MsgBox versionList, vbInformation, "Version Control - Available Versions"
    
    Dim selectedVersion As String
    Dim latestVersion As String
    If versions.Count > 0 Then
        Set version = versions(versions.Count)  ' Get the latest version as default
        latestVersion = version("Version")
    Else
        latestVersion = "v003"  ' Fallback
    End If
    
    selectedVersion = InputBox("Enter version name to compare with:", _
                              "Compare Versions", latestVersion)

    If selectedVersion <> "" Then
        If g_DebugMode Then Debug.Print "ShowCompareVersionDialog: User selected version: " & selectedVersion
        Call PerformBasicComparison(selectedVersion)
    Else
        If g_DebugMode Then Debug.Print "ShowCompareVersionDialog: User cancelled version selection"
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
            Dim version As Object
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
    
    If g_DebugMode Then Debug.Print "PerformBasicComparison: Starting comparison with version " & versionName

    ' Validate input
    If Trim(versionName) = "" Then
        MsgBox "Invalid version name. Please enter a valid version (e.g., v003).", vbExclamation, "Version Control"
        Exit Sub
    End If
    
    ' Check if current workbook path is set
    If g_CurrentWorkbookPath = "" Then
        MsgBox "Current workbook path not set. Please try the operation again.", vbExclamation, "Version Control"
        Exit Sub
    End If
    
    ' Check if current workbook exists
    If Dir(g_CurrentWorkbookPath) = "" Then
        MsgBox "Current workbook file not found: " & g_CurrentWorkbookPath, vbExclamation, "Version Control"
        Exit Sub
    End If

    ' Find version metadata file
    Dim metadataPath As String
    metadataPath = METADATA_FOLDER & versionName & ".txt"
    
    If g_DebugMode Then Debug.Print "PerformBasicComparison: Looking for metadata at " & metadataPath
    
    If Dir(metadataPath) = "" Then
        Dim errorMsg As String
        errorMsg = "Version metadata not found for: " & versionName & vbCrLf & vbCrLf
        errorMsg = errorMsg & "Expected location: " & metadataPath & vbCrLf & vbCrLf
        
        ' Show available versions
        Dim versions As Collection
        Set versions = GetVersionsList()
        If versions.Count > 0 Then
            errorMsg = errorMsg & "Available versions: "
            Dim i As Integer
            For i = 1 To versions.Count
                Dim ver As Dictionary
                Set ver = versions(i)
                errorMsg = errorMsg & ver("Version")
                If i < versions.Count Then errorMsg = errorMsg & ", "
            Next i
        Else
            errorMsg = errorMsg & "No versions available."
        End If
        
        MsgBox errorMsg, vbExclamation, "Version Control - Version Not Found"
        Exit Sub
    End If
    
    Dim versionInfo As Object
    Set versionInfo = ParseMetadataFile(metadataPath)

    If versionInfo Is Nothing Then
        MsgBox "Failed to parse version metadata for: " & versionName & vbCrLf & _
               "Metadata file: " & metadataPath & vbCrLf & vbCrLf & _
               "The metadata file may be corrupted.", vbExclamation, "Version Control"
        Exit Sub
    End If

    ' Get version file path
    If Not versionInfo.Exists("File") Then
        MsgBox "Version metadata is missing file path information for: " & versionName, vbExclamation, "Version Control"
        Exit Sub
    End If
    
    Dim versionFilePath As String
    versionFilePath = versionInfo("File")
    
    If g_DebugMode Then Debug.Print "PerformBasicComparison: Version file path: " & versionFilePath

    If Dir(versionFilePath) = "" Then
        Dim fileErrorMsg As String
        fileErrorMsg = "Version snapshot file not found:" & vbCrLf & versionFilePath & vbCrLf & vbCrLf
        fileErrorMsg = fileErrorMsg & "Version: " & versionName & vbCrLf
        fileErrorMsg = fileErrorMsg & "Metadata: " & metadataPath & vbCrLf & vbCrLf
        fileErrorMsg = fileErrorMsg & "The snapshot file may have been moved or deleted."
        
        MsgBox fileErrorMsg, vbExclamation, "Version Control - Snapshot File Not Found"
        Exit Sub
    End If

    ' Perform comparison
    Dim compareMsg As String
    compareMsg = "=== VERSION COMPARISON ===" & vbCrLf & vbCrLf
    
    ' Current file info
    compareMsg = compareMsg & "CURRENT WORKBOOK:" & vbCrLf
    compareMsg = compareMsg & "Path: " & g_CurrentWorkbookPath & vbCrLf
    compareMsg = compareMsg & "Size: " & Format(FileLen(g_CurrentWorkbookPath) / 1024, "#,##0.0") & " KB" & vbCrLf
    compareMsg = compareMsg & "Modified: " & Format(FileDateTime(g_CurrentWorkbookPath), "mm/dd/yyyy hh:mm:ss") & vbCrLf & vbCrLf
    
    ' Version file info
    compareMsg = compareMsg & "VERSION " & UCase(versionName) & ":" & vbCrLf
    compareMsg = compareMsg & "Path: " & versionFilePath & vbCrLf
    compareMsg = compareMsg & "Size: " & Format(FileLen(versionFilePath) / 1024, "#,##0.0") & " KB" & vbCrLf
    compareMsg = compareMsg & "Created: " & versionInfo("Created") & vbCrLf
    
    If versionInfo.Exists("Notes") And versionInfo("Notes") <> "" Then
        compareMsg = compareMsg & "Notes: " & versionInfo("Notes") & vbCrLf
    End If
    
    compareMsg = compareMsg & vbCrLf & "COMPARISON RESULTS:" & vbCrLf
    
    Dim currentSize As Long, versionSize As Long
    currentSize = FileLen(g_CurrentWorkbookPath)
    versionSize = FileLen(versionFilePath)
    
    If currentSize = versionSize Then
        compareMsg = compareMsg & "✓ Files are identical in size" & vbCrLf
    Else
        Dim sizeDiff As Long
        sizeDiff = currentSize - versionSize
        compareMsg = compareMsg & "⚠ Files differ in size by " & Format(Abs(sizeDiff) / 1024, "#,##0.0") & " KB"
        If sizeDiff > 0 Then
            compareMsg = compareMsg & " (current is larger)" & vbCrLf
        Else
            compareMsg = compareMsg & " (version is larger)" & vbCrLf
        End If
    End If
    
    If g_DebugMode Then Debug.Print "PerformBasicComparison: Comparison completed successfully"

    MsgBox compareMsg, vbInformation, "Version Control - Comparison Results"

    ' Offer to open version for manual comparison
    Dim response As VbMsgBoxResult
    response = MsgBox("Would you like to open the version file for detailed manual comparison?", _
                     vbYesNo + vbQuestion, "Version Control - Open Version")

    If response = vbYes Then
        If g_DebugMode Then Debug.Print "PerformBasicComparison: Opening version file for manual comparison"
        On Error Resume Next
        Workbooks.Open versionFilePath
        If Err.Number <> 0 Then
            MsgBox "Failed to open version file: " & Err.Description, vbExclamation, "Version Control"
        End If
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ErrorHandler:
    Dim errorDetails As String
    errorDetails = "Error during version comparison:" & vbCrLf & vbCrLf
    errorDetails = errorDetails & "Error: " & Err.Description & " (" & Err.Number & ")" & vbCrLf
    errorDetails = errorDetails & "Version: " & versionName & vbCrLf
    errorDetails = errorDetails & "Current Path: " & g_CurrentWorkbookPath & vbCrLf
    
    If g_DebugMode Then
        errorDetails = errorDetails & vbCrLf & "Enable debug mode and check Immediate window for detailed logs."
        Debug.Print "PerformBasicComparison: ERROR - " & Err.Description & " (" & Err.Number & ")"
    End If
    
    MsgBox errorDetails, vbCritical, "Version Control Error"
End Sub

' Rollback function
Private Sub PerformRollback(versionName As String)
    On Error GoTo ErrorHandler

    ' Find version file
    Dim versionInfo As Object
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

' Debug mode toggle functions
Public Sub EnableDebugMode()
    g_DebugMode = True
    Debug.Print "=== VERSION CONTROL DEBUG MODE ENABLED ==="
    Debug.Print "Time: " & Now
    Debug.Print "Version Folder: " & VERSION_FOLDER
    Debug.Print "Metadata Folder: " & METADATA_FOLDER
    MsgBox "Debug mode enabled. Check Immediate window (Ctrl+G) for detailed logs.", vbInformation, "Version Control Debug"
End Sub

Public Sub DisableDebugMode()
    g_DebugMode = False
    Debug.Print "=== VERSION CONTROL DEBUG MODE DISABLED ==="
    MsgBox "Debug mode disabled.", vbInformation, "Version Control Debug"
End Sub

Public Sub ShowDebugInfo()
    Dim debugInfo As String
    debugInfo = "=== VERSION CONTROL DEBUG INFO ===" & vbCrLf & vbCrLf
    debugInfo = debugInfo & "Debug Mode: " & IIf(g_DebugMode, "ENABLED", "DISABLED") & vbCrLf
    debugInfo = debugInfo & "Current Workbook: " & g_CurrentWorkbookPath & vbCrLf
    debugInfo = debugInfo & "Version Folder: " & VERSION_FOLDER & vbCrLf
    debugInfo = debugInfo & "Metadata Folder: " & METADATA_FOLDER & vbCrLf & vbCrLf
    
    ' Check folder existence
    debugInfo = debugInfo & "Folder Status:" & vbCrLf
    debugInfo = debugInfo & "- Version folder exists: " & IIf(Dir(VERSION_FOLDER, vbDirectory) <> "", "YES", "NO") & vbCrLf
    debugInfo = debugInfo & "- Metadata folder exists: " & IIf(Dir(METADATA_FOLDER, vbDirectory) <> "", "YES", "NO") & vbCrLf & vbCrLf
    
    ' Version count
    Dim versions As Collection
    Set versions = GetVersionsList()
    debugInfo = debugInfo & "Available Versions: " & versions.Count & vbCrLf
    
    If versions.Count > 0 Then
        debugInfo = debugInfo & "Latest Version: " & versions(versions.Count)("Version") & vbCrLf
    End If
    
    debugInfo = debugInfo & vbCrLf & "To enable debug logging: EnableDebugMode" & vbCrLf
    debugInfo = debugInfo & "To disable debug logging: DisableDebugMode"
    
    MsgBox debugInfo, vbInformation, "Version Control Debug Information"
End Sub

' Test function
Public Function TestVBAOnlySystem() As Boolean
    On Error GoTo ErrorHandler
    
    Dim originalDebugMode As Boolean
    originalDebugMode = g_DebugMode
    g_DebugMode = True  ' Enable debug for testing
    
    Debug.Print "=== TESTING VBA-ONLY VERSION CONTROL SYSTEM ==="

    ' Test directory creation
    Debug.Print "Testing directory creation..."
    Call EnsureDirectoriesExist

    ' Test version numbering
    Debug.Print "Testing version numbering..."
    Dim nextVersion As Long
    nextVersion = GetNextVersionNumber()
    Debug.Print "Next version number: " & nextVersion
    
    ' Test version list retrieval
    Debug.Print "Testing version list retrieval..."
    Dim versions As Collection
    Set versions = GetVersionsList()
    Debug.Print "Found " & versions.Count & " versions"
    
    ' Test metadata parsing if versions exist
    If versions.Count > 0 Then
        Debug.Print "Testing metadata parsing..."
        Dim testVersion As Dictionary
        Set testVersion = versions(1)
        Debug.Print "First version: " & testVersion("Version") & " created " & testVersion("Created")
    End If
    
    g_DebugMode = originalDebugMode  ' Restore original debug mode

    TestVBAOnlySystem = True
    MsgBox "VBA-only version control system test passed!" & vbCrLf & _
           "Next version number: " & nextVersion & vbCrLf & _
           "Available versions: " & versions.Count & vbCrLf & vbCrLf & _
           "Check Immediate window for detailed test results.", vbInformation, "Version Control Test"

    Exit Function

ErrorHandler:
    g_DebugMode = originalDebugMode  ' Restore original debug mode
    TestVBAOnlySystem = False
    MsgBox "VBA-only system test failed: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number, vbExclamation, "Version Control Test"
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
  1 . 0   T y p e   L i b r a r y   P D F M a k e r   f o r   V i s i o   1 . 0   T y p e   L i b r a r y   P D F M a k e r A P I   1 . 0   T y p e   L i b r a r y   P D F M L o t u s N o t e s   1 . 0   T y p e   L i b r a r y   P D F M O u t l o o k   1 . 0   T y p e   L i b r a r y   P D F P r e v H n d l r   1 . 0   T y p e   L i b r a r y   P h   �ڂ p�"�  O~�  j e c t s   P o l i c y   T y p e   L i b r a r y   P o r t a b l e D e v i c e A p i   1 . 0   T y p e   L i b r a r y                 