' Enhanced VBA-Python Interface Module
' Provides reliable communication with Python backend using file-based exchange

Option Explicit

' Constants
Private Const PYTHON_BRIDGE_PATH As String = "C:\Users\wschoenberger\FuzzySum\VersionControl\vba_python_bridge.py"
Private Const TEMP_DIR As String = "C:\temp\VersionControl\"
Private Const TIMEOUT_SECONDS As Long = 30

' API declarations for process management
#If VBA7 Then
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
        (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

    Private Declare PtrSafe Function CloseHandle Lib "kernel32" _
        (ByVal hObject As LongPtr) As Long

    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
#Else
    Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
        (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

    Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

    Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, lpExitCode As Long) As Long
#End If

' Type definitions
#If VBA7 Then
    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As LongPtr
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    End Type

    Private Type PROCESS_INFORMATION
        hProcess As LongPtr
        hThread As LongPtr
        dwProcessId As Long
        dwThreadId As Long
    End Type
#Else
    Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type

    Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
    End Type
#End If

' Enhanced Python command execution with file-based communication
Public Function ExecutePythonCommandEx(action As String, workbookPath As String, _
                                       Optional parameters As String = "", _
                                       Optional timeoutSeconds As Long = TIMEOUT_SECONDS) As Dictionary
    On Error GoTo ErrorHandler

    ' Initialize result dictionary
    Dim result As New Dictionary
    result("success") = False
    result("error") = "Unknown error"

    ' Ensure temp directory exists
    Call EnsureTempDirectory

    ' Create unique temporary file for output
    Dim outputFile As String
    outputFile = TEMP_DIR & "output_" & Format(Now, "yyyymmdd_hhmmss_") & Int(Rnd * 1000) & ".json"

    ' Build command line
    Dim cmdLine As String
    cmdLine = "python """ & PYTHON_BRIDGE_PATH & """ --workbook """ & workbookPath & """ --action " & action & _
              " --output """ & outputFile & """"

    If Len(parameters) > 0 Then
        cmdLine = cmdLine & " " & parameters
    End If

    ' Execute Python process
    Dim success As Boolean
    success = ExecuteProcessWithTimeout(cmdLine, timeoutSeconds)

    If success Then
        ' Read result from output file
        If Dir(outputFile) <> "" Then
            Set result = ReadJsonFile(outputFile)

            ' Clean up output file
            Kill outputFile
        Else
            result("error") = "Python process completed but output file not found"
        End If
    Else
        result("error") = "Python process failed or timed out"
    End If

    Set ExecutePythonCommandEx = result
    Exit Function

ErrorHandler:
    result("error") = "VBA Error: " & Err.Description
    Set ExecutePythonCommandEx = result
End Function

' Execute process with timeout
Private Function ExecuteProcessWithTimeout(cmdLine As String, timeoutSeconds As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim success As Long
    Dim waitResult As Long
    Dim exitCode As Long

    ' Initialize startup info
    si.cb = Len(si)
    si.dwFlags = 1 ' STARTF_USESHOWWINDOW
    si.wShowWindow = 0 ' SW_HIDE

    ' Create process
    success = CreateProcess(vbNullString, cmdLine, 0, 0, 0, 0, 0, vbNullString, si, pi)

    If success = 0 Then
        ExecuteProcessWithTimeout = False
        Exit Function
    End If

    ' Wait for process completion with timeout
    waitResult = WaitForSingleObject(pi.hProcess, timeoutSeconds * 1000)

    ' Get exit code
    GetExitCodeProcess pi.hProcess, exitCode

    ' Clean up handles
    CloseHandle pi.hProcess
    CloseHandle pi.hThread

    ' Check result
    ExecuteProcessWithTimeout = (waitResult = 0 And exitCode = 0) ' WAIT_OBJECT_0 = 0

    Exit Function

ErrorHandler:
    ExecuteProcessWithTimeout = False
End Function

' Read JSON file and parse into Dictionary
Private Function ReadJsonFile(filePath As String) As Dictionary
    On Error GoTo ErrorHandler

    Dim result As New Dictionary
    Dim fileContent As String
    Dim fileNum As Integer

    ' Read file content
    fileNum = FreeFile
    Open filePath For Input As fileNum
    fileContent = Input(LOF(fileNum), fileNum)
    Close fileNum

    ' Parse JSON (simplified parser)
    Set result = ParseSimpleJson(fileContent)

    Set ReadJsonFile = result
    Exit Function

ErrorHandler:
    Dim errorResult As New Dictionary
    errorResult("success") = False
    errorResult("error") = "Failed to read JSON file: " & Err.Description
    Set ReadJsonFile = errorResult
End Function

' Simplified JSON parser for basic Python output
Private Function ParseSimpleJson(jsonText As String) As Dictionary
    On Error GoTo ErrorHandler

    Dim result As New Dictionary
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    Dim colonPos As Long
    Dim key As String
    Dim value As String

    ' Remove braces and split by lines
    jsonText = Replace(jsonText, "{", "")
    jsonText = Replace(jsonText, "}", "")
    lines = Split(jsonText, vbLf)

    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))

        ' Skip empty lines and lines without colons
        If Len(line) > 0 And InStr(line, ":") > 0 Then
            colonPos = InStr(line, ":")
            key = Trim(Mid(line, 1, colonPos - 1))
            value = Trim(Mid(line, colonPos + 1))

            ' Remove quotes and commas
            key = Replace(Replace(key, """", ""), ",", "")
            value = Replace(Replace(value, """", ""), ",", "")

            ' Convert boolean and numeric values
            If LCase(value) = "true" Then
                result(key) = True
            ElseIf LCase(value) = "false" Then
                result(key) = False
            ElseIf IsNumeric(value) Then
                result(key) = CDbl(value)
            Else
                result(key) = value
            End If
        End If
    Next i

    Set ParseSimpleJson = result
    Exit Function

ErrorHandler:
    Dim errorResult As New Dictionary
    errorResult("success") = False
    errorResult("error") = "JSON parsing error: " & Err.Description
    Set ParseSimpleJson = errorResult
End Function

' Ensure temporary directory exists
Private Sub EnsureTempDirectory()
    On Error Resume Next
    If Dir(TEMP_DIR, vbDirectory) = "" Then
        MkDir Left(TEMP_DIR, Len(TEMP_DIR) - 1) ' Remove trailing backslash
    End If
    On Error GoTo 0
End Sub

' High-level interface functions for version control operations
Public Function CreateSnapshot(workbookPath As String, Optional notes As String = "", _
                              Optional quickSave As Boolean = False) As Dictionary
    Dim parameters As String

    If Len(notes) > 0 Then
        parameters = parameters & " --notes """ & notes & """"
    End If

    If quickSave Then
        parameters = parameters & " --quick"
    End If

    Set CreateSnapshot = ExecutePythonCommandEx("create_snapshot", workbookPath, parameters)
End Function

Public Function ListVersions(workbookPath As String) As Dictionary
    Set ListVersions = ExecutePythonCommandEx("list_versions", workbookPath)
End Function

Public Function CompareToVersion(workbookPath As String, versionName As String) As Dictionary
    Dim parameters As String
    parameters = "--version """ & versionName & """"

    Set CompareToVersion = ExecutePythonCommandEx("compare", workbookPath, parameters)
End Function

Public Function RollbackToVersion(workbookPath As String, versionName As String, _
                                 Optional backupCurrent As Boolean = True) As Dictionary
    Dim parameters As String
    parameters = "--version """ & versionName & """"

    If backupCurrent Then
        parameters = parameters & " --backup-current"
    End If

    Set RollbackToVersion = ExecutePythonCommandEx("rollback", workbookPath, parameters)
End Function

Public Function GetProjectStats(workbookPath As String) As Dictionary
    Set GetProjectStats = ExecutePythonCommandEx("stats", workbookPath)
End Function

Public Function GetVersionInfo(workbookPath As String, versionName As String) As Dictionary
    Dim parameters As String
    parameters = "--version """ & versionName & """"

    Set GetVersionInfo = ExecutePythonCommandEx("get_version_info", workbookPath, parameters)
End Function

' Utility function to check if Python bridge is available
Public Function IsPythonBridgeAvailable() As Boolean
    On Error GoTo ErrorHandler

    IsPythonBridgeAvailable = (Dir(PYTHON_BRIDGE_PATH) <> "")
    Exit Function

ErrorHandler:
    IsPythonBridgeAvailable = False
End Function

' Test function to verify Python connection
Public Function TestPythonConnection() As Boolean
    On Error GoTo ErrorHandler

    Dim testResult As Dictionary
    Dim tempWorkbook As String

    ' Create a temporary test file
    tempWorkbook = Environ("TEMP") & "\test_connection.xlsx"

    ' Use current workbook if no test file
    If ActiveWorkbook Is Nothing Then
        TestPythonConnection = False
        Exit Function
    End If

    ' Test with current workbook
    Set testResult = GetProjectStats(ActiveWorkbook.FullName)

    TestPythonConnection = testResult("success")
    Exit Function

ErrorHandler:
    TestPythonConnection = False
End Function