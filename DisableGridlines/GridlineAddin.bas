Attribute VB_Name = "GridlineFormattingAddin"
'
' Gridline Formatting Add-in Module
' Works across all Excel workbooks as an add-in
' Disables gridlines, sets zoom to 85%, and returns to A1 on all sheets
' Created for FuzzySum project
'

Option Explicit

' Main subroutine to format all sheets (gridlines off, 85% zoom, return to A1)
Public Sub FormatAllSheetsComplete()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim errorCount As Integer
    Dim wb As Workbook

    ' Check if there's an active workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found. Please open a workbook first.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    If wb.Worksheets.Count = 0 Then
        MsgBox "The active workbook has no worksheets.", vbExclamation, "No Worksheets"
        Exit Sub
    End If

    ' Store the currently active sheet
    Set originalSheet = ActiveSheet
    errorCount = 0

    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    ' Loop through all worksheets in the active workbook
    For Each ws In wb.Worksheets
        ' Activate the worksheet (required to change view properties)
        ws.Activate

        ' 1. Disable gridlines for the active window
        ActiveWindow.DisplayGridlines = False

        ' 2. Set zoom to 85%
        ActiveWindow.Zoom = 85

        ' 3. Return to cell A1
        ws.Range("A1").Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1

        DoEvents ' Allow other processes to run
    Next ws

    ' Return to the original sheet
    originalSheet.Activate

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Show completion message
    MsgBox "Formatting complete on " & wb.Worksheets.Count & " worksheet(s):" & vbCrLf & _
           "• Gridlines disabled" & vbCrLf & _
           "• Zoom set to 85%" & vbCrLf & _
           "• Returned to cell A1", _
           vbInformation, "Sheet Formatting Complete"

    Exit Sub

ErrorHandler:
    errorCount = errorCount + 1

    ' Try to continue with next sheet
    Resume Next

    If errorCount > 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        originalSheet.Activate
        MsgBox "Completed with " & errorCount & " error(s). Some sheets may be protected.", _
               vbExclamation, "Formatting Warning"
    End If
End Sub

' Individual function: Disable gridlines only
Public Sub DisableAllGridlines()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim errorCount As Integer
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    Set originalSheet = ActiveSheet
    errorCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    For Each ws In wb.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        DoEvents
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Gridlines disabled on " & wb.Worksheets.Count & " worksheet(s).", _
           vbInformation, "Gridlines Disabled"

    Exit Sub

ErrorHandler:
    errorCount = errorCount + 1
    Resume Next

    If errorCount > 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        originalSheet.Activate
        MsgBox "Completed with " & errorCount & " error(s).", vbExclamation, "Warning"
    End If
End Sub

' Individual function: Set zoom to 85% on all sheets
Public Sub SetZoomToStandard()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim errorCount As Integer
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    Set originalSheet = ActiveSheet
    errorCount = 0

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    For Each ws In wb.Worksheets
        ws.Activate
        ActiveWindow.Zoom = 85
        DoEvents
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True

    MsgBox "Zoom set to 85% on " & wb.Worksheets.Count & " worksheet(s).", _
           vbInformation, "Zoom Updated"

    Exit Sub

ErrorHandler:
    errorCount = errorCount + 1
    Resume Next

    If errorCount > 0 Then
        Application.ScreenUpdating = True
        originalSheet.Activate
        MsgBox "Completed with " & errorCount & " error(s).", vbExclamation, "Warning"
    End If
End Sub

' Individual function: Return to A1 on all sheets
Public Sub ResetToHomePosition()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim errorCount As Integer
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    Set originalSheet = ActiveSheet
    errorCount = 0

    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    For Each ws In wb.Worksheets
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        DoEvents
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True

    MsgBox "Returned to cell A1 on " & wb.Worksheets.Count & " worksheet(s).", _
           vbInformation, "Position Reset"

    Exit Sub

ErrorHandler:
    errorCount = errorCount + 1
    Resume Next

    If errorCount > 0 Then
        Application.ScreenUpdating = True
        originalSheet.Activate
        MsgBox "Completed with " & errorCount & " error(s).", vbExclamation, "Warning"
    End If
End Sub

' Function to format active sheet only
Public Sub FormatActiveSheetOnly()
    Dim ws As Worksheet

    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active worksheet found.", vbExclamation, "No Worksheet"
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Format the current sheet
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    ws.Range("A1").Select
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1

    MsgBox "Current worksheet formatted:" & vbCrLf & _
           "• Gridlines disabled" & vbCrLf & _
           "• Zoom set to 85%" & vbCrLf & _
           "• Returned to cell A1", _
           vbInformation, "Sheet Formatted"

    Exit Sub

ErrorHandler:
    MsgBox "Error formatting worksheet: " & Err.Description, vbExclamation, "Formatting Error"
End Sub

' Restore gridlines on all sheets
Public Sub EnableAllGridlines()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    Set originalSheet = ActiveSheet
    Application.ScreenUpdating = False

    On Error Resume Next

    For Each ws In wb.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayHeadings = True
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True

    MsgBox "Gridlines enabled on all worksheets.", vbInformation, "Gridlines Restored"

    On Error GoTo 0
End Sub

' Interactive formatting with user choices
Public Sub FormatWithOptions()
    Dim response As VbMsgBoxResult
    Dim formatChoice As String
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "No Workbook"
        Exit Sub
    End If

    ' Ask user what they want to format
    formatChoice = InputBox("Enter formatting options (separate with commas):" & vbCrLf & vbCrLf & _
                           "Available options:" & vbCrLf & _
                           "• gridlines (disable gridlines)" & vbCrLf & _
                           "• zoom (set to 85%)" & vbCrLf & _
                           "• home (return to A1)" & vbCrLf & _
                           "• all (all of the above)" & vbCrLf & vbCrLf & _
                           "Example: gridlines,zoom", _
                           "Format Options", "all")

    If formatChoice = "" Then Exit Sub

    formatChoice = LCase(Trim(formatChoice))

    ' Apply selected formatting
    If InStr(formatChoice, "all") > 0 Then
        Call FormatAllSheetsComplete
    Else
        If InStr(formatChoice, "gridlines") > 0 Then Call DisableAllGridlines
        If InStr(formatChoice, "zoom") > 0 Then Call SetZoomToStandard
        If InStr(formatChoice, "home") > 0 Then Call ResetToHomePosition
    End If
End Sub

' Utility function to check if gridlines are enabled on current sheet
Public Function GridlinesEnabled() As Boolean
    If ActiveSheet Is Nothing Then
        GridlinesEnabled = False
    Else
        GridlinesEnabled = ActiveWindow.DisplayGridlines
    End If
End Function

' Get current zoom level of active sheet
Public Function GetCurrentZoom() As Integer
    If ActiveSheet Is Nothing Then
        GetCurrentZoom = 100
    Else
        GetCurrentZoom = ActiveWindow.Zoom
    End If
End Function

' Initialize add-in - sets up keyboard shortcuts
Public Sub InitializeGridlineAddin()
    ' Assign keyboard shortcuts
    Application.OnKey "^+f", "FormatAllSheetsComplete"    ' Ctrl+Shift+F - Full format
    Application.OnKey "^+g", "DisableAllGridlines"        ' Ctrl+Shift+G - Gridlines only
    Application.OnKey "^+z", "SetZoomToStandard"          ' Ctrl+Shift+Z - Zoom only
    Application.OnKey "^+h", "ResetToHomePosition"        ' Ctrl+Shift+H - Home position
    Application.OnKey "^+a", "FormatActiveSheetOnly"      ' Ctrl+Shift+A - Active sheet only

    MsgBox "Gridline Formatting Add-in initialized!" & vbCrLf & vbCrLf & _
           "Keyboard shortcuts:" & vbCrLf & _
           "• Ctrl+Shift+F - Format all sheets (complete)" & vbCrLf & _
           "• Ctrl+Shift+G - Disable gridlines only" & vbCrLf & _
           "• Ctrl+Shift+Z - Set zoom to 85% only" & vbCrLf & _
           "• Ctrl+Shift+H - Return to A1 only" & vbCrLf & _
           "• Ctrl+Shift+A - Format active sheet only", _
           vbInformation, "Add-in Ready"
End Sub

' Clean up keyboard shortcuts when add-in is disabled
Public Sub CleanupGridlineAddin()
    Application.OnKey "^+f"
    Application.OnKey "^+g"
    Application.OnKey "^+z"
    Application.OnKey "^+h"
    Application.OnKey "^+a"

    MsgBox "Gridline Formatting Add-in shortcuts removed.", vbInformation, "Add-in Disabled"
End Sub