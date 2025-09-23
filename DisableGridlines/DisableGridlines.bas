Attribute VB_Name = "GridlineControl"
'
' GridlineControl Module
' Disables gridlines on all worksheets in the workbook
' Created for FuzzySum project
'

Option Explicit

' Main subroutine to disable gridlines on all sheets
Public Sub DisableAllGridlines()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim errorCount As Integer

    ' Store the currently active sheet
    Set originalSheet = ActiveSheet
    errorCount = 0

    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Activate the worksheet (required to change view properties)
        ws.Activate

        ' Disable gridlines for the active window
        ActiveWindow.DisplayGridlines = False

        ' Optional: Also disable row and column headers
        ' Uncomment the next line if you want to hide headers too
        ' ActiveWindow.DisplayHeadings = False

        ' Optional: Set a pleasant background color
        ' Uncomment and modify the next line if desired
        ' ws.Cells.Interior.Color = RGB(250, 250, 250)

        DoEvents ' Allow other processes to run
    Next ws

    ' Return to the original sheet
    originalSheet.Activate

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Show completion message
    MsgBox "Gridlines have been disabled on " & ThisWorkbook.Worksheets.Count & " worksheet(s).", _
           vbInformation, "Gridline Control Complete"

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
               vbExclamation, "Gridline Control Warning"
    End If
End Sub

' Alternative method that doesn't require sheet activation
Public Sub DisableGridlinesQuiet()
    Dim ws As Worksheet
    Dim win As Window

    Application.ScreenUpdating = False

    On Error Resume Next

    ' Try to disable gridlines without activating sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Find window for this worksheet
        For Each win In Application.Windows
            If win.Parent Is ThisWorkbook Then
                If win.ActiveSheet Is ws Then
                    win.DisplayGridlines = False
                    Exit For
                End If
            End If
        Next win
    Next ws

    Application.ScreenUpdating = True

    On Error GoTo 0
End Sub

' Subroutine to enable gridlines (reverse operation)
Public Sub EnableAllGridlines()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet

    Set originalSheet = ActiveSheet
    Application.ScreenUpdating = False

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayHeadings = True
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True

    MsgBox "Gridlines have been enabled on all worksheets.", vbInformation, "Gridlines Restored"

    On Error GoTo 0
End Sub

' Auto-run when workbook opens (optional)
' Remove the apostrophe from the next line to enable auto-run
' Private Sub Workbook_Open()
'     Call DisableAllGridlines
' End Sub

' Keyboard shortcut handler (Ctrl+Shift+G)
Public Sub GridlineToggle()
    Dim currentState As Boolean

    ' Check current state of active sheet
    currentState = ActiveWindow.DisplayGridlines

    If currentState Then
        Call DisableAllGridlines
    Else
        Call EnableAllGridlines
    End If
End Sub

' Utility function to check if gridlines are enabled on current sheet
Public Function GridlinesEnabled() As Boolean
    GridlinesEnabled = ActiveWindow.DisplayGridlines
End Function

' Add a button to the ribbon/toolbar (for Excel 2007+)
Public Sub CreateGridlineButton()
    ' This would require ribbon XML customization
    ' For now, users can assign the macro to a button manually
    MsgBox "To add a button:" & vbCrLf & _
           "1. Right-click on the ribbon" & vbCrLf & _
           "2. Select 'Customize the Ribbon'" & vbCrLf & _
           "3. Add a new button and assign the DisableAllGridlines macro", _
           vbInformation, "Add Button Instructions"
End Sub

' Initialize macro - sets up keyboard shortcut
Public Sub InitializeGridlineControl()
    ' Assign Ctrl+Shift+G to toggle gridlines
    Application.OnKey "^+g", "GridlineToggle"

    MsgBox "Gridline Control initialized!" & vbCrLf & vbCrLf & _
           "Available commands:" & vbCrLf & _
           "• Press Ctrl+Shift+G to toggle gridlines" & vbCrLf & _
           "• Run DisableAllGridlines() to turn off gridlines" & vbCrLf & _
           "• Run EnableAllGridlines() to turn on gridlines", _
           vbInformation, "Gridline Control Ready"
End Sub