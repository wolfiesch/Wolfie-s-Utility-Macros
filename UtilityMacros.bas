Attribute VB_Name = "UtilityMacros"
' ===================================================================
' EXCEL UTILITY MACROS ADD-IN
' Collection of useful Excel automation functions
' Version: 1.0
' Author: Excel Utility Tools
' ===================================================================

Option Explicit

' ===================================================================
' MAIN CONVERSION FUNCTIONS
' ===================================================================

Public Sub ConvertToAbsolute()
    On Error GoTo ErrorHandler

    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to convert.", vbExclamation, "Convert to Absolute"
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Confirm action with user for large selections
    Dim cellCount As Long
    cellCount = selectedRange.Cells.Count

    If cellCount > 100 Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You are about to convert " & cellCount & " cells to absolute references. This may take a moment. Continue?", _
                         vbYesNo + vbQuestion, "Convert to Absolute")
        If response = vbNo Then Exit Sub
    End If

    ' Track conversion statistics
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    convertedCount = 0
    skippedCount = 0
    errorCount = 0

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.StatusBar = "Converting formulas to absolute references..."

    ' Process each cell in the selection
    Dim cell As Range
    For Each cell In selectedRange.Cells
        If cell.HasFormula Then
            On Error Resume Next
            Dim originalFormula As String
            originalFormula = cell.Formula

            ' Convert to absolute using Excel's built-in method
            cell.Formula = Application.ConvertFormula( _
                Formula:=cell.Formula, _
                FromReferenceStyle:=xlA1, _
                ToAbsolute:=xlAbsolute, _
                RelativeTo:=cell)

            If Err.Number = 0 Then
                ' Check if formula actually changed
                If cell.Formula <> originalFormula Then
                    convertedCount = convertedCount + 1
                Else
                    skippedCount = skippedCount + 1
                End If
            Else
                errorCount = errorCount + 1
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Else
            skippedCount = skippedCount + 1
        End If
    Next cell

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Show results summary
    Dim resultMsg As String
    resultMsg = "Conversion Complete!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "Cells processed: " & cellCount & vbCrLf
    resultMsg = resultMsg & "Formulas converted: " & convertedCount & vbCrLf
    resultMsg = resultMsg & "Cells skipped (no formula or already absolute): " & skippedCount

    If errorCount > 0 Then
        resultMsg = resultMsg & vbCrLf & "Errors encountered: " & errorCount
    End If

    MsgBox resultMsg, vbInformation, "Convert to Absolute - Results"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred during conversion: " & Err.Description, vbCritical, "Convert to Absolute - Error"
End Sub

' ===================================================================
' ADVANCED CONVERSION WITH OPTIONS
' ===================================================================

Public Sub ConvertToAbsoluteAdvanced()
    On Error GoTo ErrorHandler

    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to convert.", vbExclamation, "Convert to Absolute"
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Show options dialog
    Dim convertOption As VbMsgBoxResult
    convertOption = MsgBox("Choose conversion type:" & vbCrLf & vbCrLf & _
                          "Yes = Convert to fully absolute ($A$1)" & vbCrLf & _
                          "No = Convert columns only ($A1)" & vbCrLf & _
                          "Cancel = Convert rows only (A$1)", _
                          vbYesNoCancel + vbQuestion, "Convert to Absolute - Options")

    If convertOption = 0 Then Exit Sub ' User closed dialog

    ' Determine conversion type
    Dim conversionType As XlReferenceType
    Select Case convertOption
        Case vbYes
            conversionType = xlAbsolute ' $A$1
        Case vbNo
            conversionType = xlAbsRowRelColumn ' $A1
        Case vbCancel
            conversionType = xlRelRowAbsColumn ' A$1
    End Select

    ' Track conversion statistics
    Dim convertedCount As Long
    convertedCount = 0

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.StatusBar = "Converting formulas..."

    ' Process each cell
    Dim cell As Range
    For Each cell In selectedRange.Cells
        If cell.HasFormula Then
            On Error Resume Next
            cell.Formula = Application.ConvertFormula( _
                Formula:=cell.Formula, _
                FromReferenceStyle:=xlA1, _
                ToAbsolute:=conversionType, _
                RelativeTo:=cell)

            If Err.Number = 0 Then
                convertedCount = convertedCount + 1
            End If
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next cell

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Show results
    MsgBox "Successfully converted " & convertedCount & " formula(s) to absolute references.", _
           vbInformation, "Convert to Absolute - Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Convert to Absolute - Error"
End Sub

' ===================================================================
' QUICK CONVERSION FUNCTIONS
' ===================================================================

Public Sub QuickConvertToAbsolute()
    On Error Resume Next

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cells with formulas to convert.", vbExclamation, "Quick Convert"
        Exit Sub
    End If

    Dim cell As Range
    For Each cell In Selection.Cells
        If cell.HasFormula Then
            cell.Formula = Application.ConvertFormula( _
                Formula:=cell.Formula, _
                FromReferenceStyle:=xlA1, _
                ToAbsolute:=xlAbsolute, _
                RelativeTo:=cell)
        End If
    Next cell

    If Err.Number <> 0 Then
        MsgBox "Some formulas could not be converted.", vbExclamation, "Quick Convert"
    End If
End Sub

' ===================================================================
' CONVERT TO RELATIVE REFERENCES
' ===================================================================

Public Sub ConvertToRelative()
    On Error GoTo ErrorHandler

    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to convert.", vbExclamation, "Convert to Relative"
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Track conversion statistics
    Dim convertedCount As Long
    convertedCount = 0

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.StatusBar = "Converting formulas to relative references..."

    ' Process each cell
    Dim cell As Range
    For Each cell In selectedRange.Cells
        If cell.HasFormula Then
            On Error Resume Next
            cell.Formula = Application.ConvertFormula( _
                Formula:=cell.Formula, _
                FromReferenceStyle:=xlA1, _
                ToAbsolute:=xlRelative, _
                RelativeTo:=cell)

            If Err.Number = 0 Then
                convertedCount = convertedCount + 1
            End If
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next cell

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Show results
    MsgBox "Successfully converted " & convertedCount & " formula(s) to relative references.", _
           vbInformation, "Convert to Relative - Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Convert to Relative - Error"
End Sub

' ===================================================================
' TOGGLE BETWEEN RELATIVE AND ABSOLUTE
' ===================================================================

Public Sub ToggleReferences()
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to toggle.", vbExclamation, "Toggle References"
        Exit Sub
    End If

    Dim cell As Range
    Dim toggleCount As Long
    toggleCount = 0

    Application.ScreenUpdating = False
    Application.StatusBar = "Toggling formula references..."

    For Each cell In Selection.Cells
        If cell.HasFormula Then
            Dim currentFormula As String
            currentFormula = cell.Formula

            ' Check if formula contains absolute references
            If InStr(currentFormula, "$") > 0 Then
                ' Convert to relative
                cell.Formula = Application.ConvertFormula( _
                    Formula:=cell.Formula, _
                    FromReferenceStyle:=xlA1, _
                    ToAbsolute:=xlRelative, _
                    RelativeTo:=cell)
            Else
                ' Convert to absolute
                cell.Formula = Application.ConvertFormula( _
                    Formula:=cell.Formula, _
                    FromReferenceStyle:=xlA1, _
                    ToAbsolute:=xlAbsolute, _
                    RelativeTo:=cell)
            End If

            toggleCount = toggleCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Successfully toggled " & toggleCount & " formula(s).", _
           vbInformation, "Toggle References - Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Toggle References - Error"
End Sub

' ===================================================================
' MIXED REFERENCE CONVERSIONS
' ===================================================================

Public Sub ConvertToMixedColumn()
    ' Convert to mixed reference with absolute column ($A1)
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation, "Convert to Mixed"
        Exit Sub
    End If

    Dim cell As Range
    Dim convertedCount As Long
    convertedCount = 0

    Application.ScreenUpdating = False

    For Each cell In Selection.Cells
        If cell.HasFormula Then
            cell.Formula = Application.ConvertFormula( _
                Formula:=cell.Formula, _
                FromReferenceStyle:=xlA1, _
                ToAbsolute:=xlAbsRowRelColumn, _
                RelativeTo:=cell)
            convertedCount = convertedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Converted " & convertedCount & " formula(s) to absolute column references ($A1).", _
           vbInformation, "Convert Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub ConvertToMixedRow()
    ' Convert to mixed reference with absolute row (A$1)
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells.", vbExclamation, "Convert to Mixed"
        Exit Sub
    End If

    Dim cell As Range
    Dim convertedCount As Long
    convertedCount = 0

    Application.ScreenUpdating = False

    For Each cell In Selection.Cells
        If cell.HasFormula Then
            cell.Formula = Application.ConvertFormula( _
                Formula:=cell.Formula, _
                FromReferenceStyle:=xlA1, _
                ToAbsolute:=xlRelRowAbsColumn, _
                RelativeTo:=cell)
            convertedCount = convertedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Converted " & convertedCount & " formula(s) to absolute row references (A$1).", _
           vbInformation, "Convert Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ===================================================================
' UTILITY FUNCTIONS
' ===================================================================

' ===================================================================
' ERROR HANDLING FUNCTIONS
' ===================================================================

Public Sub WrapWithIFERROR()
    On Error GoTo ErrorHandler

    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to wrap with IFERROR.", vbExclamation, "Wrap with IFERROR"
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Get error replacement value from user
    Dim errorValue As String
    errorValue = InputBox("Enter the value to display when an error occurs:" & vbCrLf & vbCrLf & _
                         "Common options:" & vbCrLf & _
                         "0 (zero)" & vbCrLf & _
                         """" (empty text)" & vbCrLf & _
                         """N/A""" & vbCrLf & _
                         """Error""" & vbCrLf & _
                         "FALSE", _
                         "IFERROR Value", "0")

    ' Exit if user cancelled
    If errorValue = "" Then
        MsgBox "Operation cancelled.", vbInformation, "Wrap with IFERROR"
        Exit Sub
    End If

    ' Confirm action for large selections
    Dim cellCount As Long
    cellCount = selectedRange.Cells.Count

    If cellCount > 50 Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You are about to wrap " & cellCount & " cells with IFERROR. Continue?", _
                         vbYesNo + vbQuestion, "Wrap with IFERROR")
        If response = vbNo Then Exit Sub
    End If

    ' Track statistics
    Dim wrappedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    wrappedCount = 0
    skippedCount = 0
    errorCount = 0

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.StatusBar = "Wrapping cells with IFERROR..."

    ' Process each cell
    Dim cell As Range
    For Each cell In selectedRange.Cells
        On Error Resume Next

        Dim originalContent As String
        If cell.HasFormula Then
            ' Cell has a formula
            originalContent = cell.Formula
            ' Check if already wrapped with IFERROR
            If InStr(UCase(originalContent), "IFERROR(") = 1 Then
                skippedCount = skippedCount + 1
            Else
                cell.Formula = "=IFERROR(" & Mid(originalContent, 2) & "," & errorValue & ")"
                wrappedCount = wrappedCount + 1
            End If
        ElseIf Not IsEmpty(cell.Value) Then
            ' Cell has a value (not formula)
            originalContent = cell.Value
            ' Convert value to IFERROR formula
            If IsNumeric(originalContent) Then
                cell.Formula = "=IFERROR(" & originalContent & "," & errorValue & ")"
            Else
                cell.Formula = "=IFERROR(""" & Replace(originalContent, """", """"""") & """," & errorValue & ")"
            End If
            wrappedCount = wrappedCount + 1
        Else
            ' Empty cell - skip
            skippedCount = skippedCount + 1
        End If

        If Err.Number <> 0 Then
            errorCount = errorCount + 1
            Err.Clear
        End If

        On Error GoTo ErrorHandler
    Next cell

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Show results
    Dim resultMsg As String
    resultMsg = "IFERROR Wrapping Complete!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "Cells processed: " & cellCount & vbCrLf
    resultMsg = resultMsg & "Cells wrapped: " & wrappedCount & vbCrLf
    resultMsg = resultMsg & "Cells skipped: " & skippedCount

    If errorCount > 0 Then
        resultMsg = resultMsg & vbCrLf & "Errors encountered: " & errorCount
    End If

    MsgBox resultMsg, vbInformation, "Wrap with IFERROR - Results"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Wrap with IFERROR - Error"
End Sub

Public Sub WrapWithIFERRORQuick()
    ' Quick version that uses 0 as default error value
    On Error Resume Next

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select cells to wrap with IFERROR.", vbExclamation, "Quick IFERROR"
        Exit Sub
    End If

    Dim cell As Range
    For Each cell In Selection.Cells
        If cell.HasFormula Then
            If InStr(UCase(cell.Formula), "IFERROR(") <> 1 Then
                cell.Formula = "=IFERROR(" & Mid(cell.Formula, 2) & ",0)"
            End If
        ElseIf Not IsEmpty(cell.Value) Then
            If IsNumeric(cell.Value) Then
                cell.Formula = "=IFERROR(" & cell.Value & ",0)"
            Else
                cell.Formula = "=IFERROR(""" & Replace(cell.Value, """", """"""") & """,0)"
            End If
        End If
    Next cell

    If Err.Number <> 0 Then
        MsgBox "Some cells could not be wrapped.", vbExclamation, "Quick IFERROR"
    End If
End Sub

Public Sub RemoveIFERROR()
    ' Remove IFERROR wrapper from selected cells
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to unwrap.", vbExclamation, "Remove IFERROR"
        Exit Sub
    End If

    Dim removedCount As Long
    removedCount = 0

    Application.ScreenUpdating = False
    Application.StatusBar = "Removing IFERROR wrappers..."

    Dim cell As Range
    For Each cell In Selection.Cells
        If cell.HasFormula Then
            Dim formula As String
            formula = cell.Formula

            ' Check if formula starts with IFERROR
            If InStr(UCase(formula), "=IFERROR(") = 1 Then
                ' Extract the original formula from IFERROR(original_formula,error_value)
                Dim startPos As Integer
                Dim endPos As Integer
                Dim parenCount As Integer

                startPos = 10 ' Position after "=IFERROR("
                parenCount = 1
                endPos = startPos

                ' Find the comma that separates formula from error value
                Do While endPos <= Len(formula) And parenCount > 0
                    endPos = endPos + 1
                    If Mid(formula, endPos, 1) = "(" Then
                        parenCount = parenCount + 1
                    ElseIf Mid(formula, endPos, 1) = ")" Then
                        parenCount = parenCount - 1
                    ElseIf Mid(formula, endPos, 1) = "," And parenCount = 1 Then
                        Exit Do
                    End If
                Loop

                If endPos <= Len(formula) Then
                    Dim originalFormula As String
                    originalFormula = Mid(formula, startPos, endPos - startPos)
                    cell.Formula = "=" & originalFormula
                    removedCount = removedCount + 1
                End If
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Removed IFERROR from " & removedCount & " cell(s).", _
           vbInformation, "Remove IFERROR - Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Remove IFERROR - Error"
End Sub

' ===================================================================
' DATE FORMATTING FUNCTIONS
' ===================================================================

Public Sub FormatDatesToCalendar()
    On Error GoTo ErrorHandler

    ' Check if a range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to format.", vbExclamation, "Format Dates"
        Exit Sub
    End If

    Dim selectedRange As Range
    Set selectedRange = Selection

    ' Track statistics
    Dim formattedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long
    formattedCount = 0
    skippedCount = 0
    errorCount = 0

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.StatusBar = "Formatting date cells..."

    ' Process each cell
    Dim cell As Range
    For Each cell In selectedRange.Cells
        On Error Resume Next

        ' Check if cell contains a date value
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            ' Apply MMM-YYYY format and bold formatting
            cell.NumberFormat = "mmm-yyyy"
            cell.Font.Bold = True

            If Err.Number = 0 Then
                formattedCount = formattedCount + 1
            Else
                errorCount = errorCount + 1
                Err.Clear
            End If
        Else
            ' Skip non-date cells
            skippedCount = skippedCount + 1
        End If

        On Error GoTo ErrorHandler
    Next cell

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' Show results
    Dim resultMsg As String
    resultMsg = "Date Formatting Complete!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "Cells processed: " & selectedRange.Cells.Count & vbCrLf
    resultMsg = resultMsg & "Dates formatted: " & formattedCount & vbCrLf
    resultMsg = resultMsg & "Cells skipped (non-dates): " & skippedCount

    If errorCount > 0 Then
        resultMsg = resultMsg & vbCrLf & "Errors encountered: " & errorCount
    End If

    If formattedCount > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & "Format applied: MMM-YYYY (bold)"
    End If

    MsgBox resultMsg, vbInformation, "Format Dates - Results"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred during formatting: " & Err.Description, vbCritical, "Format Dates - Error"
End Sub

Public Sub FormatDatesToCalendarAdvanced()
    ' Advanced version with format options
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to format.", vbExclamation, "Format Dates"
        Exit Sub
    End If

    ' Show format options
    Dim formatChoice As VbMsgBoxResult
    formatChoice = MsgBox("Choose date format:" & vbCrLf & vbCrLf & _
                          "Yes = MMM-YYYY (Jan-2024)" & vbCrLf & _
                          "No = MMM YYYY (Jan 2024)" & vbCrLf & _
                          "Cancel = MMMM YYYY (January 2024)", _
                          vbYesNoCancel + vbQuestion, "Date Format Options")

    If formatChoice = 0 Then Exit Sub ' User closed dialog

    ' Determine format string
    Dim dateFormat As String
    Select Case formatChoice
        Case vbYes
            dateFormat = "mmm-yyyy"
        Case vbNo
            dateFormat = "mmm yyyy"
        Case vbCancel
            dateFormat = "mmmm yyyy"
    End Select

    ' Ask about bold formatting
    Dim applyBold As VbMsgBoxResult
    applyBold = MsgBox("Apply bold formatting to formatted dates?", _
                       vbYesNo + vbQuestion, "Bold Formatting")

    Dim formattedCount As Long
    formattedCount = 0

    Application.ScreenUpdating = False
    Application.StatusBar = "Formatting date cells..."

    Dim cell As Range
    For Each cell In Selection.Cells
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            cell.NumberFormat = dateFormat
            If applyBold = vbYes Then
                cell.Font.Bold = True
            End If
            formattedCount = formattedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Successfully formatted " & formattedCount & " date cell(s) with format: " & dateFormat, _
           vbInformation, "Format Dates - Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Format Dates - Error"
End Sub

Public Sub RemoveDateFormatting()
    ' Remove custom date formatting and return to general format
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to reset formatting.", vbExclamation, "Reset Date Format"
        Exit Sub
    End If

    Dim resetCount As Long
    resetCount = 0

    Application.ScreenUpdating = False

    Dim cell As Range
    For Each cell In Selection.Cells
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            cell.NumberFormat = "General"
            cell.Font.Bold = False
            resetCount = resetCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Reset formatting for " & resetCount & " date cell(s).", _
           vbInformation, "Reset Date Format - Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Reset Date Format - Error"
End Sub

' ===================================================================
' EXPORT FUNCTIONS
' ===================================================================

Public Sub ExportSheetToJSON()
    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then
        MsgBox "No active sheet found.", vbExclamation, "Export to JSON"
        Exit Sub
    End If

    ' Get used range
    Dim usedRange As Range
    Set usedRange = ActiveSheet.UsedRange

    If usedRange Is Nothing Or usedRange.Cells.Count = 1 And IsEmpty(usedRange.Cells(1, 1)) Then
        MsgBox "No data found on the active sheet.", vbExclamation, "Export to JSON"
        Exit Sub
    End If

    ' Prompt for export options
    Dim includeHeaders As VbMsgBoxResult
    includeHeaders = MsgBox("Include first row as headers?", vbYesNoCancel + vbQuestion, "Export Options")

    If includeHeaders = vbCancel Then Exit Sub

    ' Get file path from user
    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFilename:=ActiveSheet.Name & ".json", _
        FileFilter:="JSON Files (*.json), *.json", _
        Title:="Save JSON Export As")

    If filePath = "False" Then Exit Sub

    Application.StatusBar = "Exporting to JSON..."
    Application.ScreenUpdating = False

    ' Build JSON
    Dim jsonContent As String
    jsonContent = BuildJSONFromRange(usedRange, includeHeaders = vbYes)

    ' Write to file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As fileNum
    Print #fileNum, jsonContent
    Close fileNum

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Sheet exported successfully to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "Rows exported: " & usedRange.Rows.Count & vbCrLf & _
           "Columns exported: " & usedRange.Columns.Count, _
           vbInformation, "Export Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If fileNum > 0 Then Close fileNum
    MsgBox "Error exporting to JSON: " & Err.Description, vbCritical, "Export Error"
End Sub

Private Function BuildJSONFromRange(dataRange As Range, useHeaders As Boolean) As String
    On Error Resume Next

    Dim jsonArray As String
    Dim rowData As String
    Dim cellValue As String
    Dim headers() As String
    Dim startRow As Integer

    jsonArray = "["

    ' Get headers if requested
    If useHeaders Then
        ReDim headers(1 To dataRange.Columns.Count)
        For j = 1 To dataRange.Columns.Count
            headers(j) = EscapeJSONString(CStr(dataRange.Cells(1, j).Value))
            If headers(j) = "" Then headers(j) = "Column" & j
        Next j
        startRow = 2
    Else
        startRow = 1
    End If

    ' Process each row
    For i = startRow To dataRange.Rows.Count
        If i > startRow Then jsonArray = jsonArray & ","

        rowData = "{"
        For j = 1 To dataRange.Columns.Count
            If j > 1 Then rowData = rowData & ","

            ' Get cell value
            cellValue = CStr(dataRange.Cells(i, j).Value)

            ' Determine field name
            Dim fieldName As String
            If useHeaders Then
                fieldName = headers(j)
            Else
                fieldName = "Column" & j
            End If

            ' Add to JSON (treat all as strings for simplicity)
            rowData = rowData & """" & fieldName & """:""" & EscapeJSONString(cellValue) & """"
        Next j
        rowData = rowData & "}"

        jsonArray = jsonArray & rowData
    Next i

    jsonArray = jsonArray & "]"
    BuildJSONFromRange = jsonArray
End Function

Private Function EscapeJSONString(inputStr As String) As String
    Dim result As String
    result = inputStr

    ' Escape special characters
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbTab, "\t")

    EscapeJSONString = result
End Function

Public Sub ExportSheetToPDF()
    On Error GoTo ErrorHandler

    If ActiveSheet Is Nothing Then
        MsgBox "No active sheet found.", vbExclamation, "Export to PDF"
        Exit Sub
    End If

    ' Get file path from user
    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFilename:=ActiveSheet.Name & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save PDF Export As")

    If filePath = "False" Then Exit Sub

    ' Export options
    Dim exportOption As VbMsgBoxResult
    exportOption = MsgBox("Export options:" & vbCrLf & vbCrLf & _
                          "Yes = Entire sheet" & vbCrLf & _
                          "No = Used range only" & vbCrLf & _
                          "Cancel = Current selection", _
                          vbYesNoCancel + vbQuestion, "PDF Export Options")

    If exportOption = 0 Then Exit Sub ' User closed dialog

    Application.StatusBar = "Exporting to PDF..."
    Application.ScreenUpdating = False

    ' Determine what to export
    Select Case exportOption
        Case vbYes ' Entire sheet
            ActiveSheet.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath, _
                Quality:=xlQualityStandard, _
                IncludeDocProps:=True, _
                IgnorePrintAreas:=True, _
                OpenAfterPublish:=False

        Case vbNo ' Used range only
            Dim usedRange As Range
            Set usedRange = ActiveSheet.UsedRange
            If Not usedRange Is Nothing Then
                usedRange.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=filePath, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProps:=True, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=False
            End If

        Case vbCancel ' Current selection
            If TypeName(Selection) = "Range" Then
                Selection.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=filePath, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProps:=True, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=False
            Else
                MsgBox "Please select a range to export.", vbExclamation, "Export to PDF"
                GoTo ErrorHandler
            End If
    End Select

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim openFile As VbMsgBoxResult
    openFile = MsgBox("PDF exported successfully!" & vbCrLf & vbCrLf & _
                      "Location: " & filePath & vbCrLf & vbCrLf & _
                      "Would you like to open the PDF file?", _
                      vbYesNo + vbInformation, "Export Complete")

    If openFile = vbYes Then
        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL """ & filePath & """", vbNormalFocus
    End If

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error exporting to PDF: " & Err.Description, vbCritical, "Export Error"
End Sub

Public Sub ExportWorkbookToPDF()
    On Error GoTo ErrorHandler

    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation, "Export Workbook to PDF"
        Exit Sub
    End If

    ' Get file path from user
    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFilename:=ActiveWorkbook.Name & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save Workbook PDF As")

    If filePath = "False" Then Exit Sub

    ' Show sheet selection options
    Dim sheetOption As VbMsgBoxResult
    sheetOption = MsgBox("Select sheets to export:" & vbCrLf & vbCrLf & _
                         "Yes = All worksheets" & vbCrLf & _
                         "No = All visible worksheets only" & vbCrLf & _
                         "Cancel = Active sheet only", _
                         vbYesNoCancel + vbQuestion, "Sheet Selection")

    If sheetOption = 0 Then Exit Sub ' User closed dialog

    Application.StatusBar = "Exporting workbook to PDF..."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Select Case sheetOption
        Case vbYes ' All worksheets
            ActiveWorkbook.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath, _
                Quality:=xlQualityStandard, _
                IncludeDocProps:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False

        Case vbNo ' All visible worksheets only
            Dim visibleSheets As String
            Dim ws As Worksheet
            Dim firstSheet As Boolean
            firstSheet = True

            ' Get list of visible sheets
            For Each ws In ActiveWorkbook.Worksheets
                If ws.Visible = xlSheetVisible Then
                    If Not firstSheet Then
                        visibleSheets = visibleSheets & ","
                    End If
                    visibleSheets = visibleSheets & ws.Name
                    firstSheet = False
                End If
            Next ws

            ' Export visible sheets
            If visibleSheets <> "" Then
                ActiveWorkbook.Worksheets(Split(visibleSheets, ",")).Select
                ActiveWorkbook.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=filePath, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProps:=True, _
                    IgnorePrintAreas:=False, _
                    OpenAfterPublish:=False
            End If

        Case vbCancel ' Active sheet only
            ActiveSheet.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath, _
                Quality:=xlQualityStandard, _
                IncludeDocProps:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
    End Select

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim openFile As VbMsgBoxResult
    openFile = MsgBox("Workbook exported successfully!" & vbCrLf & vbCrLf & _
                      "Location: " & filePath & vbCrLf & vbCrLf & _
                      "Would you like to open the PDF file?", _
                      vbYesNo + vbInformation, "Export Complete")

    If openFile = vbYes Then
        Shell "rundll32.exe shell32.dll,ShellExec_RunDLL """ & filePath & """", vbNormalFocus
    End If

    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error exporting workbook to PDF: " & Err.Description, vbCritical, "Export Error"
End Sub

' ===================================================================
' DASHBOARD LAUNCHER AND PROGRESS BAR
' ===================================================================

' Global variables for progress bar
Public g_ProgressForm As Object
Public g_ProgressMax As Long
Public g_ProgressCurrent As Long

Public Sub ShowUtilityDashboard()
    On Error GoTo ErrorHandler

    ' Create and show the dashboard form
    Call CreateDashboardForm

    Exit Sub

ErrorHandler:
    MsgBox "Error opening dashboard: " & Err.Description, vbCritical, "Dashboard Error"
End Sub

Private Sub CreateDashboardForm()
    On Error GoTo ErrorHandler

    ' Create a new workbook for the dashboard (temporary)
    Dim dashWB As Workbook
    Set dashWB = Workbooks.Add

    ' Create the dashboard interface on the first sheet
    Dim dashSheet As Worksheet
    Set dashSheet = dashWB.Worksheets(1)
    dashSheet.Name = "Utility Dashboard"

    ' Set up the dashboard layout
    Application.ScreenUpdating = False

    With dashSheet
        ' Title
        .Range("B2").Value = "Excel Utility Macros Dashboard"
        .Range("B2").Font.Size = 16
        .Range("B2").Font.Bold = True
        .Range("B2:G2").Merge

        ' Progress bar area
        .Range("B4").Value = "Progress:"
        .Range("C4:G4").Merge
        .Range("C4").Value = "Ready"
        .Range("C4").Name = "ProgressText"

        ' Create sections with buttons
        Dim currentRow As Integer
        currentRow = 6

        ' Formula Reference Functions
        .Range("B" & currentRow).Value = "📊 FORMULA REFERENCE FUNCTIONS"
        .Range("B" & currentRow).Font.Bold = True
        currentRow = currentRow + 1

        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Convert to Absolute ($A$1)", "ConvertToAbsoluteFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Convert to Relative (A1)", "ConvertToRelativeFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Toggle References", "ToggleReferencesFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Advanced Options", "ConvertToAbsoluteAdvancedFromDash")
        currentRow = currentRow + 2

        ' Error Handling Functions
        .Range("B" & currentRow).Value = "🛡️ ERROR HANDLING FUNCTIONS"
        .Range("B" & currentRow).Font.Bold = True
        currentRow = currentRow + 1

        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Wrap with IFERROR", "WrapWithIFERRORFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Quick IFERROR (0)", "WrapWithIFERRORQuickFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Remove IFERROR", "RemoveIFERRORFromDash")
        currentRow = currentRow + 2

        ' Date Formatting Functions
        .Range("B" & currentRow).Value = "📅 DATE FORMATTING FUNCTIONS"
        .Range("B" & currentRow).Font.Bold = True
        currentRow = currentRow + 1

        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Format to Calendar (MMM-YYYY)", "FormatDatesToCalendarFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Advanced Date Options", "FormatDatesToCalendarAdvancedFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Remove Date Formatting", "RemoveDateFormattingFromDash")
        currentRow = currentRow + 2

        ' Export Functions
        .Range("B" & currentRow).Value = "📤 EXPORT FUNCTIONS"
        .Range("B" & currentRow).Font.Bold = True
        currentRow = currentRow + 1

        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Export Sheet to JSON", "ExportSheetToJSONFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Export Sheet to PDF", "ExportSheetToPDFFromDash")
        currentRow = currentRow + 1
        Call CreateDashboardButton(dashSheet, "C" & currentRow, "Export Workbook to PDF", "ExportWorkbookToPDFFromDash")
        currentRow = currentRow + 2

        ' Utility buttons
        Call CreateDashboardButton(dashSheet, "B" & currentRow, "Show Help", "ShowUtilityMacrosInfoFromDash")
        Call CreateDashboardButton(dashSheet, "D" & currentRow, "Close Dashboard", "CloseDashboard")

        ' Format the sheet
        .Columns("A:H").AutoFit
        .Range("A1:H" & currentRow + 2).Font.Name = "Segoe UI"

        ' Add borders and styling
        With .Range("B2:G" & currentRow + 1)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        ' Color the header
        .Range("B2:G2").Interior.Color = RGB(70, 130, 180)
        .Range("B2:G2").Font.Color = RGB(255, 255, 255)

        ' Color section headers
        Dim sectionHeaders As Range
        Set sectionHeaders = .Range("B6,B11,B16,B21")
        sectionHeaders.Interior.Color = RGB(240, 240, 240)
        sectionHeaders.Font.Color = RGB(50, 50, 50)
    End With

    Application.ScreenUpdating = True

    ' Show the dashboard workbook
    dashWB.Activate

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error creating dashboard: " & Err.Description, vbCritical, "Dashboard Error"
End Sub

Private Sub CreateDashboardButton(ws As Worksheet, cellAddress As String, buttonText As String, macroName As String)
    On Error Resume Next

    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range(cellAddress).Left, ws.Range(cellAddress).Top, ws.Range(cellAddress).Width * 3, 25)

    With btn
        .Caption = buttonText
        .OnAction = macroName
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
End Sub

' Progress bar functions
Public Sub InitializeProgressBar(maxValue As Long, taskDescription As String)
    g_ProgressMax = maxValue
    g_ProgressCurrent = 0

    ' Update progress text
    On Error Resume Next
    ActiveWorkbook.Names("ProgressText").RefersToRange.Value = taskDescription & " (0%)"

    ' Update status bar
    Application.StatusBar = taskDescription & " - 0% complete"
End Sub

Public Sub UpdateProgressBar(currentValue As Long, Optional statusText As String = "")
    g_ProgressCurrent = currentValue

    Dim percentage As Integer
    If g_ProgressMax > 0 Then
        percentage = Int((g_ProgressCurrent / g_ProgressMax) * 100)
    Else
        percentage = 0
    End If

    ' Update progress text in dashboard
    On Error Resume Next
    Dim progressMsg As String
    If statusText <> "" Then
        progressMsg = statusText & " (" & percentage & "%)"
    Else
        progressMsg = "Processing... (" & percentage & "%)"
    End If

    ActiveWorkbook.Names("ProgressText").RefersToRange.Value = progressMsg

    ' Update status bar
    Application.StatusBar = progressMsg

    ' Refresh screen
    DoEvents
End Sub

Public Sub CompleteProgressBar(Optional completionMessage As String = "Task completed successfully!")
    On Error Resume Next
    ActiveWorkbook.Names("ProgressText").RefersToRange.Value = completionMessage
    Application.StatusBar = completionMessage

    ' Clear after 2 seconds
    Application.Wait Now + TimeValue("0:00:02")
    Application.StatusBar = False
    ActiveWorkbook.Names("ProgressText").RefersToRange.Value = "Ready"
End Sub

' Dashboard wrapper functions that include progress tracking
Public Sub ConvertToAbsoluteFromDash()
    Call InitializeProgressBar(100, "Converting to absolute references")
    Call UpdateProgressBar(20, "Preparing conversion")
    Call ConvertToAbsolute
    Call CompleteProgressBar("Conversion to absolute references completed!")
End Sub

Public Sub ConvertToRelativeFromDash()
    Call InitializeProgressBar(100, "Converting to relative references")
    Call UpdateProgressBar(20, "Preparing conversion")
    Call ConvertToRelative
    Call CompleteProgressBar("Conversion to relative references completed!")
End Sub

Public Sub ToggleReferencesFromDash()
    Call InitializeProgressBar(100, "Toggling references")
    Call UpdateProgressBar(20, "Analyzing current references")
    Call ToggleReferences
    Call CompleteProgressBar("Reference toggling completed!")
End Sub

Public Sub ConvertToAbsoluteAdvancedFromDash()
    Call InitializeProgressBar(100, "Advanced conversion options")
    Call UpdateProgressBar(10, "Loading options")
    Call ConvertToAbsoluteAdvanced
    Call CompleteProgressBar("Advanced conversion completed!")
End Sub

Public Sub WrapWithIFERRORFromDash()
    Call InitializeProgressBar(100, "Wrapping cells with IFERROR")
    Call UpdateProgressBar(20, "Preparing IFERROR wrapper")
    Call WrapWithIFERROR
    Call CompleteProgressBar("IFERROR wrapping completed!")
End Sub

Public Sub WrapWithIFERRORQuickFromDash()
    Call InitializeProgressBar(100, "Quick IFERROR wrapping")
    Call UpdateProgressBar(30, "Applying IFERROR with 0")
    Call WrapWithIFERRORQuick
    Call CompleteProgressBar("Quick IFERROR wrapping completed!")
End Sub

Public Sub RemoveIFERRORFromDash()
    Call InitializeProgressBar(100, "Removing IFERROR wrappers")
    Call UpdateProgressBar(25, "Analyzing IFERROR functions")
    Call RemoveIFERROR
    Call CompleteProgressBar("IFERROR removal completed!")
End Sub

Public Sub FormatDatesToCalendarFromDash()
    Call InitializeProgressBar(100, "Formatting dates to calendar")
    Call UpdateProgressBar(20, "Detecting date cells")
    Call FormatDatesToCalendar
    Call CompleteProgressBar("Date formatting completed!")
End Sub

Public Sub FormatDatesToCalendarAdvancedFromDash()
    Call InitializeProgressBar(100, "Advanced date formatting")
    Call UpdateProgressBar(15, "Loading format options")
    Call FormatDatesToCalendarAdvanced
    Call CompleteProgressBar("Advanced date formatting completed!")
End Sub

Public Sub RemoveDateFormattingFromDash()
    Call InitializeProgressBar(100, "Removing date formatting")
    Call UpdateProgressBar(30, "Resetting date formats")
    Call RemoveDateFormatting
    Call CompleteProgressBar("Date formatting reset completed!")
End Sub

Public Sub ExportSheetToJSONFromDash()
    Call InitializeProgressBar(100, "Exporting sheet to JSON")
    Call UpdateProgressBar(20, "Preparing JSON export")
    Call ExportSheetToJSON
    Call CompleteProgressBar("JSON export completed!")
End Sub

Public Sub ExportSheetToPDFFromDash()
    Call InitializeProgressBar(100, "Exporting sheet to PDF")
    Call UpdateProgressBar(25, "Preparing PDF export")
    Call ExportSheetToPDF
    Call CompleteProgressBar("PDF export completed!")
End Sub

Public Sub ExportWorkbookToPDFFromDash()
    Call InitializeProgressBar(100, "Exporting workbook to PDF")
    Call UpdateProgressBar(20, "Preparing workbook export")
    Call ExportWorkbookToPDF
    Call CompleteProgressBar("Workbook PDF export completed!")
End Sub

Public Sub ShowUtilityMacrosInfoFromDash()
    Call InitializeProgressBar(100, "Loading help information")
    Call UpdateProgressBar(50, "Preparing help content")
    Call ShowUtilityMacrosInfo
    Call CompleteProgressBar("Help displayed!")
End Sub

Public Sub CloseDashboard()
    On Error Resume Next
    Dim response As VbMsgBoxResult
    response = MsgBox("Close the Utility Dashboard?", vbYesNo + vbQuestion, "Close Dashboard")

    If response = vbYes Then
        ActiveWorkbook.Close SaveChanges:=False
    End If
End Sub

Public Sub ShowUtilityMacrosInfo()
    Dim infoMsg As String
    infoMsg = "Excel Utility Macros Add-in v1.0" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "FORMULA REFERENCE FUNCTIONS:" & vbCrLf
    infoMsg = infoMsg & "• ConvertToAbsolute - Convert to $A$1" & vbCrLf
    infoMsg = infoMsg & "• ConvertToRelative - Convert to A1" & vbCrLf
    infoMsg = infoMsg & "• ConvertToMixedColumn - Convert to $A1" & vbCrLf
    infoMsg = infoMsg & "• ConvertToMixedRow - Convert to A$1" & vbCrLf
    infoMsg = infoMsg & "• ToggleReferences - Toggle between absolute/relative" & vbCrLf
    infoMsg = infoMsg & "• ConvertToAbsoluteAdvanced - Choose conversion type" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "ERROR HANDLING FUNCTIONS:" & vbCrLf
    infoMsg = infoMsg & "• WrapWithIFERROR - Wrap cells with IFERROR function" & vbCrLf
    infoMsg = infoMsg & "• WrapWithIFERRORQuick - Quick wrap with 0 as error value" & vbCrLf
    infoMsg = infoMsg & "• RemoveIFERROR - Remove IFERROR wrapper from cells" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "DATE FORMATTING FUNCTIONS:" & vbCrLf
    infoMsg = infoMsg & "• FormatDatesToCalendar - Format dates as MMM-YYYY (bold)" & vbCrLf
    infoMsg = infoMsg & "• FormatDatesToCalendarAdvanced - Choose date format options" & vbCrLf
    infoMsg = infoMsg & "• RemoveDateFormatting - Reset date formatting to General" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "EXPORT FUNCTIONS:" & vbCrLf
    infoMsg = infoMsg & "• ExportSheetToJSON - Export active sheet as JSON file" & vbCrLf
    infoMsg = infoMsg & "• ExportSheetToPDF - Export sheet/selection to PDF" & vbCrLf
    infoMsg = infoMsg & "• ExportWorkbookToPDF - Export entire workbook to PDF" & vbCrLf & vbCrLf
    infoMsg = infoMsg & "Use the dashboard for easy access to all functions!" & vbCrLf
    infoMsg = infoMsg & "Run ShowUtilityDashboard to open the control panel."

    MsgBox infoMsg, vbInformation, "Utility Macros Info"
End Sub

' ===================================================================
' ADD-IN EVENT HANDLERS
' ===================================================================

Private Sub Workbook_AddinInstall()
    MsgBox "Excel Utility Macros Add-in installed successfully!" & vbCrLf & vbCrLf & _
           "Access functions through Developer > Macros or assign to buttons/shortcuts." & vbCrLf & vbCrLf & _
           "Available categories:" & vbCrLf & _
           "• Formula Reference Conversion" & vbCrLf & _
           "• Error Handling (IFERROR)" & vbCrLf & _
           "• More utilities coming soon!", _
           vbInformation, "Utility Macros"
End Sub

Private Sub Workbook_AddinUninstall()
    MsgBox "Excel Utility Macros Add-in uninstalled.", vbInformation, "Utility Macros"
End Sub