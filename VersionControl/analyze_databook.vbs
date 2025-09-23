' VBScript to analyze databook structure and formatting patterns
Option Explicit

Dim xlApp, xlWorkbook
Dim fso, outputFile
Dim scriptPath

Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Create Excel application
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

' Open the databook
Set xlWorkbook = xlApp.Workbooks.Open(scriptPath & "\Project Nexpera_QofE Databook.xlsx")

' Create output file
Set outputFile = fso.CreateTextFile(scriptPath & "\databook_analysis.txt", True)

outputFile.WriteLine "=== DATABOOK ANALYSIS ==="
outputFile.WriteLine "File: " & xlWorkbook.Name
outputFile.WriteLine "Total Sheets: " & xlWorkbook.Worksheets.Count
outputFile.WriteLine "Date Analyzed: " & Now()
outputFile.WriteLine ""

' Analyze each sheet
Dim ws, i
For i = 1 To xlWorkbook.Worksheets.Count
    Set ws = xlWorkbook.Worksheets(i)

    outputFile.WriteLine "SHEET " & i & ": " & ws.Name
    outputFile.WriteLine "  Used Range: " & ws.UsedRange.Address
    outputFile.WriteLine "  Last Row: " & ws.UsedRange.Rows.Count
    outputFile.WriteLine "  Last Column: " & ws.UsedRange.Columns.Count

    ' Check for specific patterns
    If ws.UsedRange.Rows.Count > 0 Then
        ' Check header patterns
        Dim cell1, cell2, cell3
        On Error Resume Next
        Set cell1 = ws.Range("A1")
        Set cell2 = ws.Range("A2")
        Set cell3 = ws.Range("A3")

        If Not cell1 Is Nothing Then
            outputFile.WriteLine "  A1 Content: " & cell1.Value
        End If
        If Not cell2 Is Nothing Then
            outputFile.WriteLine "  A2 Content: " & cell2.Value
        End If
        If Not cell3 Is Nothing Then
            outputFile.WriteLine "  A3 Content: " & cell3.Value
        End If

        ' Check for merged cells in header area
        If cell1.MergeCells Then
            outputFile.WriteLine "  A1 is merged: " & cell1.MergeArea.Address
        End If

        On Error GoTo 0
    End If

    outputFile.WriteLine ""

    ' Only analyze first 20 sheets in detail to avoid timeout
    If i >= 20 Then
        outputFile.WriteLine "... (truncated for performance)"
        Exit For
    End If
Next

outputFile.WriteLine ""
outputFile.WriteLine "=== COMMON PATTERNS DETECTED ==="

' Look for common sheet naming patterns
outputFile.WriteLine ""
outputFile.WriteLine "Sheet Categories:"
For i = 1 To xlWorkbook.Worksheets.Count
    Set ws = xlWorkbook.Worksheets(i)
    If InStr(ws.Name, "EBITDA") > 0 Then
        outputFile.WriteLine "  EBITDA: " & ws.Name
    ElseIf InStr(ws.Name, "Detail") > 0 Then
        outputFile.WriteLine "  Detail: " & ws.Name
    ElseIf InStr(ws.Name, "Summary") > 0 Then
        outputFile.WriteLine "  Summary: " & ws.Name
    ElseIf InStr(ws.Name, "_db") > 0 Then
        outputFile.WriteLine "  Database: " & ws.Name
    ElseIf InStr(ws.Name, ".") > 0 Then
        outputFile.WriteLine "  Numbered: " & ws.Name
    End If
Next

outputFile.Close

' Close Excel
xlWorkbook.Close False
xlApp.Quit

WScript.Echo "Analysis complete! Check databook_analysis.txt for results."

' Clean up
Set ws = Nothing
Set xlWorkbook = Nothing
Set xlApp = Nothing
Set outputFile = Nothing
Set fso = Nothing