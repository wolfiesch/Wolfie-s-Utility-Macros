' VBScript to analyze formula dependencies in Detail_BS and Detail_PL sheets
Option Explicit

Dim xlApp, xlWorkbook, ws
Dim fso, outputFile
Dim scriptPath, cell, formula
Dim referencedSheets
Dim i, j

Set referencedSheets = CreateObject("Scripting.Dictionary")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Create Excel application
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

' Open the Goldfish databook
Set xlWorkbook = xlApp.Workbooks.Open(scriptPath & "\Project Goldfish_QofE Databook_Draft.xlsx")

' Create output file
Set outputFile = fso.CreateTextFile(scriptPath & "\dependencies_analysis.txt", True)

outputFile.WriteLine "=== DEPENDENCY ANALYSIS ==="
outputFile.WriteLine "Analyzing: Project Goldfish_QofE Databook_Draft.xlsx"
outputFile.WriteLine ""

' Function to extract sheet references from formula
Function ExtractSheetRefs(formula)
    Dim matches, match
    Dim sheetName

    ' Pattern to find sheet references like 'SheetName'! or SheetName!
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "'([^']+)'!|([A-Za-z0-9_]+)!"

    Set matches = re.Execute(formula)
    For Each match In matches
        If match.SubMatches(0) <> "" Then
            sheetName = match.SubMatches(0)
        ElseIf match.SubMatches(1) <> "" Then
            sheetName = match.SubMatches(1)
        End If

        If sheetName <> "" Then
            If Not referencedSheets.Exists(sheetName) Then
                referencedSheets.Add sheetName, 1
            End If
        End If
    Next
End Function

' Analyze Detail_PL sheet
outputFile.WriteLine "ANALYZING: Detail_PL"
outputFile.WriteLine "=================="

On Error Resume Next
Set ws = xlWorkbook.Worksheets("Detail_PL")
If Not ws Is Nothing Then
    outputFile.WriteLine "Sheet found: Detail_PL"
    outputFile.WriteLine "Used Range: " & ws.UsedRange.Address

    ' Sample formulas from key areas
    For i = 1 To 100
        For j = 1 To 20
            Set cell = ws.Cells(i, j)
            If cell.HasFormula Then
                formula = cell.Formula
                ExtractSheetRefs formula
            End If
        Next
    Next
    outputFile.WriteLine "Sample analysis complete"
Else
    outputFile.WriteLine "Detail_PL sheet not found!"
End If
On Error GoTo 0

outputFile.WriteLine ""

' Analyze Detail_BS sheet
outputFile.WriteLine "ANALYZING: Detail_BS"
outputFile.WriteLine "=================="

On Error Resume Next
Set ws = xlWorkbook.Worksheets("Detail_BS")
If Not ws Is Nothing Then
    outputFile.WriteLine "Sheet found: Detail_BS"
    outputFile.WriteLine "Used Range: " & ws.UsedRange.Address

    ' Sample formulas from key areas
    For i = 1 To 100
        For j = 1 To 20
            Set cell = ws.Cells(i, j)
            If cell.HasFormula Then
                formula = cell.Formula
                ExtractSheetRefs formula
            End If
        Next
    Next
    outputFile.WriteLine "Sample analysis complete"
Else
    outputFile.WriteLine "Detail_BS sheet not found!"
End If
On Error GoTo 0

outputFile.WriteLine ""
outputFile.WriteLine "REFERENCED SHEETS DETECTED:"
outputFile.WriteLine "==========================="

' List all referenced sheets
Dim key
For Each key In referencedSheets.Keys
    outputFile.WriteLine "  - " & key

    ' Check if sheet exists in workbook
    On Error Resume Next
    Set ws = Nothing
    Set ws = xlWorkbook.Worksheets(key)
    If Not ws Is Nothing Then
        outputFile.WriteLine "    [EXISTS in workbook]"
    Else
        outputFile.WriteLine "    [EXTERNAL or MISSING]"
    End If
    On Error GoTo 0
Next

outputFile.WriteLine ""
outputFile.WriteLine "WORKBOOK SHEETS:"
outputFile.WriteLine "================"

' List all sheets in workbook for reference
For i = 1 To xlWorkbook.Worksheets.Count
    outputFile.WriteLine i & ". " & xlWorkbook.Worksheets(i).Name
Next

outputFile.Close

' Close Excel
xlWorkbook.Close False
xlApp.Quit

WScript.Echo "Analysis complete! Check dependencies_analysis.txt for results."

' Clean up
Set ws = Nothing
Set xlWorkbook = Nothing
Set xlApp = Nothing
Set outputFile = Nothing
Set fso = Nothing
Set referencedSheets = Nothing