' VBScript to create Excel .xlsb file with gridline control macro
' This script automates Excel to create a macro-enabled binary workbook

Option Explicit

Dim xlApp, xlWorkbook, xlModule
Dim scriptPath, basFilePath, xlsbFilePath
Dim fso, basFile, macroCode

' Constants for Excel
Const xlExcelBinary = 50  ' xlWorkbookBinary
Const vbext_ct_StdModule = 1

' Get script directory
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
basFilePath = scriptPath & "\DisableGridlines.bas"
xlsbFilePath = scriptPath & "\GridlineControl.xlsb"

' Check if .bas file exists
If Not fso.FileExists(basFilePath) Then
    WScript.Echo "Error: DisableGridlines.bas not found in " & scriptPath
    WScript.Quit 1
End If

On Error Resume Next

' Create Excel application
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not create Excel application. Is Excel installed?"
    WScript.Quit 1
End If

On Error GoTo 0

' Make Excel visible (optional - set to False for background operation)
xlApp.Visible = True
xlApp.DisplayAlerts = False

WScript.Echo "Creating new Excel workbook..."

' Create new workbook
Set xlWorkbook = xlApp.Workbooks.Add

' Add some sample worksheets
xlWorkbook.Worksheets.Add
xlWorkbook.Worksheets.Add

' Rename the worksheets
xlWorkbook.Worksheets(1).Name = "Data"
xlWorkbook.Worksheets(2).Name = "Analysis"
xlWorkbook.Worksheets(3).Name = "Results"

' Add some sample data to demonstrate the gridline effect
With xlWorkbook.Worksheets("Data")
    .Range("A1").Value = "Sample Data"
    .Range("A2").Value = "Item"
    .Range("B2").Value = "Value"
    .Range("A3").Value = "Item 1"
    .Range("A4").Value = "Item 2"
    .Range("A5").Value = "Item 3"
    .Range("A6").Value = "Item 4"
    .Range("A7").Value = "Item 5"
    .Range("B3").Value = 125.5
    .Range("B4").Value = 234.75
    .Range("B5").Value = 89.25
    .Range("B6").Value = 456.8
    .Range("B7").Value = 178.9
    .Range("A1:B1").MergeCells = True
    .Range("A1").HorizontalAlignment = -4108 ' xlCenter
End With

WScript.Echo "Reading macro code from " & basFilePath

' Read the .bas file content
Set basFile = fso.OpenTextFile(basFilePath, 1) ' ForReading
macroCode = basFile.ReadAll()
basFile.Close

WScript.Echo "Adding VBA module to workbook..."

On Error Resume Next

' Add VBA module
Set xlModule = xlWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
If Err.Number <> 0 Then
    WScript.Echo "Error adding VBA module. You may need to:"
    WScript.Echo "1. Enable 'Trust access to the VBA project object model' in Excel Options"
    WScript.Echo "2. Run Excel as Administrator"
    xlWorkbook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    WScript.Quit 1
End If

On Error GoTo 0

' Set module name and add code
xlModule.Name = "GridlineControl"
xlModule.CodeModule.AddFromString macroCode

WScript.Echo "VBA module added successfully!"

' Add a button to the first worksheet for easy access
WScript.Echo "Adding control button..."
On Error Resume Next
With xlWorkbook.Worksheets("Data")
    Dim btn
    Set btn = .Shapes.AddFormControl(1, 50, 10, 120, 30) ' xlButtonControl
    If Not btn Is Nothing Then
        btn.OnAction = "DisableAllGridlines"
        btn.Name = "GridlineButton"
        btn.TextFrame.Characters.Text = "Disable Gridlines"
    End If
End With
On Error GoTo 0

WScript.Echo "Added control button to Data worksheet"

' Save as .xlsb (Excel Binary Workbook)
WScript.Echo "Saving as " & xlsbFilePath

On Error Resume Next
xlWorkbook.SaveAs xlsbFilePath, xlExcelBinary
If Err.Number <> 0 Then
    WScript.Echo "Error saving file: " & Err.Description
    WScript.Echo "Trying to save with macro warnings..."
    Err.Clear
    xlApp.DisplayAlerts = True
    xlWorkbook.SaveAs xlsbFilePath, xlExcelBinary
    xlApp.DisplayAlerts = False
End If
On Error GoTo 0

' Optional: Initialize the macro (commented out due to potential security restrictions)
' To test the macro, open the .xlsb file manually and enable macros
WScript.Echo "Note: To test the macro, open GridlineControl.xlsb manually and enable macros"

WScript.Echo "Excel Binary Workbook created successfully!"
WScript.Echo "File saved as: " & xlsbFilePath
WScript.Echo ""
WScript.Echo "The workbook contains:"
WScript.Echo "- GridlineControl VBA module"
WScript.Echo "- DisableAllGridlines macro"
WScript.Echo "- EnableAllGridlines macro"
WScript.Echo "- Control button on Data worksheet"
WScript.Echo "- Keyboard shortcut: Ctrl+Shift+G"

' Ask if user wants to close Excel
Dim closeExcel
closeExcel = MsgBox("Close Excel now?", 4, "Close Application")
If closeExcel = 6 Then ' Yes
    xlWorkbook.Close True
    xlApp.Quit
Else
    WScript.Echo "Excel left open for manual testing"
End If

' Clean up
Set xlModule = Nothing
Set xlWorkbook = Nothing
Set xlApp = Nothing
Set fso = Nothing

WScript.Echo "Script completed successfully!"