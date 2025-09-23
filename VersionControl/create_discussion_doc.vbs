' VBScript to create Discussion Document from Goldfish databook
' Copies Detail_PL, Detail_BS, IS_Db, and BS_Db sheets with exact formatting
Option Explicit

Dim xlApp, sourceWorkbook, targetWorkbook
Dim fso, scriptPath, sourceFile, targetFile
Dim sheetsToCopy, sheetName, sourceSheet, targetSheet
Dim i

Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
sourceFile = scriptPath & "\Project Goldfish_QofE Databook_Draft.xlsx"
targetFile = scriptPath & "\Discussion Document.xlsx"

' Check if source file exists
If Not fso.FileExists(sourceFile) Then
    WScript.Echo "Error: Source file not found: " & sourceFile
    WScript.Quit 1
End If

' Define sheets to copy (in order)
Dim sheetsArray(3)
sheetsArray(0) = "Detail_PL"
sheetsArray(1) = "Detail_BS"
sheetsArray(2) = "IS_Db"
sheetsArray(3) = "BS_Db"

WScript.Echo "Creating Discussion Document from Goldfish databook..."
WScript.Echo "Source: " & sourceFile
WScript.Echo "Target: " & targetFile
WScript.Echo ""

On Error Resume Next

' Create Excel application
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not create Excel application. Is Excel installed?"
    WScript.Quit 1
End If

xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.ScreenUpdating = False

On Error GoTo 0

WScript.Echo "Opening source workbook..."

' Open source workbook
Set sourceWorkbook = xlApp.Workbooks.Open(sourceFile)

WScript.Echo "Creating new discussion document..."

' Create new workbook for discussion document
Set targetWorkbook = xlApp.Workbooks.Add

' Remove default sheets except one
Dim defaultSheetCount
defaultSheetCount = targetWorkbook.Worksheets.Count
For i = defaultSheetCount To 2 Step -1
    targetWorkbook.Worksheets(i).Delete
Next

WScript.Echo "Copying sheets with exact formatting..."

' Copy each required sheet
For i = 0 To UBound(sheetsArray)
    sheetName = sheetsArray(i)

    WScript.Echo "  Copying: " & sheetName

    On Error Resume Next
    Set sourceSheet = sourceWorkbook.Worksheets(sheetName)

    If Err.Number <> 0 Then
        WScript.Echo "    WARNING: Sheet '" & sheetName & "' not found in source workbook"
        Err.Clear
    Else
        ' Copy the entire sheet with all formatting
        sourceSheet.Copy , targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count)

        If Err.Number <> 0 Then
            WScript.Echo "    ERROR: Failed to copy sheet '" & sheetName & "'"
            Err.Clear
        Else
            WScript.Echo "    SUCCESS: " & sheetName & " copied successfully"
        End If
    End If

    On Error GoTo 0
Next

' Remove the default sheet if we successfully copied sheets
If targetWorkbook.Worksheets.Count > 1 Then
    WScript.Echo "Removing default sheet..."
    targetWorkbook.Worksheets("Sheet1").Delete
End If

WScript.Echo ""
WScript.Echo "Verifying copied sheets..."

' Verify all sheets were copied correctly
For i = 0 To UBound(sheetsArray)
    sheetName = sheetsArray(i)

    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(sheetName)

    If Err.Number <> 0 Then
        WScript.Echo "  MISSING: " & sheetName
        Err.Clear
    Else
        WScript.Echo "  VERIFIED: " & sheetName & " (" & targetSheet.UsedRange.Address & ")"
    End If

    On Error GoTo 0
Next

WScript.Echo ""
WScript.Echo "Saving discussion document..."

' Delete existing target file if it exists
If fso.FileExists(targetFile) Then
    fso.DeleteFile targetFile, True
    WScript.Echo "Deleted existing file: " & targetFile
End If

' Save the discussion document
On Error Resume Next
targetWorkbook.SaveAs targetFile

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Failed to save discussion document: " & Err.Description
    Err.Clear

    ' Try alternative save location
    Dim altFile
    altFile = scriptPath & "\Discussion Document (Alt).xlsx"
    targetWorkbook.SaveAs altFile

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to save to alternative location: " & Err.Description
    Else
        WScript.Echo "Discussion document saved to alternative location: " & altFile
    End If
Else
    WScript.Echo "Discussion document saved successfully: " & targetFile
End If

On Error GoTo 0

WScript.Echo ""
WScript.Echo "=== DISCUSSION DOCUMENT SUMMARY ==="
WScript.Echo "Workbook: " & targetWorkbook.Name
WScript.Echo "Total sheets: " & targetWorkbook.Worksheets.Count
WScript.Echo ""
WScript.Echo "Sheets included:"

For i = 1 To targetWorkbook.Worksheets.Count
    Set targetSheet = targetWorkbook.Worksheets(i)
    WScript.Echo "  " & i & ". " & targetSheet.Name
    WScript.Echo "     Used Range: " & targetSheet.UsedRange.Address
    WScript.Echo "     Last Row: " & targetSheet.UsedRange.Rows.Count
    WScript.Echo "     Last Column: " & targetSheet.UsedRange.Columns.Count
    WScript.Echo ""
Next

' Ask if user wants to open the discussion document
Dim openDoc
openDoc = MsgBox("Discussion document created successfully!" & vbCrLf & vbCrLf & _
                "File: " & targetFile & vbCrLf & vbCrLf & _
                "Would you like to open it now?", vbYesNo, "Success")

If openDoc = vbYes Then
    xlApp.Visible = True
    targetWorkbook.Activate
    targetWorkbook.Worksheets(1).Activate
    WScript.Echo "Discussion document opened in Excel"
Else
    ' Close the workbooks
    targetWorkbook.Close True
    sourceWorkbook.Close False
    xlApp.Quit
    WScript.Echo "Excel closed. Discussion document is ready for use."
End If

' Clean up
Set targetSheet = Nothing
Set sourceSheet = Nothing
Set targetWorkbook = Nothing
Set sourceWorkbook = Nothing
Set xlApp = Nothing
Set fso = Nothing

WScript.Echo "Script completed successfully!"