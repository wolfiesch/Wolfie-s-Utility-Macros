' VBScript to create Excel Add-in (.xlam) file with gridline formatting macros
' This script automates Excel to create a macro-enabled add-in for cross-workbook use

Option Explicit

Dim xlApp, xlWorkbook, xlModule
Dim scriptPath, basFilePath, xlamFilePath
Dim fso, basFile, macroCode

' Constants for Excel
Const xlAddIn = 18  ' xlAddIn format
Const vbext_ct_StdModule = 1

' Get script directory
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
basFilePath = scriptPath & "\GridlineAddin.bas"
xlamFilePath = scriptPath & "\GridlineFormatting.xlam"

' Check if .bas file exists
If Not fso.FileExists(basFilePath) Then
    WScript.Echo "Error: GridlineAddin.bas not found in " & scriptPath
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

' Make Excel visible for debugging (set to False for background operation)
xlApp.Visible = True
xlApp.DisplayAlerts = False

WScript.Echo "Creating new Excel add-in workbook..."

' Create new workbook for the add-in
Set xlWorkbook = xlApp.Workbooks.Add

WScript.Echo "Reading add-in macro code from " & basFilePath

' Read the .bas file content
Set basFile = fso.OpenTextFile(basFilePath, 1) ' ForReading
macroCode = basFile.ReadAll()
basFile.Close

WScript.Echo "Adding VBA module to add-in workbook..."

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
xlModule.Name = "GridlineFormattingAddin"
xlModule.CodeModule.AddFromString macroCode

WScript.Echo "VBA module added successfully!"

' Set add-in properties
With xlWorkbook
    .Title = "Gridline Formatting Add-in"
    .Subject = "Excel formatting tools for gridlines, zoom, and navigation"
    .Comments = "Created for FuzzySum project - Formats worksheets across all workbooks"
    .Author = "FuzzySum Development Team"
    .Keywords = "Excel, Add-in, Gridlines, Formatting, Zoom, Navigation"
End With

' Add ThisWorkbook code for add-in events
Dim thisWBModule
Set thisWBModule = xlWorkbook.VBProject.VBComponents("ThisWorkbook")

' Add add-in install/uninstall event handlers
thisWBModule.CodeModule.AddFromString _
"Private Sub Workbook_AddinInstall()" & vbCrLf & _
"    ' Called when add-in is installed" & vbCrLf & _
"    Call InitializeGridlineAddin" & vbCrLf & _
"End Sub" & vbCrLf & vbCrLf & _
"Private Sub Workbook_AddinUninstall()" & vbCrLf & _
"    ' Called when add-in is uninstalled" & vbCrLf & _
"    Call CleanupGridlineAddin" & vbCrLf & _
"End Sub" & vbCrLf & vbCrLf & _
"Private Sub Workbook_Open()" & vbCrLf & _
"    ' Called when add-in is loaded" & vbCrLf & _
"    If Me.IsAddin Then" & vbCrLf & _
"        Call InitializeGridlineAddin" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub"

WScript.Echo "Add-in event handlers added!"

' Save as .xlam (Excel Add-in)
WScript.Echo "Saving as " & xlamFilePath

On Error Resume Next

' Set as add-in before saving
xlWorkbook.IsAddin = True

' Save as Excel Add-in format
xlWorkbook.SaveAs xlamFilePath, xlAddIn

If Err.Number <> 0 Then
    WScript.Echo "Error saving add-in file: " & Err.Description
    WScript.Echo "Trying alternative save method..."
    Err.Clear
    xlApp.DisplayAlerts = True
    xlWorkbook.SaveAs xlamFilePath, xlAddIn
    xlApp.DisplayAlerts = False
End If

On Error GoTo 0

WScript.Echo "Excel Add-in (.xlam) created successfully!"
WScript.Echo "File saved as: " & xlamFilePath
WScript.Echo ""
WScript.Echo "The add-in contains:"
WScript.Echo "- FormatAllSheetsComplete() - Main formatting macro"
WScript.Echo "- DisableAllGridlines() - Gridlines only"
WScript.Echo "- SetZoomToStandard() - 85% zoom only"
WScript.Echo "- ResetToHomePosition() - Return to A1 only"
WScript.Echo "- FormatActiveSheetOnly() - Current sheet only"
WScript.Echo "- FormatWithOptions() - Interactive formatting"
WScript.Echo ""
WScript.Echo "Keyboard shortcuts (when add-in is loaded):"
WScript.Echo "- Ctrl+Shift+F - Format all sheets (complete)"
WScript.Echo "- Ctrl+Shift+G - Disable gridlines only"
WScript.Echo "- Ctrl+Shift+Z - Set zoom to 85% only"
WScript.Echo "- Ctrl+Shift+H - Return to A1 only"
WScript.Echo "- Ctrl+Shift+A - Format active sheet only"
WScript.Echo ""

' Provide installation instructions
WScript.Echo "INSTALLATION INSTRUCTIONS:"
WScript.Echo "1. Close this Excel instance"
WScript.Echo "2. Copy " & xlamFilePath & " to Excel's AddIns folder"
WScript.Echo "3. Open Excel → File → Options → Add-ins"
WScript.Echo "4. Click 'Go...' next to 'Manage: Excel Add-ins'"
WScript.Echo "5. Check 'GridlineFormatting' in the list"
WScript.Echo "6. Click OK"
WScript.Echo ""
WScript.Echo "The add-in will then be available in all workbooks!"

' Ask if user wants to install the add-in now
Dim installNow
installNow = MsgBox("Would you like to install the add-in now?", 4, "Install Add-in")

If installNow = 6 Then ' Yes
    WScript.Echo "Installing add-in..."

    ' Get Excel's AddIns folder
    Dim addInsPath
    addInsPath = xlApp.Application.UserLibraryPath

    ' Copy file to AddIns folder
    Dim targetPath
    targetPath = addInsPath & "GridlineFormatting.xlam"

    On Error Resume Next
    fso.CopyFile xlamFilePath, targetPath, True

    If Err.Number = 0 Then
        WScript.Echo "Add-in copied to: " & targetPath

        ' Try to install the add-in programmatically
        xlApp.AddIns.Add(targetPath).Installed = True

        If Err.Number = 0 Then
            WScript.Echo "Add-in installed successfully!"
            WScript.Echo "The formatting macros are now available in all workbooks."
        Else
            WScript.Echo "Add-in copied but manual installation required."
            WScript.Echo "Please follow the manual installation steps above."
        End If
    Else
        WScript.Echo "Could not copy to AddIns folder. Manual installation required."
        WScript.Echo "Copy " & xlamFilePath & " to " & addInsPath & " manually."
    End If

    On Error GoTo 0
End If

' Ask if user wants to close Excel
Dim closeExcel
closeExcel = MsgBox("Close Excel now?", 4, "Close Application")
If closeExcel = 6 Then ' Yes
    xlWorkbook.Close True
    xlApp.Quit
Else
    WScript.Echo "Excel left open. You can test the add-in manually."
End If

' Clean up
Set thisWBModule = Nothing
Set xlModule = Nothing
Set xlWorkbook = Nothing
Set xlApp = Nothing
Set fso = Nothing

WScript.Echo "Add-in creation script completed successfully!"