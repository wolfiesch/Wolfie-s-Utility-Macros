' Fixed Excel Add-in Creator
' Creates a proper Excel add-in without VBA automation issues

Option Explicit

Dim xl, wb, ws
Dim fso, scriptPath, addinPath, tempPath

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

WScript.Echo "Creating Excel Version Control Add-in..."
WScript.Echo "Script path: " & scriptPath

' Create Excel application
On Error Resume Next
Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Could not start Excel. Error: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

xl.Visible = False
xl.DisplayAlerts = False

' Enable programmatic access to VBA project
On Error Resume Next
xl.Application.VBE.MainWindow.Visible = False
If Err.Number <> 0 Then
    WScript.Echo "Warning: VBA access may be restricted. Will create basic add-in structure."
    Err.Clear
End If
On Error GoTo 0

' Create new workbook
Set wb = xl.Workbooks.Add

' Remove default worksheets except one
Do While wb.Worksheets.Count > 1
    wb.Worksheets(wb.Worksheets.Count).Delete
Loop

' Set up the remaining worksheet
Set ws = wb.Worksheets(1)
ws.Name = "VersionControl_Info"

' Add documentation to the sheet
ws.Range("A1").Value = "Excel Version Control System Add-in"
ws.Range("A2").Value = "Version: 1.0"
ws.Range("A3").Value = "Created: " & Now()
ws.Range("A4").Value = "Description: Hybrid VBA/Python version control system"

' Set workbook properties for add-in
wb.IsAddin = True
wb.Title = "Excel Version Control System"
wb.Subject = "Version Control Add-in"
wb.Comments = "Hybrid VBA/Python version control system for Excel databooks"

' Hide the worksheet now that it's an add-in
On Error Resume Next
ws.Visible = 2 ' xlSheetVeryHidden
On Error GoTo 0

' Try to add basic VBA code structure (may fail due to security)
On Error Resume Next
Dim vbModule
Set vbModule = wb.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule
If Err.Number = 0 Then
    vbModule.Name = "VersionControlMain"

    ' Add basic module structure
    Dim basicCode
    basicCode = "' Excel Version Control System - Main Module" & vbCrLf
    basicCode = basicCode & "' This module needs to be populated with the VBA code" & vbCrLf
    basicCode = basicCode & "' Import the code from VersionControlAddin.bas" & vbCrLf & vbCrLf
    basicCode = basicCode & "Option Explicit" & vbCrLf & vbCrLf
    basicCode = basicCode & "Public Sub CreateVersionSnapshot()" & vbCrLf
    basicCode = basicCode & "    MsgBox ""Version Control System - Please import the full VBA code""" & vbCrLf
    basicCode = basicCode & "End Sub"

    vbModule.CodeModule.AddFromString basicCode
    WScript.Echo "Added basic VBA module structure."
Else
    WScript.Echo "Could not add VBA modules automatically. VBA code will need to be imported manually."
    Err.Clear
End If
On Error GoTo 0

' Save as Excel Add-in using the correct format
tempPath = scriptPath & "\VersionControl_Fixed.xlam"

On Error Resume Next
' Try multiple save formats to ensure compatibility
wb.SaveAs tempPath, 55 ' xlOpenXMLAddIn (Excel 2007+ format)
If Err.Number <> 0 Then
    Err.Clear
    wb.SaveAs tempPath, 18 ' xlAddIn (older format)
    If Err.Number <> 0 Then
        WScript.Echo "Error saving add-in: " & Err.Description
        wb.Close False
        xl.Quit
        WScript.Quit 1
    End If
End If
On Error GoTo 0

' Close and clean up
wb.Close False
xl.Quit

Set ws = Nothing
Set wb = Nothing
Set xl = Nothing

' Verify the file was created
If fso.FileExists(tempPath) Then
    WScript.Echo "Add-in created successfully!"
    WScript.Echo "Saved as: " & tempPath

    ' Get file size for verification
    Dim fileSize
    fileSize = fso.GetFile(tempPath).Size
    WScript.Echo "File size: " & fileSize & " bytes"

    ' Copy to AddIns directory
    Dim addinDir, finalPath
    addinDir = CreateObject("WScript.Shell").SpecialFolders("AppData") & "\Microsoft\AddIns"
    finalPath = addinDir & "\VersionControl.xlam"

    ' Ensure AddIns directory exists
    If Not fso.FolderExists(addinDir) Then
        fso.CreateFolder addinDir
    End If

    ' Remove existing file if it exists
    If fso.FileExists(finalPath) Then
        fso.DeleteFile finalPath
    End If

    ' Copy to AddIns directory
    fso.CopyFile tempPath, finalPath

    If fso.FileExists(finalPath) Then
        WScript.Echo "Add-in copied to: " & finalPath
        WScript.Echo ""
        WScript.Echo "Installation Instructions:"
        WScript.Echo "1. Open Excel"
        WScript.Echo "2. Go to File > Options > Add-ins"
        WScript.Echo "3. Click 'Go...' next to Excel Add-ins"
        WScript.Echo "4. The add-in should appear in the list as 'Excel Version Control System'"
        WScript.Echo "5. Check the box to enable it"
        WScript.Echo ""
        WScript.Echo "IMPORTANT: You will need to manually import the VBA code:"
        WScript.Echo "1. Open the Developer tab in Excel"
        WScript.Echo "2. Click 'Visual Basic'"
        WScript.Echo "3. Find the VersionControl add-in project"
        WScript.Echo "4. Import the .bas files from the VersionControl folder"
        WScript.Echo "   - VersionControlAddin.bas"
        WScript.Echo "   - VBAPythonInterface.bas"
    Else
        WScript.Echo "Error: Could not copy to AddIns directory"
    End If
Else
    WScript.Echo "Error: Add-in file was not created"
End If

Set fso = Nothing

WScript.Echo ""
WScript.Echo "Script completed."