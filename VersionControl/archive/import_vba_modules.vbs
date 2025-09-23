' VBA Module Import Helper
' Helps import VBA modules into the Version Control add-in

Option Explicit

Dim xl, wb, vbProj
Dim fso, scriptPath

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

WScript.Echo "VBA Module Import Helper for Version Control Add-in"
WScript.Echo "=================================================="

' Check if VBA files exist
Dim vbaFile1, vbaFile2
vbaFile1 = scriptPath & "\VersionControlAddin.bas"
vbaFile2 = scriptPath & "\VBAPythonInterface.bas"

If Not fso.FileExists(vbaFile1) Then
    WScript.Echo "Error: " & vbaFile1 & " not found"
    WScript.Quit 1
End If

If Not fso.FileExists(vbaFile2) Then
    WScript.Echo "Error: " & vbaFile2 & " not found"
    WScript.Quit 1
End If

WScript.Echo "Found VBA source files:"
WScript.Echo "- " & vbaFile1
WScript.Echo "- " & vbaFile2
WScript.Echo ""

' Try to connect to Excel
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        WScript.Echo "Error: Could not connect to Excel"
        WScript.Quit 1
    End If
End If
On Error GoTo 0

xl.Visible = True

' Look for the Version Control add-in
Dim foundAddin
foundAddin = False

Dim i
For i = 1 To xl.Workbooks.Count
    Set wb = xl.Workbooks(i)
    If wb.IsAddin And (InStr(LCase(wb.Name), "versioncontrol") > 0 Or InStr(LCase(wb.Title), "version control") > 0) Then
        WScript.Echo "Found Version Control add-in: " & wb.Name
        foundAddin = True
        Exit For
    End If
Next

If Not foundAddin Then
    WScript.Echo "Version Control add-in not found in open workbooks."
    WScript.Echo "Please ensure the add-in is enabled and Excel is open."
    WScript.Echo ""
    WScript.Echo "To enable the add-in:"
    WScript.Echo "1. Go to File > Options > Add-ins"
    WScript.Echo "2. Click 'Go...' next to Excel Add-ins"
    WScript.Echo "3. Check 'Excel Version Control System'"
    WScript.Quit 1
End If

' Get VBA project
Set vbProj = wb.VBProject

' Check if we can access VBA project
On Error Resume Next
Dim testAccess
testAccess = vbProj.VBComponents.Count
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot access VBA project. This may be due to macro security settings."
    WScript.Echo ""
    WScript.Echo "To enable VBA access:"
    WScript.Echo "1. Go to File > Options > Trust Center > Trust Center Settings"
    WScript.Echo "2. Click 'Macro Settings'"
    WScript.Echo "3. Check 'Trust access to the VBA project object model'"
    WScript.Echo "4. Restart Excel and try again"
    Err.Clear
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "VBA project access confirmed. Importing modules..."

' Remove existing modules if they exist
On Error Resume Next
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlMain")
vbProj.VBComponents.Remove vbProj.VBComponents("VBAPythonInterface")
On Error GoTo 0

' Import the VBA modules
On Error Resume Next

' Import main module
vbProj.VBComponents.Import vbaFile1
If Err.Number = 0 Then
    WScript.Echo "✓ Imported VersionControlAddin.bas"
Else
    WScript.Echo "✗ Error importing VersionControlAddin.bas: " & Err.Description
    Err.Clear
End If

' Import interface module
vbProj.VBComponents.Import vbaFile2
If Err.Number = 0 Then
    WScript.Echo "✓ Imported VBAPythonInterface.bas"
Else
    WScript.Echo "✗ Error importing VBAPythonInterface.bas: " & Err.Description
    Err.Clear
End If

On Error GoTo 0

' Save the add-in
On Error Resume Next
wb.Save
If Err.Number = 0 Then
    WScript.Echo "✓ Add-in saved successfully"
Else
    WScript.Echo "✗ Error saving add-in: " & Err.Description
End If
On Error GoTo 0

WScript.Echo ""
WScript.Echo "Import process completed!"
WScript.Echo ""
WScript.Echo "Next steps:"
WScript.Echo "1. Check that the modules were imported correctly"
WScript.Echo "2. Test the version control functions"
WScript.Echo "3. If ribbon buttons don't appear, try restarting Excel"

' Clean up
Set vbProj = Nothing
Set wb = Nothing
Set xl = Nothing
Set fso = Nothing