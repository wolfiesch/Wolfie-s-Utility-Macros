' Import WSH-based VBA modules that bypass Shell restrictions

Option Explicit

Dim xl, wb, vbProj
Dim fso, scriptPath

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

WScript.Echo "VBA Module Import Helper - WSH Version"
WScript.Echo "============================================"

' Check if VBA files exist
Dim vbaFile1, vbaFile2
vbaFile1 = scriptPath & "\VersionControlAddin_WSH.bas"
vbaFile2 = scriptPath & "\RibbonCustomization.bas"

If Not fso.FileExists(vbaFile1) Then
    WScript.Echo "Error: " & vbaFile1 & " not found"
    WScript.Quit 1
End If

If Not fso.FileExists(vbaFile2) Then
    WScript.Echo "Error: " & vbaFile2 & " not found"
    WScript.Quit 1
End If

WScript.Echo "Found WSH-based VBA source files:"
WScript.Echo "- " & vbaFile1
WScript.Echo "- " & vbaFile2
WScript.Echo ""

' Connect to Excel
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

' Find Version Control add-in
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
    WScript.Echo "Version Control add-in not found. Please enable it first."
    WScript.Quit 1
End If

' Get VBA project
Set vbProj = wb.VBProject

' Check VBA access
On Error Resume Next
Dim testAccess
testAccess = vbProj.VBComponents.Count
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot access VBA project. Enable 'Trust access to VBA project object model'"
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "Importing WSH-based modules..."

' Remove existing modules
On Error Resume Next
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlMain")
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlAddin_Simple")
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlAddin_WSH")
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlAddin_VBAOnly")
vbProj.VBComponents.Remove vbProj.VBComponents("VBAPythonInterface")
vbProj.VBComponents.Remove vbProj.VBComponents("RibbonCustomization")
On Error GoTo 0

' Import WSH modules
On Error Resume Next

vbProj.VBComponents.Import vbaFile1
If Err.Number = 0 Then
    WScript.Echo "✓ Imported VersionControlAddin_WSH.bas"
Else
    WScript.Echo "✗ Error importing VersionControlAddin_WSH.bas: " & Err.Description
    Err.Clear
End If

vbProj.VBComponents.Import vbaFile2
If Err.Number = 0 Then
    WScript.Echo "✓ Imported RibbonCustomization.bas"
Else
    WScript.Echo "✗ Error importing RibbonCustomization.bas: " & Err.Description
    Err.Clear
End If

On Error GoTo 0

' Update ribbon to use WSH module
WScript.Echo "Updating ribbon callbacks..."
On Error Resume Next
Dim ribbonModule
Set ribbonModule = vbProj.VBComponents("RibbonCustomization")
If Not ribbonModule Is Nothing Then
    ' Update the code to call WSH module instead of Simple module
    Dim codeModule
    Set codeModule = ribbonModule.CodeModule

    ' Replace function calls
    Dim i2, lineCount
    lineCount = codeModule.CountOfLines

    For i2 = 1 To lineCount
        Dim lineText
        lineText = codeModule.Lines(i2, 1)

        If InStr(lineText, "VersionControlAddin_Simple.") > 0 Then
            lineText = Replace(lineText, "VersionControlAddin_Simple.", "VersionControlAddin_WSH.")
            codeModule.ReplaceLine i2, lineText
        End If
    Next i2
End If
On Error GoTo 0

' Save add-in
On Error Resume Next
wb.Save
If Err.Number = 0 Then
    WScript.Echo "✓ Add-in saved successfully"
Else
    WScript.Echo "✗ Error saving add-in: " & Err.Description
End If
On Error GoTo 0

WScript.Echo ""
WScript.Echo "WSH-based import completed!"
WScript.Echo ""
WScript.Echo "This version uses Windows Script Host to bypass Shell restrictions:"
WScript.Echo "- Uses CreateObject(""WScript.Shell"") instead of Shell()"
WScript.Echo "- Tries PowerShell execution as first option"
WScript.Echo "- Falls back to batch file execution"
WScript.Echo "- Provides better error handling"
WScript.Echo ""
WScript.Echo "Test the connection with the 'Test Connection' button"

' Clean up
Set vbProj = Nothing
Set wb = Nothing
Set xl = Nothing
Set fso = Nothing