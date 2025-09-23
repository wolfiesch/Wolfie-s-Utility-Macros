' Import VBA-only modules (no external dependencies)

Option Explicit

Dim xl, wb, vbProj
Dim fso, scriptPath

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

WScript.Echo "VBA Module Import Helper - VBA-Only Version"
WScript.Echo "=============================================="

' Check if VBA files exist
Dim vbaFile1, vbaFile2
vbaFile1 = scriptPath & "\VersionControlAddin_VBAOnly.bas"
vbaFile2 = scriptPath & "\RibbonCustomization.bas"

If Not fso.FileExists(vbaFile1) Then
    WScript.Echo "Error: " & vbaFile1 & " not found"
    WScript.Quit 1
End If

If Not fso.FileExists(vbaFile2) Then
    WScript.Echo "Error: " & vbaFile2 & " not found"
    WScript.Quit 1
End If

WScript.Echo "Found VBA-only source files:"
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

WScript.Echo "Importing VBA-only modules..."

' Remove existing modules
On Error Resume Next
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlMain")
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlAddin_Simple")
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlAddin_WSH")
vbProj.VBComponents.Remove vbProj.VBComponents("VersionControlAddin_VBAOnly")
vbProj.VBComponents.Remove vbProj.VBComponents("VBAPythonInterface")
vbProj.VBComponents.Remove vbProj.VBComponents("RibbonCustomization")
On Error GoTo 0

' Import VBA-only modules
On Error Resume Next

vbProj.VBComponents.Import vbaFile1
If Err.Number = 0 Then
    WScript.Echo "✓ Imported VersionControlAddin_VBAOnly.bas"
Else
    WScript.Echo "✗ Error importing VersionControlAddin_VBAOnly.bas: " & Err.Description
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

' Update ribbon to use VBA-only module
WScript.Echo "Updating ribbon callbacks..."
On Error Resume Next
Dim ribbonModule
Set ribbonModule = vbProj.VBComponents("RibbonCustomization")
If Not ribbonModule Is Nothing Then
    Dim codeModule
    Set codeModule = ribbonModule.CodeModule

    ' Replace function calls to use VBAOnly module
    Dim i2, lineCount
    lineCount = codeModule.CountOfLines

    For i2 = 1 To lineCount
        Dim lineText
        lineText = codeModule.Lines(i2, 1)

        If InStr(lineText, "VersionControlAddin_Simple.") > 0 Then
            lineText = Replace(lineText, "VersionControlAddin_Simple.", "VersionControlAddin_VBAOnly.")
            codeModule.ReplaceLine i2, lineText
        ElseIf InStr(lineText, "VersionControlAddin_WSH.") > 0 Then
            lineText = Replace(lineText, "VersionControlAddin_WSH.", "VersionControlAddin_VBAOnly.")
            codeModule.ReplaceLine i2, lineText
        End If
    Next i2

    ' Add test function callback
    codeModule.InsertLines lineCount + 1, ""
    codeModule.InsertLines lineCount + 2, "Public Sub OnTestVBASystem(control As IRibbonControl)"
    codeModule.InsertLines lineCount + 3, "    If VersionControlAddin_VBAOnly.TestVBAOnlySystem() Then"
    codeModule.InsertLines lineCount + 4, "        MsgBox ""VBA-only system is working correctly!"", vbInformation"
    codeModule.InsertLines lineCount + 5, "    End If"
    codeModule.InsertLines lineCount + 6, "End Sub"
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
WScript.Echo "VBA-only import completed!"
WScript.Echo ""
WScript.Echo "This version works entirely within Excel:"
WScript.Echo "- No external process execution"
WScript.Echo "- No Python dependencies"
WScript.Echo "- Simple file-based version control"
WScript.Echo "- Maximum compatibility with restricted environments"
WScript.Echo ""
WScript.Echo "Features available:"
WScript.Echo "- Create snapshots (copies of Excel files)"
WScript.Echo "- List versions with metadata"
WScript.Echo "- Basic file comparison"
WScript.Echo "- Rollback to previous versions"
WScript.Echo "- Project statistics"
WScript.Echo ""
WScript.Echo "Run the test function to verify everything works"

' Clean up
Set vbProj = Nothing
Set wb = Nothing
Set xl = Nothing
Set fso = Nothing