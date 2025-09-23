' Create Version Control Excel Add-in
' Builds comprehensive Excel add-in for the version control system

Option Explicit

Dim xl, wb, ws, module1, module2, userForm
Dim fso, scriptPath, addinPath

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Create Excel application
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

' Create new workbook
Set wb = xl.Workbooks.Add

' Remove default worksheets (keep one)
Do While wb.Worksheets.Count > 1
    wb.Worksheets(wb.Worksheets.Count).Delete
Loop

' Rename remaining worksheet
wb.Worksheets(1).Name = "VersionControl"
Set ws = wb.Worksheets("VersionControl")

' Add header information
ws.Range("A1").Value = "Excel Version Control System"
ws.Range("A2").Value = "Hybrid VBA/Python Solution"
ws.Range("A3").Value = "Created: " & Now()
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Size = 14

' Add usage instructions
ws.Range("A5").Value = "Instructions:"
ws.Range("A6").Value = "1. Use Ribbon buttons or Developer tab for version control functions"
ws.Range("A7").Value = "2. Ensure Python backend is installed and configured"
ws.Range("A8").Value = "3. Python script location: " & scriptPath & "\version_control.py"

' Format instruction cells
With ws.Range("A5:A8")
    .WrapText = True
    .RowHeight = 15
End With

' Add main VBA module
Set module1 = wb.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule
module1.Name = "VersionControlMain"

' Read VBA code from file
Dim vbaCode
vbaCode = ReadFile(scriptPath & "\VersionControlAddin.bas")
If vbaCode <> "" Then
    module1.CodeModule.AddFromString vbaCode
End If

' Add ribbon customization module
Set module2 = wb.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule
module2.Name = "RibbonCustomization"

' Add ribbon XML code
Dim ribbonCode
ribbonCode = GetRibbonCode()
module2.CodeModule.AddFromString ribbonCode

' Add UserForm for version selection
' Note: VBScript can't directly create UserForms with controls
' This would need to be done manually or with additional automation

' Set workbook properties for add-in
wb.IsAddin = True
wb.Title = "Excel Version Control System"
wb.Subject = "Version Control Add-in"
wb.Comments = "Hybrid VBA/Python version control system for Excel databooks"

' Create Ribbon XML (stored as custom property)
Dim ribbonXML
ribbonXML = GetRibbonXML()
wb.CustomDocumentProperties.Add "RibbonXML", False, 4, ribbonXML

' Save as Excel Add-in
addinPath = scriptPath & "\VersionControl.xlam"
wb.SaveAs addinPath, 18 ' xlAddIn

' Clean up
wb.Close False
xl.Quit

Set ws = Nothing
Set wb = Nothing
Set xl = Nothing
Set fso = Nothing

WScript.Echo "Version Control add-in created successfully!"
WScript.Echo "Saved as: " & addinPath
WScript.Echo ""
WScript.Echo "To install:"
WScript.Echo "1. Open Excel"
WScript.Echo "2. Go to File > Options > Add-ins"
WScript.Echo "3. Click 'Go...' next to Excel Add-ins"
WScript.Echo "4. Click 'Browse...' and select: " & addinPath
WScript.Echo "5. Check the box next to 'Version Control System'"

' Functions
Function ReadFile(filePath)
    On Error Resume Next
    Dim file, content
    Set file = fso.OpenTextFile(filePath, 1)
    If Err.Number = 0 Then
        content = file.ReadAll
        file.Close
        ReadFile = content
    Else
        ReadFile = ""
    End If
    On Error GoTo 0
End Function

Function GetRibbonCode()
    Dim code
    code = "' Ribbon Customization for Version Control" & vbCrLf
    code = code & "' Implements custom ribbon tab with version control commands" & vbCrLf & vbCrLf
    code = code & "Option Explicit" & vbCrLf & vbCrLf

    code = code & "' Ribbon callback procedures" & vbCrLf
    code = code & "Public Sub OnCreateSnapshot(control As IRibbonControl)" & vbCrLf
    code = code & "    Call VersionControlMain.CreateVersionSnapshot" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Public Sub OnCompareVersions(control As IRibbonControl)" & vbCrLf
    code = code & "    Call VersionControlMain.CompareToVersion" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Public Sub OnListVersions(control As IRibbonControl)" & vbCrLf
    code = code & "    Call VersionControlMain.ListVersions" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Public Sub OnRollback(control As IRibbonControl)" & vbCrLf
    code = code & "    Call VersionControlMain.RollbackToVersion" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "Public Sub OnShowStats(control As IRibbonControl)" & vbCrLf
    code = code & "    Call VersionControlMain.ShowProjectStats" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    code = code & "' Get image for ribbon buttons" & vbCrLf
    code = code & "Public Sub GetButtonImage(control As IRibbonControl, ByRef image)" & vbCrLf
    code = code & "    ' Return built-in Office icons" & vbCrLf
    code = code & "    Select Case control.Id" & vbCrLf
    code = code & "        Case ""btnCreateSnapshot""" & vbCrLf
    code = code & "            Set image = ""FileSave""" & vbCrLf
    code = code & "        Case ""btnCompareVersions""" & vbCrLf
    code = code & "            Set image = ""ReviewTrackChangesMenu""" & vbCrLf
    code = code & "        Case ""btnListVersions""" & vbCrLf
    code = code & "            Set image = ""FileDocumentManagementMenu""" & vbCrLf
    code = code & "        Case ""btnRollback""" & vbCrLf
    code = code & "            Set image = ""Undo""" & vbCrLf
    code = code & "        Case ""btnShowStats""" & vbCrLf
    code = code & "            Set image = ""ChartInsertMenu""" & vbCrLf
    code = code & "    End Select" & vbCrLf
    code = code & "End Sub" & vbCrLf

    GetRibbonCode = code
End Function

Function GetRibbonXML()
    Dim xml
    xml = "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    xml = xml & "  <ribbon>" & vbCrLf
    xml = xml & "    <tabs>" & vbCrLf
    xml = xml & "      <tab id=""tabVersionControl"" label=""Version Control"">" & vbCrLf
    xml = xml & "        <group id=""grpSnapshot"" label=""Snapshots"">" & vbCrLf
    xml = xml & "          <button id=""btnCreateSnapshot"" label=""Create Snapshot"" size=""large""" & vbCrLf
    xml = xml & "                  onAction=""RibbonCustomization.OnCreateSnapshot""" & vbCrLf
    xml = xml & "                  getImage=""RibbonCustomization.GetButtonImage""" & vbCrLf
    xml = xml & "                  screentip=""Create Version Snapshot""" & vbCrLf
    xml = xml & "                  supertip=""Save current workbook state as a new version""/>" & vbCrLf
    xml = xml & "          <button id=""btnListVersions"" label=""List Versions"" size=""normal""" & vbCrLf
    xml = xml & "                  onAction=""RibbonCustomization.OnListVersions""" & vbCrLf
    xml = xml & "                  getImage=""RibbonCustomization.GetButtonImage""" & vbCrLf
    xml = xml & "                  screentip=""List All Versions""" & vbCrLf
    xml = xml & "                  supertip=""View all saved versions of this workbook""/>" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "        <group id=""grpCompare"" label=""Compare"">" & vbCrLf
    xml = xml & "          <button id=""btnCompareVersions"" label=""Compare Versions"" size=""large""" & vbCrLf
    xml = xml & "                  onAction=""RibbonCustomization.OnCompareVersions""" & vbCrLf
    xml = xml & "                  getImage=""RibbonCustomization.GetButtonImage""" & vbCrLf
    xml = xml & "                  screentip=""Compare to Version""" & vbCrLf
    xml = xml & "                  supertip=""Compare current workbook to a previous version""/>" & vbCrLf
    xml = xml & "          <button id=""btnShowStats"" label=""Statistics"" size=""normal""" & vbCrLf
    xml = xml & "                  onAction=""RibbonCustomization.OnShowStats""" & vbCrLf
    xml = xml & "                  getImage=""RibbonCustomization.GetButtonImage""" & vbCrLf
    xml = xml & "                  screentip=""Project Statistics""" & vbCrLf
    xml = xml & "                  supertip=""View version control statistics for this project""/>" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "        <group id=""grpRestore"" label=""Restore"">" & vbCrLf
    xml = xml & "          <button id=""btnRollback"" label=""Rollback"" size=""large""" & vbCrLf
    xml = xml & "                  onAction=""RibbonCustomization.OnRollback""" & vbCrLf
    xml = xml & "                  getImage=""RibbonCustomization.GetButtonImage""" & vbCrLf
    xml = xml & "                  screentip=""Rollback to Version""" & vbCrLf
    xml = xml & "                  supertip=""Restore workbook to a previous version""/>" & vbCrLf
    xml = xml & "        </group>" & vbCrLf
    xml = xml & "      </tab>" & vbCrLf
    xml = xml & "    </tabs>" & vbCrLf
    xml = xml & "  </ribbon>" & vbCrLf
    xml = xml & "</customUI>"

    GetRibbonXML = xml
End Function