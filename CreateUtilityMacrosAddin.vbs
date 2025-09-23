' VBScript to create Excel Utility Macros Add-in
' This script creates an Excel Add-in file (.xlam) with various Excel utility functions

Option Explicit

Dim objExcel, objWorkbook, objVBProject, objVBComponent
Dim strBasPath, strXlamPath, strCode
Dim fso, file

' File paths
strBasPath = "C:\Users\wschoenberger\FuzzySum\UtilityMacros.bas"
strXlamPath = "C:\Users\wschoenberger\AppData\Roaming\Microsoft\AddIns\UtilityMacros.xlam"

' Create Excel application
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

' Create new workbook
Set objWorkbook = objExcel.Workbooks.Add

' Access VBA project
Set objVBProject = objWorkbook.VBProject

' Remove default modules if any
On Error Resume Next
Dim i
For i = objVBProject.VBComponents.Count To 1 Step -1
    If objVBProject.VBComponents(i).Type = 1 Then ' vbext_ct_StdModule
        objVBProject.VBComponents.Remove objVBProject.VBComponents(i)
    End If
Next
On Error GoTo 0

' Import the .bas file
objVBProject.VBComponents.Import strBasPath

' Set workbook properties
objWorkbook.Title = "Excel Utility Macros"
objWorkbook.Subject = "Collection of Excel Utility Functions"
objWorkbook.Author = "Excel Utility Tools"
objWorkbook.Comments = "Formula conversion, error handling, and other Excel automation utilities"

' Save as Excel Add-in
objWorkbook.SaveAs strXlamPath, 55 ' xlOpenXMLAddIn

' Install the add-in
objExcel.AddIns.Add(strXlamPath).Installed = True

' Clean up
objWorkbook.Close False
objExcel.Quit

Set objVBComponent = Nothing
Set objVBProject = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

' Display success message
MsgBox "Excel Utility Macros Add-in created and installed successfully!" & vbCrLf & vbCrLf & _
       "Location: " & strXlamPath & vbCrLf & vbCrLf & _
       "The add-in is now available in Excel. Access functions through:" & vbCrLf & _
       "Developer > Macros > Select a function > Run" & vbCrLf & vbCrLf & _
       "Available categories:" & vbCrLf & _
       "• Formula Reference Conversion" & vbCrLf & _
       "• Error Handling (IFERROR)" & vbCrLf & _
       "• More utilities coming soon!", _
       vbInformation, "Add-in Created"