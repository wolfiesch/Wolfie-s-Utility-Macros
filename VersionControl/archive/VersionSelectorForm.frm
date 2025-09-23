VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VersionSelectorForm
   Caption         =   "Select Version"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535
   OleObjectBlob   =   "VersionSelectorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VersionSelectorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version Selector Dialog Form
' Provides enhanced UI for version selection with preview capabilities

Option Explicit

' Form variables
Private m_SelectedVersion As String
Private m_Versions As Collection
Private m_DialogResult As VbMsgBoxResult

' Properties
Public Property Get SelectedVersion() As String
    SelectedVersion = m_SelectedVersion
End Property

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_DialogResult
End Property

' Form events
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    ' Initialize form
    m_SelectedVersion = ""
    m_DialogResult = vbCancel

    ' Setup controls
    Call SetupControls

    ' Load versions
    Call LoadVersions

    Exit Sub

ErrorHandler:
    MsgBox "Error initializing form: " & Err.Description, vbCritical, "Version Selector"
End Sub

Private Sub SetupControls()
    ' Configure listbox
    With lstVersions
        .ColumnCount = 4
        .ColumnWidths = "80;120;80;200"
        .ColumnHeads = True
    End With

    ' Setup preview text box
    With txtPreview
        .MultiLine = True
        .ScrollBars = fmScrollBarsBoth
        .Locked = True
    End With

    ' Setup buttons
    btnOK.Enabled = False
    btnCancel.Default = True
End Sub

Private Sub LoadVersions()
    On Error GoTo ErrorHandler

    ' Get versions from Python backend
    Dim versionsJson As String
    versionsJson = ExecutePythonCommand("list_versions", "")

    If Len(versionsJson) > 0 Then
        Call PopulateVersionsList(versionsJson)
    Else
        MsgBox "No versions found.", vbInformation, "Version Selector"
        m_DialogResult = vbCancel
        Me.Hide
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error loading versions: " & Err.Description, vbCritical, "Version Selector"
    m_DialogResult = vbCancel
    Me.Hide
End Sub

Private Sub PopulateVersionsList(jsonText As String)
    On Error GoTo ErrorHandler

    ' Clear existing items
    lstVersions.Clear

    ' Parse JSON and populate listbox
    ' This is a simplified parser - in production would use proper JSON library
    Dim lines() As String
    lines = Split(jsonText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), """value"":") > 0 Then
            Dim versionInfo As String
            versionInfo = lines(i)

            ' Extract version details
            Dim version As String
            Dim displayText As String
            Dim timestamp As String
            Dim fileSize As String

            version = ExtractJsonValue(versionInfo, "value")
            displayText = ExtractJsonValue(versionInfo, "display")
            timestamp = ExtractJsonValue(versionInfo, "timestamp")
            fileSize = ExtractJsonValue(versionInfo, "file_size")

            If version <> "" Then
                ' Add to listbox
                lstVersions.AddItem
                Dim rowIndex As Long
                rowIndex = lstVersions.ListCount - 1

                lstVersions.List(rowIndex, 0) = version
                lstVersions.List(rowIndex, 1) = timestamp
                lstVersions.List(rowIndex, 2) = Format(CLng(fileSize) / 1048576, "0.0") & " MB"
                lstVersions.List(rowIndex, 3) = displayText
            End If
        End If
    Next i

    ' Select first item if available
    If lstVersions.ListCount > 0 Then
        lstVersions.ListIndex = 0
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error populating versions list: " & Err.Description, vbCritical, "Version Selector"
End Sub

Private Sub lstVersions_Click()
    On Error GoTo ErrorHandler

    ' Enable OK button when selection is made
    btnOK.Enabled = (lstVersions.ListIndex >= 0)

    ' Load version preview
    If lstVersions.ListIndex >= 0 Then
        Call LoadVersionPreview
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error processing selection: " & Err.Description, vbCritical, "Version Selector"
End Sub

Private Sub lstVersions_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Double-click selects and closes
    If lstVersions.ListIndex >= 0 Then
        Call btnOK_Click
    End If
End Sub

Private Sub LoadVersionPreview()
    On Error GoTo ErrorHandler

    If lstVersions.ListIndex < 0 Then Exit Sub

    Dim selectedVersion As String
    selectedVersion = lstVersions.List(lstVersions.ListIndex, 0)

    ' Get detailed version info
    Dim versionJson As String
    versionJson = ExecutePythonCommand("get_version_info", "--version """ & selectedVersion & """")

    If Len(versionJson) > 0 Then
        ' Parse and display version details
        Dim previewText As String
        previewText = "Version: " & selectedVersion & vbCrLf
        previewText = previewText & "Date: " & ExtractJsonValue(versionJson, "datetime") & vbCrLf
        previewText = previewText & "File Size: " & Format(CLng(ExtractJsonValue(versionJson, "file_size")) / 1048576, "0.0") & " MB" & vbCrLf
        previewText = previewText & "Notes: " & ExtractJsonValue(versionJson, "notes") & vbCrLf & vbCrLf

        ' Add metrics if available
        previewText = previewText & "Key Metrics:" & vbCrLf

        ' Extract common metrics
        Dim metrics() As String
        metrics = Split("ebitda,revenue,total_assets,total_liabilities,working_capital,net_debt", ",")

        Dim j As Long
        For j = LBound(metrics) To UBound(metrics)
            Dim metricValue As String
            metricValue = ExtractJsonValue(versionJson, metrics(j))
            If metricValue <> "" And IsNumeric(metricValue) Then
                previewText = previewText & "  " & UCase(Left(metrics(j), 1)) & Mid(metrics(j), 2) & ": " & _
                            Format(CDbl(metricValue), "#,##0") & vbCrLf
            End If
        Next j

        txtPreview.Text = previewText
    Else
        txtPreview.Text = "Preview not available"
    End If

    Exit Sub

ErrorHandler:
    txtPreview.Text = "Error loading preview: " & Err.Description
End Sub

Private Sub btnOK_Click()
    On Error GoTo ErrorHandler

    If lstVersions.ListIndex >= 0 Then
        m_SelectedVersion = lstVersions.List(lstVersions.ListIndex, 0)
        m_DialogResult = vbOK
        Me.Hide
    Else
        MsgBox "Please select a version.", vbExclamation, "Version Selector"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error confirming selection: " & Err.Description, vbCritical, "Version Selector"
End Sub

Private Sub btnCancel_Click()
    m_SelectedVersion = ""
    m_DialogResult = vbCancel
    Me.Hide
End Sub

Private Sub btnRefresh_Click()
    On Error GoTo ErrorHandler

    ' Reload versions list
    Call LoadVersions

    Exit Sub

ErrorHandler:
    MsgBox "Error refreshing versions: " & Err.Description, vbCritical, "Version Selector"
End Sub

Private Sub btnCompare_Click()
    On Error GoTo ErrorHandler

    ' Allow comparison between two selected versions
    If lstVersions.ListIndex >= 0 Then
        Dim selectedVersion As String
        selectedVersion = lstVersions.List(lstVersions.ListIndex, 0)

        ' For now, just show a message
        MsgBox "Version comparison feature will be implemented in future release." & vbCrLf & _
               "Selected version: " & selectedVersion, vbInformation, "Version Selector"
    Else
        MsgBox "Please select a version to compare.", vbExclamation, "Version Selector"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error initiating comparison: " & Err.Description, vbCritical, "Version Selector"
End Sub

' Utility functions
Private Function ExtractJsonValue(jsonText As String, key As String) As String
    On Error GoTo ErrorHandler

    Dim startPos As Long
    Dim endPos As Long
    Dim searchPattern As String

    searchPattern = """" & key & """: """
    startPos = InStr(jsonText, searchPattern)

    If startPos > 0 Then
        startPos = startPos + Len(searchPattern)
        endPos = InStr(startPos, jsonText, """")

        If endPos > startPos Then
            ExtractJsonValue = Mid(jsonText, startPos, endPos - startPos)
        End If
    End If

    Exit Function

ErrorHandler:
    ExtractJsonValue = ""
End Function

Private Function ExecutePythonCommand(action As String, parameters As String) As String
    ' This would call the main module's ExecutePythonCommand function
    ' For now, return empty string as placeholder
    ExecutePythonCommand = ""
End Function

' Public methods for external access
Public Function ShowDialog() As VbMsgBoxResult
    Me.Show vbModal
    ShowDialog = m_DialogResult
End Function

Public Sub SetPrompt(prompt As String)
    lblPrompt.Caption = prompt
End Sub