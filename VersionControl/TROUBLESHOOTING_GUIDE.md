# Version Control Troubleshooting Guide

## Problem: Windows Blocked Python Execution

### Root Cause
Corporate security policies are blocking Excel from executing external processes (Python scripts) via the `Shell()` command.

### Solution Options (Try in Order)

## Option 1: WSH-Based Approach (Recommended First Try)

### Install WSH Version:
```cmd
cd C:\Users\wschoenberger\FuzzySum\VersionControl
cscript import_vba_wsh.vbs
```

### What it does:
- Uses `CreateObject("WScript.Shell")` instead of `Shell()`
- Tries PowerShell execution first (often less restricted)
- Falls back to batch file execution
- Better error handling and timeout management

### Test it:
1. Open Excel
2. Go to Version Control ribbon tab
3. Click "Test Connection" button
4. Should show success message if working

---

## Option 2: VBA-Only Version (Maximum Compatibility)

### Install VBA-Only Version:
```cmd
cd C:\Users\wschoenberger\FuzzySum\VersionControl
cscript import_vba_only.vbs
```

### What it does:
- **No external processes** - works entirely within Excel
- **No Python dependencies** - pure VBA implementation
- **Simple file-based version control**
- **Maximum compatibility** with restricted environments

### Features:
- ✅ Create snapshots (copies of Excel files)
- ✅ List versions with metadata
- ✅ Basic file comparison
- ✅ Rollback to previous versions
- ✅ Project statistics
- ❌ Advanced metrics extraction (Python-only feature)
- ❌ Detailed workbook comparison (Python-only feature)

### Test it:
1. Open Excel
2. Go to Version Control ribbon tab
3. Click "Create Snapshot"
4. Should create version without any security warnings

---

## Option 3: Manual File-Based Workflow

If both automated approaches fail, use manual workflow:

### Setup:
1. Create folders:
   ```
   C:\temp\VersionControl\
   ├── Versions\
   ├── Metadata\
   └── Requests\
   ```

2. Use VBA to write request files
3. Manually run Python scripts to process requests

### Manual Process:
1. **Create Snapshot**: Save Excel file with version number
2. **List Versions**: Check folder contents
3. **Compare**: Open two Excel files side by side
4. **Rollback**: Copy old version over current file

---

## Detailed Troubleshooting

### Issue: "Call was rejected by callee"
**Cause**: VBA trying to access restricted Excel automation
**Solution**:
- Restart Excel
- Try importing with Excel open but no other workbooks
- Use Task Manager to kill all Excel processes first

### Issue: Python not found
**Symptoms**: "python is not recognized" or similar
**Solutions**:
1. Check if Python is in PATH: `python --version`
2. Use full Python path in VBA constants
3. Try `py` instead of `python`
4. Install Python if missing

### Issue: PowerShell execution blocked
**Symptoms**: "Execution policy" errors
**Solutions**:
- WSH version uses `-ExecutionPolicy Bypass`
- IT may still block PowerShell entirely
- Fall back to VBA-only version

### Issue: Temp directory access denied
**Symptoms**: Cannot create files in C:\temp\
**Solutions**:
1. Change temp directory in VBA constants:
   ```vba
   Private Const TEMP_DIR As String = "C:\Users\wschoenberger\Documents\VersionControl\"
   ```
2. Use user profile directory instead
3. Use network drive if available

### Issue: File sharing violations
**Symptoms**: "File is in use" errors
**Solutions**:
- Close all instances of the workbook
- Use Task Manager to end Excel processes
- Restart computer if necessary
- Check for hidden Excel processes

---

## Testing Each Version

### Test WSH Version:
```vba
Sub TestWSH()
    If VersionControlAddin_WSH.TestPythonConnection() Then
        MsgBox "WSH version working!"
    Else
        MsgBox "WSH version failed - try VBA-only"
    End If
End Sub
```

### Test VBA-Only Version:
```vba
Sub TestVBAOnly()
    If VersionControlAddin_VBAOnly.TestVBAOnlySystem() Then
        MsgBox "VBA-only version working!"
    Else
        MsgBox "VBA-only version failed - check permissions"
    End If
End Sub
```

---

## Security Policy Workarounds

### For IT Departments:
If you need to request exceptions from IT:

1. **WSH Approach**: Request permission for `WScript.Shell` COM object
2. **Python Execution**: Request permission for Python.exe execution from Office
3. **PowerShell**: Request permission for PowerShell execution with Bypass policy
4. **File System**: Request write access to temp directories

### Alternative Locations:
If C:\temp\ is blocked, try:
- `%USERPROFILE%\Documents\VersionControl\`
- `%APPDATA%\VersionControl\`
- Network drives (if available)
- USB drives (if allowed)

---

## Recommended Approach for Work Laptop

Given your corporate environment restrictions:

### Step 1: Try VBA-Only Version First
- Most likely to work without security issues
- No external processes or dependencies
- Provides core functionality

### Step 2: If VBA-Only Works, Optionally Try WSH
- Better integration with Python backend
- More advanced features
- May work if security is less restrictive

### Step 3: Document What Works
- Note which version works in your environment
- Share findings with colleagues
- Consider requesting IT exceptions if needed

---

## Quick Reference Commands

```cmd
# Try WSH version (bypass Shell restrictions)
cscript import_vba_wsh.vbs

# Try VBA-only version (no external dependencies)
cscript import_vba_only.vbs

# Test Python connection (if WSH version installed)
# Use Test Connection button in Excel ribbon

# Check Python installation
python --version
py --version

# Check PowerShell access
powershell -ExecutionPolicy Bypass -Command "Write-Host 'PowerShell works'"
```

---

## Success Indicators

### WSH Version Working:
- ✅ Test Connection button shows success
- ✅ Create Snapshot works without errors
- ✅ Python backend processes requests
- ✅ Temp files are created and cleaned up

### VBA-Only Version Working:
- ✅ Create Snapshot creates .xlsx files in Versions folder
- ✅ Metadata files are created
- ✅ List Versions shows created snapshots
- ✅ No security warnings or blocks

### If Nothing Works:
- Use manual file copying for version control
- Save Excel files with version numbers
- Use Excel's built-in comparison tools
- Consider cloud-based version control (if allowed)