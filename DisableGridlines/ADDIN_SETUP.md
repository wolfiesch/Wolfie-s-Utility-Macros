# Excel Gridline Formatting Add-in Setup Guide

This guide explains how to install and use the GridlineFormatting.xlam Excel add-in that works across all your workbooks.

## 🎯 What This Add-in Does

The **Gridline Formatting Add-in** provides powerful formatting tools that work on **any open workbook**:

- **Disable gridlines** on all worksheets
- **Set zoom to 85%** for comfortable viewing
- **Return to cell A1** on all worksheets
- **Combine all three** in one command
- **Format active sheet only** option
- **Interactive formatting** with user choices

## 🚀 Quick Installation (Automated)

### Step 1: Generate the Add-in
```cmd
cd C:\Users\wschoenberger\FuzzySum
cscript create_xlam_addon.vbs
```

### Step 2: Install When Prompted
- The script will ask if you want to install the add-in automatically
- Click **Yes** to install immediately
- If automatic installation fails, follow the manual steps below

## 📋 Manual Installation

### Method 1: Excel Add-ins Manager (Recommended)

1. **Copy the add-in file:**
   - Copy `GridlineFormatting.xlam` to your clipboard

2. **Find Excel's AddIns folder:**
   - Open Excel
   - File → Options → Add-ins
   - Note the path shown at the bottom (usually `C:\Users\[username]\AppData\Roaming\Microsoft\AddIns\`)

3. **Install the add-in:**
   - Paste `GridlineFormatting.xlam` into the AddIns folder
   - In Excel: File → Options → Add-ins
   - Click **Go...** next to "Manage: Excel Add-ins"
   - Check **"GridlineFormatting"** in the list
   - Click **OK**

### Method 2: Direct Installation

1. **Open Excel**
2. **File → Options → Add-ins**
3. **Click "Go..." next to "Manage: Excel Add-ins"**
4. **Click "Browse..."**
5. **Navigate to and select `GridlineFormatting.xlam`**
6. **Click OK to install**

## 🎮 Using the Add-in

### Quick Access - Keyboard Shortcuts

Once installed, these shortcuts work in **any workbook**:

| Shortcut | Function |
|----------|----------|
| **Ctrl+Shift+F** | Format all sheets (gridlines off, 85% zoom, A1) |
| **Ctrl+Shift+G** | Disable gridlines only |
| **Ctrl+Shift+Z** | Set zoom to 85% only |
| **Ctrl+Shift+H** | Return to A1 on all sheets |
| **Ctrl+Shift+A** | Format active sheet only |

### Available Macros

You can also run macros directly (Alt+F8):

#### Main Functions
- **`FormatAllSheetsComplete`** - Complete formatting (gridlines, zoom, A1)
- **`FormatActiveSheetOnly`** - Format current sheet only
- **`FormatWithOptions`** - Interactive menu to choose what to format

#### Individual Functions
- **`DisableAllGridlines`** - Turn off gridlines on all sheets
- **`SetZoomToStandard`** - Set 85% zoom on all sheets
- **`ResetToHomePosition`** - Return to A1 on all sheets
- **`EnableAllGridlines`** - Restore gridlines on all sheets

## 🛠️ Advanced Usage

### Interactive Formatting

Run `FormatWithOptions` for a dialog that lets you choose:
- **gridlines** - Disable gridlines
- **zoom** - Set to 85%
- **home** - Return to A1
- **all** - All of the above

Example inputs:
- `all` - Complete formatting
- `gridlines,zoom` - Gridlines and zoom only
- `home` - Return to A1 only

### Checking Status

Use these utility functions to check current state:
- **`GridlinesEnabled()`** - Returns True if gridlines are on
- **`GetCurrentZoom()`** - Returns current zoom percentage

## 🔧 File Structure

After installation, you'll have:

```
FuzzySum/
├── GridlineAddin.bas           # VBA source code for add-in
├── create_xlam_addon.vbs      # Automation script
├── GridlineFormatting.xlam    # The Excel add-in file
└── ADDIN_SETUP.md            # This documentation

Excel AddIns Folder/
└── GridlineFormatting.xlam    # Installed add-in (copy)
```

## ⚠️ Troubleshooting

### Add-in Not Appearing in List
- **Check file location**: Ensure `.xlam` file is in the AddIns folder
- **Restart Excel**: Close and reopen Excel completely
- **Security settings**: File → Options → Trust Center → Trusted Locations → Add the AddIns folder

### Macros Not Working
- **Enable macros**: When opening Excel, click "Enable Content"
- **Macro security**: File → Options → Trust Center → Macro Settings → Enable all macros
- **Check add-in is loaded**: File → Options → Add-ins → Verify "GridlineFormatting" is checked

### Keyboard Shortcuts Not Working
- **Initialize add-in**: Run `InitializeGridlineAddin` macro once
- **Conflicting shortcuts**: Other add-ins might use the same keys
- **Alternative access**: Use Alt+F8 to run macros directly

### "No Active Workbook" Error
- **Open a workbook first**: The add-in needs an open workbook to work on
- **Check workbook protection**: Some protected workbooks may block changes

## 🎯 Usage Examples

### Example 1: Quick Format New Workbook
1. Open any Excel workbook
2. Press **Ctrl+Shift+F**
3. All sheets now have: no gridlines, 85% zoom, cursor at A1

### Example 2: Format Specific Aspects
1. Open workbook
2. Press **Alt+F8** to open macro dialog
3. Run `FormatWithOptions`
4. Type `gridlines,zoom` and press OK

### Example 3: Format Just the Current Sheet
1. Navigate to the sheet you want to format
2. Press **Ctrl+Shift+A**
3. Only the current sheet is formatted

## 🔄 Uninstalling the Add-in

1. **File → Options → Add-ins**
2. **Click "Go..." next to "Manage: Excel Add-ins"**
3. **Uncheck "GridlineFormatting"**
4. **Click OK**
5. **Optionally delete** `GridlineFormatting.xlam` from the AddIns folder

## 🆚 Add-in vs. Workbook Macros

### Excel Add-in (.xlam) - ✅ Recommended
- **Works across all workbooks**
- **Always available** when Excel is open
- **No need to copy macros** to each workbook
- **Professional distribution** - single file
- **Automatic keyboard shortcuts**

### Workbook Macros (.xlsb) - Limited Use
- **Only works in that specific workbook**
- **Must be opened** to access macros
- **Need to copy/paste** to other workbooks
- **Good for workbook-specific** automation

## 🔒 Security Considerations

- **Only install add-ins from trusted sources**
- **Review VBA code** before installation if security is critical
- **Use digital signatures** for enterprise distribution
- **Test in non-production environment** first

## 🎨 Customization

### Modify Zoom Level
Edit the `SetZoomToStandard` function in the VBA code:
```vba
ActiveWindow.Zoom = 90  ' Change from 85 to your preference
```

### Add More Shortcuts
Edit the `InitializeGridlineAddin` function:
```vba
Application.OnKey "^+r", "YourCustomMacro"  ' Ctrl+Shift+R
```

### Change Default Cell Position
Edit the `ResetToHomePosition` function:
```vba
ws.Range("B2").Select  ' Change from A1 to B2
```

## 📞 Support

### Common Issues
1. **Add-in not loading**: Check Excel version (2007+) and macro security
2. **Permission errors**: Run Excel as Administrator once
3. **File path issues**: Ensure no special characters in folder names
4. **VBA access denied**: Enable "Trust access to VBA project object model"

### Testing the Add-in
1. **Create a new workbook** with multiple sheets
2. **Add some data** and scroll away from A1
3. **Set different zoom levels** on different sheets
4. **Run `FormatAllSheetsComplete`**
5. **Verify**: All sheets should have no gridlines, 85% zoom, cursor at A1

---

**Created for FuzzySum Project**
*Enhanced subset sum solver with Excel integration*

**Add-in Version**: 1.0
**Compatibility**: Excel 2007+
**File Size**: ~25KB