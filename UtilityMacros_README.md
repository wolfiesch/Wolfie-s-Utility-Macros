# Excel Utility Macros Add-in

A comprehensive collection of Excel automation utilities including formula conversion, error handling, and more productivity tools.

## Files Created

1. **UtilityMacros.bas** - The VBA module with all utility functions
2. **CreateUtilityMacrosAddin.vbs** - Script to create and install the Excel Add-in
3. **UtilityMacros_README.md** - This documentation file

## Installation

### Method 1: Run the VBScript (Automatic)
1. Double-click `CreateUtilityMacrosAddin.vbs`
2. This will automatically create and install the add-in to your Excel Add-ins folder
3. The add-in will be available in Excel immediately

### Method 2: Manual Installation
1. Open Excel
2. Press Alt+F11 to open the VBA Editor
3. Go to File > Import File and select `UtilityMacros.bas`
4. Save the workbook as an Excel Add-in (.xlam) file
5. Go to File > Options > Add-ins > Manage Excel Add-ins > Browse
6. Select your saved .xlam file and check it to activate

## Available Functions

### 📊 Formula Reference Conversion
- **ConvertToAbsolute()** - Convert all formulas in selection to absolute references ($A$1)
- **ConvertToRelative()** - Convert all formulas in selection to relative references (A1)
- **ConvertToAbsoluteAdvanced()** - Choose conversion type with dialog
- **ConvertToMixedColumn()** - Convert to absolute column references ($A1)
- **ConvertToMixedRow()** - Convert to absolute row references (A$1)
- **QuickConvertToAbsolute()** - Fast conversion without dialogs
- **ToggleReferences()** - Toggle between absolute and relative

### 🛡️ Error Handling Functions
- **WrapWithIFERROR()** - Wrap selected cells with IFERROR function (custom error value)
- **WrapWithIFERRORQuick()** - Quick wrap with 0 as default error value
- **RemoveIFERROR()** - Remove IFERROR wrapper from selected cells

### 📅 Date Formatting Functions
- **FormatDatesToCalendar()** - Format date cells to MMM-YYYY with bold styling
- **FormatDatesToCalendarAdvanced()** - Choose date format options (MMM-YYYY, MMM YYYY, MMMM YYYY)
- **RemoveDateFormatting()** - Reset date formatting to General format

### 📤 Export Functions
- **ExportSheetToJSON()** - Export active sheet data as JSON file with header options
- **ExportSheetToPDF()** - Export sheet to PDF (entire sheet, used range, or selection)
- **ExportWorkbookToPDF()** - Export workbook to PDF (all sheets, visible sheets, or active sheet)

### 🎛️ Dashboard & Control Functions
- **ShowUtilityDashboard()** - Launch the interactive dashboard with all functions
- **ShowUtilityMacrosInfo()** - Display help information about all available functions

## How to Use

### Method 1: Dashboard (Recommended)
1. Press Alt+F8 and run **ShowUtilityDashboard**
2. The dashboard will open in a new workbook with organized buttons
3. Click any button to run the corresponding function with progress tracking
4. Monitor progress in real-time via the progress bar and status updates

### Method 2: Direct Function Access
1. Select a range of cells
2. Press Alt+F8 (or go to Developer > Macros)
3. Choose the desired function
4. Click Run

### Formula Reference Examples

**Before Conversion:**
```
=A1+B2
=SUM(C1:C10)
=VLOOKUP(D1,A:B,2,FALSE)
```

**After ConvertToAbsolute:**
```
=$A$1+$B$2
=SUM($C$1:$C$10)
=VLOOKUP($D$1,$A:$B,2,FALSE)
```

### IFERROR Examples

**Before IFERROR Wrapping:**
```
=A1/B1
=VLOOKUP(D1,A:B,2,FALSE)
=INDEX(A:A,MATCH(E1,B:B,0))
```

**After WrapWithIFERROR (with "N/A" as error value):**
```
=IFERROR(A1/B1,"N/A")
=IFERROR(VLOOKUP(D1,A:B,2,FALSE),"N/A")
=IFERROR(INDEX(A:A,MATCH(E1,B:B,0)),"N/A")
```

**Quick IFERROR (with 0 as error value):**
```
=IFERROR(A1/B1,0)
=IFERROR(VLOOKUP(D1,A:B,2,FALSE),0)
=IFERROR(INDEX(A:A,MATCH(E1,B:B,0)),0)
```

### Date Formatting Examples

**Before Date Formatting:**
```
1/15/2024    (or any date value)
2024-03-20   (or any date format)
March 2024   (or any recognizable date)
```

**After FormatDatesToCalendar:**
```
Jan-2024     (bold formatting applied)
Mar-2024     (bold formatting applied)
Mar-2024     (bold formatting applied)
```

**Non-date cells are skipped:**
```
"Hello"      (unchanged - not a date)
123          (unchanged - not a date)
""           (unchanged - empty cell)
```

### Export Examples

**JSON Export:**
```json
[
  {"Name":"John","Age":"30","City":"New York"},
  {"Name":"Jane","Age":"25","City":"Chicago"},
  {"Name":"Bob","Age":"35","City":"Miami"}
]
```

**PDF Export Options:**
- **Sheet to PDF**: Entire sheet, used range only, or current selection
- **Workbook to PDF**: All worksheets, visible worksheets only, or active sheet only
- Automatic file naming with sheet/workbook names
- Option to open PDF after export

### Dashboard Interface

**Main Dashboard Features:**
- Clean, organized layout with categorized function buttons
- Real-time progress tracking for all operations
- Professional styling with color-coded sections
- Integrated help and close buttons
- Progress bar with percentage and status messages

**Dashboard Layout:**
```
📊 FORMULA REFERENCE FUNCTIONS
  ▪ Convert to Absolute ($A$1)
  ▪ Convert to Relative (A1)
  ▪ Toggle References
  ▪ Advanced Options

🛡️ ERROR HANDLING FUNCTIONS
  ▪ Wrap with IFERROR
  ▪ Quick IFERROR (0)
  ▪ Remove IFERROR

📅 DATE FORMATTING FUNCTIONS
  ▪ Format to Calendar (MMM-YYYY)
  ▪ Advanced Date Options
  ▪ Remove Date Formatting

📤 EXPORT FUNCTIONS
  ▪ Export Sheet to JSON
  ▪ Export Sheet to PDF
  ▪ Export Workbook to PDF
```

## Features

### Formula Reference Conversion
- ✅ Batch conversion of multiple cells
- ✅ Progress indicators for large selections
- ✅ Error handling and validation
- ✅ Detailed conversion statistics
- ✅ Support for all Excel formula types

### Error Handling
- ✅ Wrap formulas and values with IFERROR
- ✅ Custom error value selection
- ✅ Quick wrap with default values
- ✅ Remove IFERROR wrapper functionality
- ✅ Smart detection of existing IFERROR wrapping

### Date Formatting
- ✅ Smart date detection (only formats actual dates)
- ✅ Multiple format options (MMM-YYYY, MMM YYYY, MMMM YYYY)
- ✅ Automatic bold formatting application
- ✅ Skip non-date cells automatically
- ✅ Reset formatting functionality

### Export Functions
- ✅ JSON export with proper escaping and formatting
- ✅ Header row detection and naming options
- ✅ Multiple PDF export options (sheet/workbook/selection)
- ✅ Automatic file naming and save dialogs
- ✅ Option to open exported files immediately
- ✅ Standard quality PDF generation

### Dashboard & User Experience
- ✅ Interactive dashboard with organized function buttons
- ✅ Real-time progress tracking with percentage indicators
- ✅ Professional UI design with color-coded sections
- ✅ Status bar and visual progress feedback
- ✅ One-click access to all utility functions
- ✅ Integrated help and documentation

### General
- ✅ Performance optimization for large ranges
- ✅ Undo support (Ctrl+Z)
- ✅ Comprehensive error handling
- ✅ User-friendly progress feedback

## Common Use Cases

### Error Handling Scenarios
- **Division by zero**: `=A1/B1` → `=IFERROR(A1/B1,0)`
- **VLOOKUP errors**: `=VLOOKUP(...)` → `=IFERROR(VLOOKUP(...),"Not Found")`
- **INDEX/MATCH errors**: `=INDEX(...)` → `=IFERROR(INDEX(...),"N/A")`
- **Array formulas**: Wrap entire array operations to handle errors gracefully

### Formula Reference Scenarios
- **Template creation**: Convert to absolute references for consistent copying
- **Dynamic ranges**: Use mixed references for flexible formulas
- **Report automation**: Toggle between reference types as needed

### Date Formatting Scenarios
- **Monthly reports**: Convert date columns to consistent MMM-YYYY format
- **Dashboard headers**: Format period dates with bold styling for emphasis
- **Calendar views**: Standardize date display across spreadsheets
- **Mixed data cleanup**: Format only date cells while preserving other data types

### Export Scenarios
- **Data sharing**: Export Excel data as JSON for web applications or APIs
- **Report distribution**: Generate PDF reports from Excel worksheets
- **Documentation**: Create PDF documentation from multiple worksheet tabs
- **Data archiving**: Export data in portable formats for long-term storage
- **Presentation**: Share formatted data without requiring Excel installation

### Dashboard Usage Scenarios
- **New users**: Easy discovery and access to all available functions
- **Batch operations**: Run multiple utilities in sequence with progress tracking
- **Training**: Demonstrate utility functions with visual feedback
- **Productivity**: Quick access to frequently used functions
- **Process monitoring**: Track long-running operations with real-time updates

## Keyboard Shortcuts (Optional)

You can assign keyboard shortcuts to frequently used functions:
1. Go to Developer > Macros
2. Select a function and click Options
3. Assign a shortcut key

### Suggested Shortcuts
- **Ctrl+Shift+A** - ConvertToAbsolute
- **Ctrl+Shift+R** - ConvertToRelative
- **Ctrl+Shift+E** - WrapWithIFERROR
- **Ctrl+Shift+Q** - WrapWithIFERRORQuick
- **Ctrl+Shift+T** - ToggleReferences
- **Ctrl+Shift+D** - FormatDatesToCalendar
- **Ctrl+Shift+J** - ExportSheetToJSON
- **Ctrl+Shift+P** - ExportSheetToPDF
- **Ctrl+Shift+U** - ShowUtilityDashboard (Launch Dashboard)

## Troubleshooting

**Functions not appearing:**
- Make sure macros are enabled in Excel
- Check if the add-in is activated in File > Options > Add-ins

**IFERROR wrapping issues:**
- Ensure selected cells contain formulas or values
- Check for nested IFERROR functions (tool will skip already wrapped cells)
- Verify error values are properly quoted for text ("N/A", "Error")

**Date formatting issues:**
- Function only formats cells that contain actual date values
- Text that looks like dates but isn't recognized by Excel will be skipped
- Use Excel's DATEVALUE function to convert text dates if needed
- Mixed content ranges will only format the date cells

**Export issues:**
- JSON export treats all values as strings for maximum compatibility
- PDF export requires Excel 2010 or later for ExportAsFixedFormat support
- Large sheets may take time to export - progress is shown in status bar
- JSON files are UTF-8 encoded and work with most programming languages

**Dashboard issues:**
- If dashboard buttons don't respond, ensure macros are enabled
- Dashboard creates a temporary workbook - close it when finished
- Progress bar updates in real-time during operations
- If dashboard fails to load, try running ShowUtilityDashboard again

**Performance issues:**
- The add-in automatically optimizes for large selections
- For very large ranges (>1000 cells), consider processing in smaller batches
- Use Quick functions for faster processing without dialogs
- Dashboard functions include progress tracking for better user experience

**Undo not working:**
- Excel's Undo (Ctrl+Z) works for all functions
- For complex operations, consider testing on a copy first

## System Requirements

- Microsoft Excel 2010 or later
- Macros must be enabled
- Windows operating system
- VBA environment accessible

## Version History

- **v1.0** - Initial release
  - Interactive dashboard with progress tracking
  - Formula reference conversion functions
  - IFERROR wrapping and unwrapping
  - Date formatting with smart detection
  - Export functions (JSON and PDF)
  - Mixed reference support
  - Advanced options dialogs
  - Error handling and statistics
  - Performance optimization

## Planned Features

- 📋 Copy/Paste special utilities
- 🔍 Advanced find/replace functions
- 📈 Data validation helpers
- 🎨 Formatting utilities
- 📊 Chart automation tools

---

*Created with Excel Utility Tools - Making Excel work smarter, not harder*