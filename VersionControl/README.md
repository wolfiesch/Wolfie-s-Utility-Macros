# Excel Version Control System

A comprehensive hybrid VBA/Python solution for version control of Excel databooks with intelligent metrics extraction and comparison capabilities.

## System Architecture

```
Excel Workbook (VBA UI)
        ↓
VBA-Python Bridge
        ↓
Python Backend
├── Version Controller
├── Metrics Extractor
├── Workbook Comparator
└── Storage Manager
```

## Features

### Core Functionality
- **Version Snapshots**: Create timestamped snapshots with automatic optimization
- **Version Comparison**: Intelligent comparison between workbook versions
- **Metrics Extraction**: Automatic extraction of key financial metrics
- **Rollback**: Safe rollback to previous versions with automatic backup
- **Project Statistics**: Comprehensive analytics and reporting

### Advanced Features
- **Smart Metrics Detection**: Automatically finds EBITDA, Revenue, Assets, etc.
- **Formula Change Detection**: Tracks changes in Excel formulas
- **Configurable Patterns**: Supports multiple databook naming conventions
- **Compression & Optimization**: Reduces file sizes while preserving data
- **Excel Integration**: Native Excel ribbon interface

## Installation

### Prerequisites
```bash
# Python 3.7+ required
python --version

# Install Python dependencies
pip install -r requirements.txt
```

### Dependencies
- `openpyxl>=3.1.0` - Excel file manipulation
- `pandas>=2.0.0` - Data processing
- `pyyaml>=6.0` - Configuration management
- `deepdiff>=6.0.0` - Advanced comparison
- `numpy>=1.24.0` - Numerical operations

### Excel Add-in Setup
1. Run the add-in creation script:
   ```cmd
   cscript create_version_control_addin.vbs
   ```

2. Install the generated add-in:
   - Open Excel
   - File → Options → Add-ins
   - Click "Go..." next to Excel Add-ins
   - Browse and select `VersionControl.xlam`
   - Check the box to enable

## Usage

### From Excel (Recommended)
The system adds a "Version Control" ribbon tab with buttons for:

- **Create Snapshot**: Save current workbook state
- **Compare Versions**: Compare with previous versions
- **List Versions**: View all saved versions
- **Rollback**: Restore to previous version
- **Statistics**: View project analytics

### From Command Line
```bash
# Create snapshot
python version_control.py --action create_snapshot --workbook "path/to/workbook.xlsx" --notes "Description"

# List versions
python version_control.py --action list_versions --workbook "path/to/workbook.xlsx"

# Compare versions
python version_control.py --action compare --workbook "path/to/workbook.xlsx" --version "v001"

# Rollback
python version_control.py --action rollback --workbook "path/to/workbook.xlsx" --version "v001"
```

### Via VBA Bridge
```vba
' Create snapshot
Dim result As Dictionary
Set result = CreateSnapshot(ActiveWorkbook.FullName, "My snapshot notes")

' Compare versions
Set result = CompareToVersion(ActiveWorkbook.FullName, "v001")

' List versions
Set result = ListVersions(ActiveWorkbook.FullName)
```

## Configuration

### Metrics Configuration (`config.yaml`)
```yaml
metrics:
  locations:
    ebitda:
      sheets: ["EBITDA", "EBITDA_Consol", "1.0_EBITDA"]
      patterns: ["EBITDA", "Adjusted EBITDA", "AEBITDA"]
      search_range: "A1:Z100"

    revenue:
      sheets: ["Detail_PL", "Summary_PL", "P&L"]
      patterns: ["Revenue", "Total Revenue", "Net Revenue"]
      search_range: "A1:AZ200"
```

### Version Control Settings
```yaml
version_control:
  max_versions: 50          # Maximum versions to keep
  auto_cleanup: true        # Automatically remove old versions
  backup_current_on_rollback: true

comparison:
  tolerance: 0.01           # Ignore differences less than 1%
  ignore_sheets: ["Temp", "Scratch", "Notes"]

storage:
  compression: true         # Compress snapshots
  optimize_snapshots: true  # Remove empty rows/columns
```

## File Structure

```
VersionControl/
├── version_control.py          # Main controller
├── metrics_extractor.py        # Metrics extraction engine
├── comparator.py              # Workbook comparison engine
├── storage_manager.py         # Version metadata management
├── vba_python_bridge.py       # VBA-Python communication
├── config.yaml               # Configuration file
├── requirements.txt           # Python dependencies
├── test_integration.py        # Integration test suite
│
├── VBA Modules & Add-ins:
├── VersionControlAddin_Simple.bas    # Simplified VBA module
├── VersionControlAddin_VBAOnly.bas   # VBA-only implementation
├── VersionControlAddin_WSH.bas       # Windows Script Host version
├── VBAPythonInterface.bas            # Enhanced VBA-Python interface
├── VersionSelectorForm.frm           # Version selection dialog
├── RibbonCustomization.bas           # Excel ribbon customization
│
├── Add-in Files:
├── VersionControl.xlam               # Main Excel add-in
├── VersionControl_Fixed.xlam         # Fixed version of add-in
│
├── Build & Setup Scripts:
├── create_version_control_addin.vbs  # Main add-in builder
├── create_addin_fixed.vbs           # Fixed add-in builder
├── import_vba_modules.vbs           # VBA module importer
├── import_vba_only.vbs              # VBA-only module importer
├── import_vba_simple.vbs            # Simple module importer
├── import_vba_wsh.vbs               # WSH module importer
│
├── Analysis Tools:
├── analyze_databook.vbs             # Databook analysis script
├── analyze_dependencies.vbs         # Dependency analysis script
├── create_discussion_doc.vbs        # Discussion document generator
├── databook_analysis.txt            # Analysis results
├── dependencies_analysis.txt        # Dependency analysis results
│
└── Documentation:
    ├── README.md                    # This file
    ├── INSTALLATION_GUIDE.md        # Detailed installation guide
    └── TROUBLESHOOTING_GUIDE.md     # Troubleshooting reference

Project Files (Examples):
├── Project Goldfish_QofE Databook_Draft.xlsx
└── Project Nexpera_QofE Databook.xlsx

Generated Files:
├── Versions/
│   └── [project_name]/
│       ├── snapshots/         # Version snapshots
│       └── metadata/          # Version metadata
├── Reports/                   # Comparison reports
└── VersionControl/
    └── logs/                  # System logs
```

## Key Components

### 1. Version Controller (`version_control.py`)
- Main orchestration class
- Handles snapshot creation, comparison, rollback
- Coordinates between all system components

### 2. Metrics Extractor (`metrics_extractor.py`)
- Intelligent pattern-based metric detection
- Supports multiple databook naming conventions
- Configurable search patterns and ranges

### 3. Workbook Comparator (`comparator.py`)
- Cell-by-cell comparison with tolerance
- Formula change detection
- Structural change analysis
- Excel report generation

### 4. Storage Manager (`storage_manager.py`)
- Version metadata storage and retrieval
- File organization and cleanup
- Project statistics and analytics

### 5. VBA-Python Bridge (`vba_python_bridge.py`)
- Reliable VBA-Python communication
- File-based data exchange
- Process timeout handling

## Testing

Run the integration test suite:
```bash
python test_integration.py
```

The test suite verifies:
- Python backend functionality
- VBA-Python communication
- Version control workflows
- Error handling
- System requirements

## Troubleshooting

### Common Issues

**Python not found**
- Ensure Python is in system PATH
- Verify Python version: `python --version`

**Permission errors**
- Run Excel as administrator if needed
- Check file/folder permissions

**Module import errors**
- Install dependencies: `pip install -r requirements.txt`
- Verify Python path in VBA constants

**VBA Dictionary object not found**
- Enable Microsoft Scripting Runtime reference
- Tools → References → Microsoft Scripting Runtime

### Debug Mode

Enable detailed logging by setting log level to DEBUG in Python modules:
```python
logging.basicConfig(level=logging.DEBUG)
```

### Log Files
- VBA Bridge: `VersionControl/logs/vba_bridge.log`
- Version Control: `VersionControl/logs/version_control_[date].log`

## API Reference

### Version Controller Methods

| Method | Description | Parameters |
|--------|-------------|------------|
| `create_snapshot()` | Create version snapshot | `notes`, `quick_save` |
| `list_versions()` | Get all versions | None |
| `compare_to_version()` | Compare to version | `version_name` |
| `rollback_to_version()` | Rollback to version | `version_name`, `backup_current` |
| `get_project_stats()` | Get project statistics | None |

### VBA Interface Functions

| Function | Description | Returns |
|----------|-------------|---------|
| `CreateSnapshot()` | Create snapshot from VBA | Dictionary |
| `ListVersions()` | List versions from VBA | Dictionary |
| `CompareToVersion()` | Compare versions from VBA | Dictionary |
| `RollbackToVersion()` | Rollback from VBA | Dictionary |
| `GetProjectStats()` | Get stats from VBA | Dictionary |

## Performance Optimization

### Large Workbooks
- Use `quick_save=True` for faster snapshots
- Increase timeout values for large files
- Enable compression in configuration

### Storage Management
- Configure automatic cleanup
- Set appropriate `max_versions` limit
- Use `optimize_snapshots` for space savings

## Security Considerations

- Snapshots contain full workbook data
- Configure appropriate file permissions
- Consider encryption for sensitive data
- Regularly backup version metadata

## Future Enhancements

- Cloud storage integration
- Multi-user collaboration features
- Advanced visualization dashboards
- Integration with Git for hybrid workflows
- Real-time collaboration tracking

## License

This project is developed for internal use. All rights reserved.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review log files for error details
3. Run integration tests to verify system health
4. Contact the development team for technical support

---

**Version**: 1.0.0
**Last Updated**: January 2025
**Python Version**: 3.7+
**Excel Version**: 2016+ (Office 365 recommended)