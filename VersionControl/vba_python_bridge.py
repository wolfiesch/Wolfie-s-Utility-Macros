#!/usr/bin/env python3
"""
VBA-Python Communication Bridge
Provides a reliable interface between Excel VBA and Python backend
"""

import sys
import json
import logging
import tempfile
import os
from pathlib import Path
from typing import Dict, Any, Optional
import argparse

# Add the VersionControl directory to Python path
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from version_control import VersionController


class VBAPythonBridge:
    """Bridge class to handle VBA-Python communication"""

    def __init__(self):
        self.setup_logging()
        self.logger = logging.getLogger(__name__)

    def setup_logging(self):
        """Setup logging for the bridge"""
        log_dir = Path("VersionControl") / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_dir / "vba_bridge.log"),
                logging.StreamHandler()
            ]
        )

    def execute_command(self, workbook_path: str, action: str, **kwargs) -> Dict[str, Any]:
        """
        Execute a version control command and return structured result

        Args:
            workbook_path: Path to the Excel workbook
            action: Action to perform
            **kwargs: Additional parameters

        Returns:
            Dictionary with execution results
        """
        try:
            self.logger.info(f"Executing command: {action} for workbook: {workbook_path}")

            # Validate workbook path
            if not Path(workbook_path).exists():
                return {
                    "success": False,
                    "error": f"Workbook not found: {workbook_path}"
                }

            # Initialize version controller
            vc = VersionController(workbook_path)

            # Execute the requested action
            if action == "create_snapshot":
                return self._create_snapshot(vc, kwargs)

            elif action == "list_versions":
                return self._list_versions(vc)

            elif action == "compare":
                return self._compare_versions(vc, kwargs)

            elif action == "rollback":
                return self._rollback(vc, kwargs)

            elif action == "stats":
                return self._get_stats(vc)

            elif action == "get_version_info":
                return self._get_version_info(vc, kwargs)

            else:
                return {
                    "success": False,
                    "error": f"Unknown action: {action}"
                }

        except Exception as e:
            self.logger.error(f"Error executing command {action}: {str(e)}")
            return {
                "success": False,
                "error": str(e)
            }

    def _create_snapshot(self, vc: VersionController, params: Dict[str, Any]) -> Dict[str, Any]:
        """Create version snapshot"""
        notes = params.get("notes", "")
        quick_save = params.get("quick", False)

        result = vc.create_snapshot(notes, quick_save)
        return result

    def _list_versions(self, vc: VersionController) -> Dict[str, Any]:
        """List all versions"""
        try:
            versions = vc.list_versions()
            return {
                "success": True,
                "versions": versions,
                "count": len(versions)
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }

    def _compare_versions(self, vc: VersionController, params: Dict[str, Any]) -> Dict[str, Any]:
        """Compare current workbook to a version"""
        version_name = params.get("version")
        if not version_name:
            return {
                "success": False,
                "error": "Version name required for comparison"
            }

        result = vc.compare_to_version(version_name)
        return result

    def _rollback(self, vc: VersionController, params: Dict[str, Any]) -> Dict[str, Any]:
        """Rollback to a specific version"""
        version_name = params.get("version")
        if not version_name:
            return {
                "success": False,
                "error": "Version name required for rollback"
            }

        backup_current = params.get("backup_current", True)
        result = vc.rollback_to_version(version_name, backup_current)
        return result

    def _get_stats(self, vc: VersionController) -> Dict[str, Any]:
        """Get project statistics"""
        try:
            stats = vc.get_project_stats()
            return {
                "success": True,
                **stats
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }

    def _get_version_info(self, vc: VersionController, params: Dict[str, Any]) -> Dict[str, Any]:
        """Get detailed information about a specific version"""
        version_name = params.get("version")
        if not version_name:
            return {
                "success": False,
                "error": "Version name required"
            }

        try:
            version_info = vc.storage.get_version_info(version_name)
            if version_info:
                return {
                    "success": True,
                    **version_info
                }
            else:
                return {
                    "success": False,
                    "error": f"Version {version_name} not found"
                }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }

    def write_result_to_file(self, result: Dict[str, Any], output_file: str) -> bool:
        """
        Write result to a temporary file for VBA to read

        Args:
            result: Result dictionary to write
            output_file: Path to output file

        Returns:
            True if successful, False otherwise
        """
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2, default=str, ensure_ascii=False)
            return True
        except Exception as e:
            self.logger.error(f"Error writing result to file: {str(e)}")
            return False


def main():
    """Command line interface for VBA-Python bridge"""
    parser = argparse.ArgumentParser(description='VBA-Python Bridge for Version Control')
    parser.add_argument('--workbook', required=True, help='Path to Excel workbook')
    parser.add_argument('--action', required=True, help='Action to perform')
    parser.add_argument('--output', help='Output file for results')
    parser.add_argument('--notes', default='', help='Notes for snapshot creation')
    parser.add_argument('--version', help='Version name for comparison or rollback')
    parser.add_argument('--quick', action='store_true', help='Quick save (skip optimization)')
    parser.add_argument('--backup-current', action='store_true', default=True,
                       help='Backup current version before rollback')

    args = parser.parse_args()

    # Initialize bridge
    bridge = VBAPythonBridge()

    # Prepare parameters
    params = {
        'notes': args.notes,
        'version': args.version,
        'quick': args.quick,
        'backup_current': args.backup_current
    }

    # Remove None values
    params = {k: v for k, v in params.items() if v is not None}

    # Execute command
    result = bridge.execute_command(args.workbook, args.action, **params)

    # Output result
    if args.output:
        # Write to specified file
        success = bridge.write_result_to_file(result, args.output)
        if not success:
            sys.exit(1)
    else:
        # Print to stdout for VBA to capture
        print(json.dumps(result, indent=2, default=str, ensure_ascii=False))

    # Exit with appropriate code
    if result.get('success', False):
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main()