#!/usr/bin/env python3
"""
Storage Manager for Excel Databook Version Control
Handles version metadata storage, retrieval, and organization
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import datetime
import shutil
import os


class StorageManager:
    """Manages version metadata and file organization for the version control system"""

    def __init__(self, project_name: str, config: Dict[str, Any]):
        """
        Initialize storage manager

        Args:
            project_name: Name of the project/workbook
            config: Configuration dictionary
        """
        self.project_name = project_name
        self.config = config
        self.logger = logging.getLogger(__name__)

        # Set up storage paths
        self.base_dir = Path("Versions")
        self.project_dir = self.base_dir / self.project_name
        self.metadata_dir = self.project_dir / "metadata"
        self.snapshots_dir = self.project_dir / "snapshots"

        # Create directory structure
        self._setup_directories()

        # Metadata file paths
        self.versions_index_file = self.metadata_dir / "versions_index.json"
        self.project_info_file = self.metadata_dir / "project_info.json"

        # Load or initialize metadata
        self._load_metadata()

    def _setup_directories(self):
        """Create necessary directory structure"""
        directories = [
            self.base_dir,
            self.project_dir,
            self.metadata_dir,
            self.snapshots_dir
        ]

        for directory in directories:
            directory.mkdir(parents=True, exist_ok=True)

    def _load_metadata(self):
        """Load existing metadata or initialize new structures"""
        # Load versions index
        if self.versions_index_file.exists():
            try:
                with open(self.versions_index_file, 'r') as f:
                    self.versions_index = json.load(f)
            except Exception as e:
                self.logger.warning(f"Failed to load versions index: {str(e)}")
                self.versions_index = {"versions": [], "next_version_number": 1}
        else:
            self.versions_index = {"versions": [], "next_version_number": 1}

        # Load project info
        if self.project_info_file.exists():
            try:
                with open(self.project_info_file, 'r') as f:
                    self.project_info = json.load(f)
            except Exception as e:
                self.logger.warning(f"Failed to load project info: {str(e)}")
                self.project_info = self._get_default_project_info()
        else:
            self.project_info = self._get_default_project_info()

    def _get_default_project_info(self) -> Dict[str, Any]:
        """Return default project information structure"""
        return {
            "project_name": self.project_name,
            "created_date": datetime.now().isoformat(),
            "last_accessed": datetime.now().isoformat(),
            "total_versions": 0,
            "total_size_bytes": 0,
            "description": "",
            "tags": [],
            "settings": {
                "auto_cleanup": self.config.get("version_control", {}).get("auto_cleanup", True),
                "max_versions": self.config.get("version_control", {}).get("max_versions", 50),
                "compression_enabled": self.config.get("storage", {}).get("compression", True)
            }
        }

    def save_version_metadata(self, version_info: Dict[str, Any]) -> bool:
        """
        Save metadata for a new version

        Args:
            version_info: Dictionary containing version information

        Returns:
            True if successful, False otherwise
        """
        try:
            version_name = version_info["version"]

            # Create individual version metadata file
            version_metadata_file = self.metadata_dir / f"{version_name}.json"
            with open(version_metadata_file, 'w') as f:
                json.dump(version_info, f, indent=2, default=str)

            # Update versions index
            version_entry = {
                "version": version_name,
                "timestamp": version_info["timestamp"],
                "datetime": version_info["datetime"],
                "file_size": version_info["file_size"],
                "notes": version_info.get("notes", ""),
                "metadata_file": str(version_metadata_file),
                "snapshot_file": version_info["file_path"]
            }

            # Remove existing entry if it exists (for updates)
            self.versions_index["versions"] = [
                v for v in self.versions_index["versions"]
                if v["version"] != version_name
            ]

            # Add new entry
            self.versions_index["versions"].append(version_entry)

            # Sort by version name for consistent ordering
            self.versions_index["versions"].sort(key=lambda x: x["version"])

            # Update next version number
            current_version_num = int(version_name[1:])  # Remove 'v' prefix
            self.versions_index["next_version_number"] = max(
                self.versions_index["next_version_number"],
                current_version_num + 1
            )

            # Save updated index
            self._save_versions_index()

            # Update project info
            self._update_project_info(version_info)

            self.logger.info(f"Saved metadata for version {version_name}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to save version metadata: {str(e)}")
            return False

    def get_version_info(self, version_name: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information for a specific version

        Args:
            version_name: Name of the version (e.g., "v001")

        Returns:
            Version information dictionary or None if not found
        """
        try:
            version_metadata_file = self.metadata_dir / f"{version_name}.json"

            if not version_metadata_file.exists():
                return None

            with open(version_metadata_file, 'r') as f:
                return json.load(f)

        except Exception as e:
            self.logger.error(f"Failed to get version info for {version_name}: {str(e)}")
            return None

    def get_all_versions(self) -> List[Dict[str, Any]]:
        """
        Get list of all versions with basic information

        Returns:
            List of version dictionaries
        """
        try:
            return self.versions_index["versions"].copy()
        except Exception as e:
            self.logger.error(f"Failed to get versions list: {str(e)}")
            return []

    def get_next_version_number(self) -> int:
        """
        Get the next version number to use

        Returns:
            Next version number
        """
        return self.versions_index["next_version_number"]

    def remove_version_metadata(self, version_name: str) -> bool:
        """
        Remove metadata for a specific version

        Args:
            version_name: Name of the version to remove

        Returns:
            True if successful, False otherwise
        """
        try:
            # Remove from versions index
            original_count = len(self.versions_index["versions"])
            self.versions_index["versions"] = [
                v for v in self.versions_index["versions"]
                if v["version"] != version_name
            ]

            if len(self.versions_index["versions"]) == original_count:
                self.logger.warning(f"Version {version_name} not found in index")
                return False

            # Remove individual metadata file
            version_metadata_file = self.metadata_dir / f"{version_name}.json"
            if version_metadata_file.exists():
                version_metadata_file.unlink()

            # Save updated index
            self._save_versions_index()

            # Update project statistics
            self._recalculate_project_stats()

            self.logger.info(f"Removed metadata for version {version_name}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to remove version metadata for {version_name}: {str(e)}")
            return False

    def get_project_statistics(self) -> Dict[str, Any]:
        """
        Get comprehensive project statistics

        Returns:
            Dictionary with project statistics
        """
        try:
            stats = self.project_info.copy()

            # Real-time calculations
            versions = self.get_all_versions()
            stats["current_version_count"] = len(versions)

            if versions:
                # Calculate total size
                total_size = 0
                for version in versions:
                    try:
                        snapshot_path = Path(version["snapshot_file"])
                        if snapshot_path.exists():
                            total_size += snapshot_path.stat().st_size
                    except Exception:
                        continue

                stats["current_total_size_bytes"] = total_size
                stats["current_total_size_mb"] = round(total_size / (1024 * 1024), 2)

                # Latest version info
                latest_version = max(versions, key=lambda x: x["timestamp"])
                stats["latest_version"] = latest_version["version"]
                stats["latest_version_date"] = latest_version["datetime"]

                # Oldest version info
                oldest_version = min(versions, key=lambda x: x["timestamp"])
                stats["oldest_version"] = oldest_version["version"]
                stats["oldest_version_date"] = oldest_version["datetime"]

                # Average file size
                stats["average_file_size_mb"] = round(
                    stats["current_total_size_mb"] / len(versions), 2
                )

            return stats

        except Exception as e:
            self.logger.error(f"Failed to get project statistics: {str(e)}")
            return self.project_info

    def search_versions(self, search_criteria: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Search versions based on criteria

        Args:
            search_criteria: Dictionary with search parameters
                - date_from: Start date for search
                - date_to: End date for search
                - notes_contain: Text that must be in notes
                - min_size: Minimum file size
                - max_size: Maximum file size

        Returns:
            List of matching versions
        """
        try:
            versions = self.get_all_versions()
            matching_versions = []

            for version in versions:
                matches = True

                # Date range filter
                if search_criteria.get("date_from"):
                    version_date = datetime.fromisoformat(version["datetime"].replace('Z', '+00:00'))
                    if version_date < search_criteria["date_from"]:
                        matches = False

                if search_criteria.get("date_to"):
                    version_date = datetime.fromisoformat(version["datetime"].replace('Z', '+00:00'))
                    if version_date > search_criteria["date_to"]:
                        matches = False

                # Notes filter
                if search_criteria.get("notes_contain"):
                    if search_criteria["notes_contain"].lower() not in version.get("notes", "").lower():
                        matches = False

                # Size filters
                if search_criteria.get("min_size"):
                    if version["file_size"] < search_criteria["min_size"]:
                        matches = False

                if search_criteria.get("max_size"):
                    if version["file_size"] > search_criteria["max_size"]:
                        matches = False

                if matches:
                    matching_versions.append(version)

            return matching_versions

        except Exception as e:
            self.logger.error(f"Failed to search versions: {str(e)}")
            return []

    def export_project_report(self, output_path: str) -> bool:
        """
        Export comprehensive project report

        Args:
            output_path: Path where to save the report

        Returns:
            True if successful, False otherwise
        """
        try:
            import pandas as pd

            # Get all version data
            versions = []
            for version_entry in self.get_all_versions():
                version_data = self.get_version_info(version_entry["version"])
                if version_data:
                    # Flatten metrics for the report
                    flat_version = {
                        "Version": version_data["version"],
                        "Date": version_data["datetime"],
                        "File Size (MB)": round(version_data["file_size"] / (1024 * 1024), 2),
                        "Notes": version_data.get("notes", ""),
                        "Quick Save": version_data.get("quick_save", False)
                    }

                    # Add metrics if available
                    metrics = version_data.get("metrics", {})
                    for metric_name, metric_value in metrics.items():
                        if not metric_name.startswith("_") and isinstance(metric_value, (int, float)):
                            flat_version[f"Metric_{metric_name}"] = metric_value

                    versions.append(flat_version)

            # Create Excel report
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Versions summary
                if versions:
                    versions_df = pd.DataFrame(versions)
                    versions_df.to_excel(writer, sheet_name='Versions', index=False)

                # Project statistics
                stats = self.get_project_statistics()
                stats_data = [{"Metric": k, "Value": v} for k, v in stats.items()
                             if not isinstance(v, dict)]
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='Project Stats', index=False)

            self.logger.info(f"Exported project report to {output_path}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to export project report: {str(e)}")
            return False

    def backup_metadata(self, backup_path: str) -> bool:
        """
        Create backup of all metadata

        Args:
            backup_path: Path where to save the backup

        Returns:
            True if successful, False otherwise
        """
        try:
            backup_dir = Path(backup_path)
            backup_dir.mkdir(parents=True, exist_ok=True)

            # Copy metadata directory
            metadata_backup = backup_dir / "metadata"
            if metadata_backup.exists():
                shutil.rmtree(metadata_backup)

            shutil.copytree(self.metadata_dir, metadata_backup)

            self.logger.info(f"Backed up metadata to {backup_path}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to backup metadata: {str(e)}")
            return False

    def restore_metadata(self, backup_path: str) -> bool:
        """
        Restore metadata from backup

        Args:
            backup_path: Path to the backup

        Returns:
            True if successful, False otherwise
        """
        try:
            backup_metadata = Path(backup_path) / "metadata"
            if not backup_metadata.exists():
                raise FileNotFoundError(f"Backup metadata not found at {backup_metadata}")

            # Remove current metadata
            if self.metadata_dir.exists():
                shutil.rmtree(self.metadata_dir)

            # Restore from backup
            shutil.copytree(backup_metadata, self.metadata_dir)

            # Reload metadata
            self._load_metadata()

            self.logger.info(f"Restored metadata from {backup_path}")
            return True

        except Exception as e:
            self.logger.error(f"Failed to restore metadata: {str(e)}")
            return False

    def _save_versions_index(self):
        """Save the versions index to file"""
        try:
            with open(self.versions_index_file, 'w') as f:
                json.dump(self.versions_index, f, indent=2, default=str)
        except Exception as e:
            self.logger.error(f"Failed to save versions index: {str(e)}")

    def _update_project_info(self, version_info: Dict[str, Any]):
        """Update project information with new version data"""
        try:
            self.project_info["last_accessed"] = datetime.now().isoformat()
            self.project_info["total_versions"] = len(self.versions_index["versions"])

            # Save updated project info
            with open(self.project_info_file, 'w') as f:
                json.dump(self.project_info, f, indent=2, default=str)

        except Exception as e:
            self.logger.error(f"Failed to update project info: {str(e)}")

    def _recalculate_project_stats(self):
        """Recalculate project statistics after version removal"""
        try:
            versions = self.get_all_versions()

            # Recalculate total size
            total_size = 0
            for version in versions:
                try:
                    snapshot_path = Path(version["snapshot_file"])
                    if snapshot_path.exists():
                        total_size += snapshot_path.stat().st_size
                except Exception:
                    continue

            self.project_info["total_size_bytes"] = total_size
            self.project_info["total_versions"] = len(versions)
            self.project_info["last_accessed"] = datetime.now().isoformat()

            # Save updated info
            with open(self.project_info_file, 'w') as f:
                json.dump(self.project_info, f, indent=2, default=str)

        except Exception as e:
            self.logger.error(f"Failed to recalculate project stats: {str(e)}")

    def cleanup_orphaned_files(self) -> Dict[str, int]:
        """
        Clean up orphaned snapshot files and metadata

        Returns:
            Dictionary with cleanup statistics
        """
        try:
            cleanup_stats = {
                "orphaned_snapshots_removed": 0,
                "orphaned_metadata_removed": 0,
                "missing_snapshots_found": 0
            }

            # Get current versions from index
            indexed_versions = {v["version"] for v in self.get_all_versions()}
            indexed_snapshots = {Path(v["snapshot_file"]).name for v in self.get_all_versions()}

            # Check for orphaned snapshot files
            if self.snapshots_dir.exists():
                for snapshot_file in self.snapshots_dir.glob("*.xlsx"):
                    if snapshot_file.name not in indexed_snapshots:
                        snapshot_file.unlink()
                        cleanup_stats["orphaned_snapshots_removed"] += 1
                        self.logger.info(f"Removed orphaned snapshot: {snapshot_file.name}")

            # Check for orphaned metadata files
            for metadata_file in self.metadata_dir.glob("v*.json"):
                version_name = metadata_file.stem
                if version_name not in indexed_versions:
                    metadata_file.unlink()
                    cleanup_stats["orphaned_metadata_removed"] += 1
                    self.logger.info(f"Removed orphaned metadata: {metadata_file.name}")

            # Check for missing snapshot files
            for version in self.get_all_versions():
                snapshot_path = Path(version["snapshot_file"])
                if not snapshot_path.exists():
                    cleanup_stats["missing_snapshots_found"] += 1
                    self.logger.warning(f"Missing snapshot file: {snapshot_path}")

            return cleanup_stats

        except Exception as e:
            self.logger.error(f"Failed to cleanup orphaned files: {str(e)}")
            return {"error": str(e)}