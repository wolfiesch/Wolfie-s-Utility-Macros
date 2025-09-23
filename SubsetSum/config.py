#!/usr/bin/env python3
"""
Configuration Management: Handle configuration files and settings for FuzzySum.

This module provides configuration management with YAML/JSON support,
default settings, and user preference handling.
"""

import os
import sys
import json
from pathlib import Path
from typing import Dict, Any, Optional, Union, List
from dataclasses import dataclass, asdict, field
import logging

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False

from fuzzy_sum import SolveMode, SolverStrategy


@dataclass
class SolverDefaults:
    """Default solver configuration."""
    mode: str = "closest"
    strategy: str = "ortools"
    scale_factor: int = 100
    time_limit: float = 5.0
    min_items: Optional[int] = None
    max_items: Optional[int] = None
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    tolerance: Optional[float] = None


@dataclass
class OutputDefaults:
    """Default output configuration."""
    level: str = "summary"
    format: str = "terminal"
    width: Optional[int] = None
    show_progress: bool = True
    show_timestamps: bool = False
    color_scheme: str = "default"


@dataclass
class ParallelDefaults:
    """Default parallel processing configuration."""
    max_workers: Optional[int] = None
    use_processes: bool = False
    chunk_size: int = 10
    timeout: Optional[float] = None
    enable_caching: bool = True
    cache_dir: Optional[str] = None


@dataclass
class ExcelDefaults:
    """Default Excel integration configuration."""
    default_sheet: Optional[str] = None
    start_address: str = "A1"
    highlight_selected: bool = True
    include_summary: bool = True
    auto_size_columns: bool = True


@dataclass
class FuzzySumConfig:
    """Main configuration class for FuzzySum."""
    solver: SolverDefaults = field(default_factory=SolverDefaults)
    output: OutputDefaults = field(default_factory=OutputDefaults)
    parallel: ParallelDefaults = field(default_factory=ParallelDefaults)
    excel: ExcelDefaults = field(default_factory=ExcelDefaults)

    # Global settings
    log_level: str = "INFO"
    debug: bool = False
    profiles: Dict[str, Dict[str, Any]] = field(default_factory=dict)


class ConfigManager:
    """Manage configuration files and settings."""

    DEFAULT_CONFIG_PATHS = [
        Path.home() / ".fuzzysum.yaml",
        Path.home() / ".fuzzysum.json",
        Path.home() / ".config" / "fuzzysum" / "config.yaml",
        Path.home() / ".config" / "fuzzysum" / "config.json",
        Path.cwd() / ".fuzzysum.yaml",
        Path.cwd() / ".fuzzysum.json",
    ]

    def __init__(self, config_path: Optional[Path] = None):
        """
        Initialize configuration manager.

        Args:
            config_path: Optional path to specific config file
        """
        self.logger = logging.getLogger(__name__)
        self.config_path = config_path
        self.config = FuzzySumConfig()

        # Load configuration
        self._load_config()

    def _load_config(self) -> None:
        """Load configuration from file or create default."""
        config_file = None

        if self.config_path:
            # Use specified config file
            if self.config_path.exists():
                config_file = self.config_path
            else:
                self.logger.warning(f"Specified config file not found: {self.config_path}")
        else:
            # Search for config files in default locations
            for path in self.DEFAULT_CONFIG_PATHS:
                if path.exists():
                    config_file = path
                    break

        if config_file:
            try:
                self.config = self._load_config_file(config_file)
                self.config_path = config_file
                self.logger.info(f"Loaded configuration from {config_file}")
            except Exception as e:
                self.logger.error(f"Failed to load config from {config_file}: {e}")
                self.logger.info("Using default configuration")
        else:
            self.logger.info("No configuration file found, using defaults")

    def _load_config_file(self, config_path: Path) -> FuzzySumConfig:
        """Load configuration from a specific file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            if config_path.suffix.lower() in ['.yaml', '.yml']:
                if not HAS_YAML:
                    raise ImportError("PyYAML required for YAML config files")
                data = yaml.safe_load(f)
            elif config_path.suffix.lower() == '.json':
                data = json.load(f)
            else:
                # Try to detect format by content
                content = f.read()
                f.seek(0)
                try:
                    if HAS_YAML:
                        data = yaml.safe_load(f)
                    else:
                        data = json.loads(content)
                except Exception:
                    data = json.loads(content)

        return self._dict_to_config(data)

    def _dict_to_config(self, data: Dict[str, Any]) -> FuzzySumConfig:
        """Convert dictionary to FuzzySumConfig object."""
        config = FuzzySumConfig()

        # Update solver defaults
        if 'solver' in data:
            solver_data = data['solver']
            config.solver = SolverDefaults(
                mode=solver_data.get('mode', config.solver.mode),
                strategy=solver_data.get('strategy', config.solver.strategy),
                scale_factor=solver_data.get('scale_factor', config.solver.scale_factor),
                time_limit=solver_data.get('time_limit', config.solver.time_limit),
                min_items=solver_data.get('min_items', config.solver.min_items),
                max_items=solver_data.get('max_items', config.solver.max_items),
                min_value=solver_data.get('min_value', config.solver.min_value),
                max_value=solver_data.get('max_value', config.solver.max_value),
                tolerance=solver_data.get('tolerance', config.solver.tolerance),
            )

        # Update output defaults
        if 'output' in data:
            output_data = data['output']
            config.output = OutputDefaults(
                level=output_data.get('level', config.output.level),
                format=output_data.get('format', config.output.format),
                width=output_data.get('width', config.output.width),
                show_progress=output_data.get('show_progress', config.output.show_progress),
                show_timestamps=output_data.get('show_timestamps', config.output.show_timestamps),
                color_scheme=output_data.get('color_scheme', config.output.color_scheme),
            )

        # Update parallel defaults
        if 'parallel' in data:
            parallel_data = data['parallel']
            config.parallel = ParallelDefaults(
                max_workers=parallel_data.get('max_workers', config.parallel.max_workers),
                use_processes=parallel_data.get('use_processes', config.parallel.use_processes),
                chunk_size=parallel_data.get('chunk_size', config.parallel.chunk_size),
                timeout=parallel_data.get('timeout', config.parallel.timeout),
                enable_caching=parallel_data.get('enable_caching', config.parallel.enable_caching),
                cache_dir=parallel_data.get('cache_dir', config.parallel.cache_dir),
            )

        # Update Excel defaults
        if 'excel' in data:
            excel_data = data['excel']
            config.excel = ExcelDefaults(
                default_sheet=excel_data.get('default_sheet', config.excel.default_sheet),
                start_address=excel_data.get('start_address', config.excel.start_address),
                highlight_selected=excel_data.get('highlight_selected', config.excel.highlight_selected),
                include_summary=excel_data.get('include_summary', config.excel.include_summary),
                auto_size_columns=excel_data.get('auto_size_columns', config.excel.auto_size_columns),
            )

        # Update global settings
        config.log_level = data.get('log_level', config.log_level)
        config.debug = data.get('debug', config.debug)
        config.profiles = data.get('profiles', config.profiles)

        return config

    def save_config(self, config_path: Optional[Path] = None) -> None:
        """
        Save current configuration to file.

        Args:
            config_path: Optional path to save config (uses current path if None)
        """
        save_path = config_path or self.config_path

        if save_path is None:
            # Use default location
            save_path = Path.home() / ".fuzzysum.yaml"

        # Ensure directory exists
        save_path.parent.mkdir(parents=True, exist_ok=True)

        # Convert config to dictionary
        config_dict = self._config_to_dict()

        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                if save_path.suffix.lower() in ['.yaml', '.yml']:
                    if not HAS_YAML:
                        raise ImportError("PyYAML required for YAML config files")
                    yaml.dump(config_dict, f, default_flow_style=False, sort_keys=False)
                else:
                    json.dump(config_dict, f, indent=2, sort_keys=False)

            self.config_path = save_path
            self.logger.info(f"Configuration saved to {save_path}")

        except Exception as e:
            self.logger.error(f"Failed to save config to {save_path}: {e}")
            raise

    def _config_to_dict(self) -> Dict[str, Any]:
        """Convert FuzzySumConfig to dictionary."""
        return {
            'solver': asdict(self.config.solver),
            'output': asdict(self.config.output),
            'parallel': asdict(self.config.parallel),
            'excel': asdict(self.config.excel),
            'log_level': self.config.log_level,
            'debug': self.config.debug,
            'profiles': self.config.profiles,
        }

    def create_profile(self, name: str, **kwargs) -> None:
        """
        Create a named configuration profile.

        Args:
            name: Profile name
            **kwargs: Configuration overrides
        """
        self.config.profiles[name] = kwargs
        self.logger.info(f"Created profile: {name}")

    def apply_profile(self, name: str) -> None:
        """
        Apply a named configuration profile.

        Args:
            name: Profile name to apply
        """
        if name not in self.config.profiles:
            raise ValueError(f"Profile '{name}' not found")

        profile_data = self.config.profiles[name]

        # Apply profile settings to current config
        for section, settings in profile_data.items():
            if section == 'solver' and hasattr(self.config, 'solver'):
                for key, value in settings.items():
                    if hasattr(self.config.solver, key):
                        setattr(self.config.solver, key, value)
            elif section == 'output' and hasattr(self.config, 'output'):
                for key, value in settings.items():
                    if hasattr(self.config.output, key):
                        setattr(self.config.output, key, value)
            elif section == 'parallel' and hasattr(self.config, 'parallel'):
                for key, value in settings.items():
                    if hasattr(self.config.parallel, key):
                        setattr(self.config.parallel, key, value)
            elif section == 'excel' and hasattr(self.config, 'excel'):
                for key, value in settings.items():
                    if hasattr(self.config.excel, key):
                        setattr(self.config.excel, key, value)
            elif hasattr(self.config, section):
                setattr(self.config, section, settings)

        self.logger.info(f"Applied profile: {name}")

    def list_profiles(self) -> List[str]:
        """Get list of available profiles."""
        return list(self.config.profiles.keys())

    def delete_profile(self, name: str) -> None:
        """
        Delete a named profile.

        Args:
            name: Profile name to delete
        """
        if name not in self.config.profiles:
            raise ValueError(f"Profile '{name}' not found")

        del self.config.profiles[name]
        self.logger.info(f"Deleted profile: {name}")

    def get_solver_config(self) -> Dict[str, Any]:
        """Get solver configuration as dictionary."""
        return asdict(self.config.solver)

    def get_output_config(self) -> Dict[str, Any]:
        """Get output configuration as dictionary."""
        return asdict(self.config.output)

    def get_parallel_config(self) -> Dict[str, Any]:
        """Get parallel processing configuration as dictionary."""
        return asdict(self.config.parallel)

    def get_excel_config(self) -> Dict[str, Any]:
        """Get Excel configuration as dictionary."""
        return asdict(self.config.excel)

    def update_solver_config(self, **kwargs) -> None:
        """Update solver configuration."""
        for key, value in kwargs.items():
            if hasattr(self.config.solver, key):
                setattr(self.config.solver, key, value)

    def update_output_config(self, **kwargs) -> None:
        """Update output configuration."""
        for key, value in kwargs.items():
            if hasattr(self.config.output, key):
                setattr(self.config.output, key, value)

    def reset_to_defaults(self) -> None:
        """Reset configuration to defaults."""
        self.config = FuzzySumConfig()
        self.logger.info("Configuration reset to defaults")

    def validate_config(self) -> List[str]:
        """
        Validate current configuration.

        Returns:
            List of validation errors (empty if valid)
        """
        errors = []

        # Validate solver settings
        try:
            SolveMode(self.config.solver.mode)
        except ValueError:
            errors.append(f"Invalid solver mode: {self.config.solver.mode}")

        try:
            SolverStrategy(self.config.solver.strategy)
        except ValueError:
            errors.append(f"Invalid solver strategy: {self.config.solver.strategy}")

        if self.config.solver.scale_factor <= 0:
            errors.append("Scale factor must be positive")

        if self.config.solver.time_limit <= 0:
            errors.append("Time limit must be positive")

        # Validate output settings
        valid_levels = ["minimal", "summary", "detailed", "verbose", "json"]
        if self.config.output.level not in valid_levels:
            errors.append(f"Invalid output level: {self.config.output.level}")

        valid_formats = ["terminal", "plain", "json", "csv", "markdown", "html"]
        if self.config.output.format not in valid_formats:
            errors.append(f"Invalid output format: {self.config.output.format}")

        # Validate parallel settings
        if self.config.parallel.max_workers is not None and self.config.parallel.max_workers <= 0:
            errors.append("Max workers must be positive")

        if self.config.parallel.chunk_size <= 0:
            errors.append("Chunk size must be positive")

        # Validate Excel settings
        if self.config.excel.start_address:
            try:
                from fuzzy_sum import A1Converter
                A1Converter.parse_a1(self.config.excel.start_address)
            except Exception:
                errors.append(f"Invalid Excel start address: {self.config.excel.start_address}")

        return errors


def load_config(config_path: Optional[Path] = None) -> ConfigManager:
    """
    Load configuration from file or create default.

    Args:
        config_path: Optional path to specific config file

    Returns:
        ConfigManager instance
    """
    return ConfigManager(config_path)


def create_default_config(output_path: Path, format: str = "yaml") -> None:
    """
    Create a default configuration file.

    Args:
        output_path: Path where to save the config file
        format: Config format ('yaml' or 'json')
    """
    config_manager = ConfigManager()

    # Set appropriate file extension
    if format.lower() == "yaml" and not output_path.suffix:
        output_path = output_path.with_suffix(".yaml")
    elif format.lower() == "json" and not output_path.suffix:
        output_path = output_path.with_suffix(".json")

    config_manager.save_config(output_path)


def get_config_template() -> str:
    """Get a YAML template for configuration file."""
    template = """# FuzzySum Configuration File
# This file contains default settings for the FuzzySum subset sum solver

# Solver defaults
solver:
  mode: "closest"           # closest, le, ge, exact, range
  strategy: "ortools"       # ortools, dynamic, greedy, hybrid
  scale_factor: 100         # Scale factor for float-to-int conversion
  time_limit: 5.0          # Maximum solve time in seconds
  min_items: null          # Minimum items in solution
  max_items: null          # Maximum items in solution
  min_value: null          # Minimum value of items to include
  max_value: null          # Maximum value of items to include
  tolerance: null          # Tolerance for range mode

# Output defaults
output:
  level: "summary"         # minimal, summary, detailed, verbose, json
  format: "terminal"       # terminal, plain, json, csv, markdown, html
  width: null              # Terminal width (null for auto-detect)
  show_progress: true      # Show progress indicators
  show_timestamps: false   # Include timestamps in output
  color_scheme: "default"  # Color scheme for terminal output

# Parallel processing defaults
parallel:
  max_workers: null        # Maximum parallel workers (null for auto)
  use_processes: false     # Use processes instead of threads
  chunk_size: 10          # Batch size for parallel processing
  timeout: null           # Timeout for parallel operations
  enable_caching: true    # Enable result caching
  cache_dir: null         # Cache directory (null for temp)

# Excel integration defaults
excel:
  default_sheet: null     # Default sheet name (null for first sheet)
  start_address: "A1"     # Default starting address
  highlight_selected: true # Highlight selected cells in output
  include_summary: true   # Include summary table in Excel output
  auto_size_columns: true # Auto-size columns in Excel output

# Global settings
log_level: "INFO"         # DEBUG, INFO, WARNING, ERROR
debug: false             # Enable debug mode

# Named profiles for common configurations
profiles:
  fast:
    solver:
      strategy: "greedy"
      time_limit: 1.0
    output:
      level: "minimal"

  accurate:
    solver:
      strategy: "ortools"
      time_limit: 30.0
    output:
      level: "detailed"

  batch:
    parallel:
      max_workers: 8
      enable_caching: true
    output:
      show_progress: true
"""
    return template