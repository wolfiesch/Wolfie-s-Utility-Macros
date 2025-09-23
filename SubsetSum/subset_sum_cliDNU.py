#!/usr/bin/env python3
"""
FuzzySum CLI: Enhanced subset sum solver with rich terminal output.

A powerful command-line tool for solving subset sum problems with multiple strategies,
Excel integration, and beautiful terminal output.

Usage:
    python subset_sum_cli.py --csv data.csv --target 12345.67 --addr A5
    python subset_sum_cli.py --excel report.xlsx --target 50000 --verbose
    python subset_sum_cli.py --batch targets.txt --data values.csv --export results.json
"""

import sys
import argparse
import csv
import time
import logging
from pathlib import Path
from typing import List, Optional, Dict, Any

try:
    import click
    HAS_CLICK = True
except ImportError:
    HAS_CLICK = False

from fuzzy_sum import (
    solve_subset_sum, SolverResult, SolveMode, SolverStrategy,
    A1Converter, SolverConstraints
)
from cli_formatter import create_formatter, OutputLevel, OutputFormat


def setup_logging(verbose: bool = False, debug: bool = False) -> None:
    """Set up logging configuration."""
    if debug:
        level = logging.DEBUG
    elif verbose:
        level = logging.INFO
    else:
        level = logging.WARNING

    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stderr)
        ]
    )


def read_csv_data(csv_path: Path) -> List[float]:
    """
    Read numeric data from CSV file.

    Args:
        csv_path: Path to CSV file

    Returns:
        List of float values, ignoring blanks and non-numeric entries
    """
    values = []

    try:
        with csv_path.open(newline='', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for row_num, row in enumerate(reader, 1):
                if not row:
                    continue

                cell = row[0].strip()
                if cell == "":
                    continue

                # Try to parse as float, handling common formats
                try:
                    # Remove commas and parse
                    clean_value = cell.replace(',', '').replace('$', '').replace('%', '')
                    values.append(float(clean_value))
                except ValueError:
                    logging.warning(f"Skipping non-numeric value '{cell}' at row {row_num}")
                    continue

    except FileNotFoundError:
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    except Exception as e:
        raise RuntimeError(f"Error reading CSV file {csv_path}: {e}")

    return values


def read_targets_file(targets_path: Path) -> List[float]:
    """
    Read target values from file.

    Args:
        targets_path: Path to file containing target values (one per line)

    Returns:
        List of target values
    """
    targets = []

    try:
        with targets_path.open('r') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue

                try:
                    targets.append(float(line.replace(',', '')))
                except ValueError:
                    logging.warning(f"Skipping invalid target '{line}' at line {line_num}")

    except FileNotFoundError:
        raise FileNotFoundError(f"Targets file not found: {targets_path}")
    except Exception as e:
        raise RuntimeError(f"Error reading targets file {targets_path}: {e}")

    return targets


def validate_a1_address(address: str) -> str:
    """
    Validate and normalize A1 address.

    Args:
        address: A1 address string (e.g., "A5", "B10")

    Returns:
        Normalized A1 address

    Raises:
        ValueError: If address is invalid
    """
    try:
        col_idx, row_num = A1Converter.parse_a1(address)
        return A1Converter.to_a1(col_idx, row_num)
    except Exception:
        raise ValueError(f"Invalid A1 address: {address}")


def build_constraints(args: argparse.Namespace) -> SolverConstraints:
    """Build solver constraints from command line arguments."""
    constraints = SolverConstraints()

    if hasattr(args, 'min_items') and args.min_items is not None:
        constraints.min_items = args.min_items

    if hasattr(args, 'max_items') and args.max_items is not None:
        constraints.max_items = args.max_items

    if hasattr(args, 'min_value') and args.min_value is not None:
        constraints.min_value = args.min_value

    if hasattr(args, 'max_value') and args.max_value is not None:
        constraints.max_value = args.max_value

    if hasattr(args, 'tolerance') and args.tolerance is not None:
        constraints.tolerance = args.tolerance

    if hasattr(args, 'required') and args.required:
        constraints.required_indices = set(args.required)

    if hasattr(args, 'forbidden') and args.forbidden:
        constraints.forbidden_indices = set(args.forbidden)

    return constraints


def solve_single_target(
    values: List[float],
    target: float,
    args: argparse.Namespace,
    formatter
) -> SolverResult:
    """Solve for a single target value."""
    constraints = build_constraints(args)

    formatter.print_progress_start(f"Solving for target {target:,.2f}")

    # Solve the problem
    result = solve_subset_sum(
        values=values,
        target=target,
        mode=args.mode,
        strategy=args.strategy,
        scale_factor=args.scale,
        time_limit=args.time,
        min_items=constraints.min_items,
        max_items=constraints.max_items,
        min_value=constraints.min_value,
        max_value=constraints.max_value,
        required_indices=constraints.required_indices,
        forbidden_indices=constraints.forbidden_indices,
        tolerance=constraints.tolerance
    )

    return result


def solve_batch_targets(
    values: List[float],
    targets: List[float],
    args: argparse.Namespace,
    formatter
) -> List[SolverResult]:
    """Solve for multiple target values."""
    results = []

    formatter.print_progress_start(f"Solving batch of {len(targets)} targets")

    with formatter.create_progress_bar("Processing targets", len(targets)) as progress:
        task = progress.add_task("Solving...", total=len(targets))

        for i, target in enumerate(targets):
            result = solve_single_target(values, target, args, formatter)
            results.append(result)

            progress.update(task, advance=1)

            if args.verbose:
                status = "✓" if result.success else "✗"
                formatter.print_info(f"{status} Target {target:,.2f}: {result.solver_status}")

    return results


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="FuzzySum: Enhanced subset sum solver with rich output",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic CSV input
  python subset_sum_cli.py --csv data.csv --target 12345.67 --addr A5

  # Excel input with constraints
  python subset_sum_cli.py --excel report.xlsx --target 50000 --max-items 10

  # Batch processing
  python subset_sum_cli.py --batch targets.txt --csv values.csv --export results.json

  # Verbose output with specific strategy
  python subset_sum_cli.py --csv data.csv --target 10000 --strategy dynamic --verbose

  # Range mode with tolerance
  python subset_sum_cli.py --csv data.csv --target 25000 --mode range --tolerance 100
        """)

    # Input sources (mutually exclusive)
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        '--csv', type=Path,
        help='Path to CSV file with numeric data (one column)'
    )
    input_group.add_argument(
        '--excel', type=Path,
        help='Path to Excel file (.xlsx, .xls)'
    )
    input_group.add_argument(
        '--data', type=Path,
        help='Path to data file (CSV or Excel)'
    )

    # Target specification (mutually exclusive)
    target_group = parser.add_mutually_exclusive_group(required=True)
    target_group.add_argument(
        '--target', type=float,
        help='Target sum to achieve'
    )
    target_group.add_argument(
        '--batch', type=Path,
        help='File with multiple target values (one per line)'
    )

    # Excel-specific options
    parser.add_argument(
        '--sheet', type=str, default=None,
        help='Excel sheet name (default: first sheet)'
    )
    parser.add_argument(
        '--addr', type=str, default='A1',
        help='Starting A1 address for Excel range (default: A1)'
    )

    # Solver options
    parser.add_argument(
        '--mode', choices=['closest', 'le', 'ge', 'exact', 'range'],
        default='closest',
        help='Solving mode (default: closest)'
    )
    parser.add_argument(
        '--strategy', choices=['ortools', 'dynamic', 'greedy', 'hybrid'],
        default='ortools',
        help='Solver strategy (default: ortools)'
    )
    parser.add_argument(
        '--scale', type=int, default=100,
        help='Scale factor for float-to-int conversion (default: 100)'
    )
    parser.add_argument(
        '--time', type=float, default=5.0,
        help='Maximum solve time in seconds (default: 5.0)'
    )

    # Constraints
    parser.add_argument(
        '--min-items', type=int,
        help='Minimum number of items in solution'
    )
    parser.add_argument(
        '--max-items', type=int,
        help='Maximum number of items in solution'
    )
    parser.add_argument(
        '--min-value', type=float,
        help='Minimum value of items to include'
    )
    parser.add_argument(
        '--max-value', type=float,
        help='Maximum value of items to include'
    )
    parser.add_argument(
        '--tolerance', type=float,
        help='Tolerance for range mode'
    )
    parser.add_argument(
        '--required', type=int, nargs='*',
        help='Required item indices (0-based)'
    )
    parser.add_argument(
        '--forbidden', type=int, nargs='*',
        help='Forbidden item indices (0-based)'
    )

    # Output options
    parser.add_argument(
        '--output', choices=['minimal', 'summary', 'detailed', 'verbose'],
        default='summary',
        help='Output verbosity level (default: summary)'
    )
    parser.add_argument(
        '--format', choices=['terminal', 'plain', 'json'],
        default='terminal',
        help='Output format (default: terminal)'
    )
    parser.add_argument(
        '--export', type=Path,
        help='Export results to file (.json, .csv, .md, .html)'
    )
    parser.add_argument(
        '--width', type=int,
        help='Terminal width for output formatting'
    )

    # Flags
    parser.add_argument(
        '--no-progress', action='store_true',
        help='Disable progress indicators'
    )
    parser.add_argument(
        '--verbose', '-v', action='store_true',
        help='Enable verbose logging'
    )
    parser.add_argument(
        '--debug', action='store_true',
        help='Enable debug logging'
    )
    parser.add_argument(
        '--quiet', '-q', action='store_true',
        help='Suppress all output except results'
    )

    args = parser.parse_args()

    # Set up logging
    setup_logging(args.verbose or args.debug, args.debug)

    # Adjust output level based on flags
    if args.quiet:
        output_level = 'minimal'
    elif args.verbose:
        output_level = 'verbose'
    else:
        output_level = args.output

    # Create formatter
    try:
        formatter = create_formatter(
            level=output_level,
            format_type=args.format,
            width=args.width,
            show_progress=not args.no_progress,
            export_path=str(args.export) if args.export else None
        )
    except Exception as e:
        print(f"Error creating formatter: {e}", file=sys.stderr)
        sys.exit(1)

    # Print header (unless quiet mode)
    if not args.quiet:
        formatter.print_header(
            "FuzzySum - Enhanced Subset Sum Solver",
            "Find optimal subsets that sum closest to target values"
        )

    try:
        # Validate A1 address if provided
        start_address = None
        if args.addr:
            try:
                start_address = validate_a1_address(args.addr)
            except ValueError as e:
                formatter.print_error(str(e))
                sys.exit(1)

        # Read input data
        if args.csv or (args.data and str(args.data).endswith('.csv')):
            data_path = args.csv or args.data
            if not args.quiet:
                formatter.print_progress_start(f"Reading CSV data from {data_path}")
            values = read_csv_data(data_path)

        elif args.excel or (args.data and str(args.data).endswith(('.xlsx', '.xls', '.xlsm'))):
            # Import Excel integration
            try:
                from excel_integration import read_excel_data
                data_path = args.excel or args.data
                if not args.quiet:
                    formatter.print_progress_start(f"Reading Excel data from {data_path}")
                values = read_excel_data(data_path, args.sheet, start_address)
            except ImportError:
                formatter.print_error("Excel integration not available. Install openpyxl: pip install openpyxl")
                sys.exit(1)
            except Exception as e:
                formatter.print_error(f"Error reading Excel file: {e}")
                sys.exit(1)

        else:
            formatter.print_error("Unsupported data file format")
            sys.exit(1)

        if not values:
            formatter.print_error("No valid numeric values found in input file")
            sys.exit(1)

        if not args.quiet:
            formatter.print_success(f"Loaded {len(values)} values from {data_path}")

        # Solve the problem(s)
        start_time = time.time()

        if args.batch:
            # Batch mode - multiple targets
            if not args.quiet:
                formatter.print_progress_start(f"Reading targets from {args.batch}")
            targets = read_targets_file(args.batch)

            if not targets:
                formatter.print_error("No valid targets found in batch file")
                sys.exit(1)

            results = solve_batch_targets(values, targets, args, formatter)

            # Print batch results summary
            if not args.quiet:
                total_time = time.time() - start_time
                successful = sum(1 for r in results if r.success)
                formatter.print_success(f"Batch complete: {successful}/{len(results)} successful in {total_time:.2f}s")

            # Print individual results
            for i, (target, result) in enumerate(zip(targets, results)):
                if not args.quiet and len(targets) > 1:
                    formatter.print_info(f"--- Result {i+1}/{len(targets)} ---")
                formatter.print_result_summary(result, start_address)
                if args.export:
                    formatter.export_result(result, start_address)

        else:
            # Single target mode
            result = solve_single_target(values, args.target, args, formatter)

            # Print results
            formatter.print_result_summary(result, start_address)

            # Export if requested
            if args.export:
                formatter.export_result(result, start_address)
                if not args.quiet:
                    formatter.print_success(f"Results exported to {args.export}")

            # Set exit code based on success
            if not result.success:
                sys.exit(1)

    except KeyboardInterrupt:
        formatter.print_warning("Operation cancelled by user")
        sys.exit(130)
    except Exception as e:
        formatter.print_error(f"Unexpected error: {e}")
        if args.debug:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()