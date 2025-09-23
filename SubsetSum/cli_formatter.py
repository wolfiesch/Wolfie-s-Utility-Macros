#!/usr/bin/env python3
"""
CLI Formatter: Rich terminal output and formatting for FuzzySum.

This module provides colorful, well-structured terminal output with tables,
progress bars, and detailed formatting for the FuzzySum subset sum solver.
"""

import sys
import time
from typing import List, Dict, Any, Optional, Union
from dataclasses import dataclass
from enum import Enum

try:
    from rich.console import Console
    from rich.table import Table
    from rich.progress import Progress, BarColumn, TextColumn, TimeElapsedColumn
    from rich.panel import Panel
    from rich.text import Text
    from rich.syntax import Syntax
    from rich.markdown import Markdown
    from rich.layout import Layout
    from rich.live import Live
    from rich.spinner import Spinner
    from rich.tree import Tree
    from rich.align import Align
    from rich.padding import Padding
    from rich import box
    HAS_RICH = True
except ImportError:
    HAS_RICH = False

try:
    from colorama import init, Fore, Back, Style
    init(autoreset=True)
    HAS_COLORAMA = True
except ImportError:
    HAS_COLORAMA = False

from fuzzy_sum import SolverResult, SolveMode, SolverStrategy


class OutputLevel(Enum):
    """Output verbosity levels."""
    MINIMAL = "minimal"      # Just the essentials
    SUMMARY = "summary"      # Key results with some formatting
    DETAILED = "detailed"    # Full results with rich formatting
    VERBOSE = "verbose"      # Everything including debug info
    JSON = "json"           # Machine-readable JSON output


class OutputFormat(Enum):
    """Output format types."""
    TERMINAL = "terminal"    # Rich terminal output
    PLAIN = "plain"         # Plain text without colors
    JSON = "json"           # JSON format
    CSV = "csv"             # CSV format
    MARKDOWN = "markdown"   # Markdown format
    HTML = "html"           # HTML format


@dataclass
class FormatConfig:
    """Configuration for output formatting."""
    level: OutputLevel = OutputLevel.SUMMARY
    format: OutputFormat = OutputFormat.TERMINAL
    width: Optional[int] = None
    show_progress: bool = True
    show_timestamps: bool = False
    color_scheme: str = "default"
    export_path: Optional[str] = None


class TerminalFormatter:
    """Rich terminal output formatter."""

    def __init__(self, config: FormatConfig = None):
        self.config = config or FormatConfig()

        if HAS_RICH and self.config.format == OutputFormat.TERMINAL:
            self.console = Console(
                width=self.config.width,
                force_terminal=True if sys.stdout.isatty() else False
            )
            self.use_rich = True
        else:
            self.console = None
            self.use_rich = False

        # Color schemes
        self.colors = self._init_colors()

    def _init_colors(self) -> Dict[str, str]:
        """Initialize color schemes."""
        if self.use_rich:
            return {
                'success': 'green',
                'error': 'red',
                'warning': 'yellow',
                'info': 'blue',
                'highlight': 'cyan',
                'muted': 'dim',
                'accent': 'magenta',
                'number': 'bright_green',
                'header': 'bold blue',
                'subheader': 'bold',
            }
        elif HAS_COLORAMA:
            return {
                'success': Fore.GREEN,
                'error': Fore.RED,
                'warning': Fore.YELLOW,
                'info': Fore.BLUE,
                'highlight': Fore.CYAN,
                'muted': Style.DIM,
                'accent': Fore.MAGENTA,
                'number': Fore.LIGHTGREEN_EX,
                'header': Fore.BLUE + Style.BRIGHT,
                'subheader': Style.BRIGHT,
                'reset': Style.RESET_ALL,
            }
        else:
            return {k: '' for k in ['success', 'error', 'warning', 'info',
                                   'highlight', 'muted', 'accent', 'number',
                                   'header', 'subheader', 'reset']}

    def print_header(self, title: str, subtitle: str = None):
        """Print a formatted header."""
        if self.use_rich:
            panel_content = f"[bold blue]{title}[/bold blue]"
            if subtitle:
                panel_content += f"\n[dim]{subtitle}[/dim]"

            panel = Panel(
                panel_content,
                box=box.DOUBLE,
                border_style="blue",
                padding=(1, 2)
            )
            self.console.print(panel)
            self.console.print()
        else:
            header_line = "=" * 60
            print(f"{self.colors.get('header', '')}{header_line}")
            print(f"{title}")
            if subtitle:
                print(f"{subtitle}")
            print(f"{header_line}{self.colors.get('reset', '')}")
            print()

    def print_result_summary(self, result: SolverResult, start_address: str = None):
        """Print a formatted summary of the solver result."""
        if self.config.level == OutputLevel.MINIMAL:
            self._print_minimal_result(result, start_address)
        elif self.config.level == OutputLevel.JSON:
            self._print_json_result(result, start_address)
        else:
            self._print_formatted_result(result, start_address)

    def _print_minimal_result(self, result: SolverResult, start_address: str = None):
        """Print minimal result output."""
        if start_address and result.success:
            from fuzzy_sum import A1Converter
            addresses = A1Converter.offset_addresses(start_address, result.selected_indices)
            print(f"PICKS: {','.join(addresses) if addresses else ''}")
        else:
            indices_str = ','.join(map(str, result.selected_indices)) if result.selected_indices else ''
            print(f"PICKS: {indices_str}")

        print(f"BEST_SUM: {result.total_sum:.10g}")
        print(f"ERROR: {result.error:.10g}")
        print(f"COUNT: {len(result.selected_indices)}")

    def _print_json_result(self, result: SolverResult, start_address: str = None):
        """Print result in JSON format."""
        import json

        result_dict = {
            'success': result.success,
            'selected_indices': result.selected_indices,
            'selected_values': result.selected_values,
            'total_sum': result.total_sum,
            'target': result.target,
            'error': result.error,
            'count': len(result.selected_indices),
            'solve_time': result.solve_time,
            'strategy': result.strategy_used.value,
            'mode': result.mode_used.value,
            'status': result.solver_status
        }

        if start_address and result.success:
            from fuzzy_sum import A1Converter
            addresses = A1Converter.offset_addresses(start_address, result.selected_indices)
            result_dict['cell_addresses'] = addresses

        print(json.dumps(result_dict, indent=2))

    def _print_formatted_result(self, result: SolverResult, start_address: str = None):
        """Print richly formatted result."""
        if self.use_rich:
            self._print_rich_result(result, start_address)
        else:
            self._print_plain_result(result, start_address)

    def _print_rich_result(self, result: SolverResult, start_address: str = None):
        """Print result using Rich formatting."""
        # Main result panel
        status_color = "green" if result.success else "red"
        status_text = "✓ SUCCESS" if result.success else "✗ FAILED"

        # Create main results table
        table = Table(
            title=f"[{status_color}]{status_text}[/{status_color}] Subset Sum Results",
            box=box.ROUNDED,
            border_style=status_color,
            show_header=True,
            header_style="bold"
        )

        table.add_column("Metric", style="cyan", min_width=15)
        table.add_column("Value", style="white", min_width=20)
        table.add_column("Details", style="dim", min_width=25)

        # Add main metrics
        table.add_row(
            "Target Sum",
            f"[bright_green]{result.target:,.2f}[/bright_green]",
            f"({result.target:.10g})"
        )

        if result.success:
            table.add_row(
                "Achieved Sum",
                f"[bright_green]{result.total_sum:,.2f}[/bright_green]",
                f"({result.total_sum:.10g})"
            )

            error_style = "green" if result.error < 0.01 else "yellow" if result.error < 1.0 else "red"
            table.add_row(
                "Error",
                f"[{error_style}]{result.error:,.4f}[/{error_style}]",
                f"{(result.error/result.target*100):.2f}% of target" if result.target != 0 else "N/A"
            )

            table.add_row(
                "Items Selected",
                f"[cyan]{len(result.selected_indices)}[/cyan]",
                f"out of {len(result.selected_values)} available" if result.selected_values else ""
            )
        else:
            table.add_row(
                "Error",
                f"[red]{result.error:,.2f}[/red]",
                result.message or "No solution found"
            )

        table.add_row(
            "Solve Time",
            f"[yellow]{result.solve_time:.3f}s[/yellow]",
            f"Strategy: {result.strategy_used.value}"
        )

        table.add_row(
            "Status",
            f"[{status_color}]{result.solver_status}[/{status_color}]",
            f"Mode: {result.mode_used.value}"
        )

        self.console.print(table)

        # Selected items details
        if result.success and result.selected_indices and self.config.level in [OutputLevel.DETAILED, OutputLevel.VERBOSE]:
            self.console.print()
            self._print_selected_items_table(result, start_address)

        # Performance details for verbose output
        if self.config.level == OutputLevel.VERBOSE:
            self.console.print()
            self._print_performance_details(result)

    def _print_selected_items_table(self, result: SolverResult, start_address: str = None):
        """Print table of selected items."""
        items_table = Table(
            title="Selected Items",
            box=box.SIMPLE,
            show_header=True,
            header_style="bold cyan"
        )

        items_table.add_column("#", style="dim", width=4, justify="right")
        if start_address:
            items_table.add_column("Cell", style="magenta", width=8)
        items_table.add_column("Index", style="yellow", width=6, justify="right")
        items_table.add_column("Value", style="green", width=15, justify="right")
        items_table.add_column("Cumulative", style="blue", width=15, justify="right")

        cumulative = 0.0
        addresses = None
        if start_address:
            from fuzzy_sum import A1Converter
            addresses = A1Converter.offset_addresses(start_address, result.selected_indices)

        for i, (idx, value) in enumerate(zip(result.selected_indices, result.selected_values)):
            cumulative += value
            row = [str(i + 1)]

            if addresses:
                row.append(addresses[i])

            row.extend([
                str(idx),
                f"{value:,.4f}",
                f"{cumulative:,.4f}"
            ])

            items_table.add_row(*row)

        self.console.print(items_table)

    def _print_performance_details(self, result: SolverResult):
        """Print performance and solver details."""
        perf_table = Table(
            title="Performance Details",
            box=box.SIMPLE,
            show_header=True,
            header_style="bold yellow"
        )

        perf_table.add_column("Metric", style="cyan")
        perf_table.add_column("Value", style="white")

        perf_table.add_row("Solver Strategy", result.strategy_used.value.title())
        perf_table.add_row("Solve Mode", result.mode_used.value.title())
        perf_table.add_row("Solve Time", f"{result.solve_time:.6f} seconds")
        perf_table.add_row("Solver Status", result.solver_status)

        if result.iterations > 0:
            perf_table.add_row("Iterations", str(result.iterations))

        if result.message:
            perf_table.add_row("Message", result.message)

        self.console.print(perf_table)

    def _print_plain_result(self, result: SolverResult, start_address: str = None):
        """Print result using plain text formatting."""
        status = "SUCCESS" if result.success else "FAILED"
        color = self.colors.get('success' if result.success else 'error', '')
        reset = self.colors.get('reset', '')

        print(f"\n{color}=== {status} - Subset Sum Results ==={reset}")
        print(f"Target Sum:     {result.target:,.2f}")

        if result.success:
            print(f"Achieved Sum:   {result.total_sum:,.2f}")
            print(f"Error:          {result.error:,.4f} ({(result.error/result.target*100):.2f}% of target)" if result.target != 0 else f"Error: {result.error:,.4f}")
            print(f"Items Selected: {len(result.selected_indices)}")

            if start_address and self.config.level != OutputLevel.SUMMARY:
                from fuzzy_sum import A1Converter
                addresses = A1Converter.offset_addresses(start_address, result.selected_indices)
                print(f"Cell Addresses: {', '.join(addresses) if addresses else 'None'}")
        else:
            print(f"Error:          {result.error:,.2f}")
            if result.message:
                print(f"Message:        {result.message}")

        print(f"Solve Time:     {result.solve_time:.3f}s")
        print(f"Strategy:       {result.strategy_used.value}")
        print(f"Status:         {result.solver_status}")

        # Selected items details
        if (result.success and result.selected_indices and
            self.config.level in [OutputLevel.DETAILED, OutputLevel.VERBOSE]):
            print(f"\n{self.colors.get('header', '')}Selected Items:{reset}")
            print(f"{'#':<4} {'Index':<6} {'Value':<15} {'Cumulative':<15}")
            print("-" * 45)

            cumulative = 0.0
            for i, (idx, value) in enumerate(zip(result.selected_indices, result.selected_values)):
                cumulative += value
                print(f"{i+1:<4} {idx:<6} {value:<15.4f} {cumulative:<15.4f}")

    def print_progress_start(self, message: str):
        """Start a progress indicator."""
        if self.use_rich and self.config.show_progress:
            self.console.print(f"[blue]⚡ {message}...[/blue]")
        else:
            print(f"⚡ {message}...")

    def print_progress_update(self, message: str, percentage: float = None):
        """Update progress indicator."""
        if percentage is not None:
            if self.use_rich:
                self.console.print(f"[yellow]   {message} ({percentage:.1f}%)[/yellow]")
            else:
                print(f"   {message} ({percentage:.1f}%)")
        else:
            if self.use_rich:
                self.console.print(f"[yellow]   {message}[/yellow]")
            else:
                print(f"   {message}")

    def print_warning(self, message: str):
        """Print a warning message."""
        if self.use_rich:
            self.console.print(f"[yellow]⚠️  WARNING: {message}[/yellow]")
        else:
            color = self.colors.get('warning', '')
            reset = self.colors.get('reset', '')
            print(f"{color}⚠️  WARNING: {message}{reset}")

    def print_error(self, message: str):
        """Print an error message."""
        if self.use_rich:
            self.console.print(f"[red]❌ ERROR: {message}[/red]")
        else:
            color = self.colors.get('error', '')
            reset = self.colors.get('reset', '')
            print(f"{color}❌ ERROR: {message}{reset}")

    def print_info(self, message: str):
        """Print an info message."""
        if self.use_rich:
            self.console.print(f"[blue]ℹ️  {message}[/blue]")
        else:
            color = self.colors.get('info', '')
            reset = self.colors.get('reset', '')
            print(f"{color}ℹ️  {message}{reset}")

    def print_success(self, message: str):
        """Print a success message."""
        if self.use_rich:
            self.console.print(f"[green]✅ {message}[/green]")
        else:
            color = self.colors.get('success', '')
            reset = self.colors.get('reset', '')
            print(f"{color}✅ {message}{reset}")

    def create_progress_bar(self, description: str, total: int = 100):
        """Create a progress bar context manager."""
        if self.use_rich and self.config.show_progress:
            return Progress(
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeElapsedColumn(),
                console=self.console
            )
        else:
            return DummyProgress()

    def export_result(self, result: SolverResult, start_address: str = None):
        """Export result to file if export path is specified."""
        if not self.config.export_path:
            return

        try:
            if self.config.export_path.endswith('.json'):
                self._export_json(result, start_address)
            elif self.config.export_path.endswith('.csv'):
                self._export_csv(result, start_address)
            elif self.config.export_path.endswith('.md'):
                self._export_markdown(result, start_address)
            elif self.config.export_path.endswith('.html'):
                self._export_html(result, start_address)
            else:
                self.print_warning(f"Unknown export format for {self.config.export_path}")
        except Exception as e:
            self.print_error(f"Failed to export to {self.config.export_path}: {e}")

    def _export_json(self, result: SolverResult, start_address: str = None):
        """Export result to JSON file."""
        import json

        result_dict = {
            'success': result.success,
            'selected_indices': result.selected_indices,
            'selected_values': result.selected_values,
            'total_sum': result.total_sum,
            'target': result.target,
            'error': result.error,
            'count': len(result.selected_indices),
            'solve_time': result.solve_time,
            'strategy': result.strategy_used.value,
            'mode': result.mode_used.value,
            'status': result.solver_status,
            'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
        }

        if start_address and result.success:
            from fuzzy_sum import A1Converter
            addresses = A1Converter.offset_addresses(start_address, result.selected_indices)
            result_dict['cell_addresses'] = addresses

        with open(self.config.export_path, 'w') as f:
            json.dump(result_dict, f, indent=2)

    def _export_csv(self, result: SolverResult, start_address: str = None):
        """Export result to CSV file."""
        import csv

        with open(self.config.export_path, 'w', newline='') as f:
            writer = csv.writer(f)

            # Header row
            headers = ['Index', 'Value']
            if start_address:
                headers.insert(0, 'Cell_Address')
            writer.writerow(headers)

            # Data rows
            if result.success:
                addresses = None
                if start_address:
                    from fuzzy_sum import A1Converter
                    addresses = A1Converter.offset_addresses(start_address, result.selected_indices)

                for i, (idx, value) in enumerate(zip(result.selected_indices, result.selected_values)):
                    row = [idx, value]
                    if addresses:
                        row.insert(0, addresses[i])
                    writer.writerow(row)

    def _export_markdown(self, result: SolverResult, start_address: str = None):
        """Export result to Markdown file."""
        content = f"""# FuzzySum Results

**Status:** {'✅ SUCCESS' if result.success else '❌ FAILED'}
**Timestamp:** {time.strftime('%Y-%m-%d %H:%M:%S')}

## Summary

| Metric | Value |
|--------|-------|
| Target Sum | {result.target:,.2f} |
| Achieved Sum | {result.total_sum:,.2f} |
| Error | {result.error:,.4f} |
| Items Selected | {len(result.selected_indices)} |
| Solve Time | {result.solve_time:.3f}s |
| Strategy | {result.strategy_used.value} |
| Status | {result.solver_status} |

"""

        if result.success and result.selected_indices:
            content += "## Selected Items\n\n"
            content += "| # | Index | Value |\n|---|-------|-------|\n"

            for i, (idx, value) in enumerate(zip(result.selected_indices, result.selected_values)):
                content += f"| {i+1} | {idx} | {value:.4f} |\n"

        with open(self.config.export_path, 'w') as f:
            f.write(content)

    def _export_html(self, result: SolverResult, start_address: str = None):
        """Export result to HTML file."""
        status_class = "success" if result.success else "error"
        status_text = "✅ SUCCESS" if result.success else "❌ FAILED"

        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>FuzzySum Results</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 40px; }}
        .{status_class} {{ color: {'green' if result.success else 'red'}; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        .number {{ text-align: right; }}
    </style>
</head>
<body>
    <h1>FuzzySum Results</h1>
    <h2 class="{status_class}">{status_text}</h2>
    <p><strong>Timestamp:</strong> {time.strftime('%Y-%m-%d %H:%M:%S')}</p>

    <h3>Summary</h3>
    <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Target Sum</td><td class="number">{result.target:,.2f}</td></tr>
        <tr><td>Achieved Sum</td><td class="number">{result.total_sum:,.2f}</td></tr>
        <tr><td>Error</td><td class="number">{result.error:,.4f}</td></tr>
        <tr><td>Items Selected</td><td class="number">{len(result.selected_indices)}</td></tr>
        <tr><td>Solve Time</td><td class="number">{result.solve_time:.3f}s</td></tr>
        <tr><td>Strategy</td><td>{result.strategy_used.value}</td></tr>
        <tr><td>Status</td><td>{result.solver_status}</td></tr>
    </table>
"""

        if result.success and result.selected_indices:
            html_content += """
    <h3>Selected Items</h3>
    <table>
        <tr><th>#</th><th>Index</th><th>Value</th></tr>
"""
            for i, (idx, value) in enumerate(zip(result.selected_indices, result.selected_values)):
                html_content += f"        <tr><td>{i+1}</td><td>{idx}</td><td class='number'>{value:.4f}</td></tr>\n"

            html_content += "    </table>"

        html_content += "\n</body>\n</html>"

        with open(self.config.export_path, 'w') as f:
            f.write(html_content)


class DummyProgress:
    """Dummy progress bar for when rich is not available."""

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def add_task(self, description, total=100):
        return 0

    def update(self, task_id, advance=1):
        pass


def create_formatter(
    level: str = "summary",
    format_type: str = "terminal",
    width: int = None,
    show_progress: bool = True,
    export_path: str = None
) -> TerminalFormatter:
    """
    Create a formatter with the specified configuration.

    Args:
        level: Output level ('minimal', 'summary', 'detailed', 'verbose', 'json')
        format_type: Output format ('terminal', 'plain', 'json', 'csv', 'markdown', 'html')
        width: Terminal width (None for auto-detect)
        show_progress: Whether to show progress indicators
        export_path: Path to export results (optional)

    Returns:
        Configured TerminalFormatter instance
    """
    config = FormatConfig(
        level=OutputLevel(level.lower()),
        format=OutputFormat(format_type.lower()),
        width=width,
        show_progress=show_progress,
        export_path=export_path
    )

    return TerminalFormatter(config)