#!/usr/bin/env python3
"""
FuzzySum Demo: Example usage of the FuzzySum library.

This script demonstrates various ways to use FuzzySum for solving
subset sum problems with different configurations and strategies.
"""

import sys
import time
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from fuzzy_sum import solve_subset_sum, SolveMode, SolverStrategy, SolverConstraints
from cli_formatter import create_formatter
from parallel_solver import solve_parallel_targets
from config import ConfigManager


def demo_basic_usage():
    """Demonstrate basic FuzzySum usage."""
    print("🔍 Demo 1: Basic Usage")
    print("-" * 30)

    # Sample data
    values = [125.50, 234.75, 89.25, 456.80, 178.90, 312.45, 67.30, 189.60]
    target = 500.00

    print(f"Values: {values}")
    print(f"Target: {target}")
    print()

    # Solve with default settings
    result = solve_subset_sum(values, target)

    formatter = create_formatter(level="detailed", format_type="plain")
    formatter.print_result_summary(result)
    print()


def demo_different_modes():
    """Demonstrate different solving modes."""
    print("🎯 Demo 2: Different Solving Modes")
    print("-" * 35)

    values = [100, 200, 300, 400, 500]
    target = 450

    modes = ["closest", "le", "ge", "exact"]

    for mode in modes:
        print(f"Mode: {mode.upper()}")
        result = solve_subset_sum(values, target, mode=mode)

        if result.success:
            print(f"  Selected: {result.selected_values}")
            print(f"  Sum: {result.total_sum}, Error: {result.error:.2f}")
        else:
            print(f"  No solution found")
        print()


def demo_constraints():
    """Demonstrate using constraints."""
    print("⚙️ Demo 3: Using Constraints")
    print("-" * 28)

    values = [50, 100, 150, 200, 250, 300, 350, 400]
    target = 600

    print(f"Values: {values}")
    print(f"Target: {target}")
    print()

    # Solve with constraints
    result = solve_subset_sum(
        values, target,
        min_items=2,
        max_items=4,
        min_value=100,
        max_value=300
    )

    print("Constraints: min_items=2, max_items=4, min_value=100, max_value=300")
    if result.success:
        print(f"Selected values: {result.selected_values}")
        print(f"Total sum: {result.total_sum}")
        print(f"Count: {len(result.selected_values)} items")
    else:
        print("No solution found with given constraints")
    print()


def demo_strategies():
    """Demonstrate different solver strategies."""
    print("🚀 Demo 4: Different Strategies")
    print("-" * 32)

    values = [15, 25, 35, 45, 55, 65, 75, 85, 95]
    target = 200

    strategies = ["ortools", "dynamic", "greedy"]

    for strategy in strategies:
        start_time = time.time()
        result = solve_subset_sum(values, target, strategy=strategy)
        solve_time = time.time() - start_time

        print(f"Strategy: {strategy.upper()}")
        if result.success:
            print(f"  Sum: {result.total_sum}, Error: {result.error:.2f}")
            print(f"  Time: {solve_time:.3f}s")
        else:
            print(f"  Failed: {result.message}")
        print()


def demo_parallel_solving():
    """Demonstrate parallel solving for multiple targets."""
    print("⚡ Demo 5: Parallel Solving")
    print("-" * 28)

    values = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
    targets = [75, 150, 225, 300, 175]

    print(f"Values: {values}")
    print(f"Targets: {targets}")
    print()

    def progress_callback(completed, total):
        print(f"Progress: {completed}/{total} targets completed")

    start_time = time.time()
    results = solve_parallel_targets(
        values, targets,
        max_workers=3,
        progress_callback=progress_callback
    )
    total_time = time.time() - start_time

    print(f"\nCompleted {len(results)} targets in {total_time:.2f}s")
    print("Results:")

    for i, result in enumerate(results):
        target = targets[i]
        if result.result.success:
            print(f"  Target {target}: Sum={result.result.total_sum:.2f}, "
                  f"Error={result.result.error:.2f}")
        else:
            print(f"  Target {target}: Failed")
    print()


def demo_excel_style_references():
    """Demonstrate Excel-style A1 references."""
    print("📊 Demo 6: Excel-Style References")
    print("-" * 35)

    from fuzzy_sum import A1Converter

    values = [125.50, 234.75, 89.25, 456.80, 178.90]
    target = 400
    start_address = "B5"  # Starting at cell B5

    result = solve_subset_sum(values, target)

    if result.success:
        # Convert indices to A1 addresses
        addresses = A1Converter.offset_addresses(start_address, result.selected_indices)

        print(f"Starting address: {start_address}")
        print(f"Selected cells: {', '.join(addresses)}")
        print(f"Values: {result.selected_values}")
        print(f"Total: {result.total_sum}")
    else:
        print("No solution found")
    print()


def demo_configuration():
    """Demonstrate configuration management."""
    print("⚙️ Demo 7: Configuration Management")
    print("-" * 37)

    # Create config manager
    config_manager = ConfigManager()

    # Show default settings
    print("Default solver settings:")
    solver_config = config_manager.get_solver_config()
    for key, value in solver_config.items():
        print(f"  {key}: {value}")
    print()

    # Create a custom profile
    config_manager.create_profile("fast_mode", solver={
        "strategy": "greedy",
        "time_limit": 1.0
    })

    print("Created profile: fast_mode")
    print("Applying profile...")

    # Apply profile
    config_manager.apply_profile("fast_mode")

    print("Updated solver settings:")
    solver_config = config_manager.get_solver_config()
    for key, value in solver_config.items():
        print(f"  {key}: {value}")
    print()


def demo_error_handling():
    """Demonstrate error handling and edge cases."""
    print("⚠️ Demo 8: Error Handling")
    print("-" * 26)

    # Test with empty values
    print("Testing with empty values:")
    result = solve_subset_sum([], 100)
    print(f"Success: {result.success}, Status: {result.solver_status}")
    print()

    # Test with impossible target
    print("Testing with impossible target:")
    result = solve_subset_sum([1, 2, 3], 1000, mode="exact")
    print(f"Success: {result.success}, Status: {result.solver_status}")
    print()

    # Test with conflicting constraints
    print("Testing with conflicting constraints:")
    result = solve_subset_sum(
        [10, 20, 30],
        50,
        min_items=5,  # Impossible - only 3 values available
        max_items=2
    )
    print(f"Success: {result.success}")
    if not result.success:
        print(f"Message: {result.message}")
    print()


def main():
    """Run all demos."""
    print("🎯 FuzzySum Library Demonstration")
    print("=" * 50)
    print()

    demos = [
        demo_basic_usage,
        demo_different_modes,
        demo_constraints,
        demo_strategies,
        demo_parallel_solving,
        demo_excel_style_references,
        demo_configuration,
        demo_error_handling
    ]

    for i, demo_func in enumerate(demos, 1):
        try:
            demo_func()
        except Exception as e:
            print(f"Demo {i} failed: {e}")
            print()

        if i < len(demos):
            input("Press Enter to continue to next demo...")
            print()

    print("🎉 All demos completed!")
    print("Try running the CLI tool:")
    print("  python subset_sum_cli.py --csv examples/sample_data.csv --target 500 --addr A1")


if __name__ == "__main__":
    main()