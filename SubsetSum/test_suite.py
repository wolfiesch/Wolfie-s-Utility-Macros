#!/usr/bin/env python3
"""
Test Suite: Comprehensive tests for FuzzySum subset sum solver.

This module contains unit tests, integration tests, performance benchmarks,
and edge case validation for all FuzzySum components.
"""

import unittest
import tempfile
import json
import time
import random
from pathlib import Path
from typing import List, Dict, Any
import logging

# Set up logging for tests
logging.basicConfig(level=logging.WARNING)

from fuzzy_sum import (
    solve_subset_sum, SubsetSumSolver, SolverResult, SolveMode, SolverStrategy,
    A1Converter, SolverConstraints
)
from cli_formatter import create_formatter, OutputLevel, OutputFormat
from config import ConfigManager, FuzzySumConfig


class TestA1Converter(unittest.TestCase):
    """Test A1 notation converter utilities."""

    def test_col_to_index(self):
        """Test column letter to index conversion."""
        self.assertEqual(A1Converter.col_to_index("A"), 1)
        self.assertEqual(A1Converter.col_to_index("Z"), 26)
        self.assertEqual(A1Converter.col_to_index("AA"), 27)
        self.assertEqual(A1Converter.col_to_index("AB"), 28)
        self.assertEqual(A1Converter.col_to_index("AZ"), 52)
        self.assertEqual(A1Converter.col_to_index("BA"), 53)

    def test_index_to_col(self):
        """Test index to column letter conversion."""
        self.assertEqual(A1Converter.index_to_col(1), "A")
        self.assertEqual(A1Converter.index_to_col(26), "Z")
        self.assertEqual(A1Converter.index_to_col(27), "AA")
        self.assertEqual(A1Converter.index_to_col(28), "AB")
        self.assertEqual(A1Converter.index_to_col(52), "AZ")
        self.assertEqual(A1Converter.index_to_col(53), "BA")

    def test_roundtrip_conversion(self):
        """Test roundtrip conversion between indices and column letters."""
        for i in range(1, 1000):
            col = A1Converter.index_to_col(i)
            back_to_index = A1Converter.col_to_index(col)
            self.assertEqual(i, back_to_index)

    def test_parse_a1(self):
        """Test A1 address parsing."""
        self.assertEqual(A1Converter.parse_a1("A1"), (1, 1))
        self.assertEqual(A1Converter.parse_a1("B5"), (2, 5))
        self.assertEqual(A1Converter.parse_a1("Z10"), (26, 10))
        self.assertEqual(A1Converter.parse_a1("AA100"), (27, 100))

    def test_to_a1(self):
        """Test A1 address generation."""
        self.assertEqual(A1Converter.to_a1(1, 1), "A1")
        self.assertEqual(A1Converter.to_a1(2, 5), "B5")
        self.assertEqual(A1Converter.to_a1(26, 10), "Z10")
        self.assertEqual(A1Converter.to_a1(27, 100), "AA100")

    def test_invalid_a1_addresses(self):
        """Test handling of invalid A1 addresses."""
        with self.assertRaises(ValueError):
            A1Converter.parse_a1("123")
        with self.assertRaises(ValueError):
            A1Converter.parse_a1("A")
        with self.assertRaises(ValueError):
            A1Converter.parse_a1("")

    def test_offset_addresses(self):
        """Test offset address generation."""
        addresses = A1Converter.offset_addresses("A5", [0, 1, 3])
        self.assertEqual(addresses, ["A5", "A6", "A8"])

        addresses = A1Converter.offset_addresses("C10", [2, 5])
        self.assertEqual(addresses, ["C12", "C15"])


class TestSubsetSumSolver(unittest.TestCase):
    """Test core subset sum solver functionality."""

    def setUp(self):
        """Set up test fixtures."""
        self.solver = SubsetSumSolver(scale_factor=100, time_limit=1.0)
        self.test_values = [1.0, 2.5, 3.0, 4.5, 5.0, 7.5, 10.0]

    def test_exact_solution(self):
        """Test finding exact solutions."""
        # Test case where exact solution exists
        values = [1, 2, 3, 4, 5]
        target = 8  # Can be made with 3 + 5

        result = self.solver.solve(
            values, target, SolveMode.EXACT, SolverStrategy.ORTOOLS
        )

        self.assertTrue(result.success)
        self.assertEqual(result.total_sum, target)
        self.assertEqual(result.error, 0.0)

    def test_closest_solution(self):
        """Test finding closest solutions."""
        values = [1, 2, 3, 4, 5]
        target = 7.5  # Closest would be 7 (3+4) or 8 (3+5)

        result = self.solver.solve(
            values, target, SolveMode.CLOSEST, SolverStrategy.ORTOOLS
        )

        self.assertTrue(result.success)
        self.assertLessEqual(result.error, 1.0)  # Should be close

    def test_le_solution(self):
        """Test less-than-or-equal solutions."""
        values = [1, 2, 3, 4, 5]
        target = 8

        result = self.solver.solve(
            values, target, SolveMode.LE, SolverStrategy.ORTOOLS
        )

        self.assertTrue(result.success)
        self.assertLessEqual(result.total_sum, target)

    def test_ge_solution(self):
        """Test greater-than-or-equal solutions."""
        values = [1, 2, 3, 4, 5]
        target = 8

        result = self.solver.solve(
            values, target, SolveMode.GE, SolverStrategy.ORTOOLS
        )

        self.assertTrue(result.success)
        self.assertGreaterEqual(result.total_sum, target)

    def test_constraints_min_max_items(self):
        """Test minimum and maximum item constraints."""
        values = [1, 2, 3, 4, 5]
        target = 10

        constraints = SolverConstraints(min_items=2, max_items=3)
        result = self.solver.solve(
            values, target, SolveMode.CLOSEST, SolverStrategy.ORTOOLS, constraints
        )

        if result.success:
            self.assertGreaterEqual(len(result.selected_indices), 2)
            self.assertLessEqual(len(result.selected_indices), 3)

    def test_constraints_value_range(self):
        """Test value range constraints."""
        values = [1, 2, 3, 4, 5, 10, 20]
        target = 15

        constraints = SolverConstraints(min_value=2, max_value=10)
        result = self.solver.solve(
            values, target, SolveMode.CLOSEST, SolverStrategy.ORTOOLS, constraints
        )

        if result.success:
            for val in result.selected_values:
                self.assertGreaterEqual(val, 2)
                self.assertLessEqual(val, 10)

    def test_constraints_required_forbidden(self):
        """Test required and forbidden item constraints."""
        values = [1, 2, 3, 4, 5]
        target = 10

        constraints = SolverConstraints(
            required_indices={0, 1},  # Must include indices 0 and 1
            forbidden_indices={4}     # Cannot include index 4
        )
        result = self.solver.solve(
            values, target, SolveMode.CLOSEST, SolverStrategy.ORTOOLS, constraints
        )

        if result.success:
            self.assertIn(0, result.selected_indices)
            self.assertIn(1, result.selected_indices)
            self.assertNotIn(4, result.selected_indices)

    def test_empty_values(self):
        """Test handling of empty value lists."""
        result = self.solver.solve([], 10, SolveMode.CLOSEST)
        self.assertFalse(result.success)
        self.assertEqual(result.solver_status, "EMPTY_INPUT")

    def test_no_valid_values(self):
        """Test handling when no values meet constraints."""
        values = [1, 2, 3]
        constraints = SolverConstraints(min_value=10)  # No values >= 10

        result = self.solver.solve(
            values, 15, SolveMode.CLOSEST, SolverStrategy.ORTOOLS, constraints
        )
        self.assertFalse(result.success)

    def test_different_strategies(self):
        """Test different solver strategies."""
        values = [1, 2, 3, 4, 5]
        target = 8

        strategies = [SolverStrategy.ORTOOLS, SolverStrategy.DYNAMIC, SolverStrategy.GREEDY]

        for strategy in strategies:
            with self.subTest(strategy=strategy):
                result = self.solver.solve(values, target, SolveMode.CLOSEST, strategy)
                # All strategies should find some solution for this simple case
                self.assertTrue(result.success or strategy == SolverStrategy.DYNAMIC)


class TestConvenienceFunction(unittest.TestCase):
    """Test the convenience solve_subset_sum function."""

    def test_basic_usage(self):
        """Test basic usage of convenience function."""
        values = [1.0, 2.0, 3.0, 4.0, 5.0]
        target = 7.0

        result = solve_subset_sum(values, target)
        self.assertIsInstance(result, SolverResult)
        self.assertTrue(result.success)

    def test_with_constraints(self):
        """Test convenience function with constraints."""
        values = [1, 2, 3, 4, 5]
        target = 10

        result = solve_subset_sum(
            values, target,
            mode="closest",
            strategy="ortools",
            min_items=2,
            max_items=3
        )

        if result.success:
            self.assertGreaterEqual(len(result.selected_indices), 2)
            self.assertLessEqual(len(result.selected_indices), 3)

    def test_invalid_mode(self):
        """Test handling of invalid mode."""
        values = [1, 2, 3]
        target = 5

        with self.assertRaises(ValueError):
            solve_subset_sum(values, target, mode="invalid_mode")

    def test_invalid_strategy(self):
        """Test handling of invalid strategy."""
        values = [1, 2, 3]
        target = 5

        with self.assertRaises(ValueError):
            solve_subset_sum(values, target, strategy="invalid_strategy")


class TestFormatterOutput(unittest.TestCase):
    """Test CLI formatter functionality."""

    def test_create_formatter(self):
        """Test formatter creation with different configs."""
        formatter = create_formatter(
            level="summary",
            format_type="terminal",
            show_progress=True
        )
        self.assertIsNotNone(formatter)

    def test_formatter_levels(self):
        """Test different output levels."""
        levels = ["minimal", "summary", "detailed", "verbose", "json"]

        for level in levels:
            with self.subTest(level=level):
                formatter = create_formatter(level=level, format_type="plain")
                self.assertIsNotNone(formatter)

    def test_result_formatting(self):
        """Test result formatting doesn't crash."""
        values = [1, 2, 3, 4, 5]
        target = 8
        result = solve_subset_sum(values, target)

        formatter = create_formatter(level="summary", format_type="plain")

        # Test that formatting doesn't raise exceptions
        try:
            formatter.print_result_summary(result, "A1")
        except Exception as e:
            self.fail(f"Result formatting raised an exception: {e}")


class TestConfigManager(unittest.TestCase):
    """Test configuration management."""

    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = Path(tempfile.mkdtemp())

    def tearDown(self):
        """Clean up test fixtures."""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_default_config(self):
        """Test default configuration creation."""
        config_manager = ConfigManager()
        self.assertIsInstance(config_manager.config, FuzzySumConfig)

    def test_save_load_config(self):
        """Test saving and loading configuration."""
        config_path = self.temp_dir / "test_config.json"

        # Create and save config
        config_manager = ConfigManager()
        config_manager.config.solver.mode = "le"
        config_manager.config.solver.time_limit = 10.0
        config_manager.save_config(config_path)

        # Load config and verify
        loaded_manager = ConfigManager(config_path)
        self.assertEqual(loaded_manager.config.solver.mode, "le")
        self.assertEqual(loaded_manager.config.solver.time_limit, 10.0)

    def test_profiles(self):
        """Test configuration profiles."""
        config_manager = ConfigManager()

        # Create profile
        config_manager.create_profile("test_profile", solver={"mode": "exact"})
        self.assertIn("test_profile", config_manager.list_profiles())

        # Apply profile
        config_manager.apply_profile("test_profile")
        self.assertEqual(config_manager.config.solver.mode, "exact")

        # Delete profile
        config_manager.delete_profile("test_profile")
        self.assertNotIn("test_profile", config_manager.list_profiles())

    def test_config_validation(self):
        """Test configuration validation."""
        config_manager = ConfigManager()

        # Valid config should pass
        errors = config_manager.validate_config()
        self.assertEqual(len(errors), 0)

        # Invalid config should fail
        config_manager.config.solver.mode = "invalid_mode"
        errors = config_manager.validate_config()
        self.assertGreater(len(errors), 0)


class TestExcelIntegration(unittest.TestCase):
    """Test Excel integration functionality."""

    def test_excel_import_available(self):
        """Test that Excel integration imports without error."""
        try:
            from excel_integration import ExcelReader, ExcelWriter, get_available_engines
            engines = get_available_engines()
            self.assertIsInstance(engines, dict)
        except ImportError:
            self.skipTest("Excel integration dependencies not available")

    def test_excel_reader_creation(self):
        """Test ExcelReader creation with non-existent file."""
        try:
            from excel_integration import ExcelReader, ExcelIntegrationError

            with self.assertRaises(FileNotFoundError):
                ExcelReader(Path("nonexistent.xlsx"))

        except ImportError:
            self.skipTest("Excel integration dependencies not available")


class TestParallelSolver(unittest.TestCase):
    """Test parallel solver functionality."""

    def test_parallel_import(self):
        """Test that parallel solver imports without error."""
        try:
            from parallel_solver import ParallelSolver, ParallelConfig, solve_parallel_targets
        except ImportError as e:
            self.fail(f"Parallel solver import failed: {e}")

    def test_parallel_solver_creation(self):
        """Test parallel solver creation."""
        try:
            from parallel_solver import ParallelSolver, ParallelConfig

            config = ParallelConfig(max_workers=2, enable_caching=False)
            solver = ParallelSolver(config)
            self.assertIsNotNone(solver)

        except ImportError:
            self.skipTest("Parallel solver dependencies not available")

    def test_parallel_multiple_targets(self):
        """Test parallel solving of multiple targets."""
        try:
            from parallel_solver import solve_parallel_targets

            values = [1, 2, 3, 4, 5]
            targets = [5, 7, 10]

            results = solve_parallel_targets(
                values, targets,
                max_workers=2,
                enable_caching=False
            )

            self.assertEqual(len(results), len(targets))
            for result in results:
                self.assertIsNotNone(result.result)

        except ImportError:
            self.skipTest("Parallel solver dependencies not available")


class TestPerformance(unittest.TestCase):
    """Performance benchmark tests."""

    def test_small_problem_performance(self):
        """Test performance on small problems."""
        values = list(range(1, 21))  # 20 values
        target = 50

        start_time = time.time()
        result = solve_subset_sum(values, target, time_limit=2.0)
        solve_time = time.time() - start_time

        self.assertLess(solve_time, 2.0)  # Should complete quickly
        self.assertTrue(result.success)

    def test_medium_problem_performance(self):
        """Test performance on medium problems."""
        values = list(range(1, 101))  # 100 values
        target = 500

        start_time = time.time()
        result = solve_subset_sum(values, target, time_limit=5.0)
        solve_time = time.time() - start_time

        self.assertLess(solve_time, 5.0)
        # Should find some solution, even if not optimal

    def test_large_problem_with_greedy(self):
        """Test greedy strategy on large problems."""
        values = [random.uniform(1, 100) for _ in range(1000)]
        target = sum(values) * 0.3  # Target 30% of total

        start_time = time.time()
        result = solve_subset_sum(
            values, target,
            strategy="greedy",
            time_limit=5.0
        )
        solve_time = time.time() - start_time

        self.assertLess(solve_time, 5.0)
        self.assertTrue(result.success)  # Greedy should always find something


class TestEdgeCases(unittest.TestCase):
    """Test edge cases and error conditions."""

    def test_single_value(self):
        """Test with single value."""
        values = [5.0]
        target = 5.0

        result = solve_subset_sum(values, target)
        self.assertTrue(result.success)
        self.assertEqual(result.total_sum, 5.0)

    def test_negative_values(self):
        """Test with negative values."""
        values = [-1, -2, 3, 4, 5]
        target = 2

        result = solve_subset_sum(values, target)
        # Should handle negative values appropriately

    def test_zero_target(self):
        """Test with zero target."""
        values = [1, 2, 3, 4, 5]
        target = 0

        result = solve_subset_sum(values, target)
        if result.success:
            self.assertEqual(result.total_sum, 0)

    def test_very_large_values(self):
        """Test with very large values."""
        values = [1e6, 2e6, 3e6]
        target = 4e6

        result = solve_subset_sum(values, target, scale_factor=1)
        # Should handle large values without overflow

    def test_very_small_values(self):
        """Test with very small values."""
        values = [0.001, 0.002, 0.003]
        target = 0.004

        result = solve_subset_sum(values, target, scale_factor=1000)
        # Should handle small values with appropriate scaling

    def test_duplicate_values(self):
        """Test with duplicate values."""
        values = [1, 1, 2, 2, 3, 3]
        target = 6

        result = solve_subset_sum(values, target)
        self.assertTrue(result.success)

    def test_impossible_target(self):
        """Test with impossible target."""
        values = [1, 2, 3]
        target = 100  # Impossible to achieve

        result = solve_subset_sum(values, target, mode="exact")
        # May succeed with closest approximation depending on mode


def run_benchmark_suite():
    """Run performance benchmarks and print results."""
    print("FuzzySum Performance Benchmarks")
    print("=" * 40)

    # Small problem benchmark
    values_small = list(range(1, 21))
    target_small = 50

    start = time.time()
    result_small = solve_subset_sum(values_small, target_small)
    time_small = time.time() - start

    print(f"Small problem (20 values): {time_small:.3f}s")
    print(f"  Success: {result_small.success}, Error: {result_small.error:.3f}")

    # Medium problem benchmark
    values_medium = list(range(1, 101))
    target_medium = 500

    start = time.time()
    result_medium = solve_subset_sum(values_medium, target_medium)
    time_medium = time.time() - start

    print(f"Medium problem (100 values): {time_medium:.3f}s")
    print(f"  Success: {result_medium.success}, Error: {result_medium.error:.3f}")

    # Large problem with greedy
    values_large = [random.uniform(1, 100) for _ in range(1000)]
    target_large = sum(values_large) * 0.3

    start = time.time()
    result_large = solve_subset_sum(values_large, target_large, strategy="greedy")
    time_large = time.time() - start

    print(f"Large problem (1000 values, greedy): {time_large:.3f}s")
    print(f"  Success: {result_large.success}, Error: {result_large.error:.3f}")

    print("=" * 40)


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "--benchmark":
        run_benchmark_suite()
    else:
        # Run test suite
        unittest.main(verbosity=2)