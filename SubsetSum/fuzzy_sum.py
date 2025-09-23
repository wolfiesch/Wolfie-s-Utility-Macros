#!/usr/bin/env python3
"""
FuzzySum: Core library for subset sum solving with multiple strategies.

This module provides the core functionality for finding optimal subsets of numbers
that sum closest to a target value, using various algorithms including OR-Tools,
dynamic programming, and greedy approaches.
"""

import re
import time
import logging
from typing import List, Tuple, Optional, Dict, Any, Set
from dataclasses import dataclass
from enum import Enum
import numpy as np

try:
    from ortools.sat.python import cp_model
    HAS_ORTOOLS = True
except ImportError:
    HAS_ORTOOLS = False
    logging.warning("OR-Tools not available. Some solver methods will be disabled.")


class SolveMode(Enum):
    """Solving modes for subset sum problems."""
    CLOSEST = "closest"           # Find subset with sum closest to target
    LE = "le"                    # Find subset with sum <= target (maximize)
    GE = "ge"                    # Find subset with sum >= target (minimize)
    EXACT = "exact"              # Find subset with exact sum (if possible)
    RANGE = "range"              # Find subset within a range of target


class SolverStrategy(Enum):
    """Available solver strategies."""
    ORTOOLS = "ortools"          # Use OR-Tools CP-SAT solver
    DYNAMIC = "dynamic"          # Dynamic programming approach
    GREEDY = "greedy"           # Greedy approximation
    HYBRID = "hybrid"           # Combination of strategies


@dataclass
class SolverConstraints:
    """Constraints for subset sum solving."""
    min_items: Optional[int] = None
    max_items: Optional[int] = None
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    required_indices: Optional[Set[int]] = None
    forbidden_indices: Optional[Set[int]] = None
    tolerance: Optional[float] = None  # For range mode


@dataclass
class SolverResult:
    """Result of a subset sum solve operation."""
    success: bool
    selected_indices: List[int]
    selected_values: List[float]
    total_sum: float
    target: float
    error: float
    solve_time: float
    strategy_used: SolverStrategy
    mode_used: SolveMode
    solver_status: str
    iterations: int = 0
    message: str = ""


class A1Converter:
    """Utilities for converting between Excel A1 notation and indices."""

    A1_PATTERN = re.compile(r"^([A-Za-z]+)(\d+)$")

    @staticmethod
    def col_to_index(col_letters: str) -> int:
        """Convert column letters to 1-based index (A=1, Z=26, AA=27, etc.)."""
        col_letters = col_letters.upper()
        result = 0
        for char in col_letters:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result

    @staticmethod
    def index_to_col(index: int) -> str:
        """Convert 1-based index to column letters (1=A, 26=Z, 27=AA, etc.)."""
        result = []
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            result.append(chr(ord('A') + remainder))
        return ''.join(reversed(result))

    @classmethod
    def parse_a1(cls, address: str) -> Tuple[int, int]:
        """Parse A1 address into (column_index, row_number)."""
        match = cls.A1_PATTERN.match(address.strip())
        if not match:
            raise ValueError(f"Invalid A1 address: {address}")
        col_letters, row_str = match.groups()
        return cls.col_to_index(col_letters), int(row_str)

    @classmethod
    def to_a1(cls, col_index: int, row_number: int) -> str:
        """Convert column index and row number to A1 address."""
        return f"{cls.index_to_col(col_index)}{row_number}"

    @classmethod
    def offset_addresses(cls, start_address: str, indices: List[int]) -> List[str]:
        """Convert array indices to A1 addresses given a starting address."""
        start_col, start_row = cls.parse_a1(start_address)
        col_letters = cls.index_to_col(start_col)
        return [f"{col_letters}{start_row + i}" for i in indices]


class SubsetSumSolver:
    """Main solver class for subset sum problems."""

    def __init__(self, scale_factor: int = 100, time_limit: float = 5.0):
        """
        Initialize the solver.

        Args:
            scale_factor: Factor to scale floats to integers (100 for cents)
            time_limit: Maximum time for solver in seconds
        """
        self.scale_factor = scale_factor
        self.time_limit = time_limit
        self.logger = logging.getLogger(__name__)

    def solve(
        self,
        values: List[float],
        target: float,
        mode: SolveMode = SolveMode.CLOSEST,
        strategy: SolverStrategy = SolverStrategy.ORTOOLS,
        constraints: Optional[SolverConstraints] = None
    ) -> SolverResult:
        """
        Solve subset sum problem with given parameters.

        Args:
            values: List of numbers to choose from
            target: Target sum to achieve
            mode: Solving mode (closest, le, ge, exact, range)
            strategy: Solver strategy to use
            constraints: Additional constraints on the solution

        Returns:
            SolverResult with solution details
        """
        start_time = time.time()

        if not values:
            return SolverResult(
                success=False,
                selected_indices=[],
                selected_values=[],
                total_sum=0.0,
                target=target,
                error=abs(target),
                solve_time=0.0,
                strategy_used=strategy,
                mode_used=mode,
                solver_status="EMPTY_INPUT",
                message="No values provided"
            )

        # Filter out invalid values and apply constraints
        filtered_values, valid_indices = self._filter_values(values, constraints)

        if not filtered_values:
            return SolverResult(
                success=False,
                selected_indices=[],
                selected_values=[],
                total_sum=0.0,
                target=target,
                error=abs(target),
                solve_time=time.time() - start_time,
                strategy_used=strategy,
                mode_used=mode,
                solver_status="NO_VALID_VALUES",
                message="No valid values after filtering"
            )

        # Choose and execute strategy
        if strategy == SolverStrategy.ORTOOLS and HAS_ORTOOLS:
            result = self._solve_ortools(filtered_values, target, mode, constraints)
        elif strategy == SolverStrategy.DYNAMIC:
            result = self._solve_dynamic(filtered_values, target, mode, constraints)
        elif strategy == SolverStrategy.GREEDY:
            result = self._solve_greedy(filtered_values, target, mode, constraints)
        elif strategy == SolverStrategy.HYBRID:
            result = self._solve_hybrid(filtered_values, target, mode, constraints)
        else:
            # Fallback to available strategy
            if HAS_ORTOOLS:
                result = self._solve_ortools(filtered_values, target, mode, constraints)
            else:
                result = self._solve_dynamic(filtered_values, target, mode, constraints)

        # Map back to original indices
        if result.success:
            original_indices = [valid_indices[i] for i in result.selected_indices]
            original_values = [values[i] for i in original_indices]
            result.selected_indices = original_indices
            result.selected_values = original_values

        result.solve_time = time.time() - start_time
        return result

    def _filter_values(
        self,
        values: List[float],
        constraints: Optional[SolverConstraints]
    ) -> Tuple[List[float], List[int]]:
        """Filter values based on constraints and validity."""
        filtered_values = []
        valid_indices = []

        for i, value in enumerate(values):
            # Skip non-numeric values
            if not isinstance(value, (int, float)) or np.isnan(value) or np.isinf(value):
                continue

            # Apply constraints
            if constraints:
                if constraints.min_value is not None and value < constraints.min_value:
                    continue
                if constraints.max_value is not None and value > constraints.max_value:
                    continue
                if constraints.forbidden_indices and i in constraints.forbidden_indices:
                    continue

            filtered_values.append(value)
            valid_indices.append(i)

        return filtered_values, valid_indices

    def _solve_ortools(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        constraints: Optional[SolverConstraints]
    ) -> SolverResult:
        """Solve using OR-Tools CP-SAT solver."""
        # Scale to integers
        scaled_values = [int(round(v * self.scale_factor)) for v in values]
        scaled_target = int(round(target * self.scale_factor))

        model = cp_model.CpModel()
        n = len(values)

        # Decision variables
        x = [model.NewBoolVar(f"x_{i}") for i in range(n)]

        # Sum variable
        total_sum = sum(scaled_values[i] * x[i] for i in range(n))

        # Constraints
        if constraints:
            if constraints.min_items is not None:
                model.Add(sum(x) >= constraints.min_items)
            if constraints.max_items is not None:
                model.Add(sum(x) <= constraints.max_items)
            if constraints.required_indices:
                for idx in constraints.required_indices:
                    if 0 <= idx < n:
                        model.Add(x[idx] == 1)

        # Objective based on mode
        if mode == SolveMode.CLOSEST:
            # Minimize |sum - target|
            abs_diff = model.NewIntVar(0, sum(abs(v) for v in scaled_values), "abs_diff")
            model.Add(abs_diff >= total_sum - scaled_target)
            model.Add(abs_diff >= scaled_target - total_sum)
            model.Minimize(abs_diff)
        elif mode == SolveMode.LE:
            # Maximize sum subject to sum <= target
            model.Add(total_sum <= scaled_target)
            model.Maximize(total_sum)
        elif mode == SolveMode.GE:
            # Minimize sum subject to sum >= target
            model.Add(total_sum >= scaled_target)
            model.Minimize(total_sum)
        elif mode == SolveMode.EXACT:
            # Find exact match
            model.Add(total_sum == scaled_target)
        elif mode == SolveMode.RANGE:
            # Within tolerance range
            tolerance = constraints.tolerance if constraints else 0.01
            scaled_tolerance = int(round(tolerance * self.scale_factor))
            model.Add(total_sum >= scaled_target - scaled_tolerance)
            model.Add(total_sum <= scaled_target + scaled_tolerance)
            model.Minimize(abs_diff)  # Still minimize difference within range

        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = self.time_limit
        solver.parameters.num_search_workers = 8

        status = solver.Solve(model)

        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            selected_indices = [i for i in range(n) if solver.Value(x[i]) == 1]
            selected_values = [values[i] for i in selected_indices]
            total = sum(selected_values)

            return SolverResult(
                success=True,
                selected_indices=selected_indices,
                selected_values=selected_values,
                total_sum=total,
                target=target,
                error=abs(target - total),
                solve_time=0.0,  # Will be set by caller
                strategy_used=SolverStrategy.ORTOOLS,
                mode_used=mode,
                solver_status=solver.StatusName(status)
            )
        else:
            return SolverResult(
                success=False,
                selected_indices=[],
                selected_values=[],
                total_sum=0.0,
                target=target,
                error=abs(target),
                solve_time=0.0,
                strategy_used=SolverStrategy.ORTOOLS,
                mode_used=mode,
                solver_status=solver.StatusName(status),
                message=f"OR-Tools solver failed: {solver.StatusName(status)}"
            )

    def _solve_dynamic(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        constraints: Optional[SolverConstraints]
    ) -> SolverResult:
        """Solve using dynamic programming approach."""
        # Scale to integers for DP
        scaled_values = [int(round(v * self.scale_factor)) for v in values]
        scaled_target = int(round(target * self.scale_factor))

        # Simple DP implementation for closest sum
        n = len(values)
        max_sum = sum(abs(v) for v in scaled_values)

        # DP table: dp[i][s] = can we achieve sum s using first i items
        dp = {}
        parent = {}

        # Initialize
        dp[(0, 0)] = True

        for i in range(n):
            new_dp = dp.copy()
            for (items, current_sum), possible in dp.items():
                if possible and items == i:
                    # Take current item
                    new_sum = current_sum + scaled_values[i]
                    new_items = items + 1

                    # Check constraints
                    valid = True
                    if constraints:
                        if constraints.max_items and new_items > constraints.max_items:
                            valid = False

                    if valid and abs(new_sum - scaled_target) <= max_sum:
                        key = (new_items, new_sum)
                        if key not in new_dp or not new_dp[key]:
                            new_dp[key] = True
                            parent[key] = (items, current_sum, i)
            dp = new_dp

        # Find best solution
        best_error = float('inf')
        best_key = None

        for (items, current_sum), possible in dp.items():
            if not possible:
                continue

            # Check constraints
            if constraints:
                if constraints.min_items and items < constraints.min_items:
                    continue
                if constraints.max_items and items > constraints.max_items:
                    continue

            error = abs(current_sum - scaled_target)
            if mode == SolveMode.LE and current_sum > scaled_target:
                continue
            if mode == SolveMode.GE and current_sum < scaled_target:
                continue
            if mode == SolveMode.EXACT and current_sum != scaled_target:
                continue

            if error < best_error:
                best_error = error
                best_key = (items, current_sum)

        if best_key is None:
            return SolverResult(
                success=False,
                selected_indices=[],
                selected_values=[],
                total_sum=0.0,
                target=target,
                error=abs(target),
                solve_time=0.0,
                strategy_used=SolverStrategy.DYNAMIC,
                mode_used=mode,
                solver_status="NO_SOLUTION",
                message="No valid solution found with dynamic programming"
            )

        # Reconstruct solution
        selected_indices = []
        current = best_key
        while current in parent:
            prev_items, prev_sum, item_idx = parent[current]
            selected_indices.append(item_idx)
            current = (prev_items, prev_sum)

        selected_indices.reverse()
        selected_values = [values[i] for i in selected_indices]
        total = sum(selected_values)

        return SolverResult(
            success=True,
            selected_indices=selected_indices,
            selected_values=selected_values,
            total_sum=total,
            target=target,
            error=abs(target - total),
            solve_time=0.0,
            strategy_used=SolverStrategy.DYNAMIC,
            mode_used=mode,
            solver_status="OPTIMAL"
        )

    def _solve_greedy(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        constraints: Optional[SolverConstraints]
    ) -> SolverResult:
        """Solve using greedy approximation."""
        # Sort by value for greedy approach
        indexed_values = [(v, i) for i, v in enumerate(values)]

        if mode == SolveMode.LE:
            # Sort descending for LE mode
            indexed_values.sort(reverse=True)
        else:
            # Sort by proximity to target/len(values) for other modes
            avg_target = target / max(1, len(values) // 4)
            indexed_values.sort(key=lambda x: abs(x[0] - avg_target))

        selected_indices = []
        current_sum = 0.0

        for value, original_idx in indexed_values:
            # Check if we can add this value
            new_sum = current_sum + value

            # Check constraints
            if constraints:
                if constraints.max_items and len(selected_indices) >= constraints.max_items:
                    break
                if constraints.forbidden_indices and original_idx in constraints.forbidden_indices:
                    continue

            # Check mode-specific conditions
            if mode == SolveMode.LE and new_sum > target:
                continue
            elif mode == SolveMode.GE and current_sum >= target:
                break
            elif mode == SolveMode.EXACT and new_sum > target:
                continue

            # Add if it improves the solution
            if mode == SolveMode.CLOSEST:
                if abs(new_sum - target) < abs(current_sum - target):
                    selected_indices.append(original_idx)
                    current_sum = new_sum
            else:
                selected_indices.append(original_idx)
                current_sum = new_sum

            # Early termination for exact match
            if mode == SolveMode.EXACT and current_sum == target:
                break

        # Check minimum items constraint
        if constraints and constraints.min_items:
            if len(selected_indices) < constraints.min_items:
                return SolverResult(
                    success=False,
                    selected_indices=[],
                    selected_values=[],
                    total_sum=0.0,
                    target=target,
                    error=abs(target),
                    solve_time=0.0,
                    strategy_used=SolverStrategy.GREEDY,
                    mode_used=mode,
                    solver_status="CONSTRAINT_VIOLATION",
                    message="Cannot satisfy minimum items constraint"
                )

        selected_values = [values[i] for i in selected_indices]
        total = sum(selected_values)

        return SolverResult(
            success=True,
            selected_indices=selected_indices,
            selected_values=selected_values,
            total_sum=total,
            target=target,
            error=abs(target - total),
            solve_time=0.0,
            strategy_used=SolverStrategy.GREEDY,
            mode_used=mode,
            solver_status="FEASIBLE"
        )

    def _solve_hybrid(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        constraints: Optional[SolverConstraints]
    ) -> SolverResult:
        """Solve using hybrid approach - try multiple strategies."""
        strategies = []

        # Add available strategies
        if HAS_ORTOOLS and len(values) <= 1000:  # OR-Tools for smaller problems
            strategies.append(SolverStrategy.ORTOOLS)

        if len(values) <= 500:  # DP for medium problems
            strategies.append(SolverStrategy.DYNAMIC)

        strategies.append(SolverStrategy.GREEDY)  # Always available

        best_result = None

        for strategy in strategies:
            if strategy == SolverStrategy.ORTOOLS:
                result = self._solve_ortools(values, target, mode, constraints)
            elif strategy == SolverStrategy.DYNAMIC:
                result = self._solve_dynamic(values, target, mode, constraints)
            else:
                result = self._solve_greedy(values, target, mode, constraints)

            if result.success:
                if best_result is None or result.error < best_result.error:
                    best_result = result
                    best_result.strategy_used = SolverStrategy.HYBRID

                # If we found exact match, stop
                if result.error < 1e-10:
                    break

        if best_result is None:
            return SolverResult(
                success=False,
                selected_indices=[],
                selected_values=[],
                total_sum=0.0,
                target=target,
                error=abs(target),
                solve_time=0.0,
                strategy_used=SolverStrategy.HYBRID,
                mode_used=mode,
                solver_status="ALL_STRATEGIES_FAILED",
                message="All strategies failed to find a solution"
            )

        return best_result


def solve_subset_sum(
    values: List[float],
    target: float,
    mode: str = "closest",
    strategy: str = "ortools",
    scale_factor: int = 100,
    time_limit: float = 5.0,
    **kwargs
) -> SolverResult:
    """
    Convenience function for solving subset sum problems.

    Args:
        values: List of numbers to choose from
        target: Target sum to achieve
        mode: Solving mode ('closest', 'le', 'ge', 'exact', 'range')
        strategy: Solver strategy ('ortools', 'dynamic', 'greedy', 'hybrid')
        scale_factor: Factor to scale floats to integers
        time_limit: Maximum solving time in seconds
        **kwargs: Additional constraints (min_items, max_items, etc.)

    Returns:
        SolverResult with solution details
    """
    solver = SubsetSumSolver(scale_factor=scale_factor, time_limit=time_limit)

    # Parse mode and strategy
    solve_mode = SolveMode(mode.lower())
    solver_strategy = SolverStrategy(strategy.lower())

    # Build constraints
    constraints = SolverConstraints()
    if 'min_items' in kwargs:
        constraints.min_items = kwargs['min_items']
    if 'max_items' in kwargs:
        constraints.max_items = kwargs['max_items']
    if 'min_value' in kwargs:
        constraints.min_value = kwargs['min_value']
    if 'max_value' in kwargs:
        constraints.max_value = kwargs['max_value']
    if 'required_indices' in kwargs:
        constraints.required_indices = set(kwargs['required_indices'])
    if 'forbidden_indices' in kwargs:
        constraints.forbidden_indices = set(kwargs['forbidden_indices'])
    if 'tolerance' in kwargs:
        constraints.tolerance = kwargs['tolerance']

    return solver.solve(values, target, solve_mode, solver_strategy, constraints)