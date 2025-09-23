#!/usr/bin/env python3
"""
Parallel Solver: Performance optimizations and parallel processing for FuzzySum.

This module provides parallel processing capabilities, caching, and performance
optimizations for large-scale subset sum solving operations.
"""

import time
import threading
import multiprocessing as mp
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
from typing import List, Dict, Tuple, Optional, Callable, Any, Union
from dataclasses import dataclass
import hashlib
import pickle
import logging
from pathlib import Path

try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    HAS_NUMPY = False

from fuzzy_sum import (
    SubsetSumSolver, SolverResult, SolveMode, SolverStrategy,
    SolverConstraints, solve_subset_sum
)


@dataclass
class ParallelConfig:
    """Configuration for parallel processing."""
    max_workers: Optional[int] = None
    use_processes: bool = False  # Use processes instead of threads
    chunk_size: int = 10
    timeout: Optional[float] = None
    enable_caching: bool = True
    cache_dir: Optional[Path] = None


@dataclass
class BatchTask:
    """Single task in a batch operation."""
    task_id: int
    values: List[float]
    target: float
    mode: SolveMode
    strategy: SolverStrategy
    constraints: Optional[SolverConstraints] = None
    scale_factor: int = 100
    time_limit: float = 5.0


@dataclass
class BatchResult:
    """Result from a batch task."""
    task_id: int
    target: float
    result: SolverResult
    cache_hit: bool = False


class ResultCache:
    """Cache for solver results to avoid recomputation."""

    def __init__(self, cache_dir: Optional[Path] = None, max_size: int = 1000):
        """
        Initialize result cache.

        Args:
            cache_dir: Directory for persistent cache (None for memory-only)
            max_size: Maximum number of entries in memory cache
        """
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self.max_size = max_size
        self.memory_cache: Dict[str, SolverResult] = {}
        self.access_times: Dict[str, float] = {}
        self.logger = logging.getLogger(__name__)

        if self.cache_dir:
            self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _generate_key(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        strategy: SolverStrategy,
        constraints: Optional[SolverConstraints],
        scale_factor: int,
        time_limit: float
    ) -> str:
        """Generate cache key for the given parameters."""
        # Create a hashable representation of the inputs
        key_data = {
            'values_hash': hashlib.md5(str(sorted(values)).encode()).hexdigest(),
            'target': target,
            'mode': mode.value,
            'strategy': strategy.value,
            'scale_factor': scale_factor,
            'time_limit': time_limit
        }

        if constraints:
            key_data['constraints'] = {
                'min_items': constraints.min_items,
                'max_items': constraints.max_items,
                'min_value': constraints.min_value,
                'max_value': constraints.max_value,
                'required_indices': tuple(sorted(constraints.required_indices)) if constraints.required_indices else None,
                'forbidden_indices': tuple(sorted(constraints.forbidden_indices)) if constraints.forbidden_indices else None,
                'tolerance': constraints.tolerance
            }

        # Generate hash of the serialized data
        serialized = pickle.dumps(key_data, protocol=pickle.HIGHEST_PROTOCOL)
        return hashlib.sha256(serialized).hexdigest()

    def get(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        strategy: SolverStrategy,
        constraints: Optional[SolverConstraints] = None,
        scale_factor: int = 100,
        time_limit: float = 5.0
    ) -> Optional[SolverResult]:
        """Get cached result if available."""
        key = self._generate_key(values, target, mode, strategy, constraints, scale_factor, time_limit)

        # Check memory cache first
        if key in self.memory_cache:
            self.access_times[key] = time.time()
            self.logger.debug(f"Cache hit (memory): {key[:8]}...")
            return self.memory_cache[key]

        # Check persistent cache
        if self.cache_dir:
            cache_file = self.cache_dir / f"{key}.pkl"
            if cache_file.exists():
                try:
                    with open(cache_file, 'rb') as f:
                        result = pickle.load(f)

                    # Add to memory cache
                    self._add_to_memory_cache(key, result)
                    self.logger.debug(f"Cache hit (disk): {key[:8]}...")
                    return result
                except Exception as e:
                    self.logger.warning(f"Failed to load cache file {cache_file}: {e}")

        return None

    def put(
        self,
        values: List[float],
        target: float,
        mode: SolveMode,
        strategy: SolverStrategy,
        result: SolverResult,
        constraints: Optional[SolverConstraints] = None,
        scale_factor: int = 100,
        time_limit: float = 5.0
    ) -> None:
        """Store result in cache."""
        key = self._generate_key(values, target, mode, strategy, constraints, scale_factor, time_limit)

        # Add to memory cache
        self._add_to_memory_cache(key, result)

        # Save to persistent cache
        if self.cache_dir:
            cache_file = self.cache_dir / f"{key}.pkl"
            try:
                with open(cache_file, 'wb') as f:
                    pickle.dump(result, f, protocol=pickle.HIGHEST_PROTOCOL)
                self.logger.debug(f"Cache stored: {key[:8]}...")
            except Exception as e:
                self.logger.warning(f"Failed to save cache file {cache_file}: {e}")

    def _add_to_memory_cache(self, key: str, result: SolverResult) -> None:
        """Add result to memory cache with LRU eviction."""
        # Remove oldest entries if cache is full
        while len(self.memory_cache) >= self.max_size:
            oldest_key = min(self.access_times.keys(), key=self.access_times.get)
            del self.memory_cache[oldest_key]
            del self.access_times[oldest_key]

        self.memory_cache[key] = result
        self.access_times[key] = time.time()

    def clear(self) -> None:
        """Clear all cached results."""
        self.memory_cache.clear()
        self.access_times.clear()

        if self.cache_dir and self.cache_dir.exists():
            for cache_file in self.cache_dir.glob("*.pkl"):
                try:
                    cache_file.unlink()
                except Exception as e:
                    self.logger.warning(f"Failed to delete cache file {cache_file}: {e}")

    def stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        disk_files = 0
        if self.cache_dir and self.cache_dir.exists():
            disk_files = len(list(self.cache_dir.glob("*.pkl")))

        return {
            'memory_entries': len(self.memory_cache),
            'disk_entries': disk_files,
            'max_size': self.max_size,
            'cache_dir': str(self.cache_dir) if self.cache_dir else None
        }


class ParallelSolver:
    """Parallel subset sum solver for batch operations."""

    def __init__(self, config: ParallelConfig = None):
        """Initialize parallel solver."""
        self.config = config or ParallelConfig()
        self.cache = ResultCache(
            cache_dir=self.config.cache_dir,
            max_size=1000
        ) if self.config.enable_caching else None
        self.logger = logging.getLogger(__name__)

    def solve_batch(
        self,
        tasks: List[BatchTask],
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> List[BatchResult]:
        """
        Solve multiple subset sum problems in parallel.

        Args:
            tasks: List of batch tasks to solve
            progress_callback: Optional callback for progress updates

        Returns:
            List of batch results
        """
        if not tasks:
            return []

        self.logger.info(f"Starting batch solve with {len(tasks)} tasks")
        start_time = time.time()

        results = []
        completed = 0

        if self.config.use_processes:
            results = self._solve_with_processes(tasks, progress_callback)
        else:
            results = self._solve_with_threads(tasks, progress_callback)

        total_time = time.time() - start_time
        cache_hits = sum(1 for r in results if r.cache_hit)

        self.logger.info(
            f"Batch complete: {len(results)} tasks in {total_time:.2f}s "
            f"({cache_hits} cache hits)"
        )

        return results

    def _solve_with_threads(
        self,
        tasks: List[BatchTask],
        progress_callback: Optional[Callable[[int, int], None]]
    ) -> List[BatchResult]:
        """Solve tasks using thread pool."""
        results = []
        max_workers = self.config.max_workers or min(32, len(tasks))

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_task = {
                executor.submit(self._solve_single_task, task): task
                for task in tasks
            }

            # Collect results as they complete
            for future in as_completed(future_to_task, timeout=self.config.timeout):
                try:
                    result = future.result()
                    results.append(result)

                    if progress_callback:
                        progress_callback(len(results), len(tasks))

                except Exception as e:
                    task = future_to_task[future]
                    self.logger.error(f"Task {task.task_id} failed: {e}")
                    # Create failed result
                    failed_result = SolverResult(
                        success=False,
                        selected_indices=[],
                        selected_values=[],
                        total_sum=0.0,
                        target=task.target,
                        error=abs(task.target),
                        solve_time=0.0,
                        strategy_used=task.strategy,
                        mode_used=task.mode,
                        solver_status="TASK_FAILED",
                        message=str(e)
                    )
                    results.append(BatchResult(
                        task_id=task.task_id,
                        target=task.target,
                        result=failed_result,
                        cache_hit=False
                    ))

        # Sort results by task_id to maintain order
        results.sort(key=lambda x: x.task_id)
        return results

    def _solve_with_processes(
        self,
        tasks: List[BatchTask],
        progress_callback: Optional[Callable[[int, int], None]]
    ) -> List[BatchResult]:
        """Solve tasks using process pool."""
        results = []
        max_workers = self.config.max_workers or min(mp.cpu_count(), len(tasks))

        # Note: Process pool doesn't share cache, so we disable caching for this method
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_task = {
                executor.submit(_solve_task_worker, task): task
                for task in tasks
            }

            # Collect results as they complete
            for future in as_completed(future_to_task, timeout=self.config.timeout):
                try:
                    result = future.result()
                    results.append(result)

                    if progress_callback:
                        progress_callback(len(results), len(tasks))

                except Exception as e:
                    task = future_to_task[future]
                    self.logger.error(f"Task {task.task_id} failed: {e}")
                    # Create failed result
                    failed_result = SolverResult(
                        success=False,
                        selected_indices=[],
                        selected_values=[],
                        total_sum=0.0,
                        target=task.target,
                        error=abs(task.target),
                        solve_time=0.0,
                        strategy_used=task.strategy,
                        mode_used=task.mode,
                        solver_status="TASK_FAILED",
                        message=str(e)
                    )
                    results.append(BatchResult(
                        task_id=task.task_id,
                        target=task.target,
                        result=failed_result,
                        cache_hit=False
                    ))

        # Sort results by task_id to maintain order
        results.sort(key=lambda x: x.task_id)
        return results

    def _solve_single_task(self, task: BatchTask) -> BatchResult:
        """Solve a single task with caching."""
        # Check cache first
        cache_hit = False
        if self.cache:
            cached_result = self.cache.get(
                task.values, task.target, task.mode, task.strategy,
                task.constraints, task.scale_factor, task.time_limit
            )
            if cached_result:
                return BatchResult(
                    task_id=task.task_id,
                    target=task.target,
                    result=cached_result,
                    cache_hit=True
                )

        # Solve the problem
        result = solve_subset_sum(
            values=task.values,
            target=task.target,
            mode=task.mode.value,
            strategy=task.strategy.value,
            scale_factor=task.scale_factor,
            time_limit=task.time_limit,
            min_items=task.constraints.min_items if task.constraints else None,
            max_items=task.constraints.max_items if task.constraints else None,
            min_value=task.constraints.min_value if task.constraints else None,
            max_value=task.constraints.max_value if task.constraints else None,
            required_indices=task.constraints.required_indices if task.constraints else None,
            forbidden_indices=task.constraints.forbidden_indices if task.constraints else None,
            tolerance=task.constraints.tolerance if task.constraints else None
        )

        # Cache the result
        if self.cache:
            self.cache.put(
                task.values, task.target, task.mode, task.strategy, result,
                task.constraints, task.scale_factor, task.time_limit
            )

        return BatchResult(
            task_id=task.task_id,
            target=task.target,
            result=result,
            cache_hit=False
        )

    def solve_multiple_targets(
        self,
        values: List[float],
        targets: List[float],
        mode: Union[str, SolveMode] = SolveMode.CLOSEST,
        strategy: Union[str, SolverStrategy] = SolverStrategy.ORTOOLS,
        constraints: Optional[SolverConstraints] = None,
        scale_factor: int = 100,
        time_limit: float = 5.0,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> List[BatchResult]:
        """
        Solve subset sum for multiple targets using the same values.

        Args:
            values: List of numbers to choose from
            targets: List of target sums
            mode: Solving mode
            strategy: Solver strategy
            constraints: Additional constraints
            scale_factor: Scale factor for float-to-int conversion
            time_limit: Maximum time per solve
            progress_callback: Optional progress callback

        Returns:
            List of batch results
        """
        # Convert string modes/strategies to enums
        if isinstance(mode, str):
            mode = SolveMode(mode.lower())
        if isinstance(strategy, str):
            strategy = SolverStrategy(strategy.lower())

        # Create batch tasks
        tasks = [
            BatchTask(
                task_id=i,
                values=values,
                target=target,
                mode=mode,
                strategy=strategy,
                constraints=constraints,
                scale_factor=scale_factor,
                time_limit=time_limit
            )
            for i, target in enumerate(targets)
        ]

        return self.solve_batch(tasks, progress_callback)

    def solve_multiple_strategies(
        self,
        values: List[float],
        target: float,
        strategies: List[Union[str, SolverStrategy]],
        mode: Union[str, SolveMode] = SolveMode.CLOSEST,
        constraints: Optional[SolverConstraints] = None,
        scale_factor: int = 100,
        time_limit: float = 5.0
    ) -> Dict[str, BatchResult]:
        """
        Solve the same problem with multiple strategies and return the best result.

        Args:
            values: List of numbers to choose from
            target: Target sum
            strategies: List of strategies to try
            mode: Solving mode
            constraints: Additional constraints
            scale_factor: Scale factor for float-to-int conversion
            time_limit: Maximum time per solve

        Returns:
            Dictionary mapping strategy names to results
        """
        # Convert string mode to enum
        if isinstance(mode, str):
            mode = SolveMode(mode.lower())

        # Create tasks for each strategy
        tasks = []
        for i, strategy in enumerate(strategies):
            if isinstance(strategy, str):
                strategy = SolverStrategy(strategy.lower())

            tasks.append(BatchTask(
                task_id=i,
                values=values,
                target=target,
                mode=mode,
                strategy=strategy,
                constraints=constraints,
                scale_factor=scale_factor,
                time_limit=time_limit
            ))

        # Solve all strategies
        results = self.solve_batch(tasks)

        # Return results mapped by strategy name
        strategy_results = {}
        for task, result in zip(tasks, results):
            strategy_results[task.strategy.value] = result

        return strategy_results

    def get_cache_stats(self) -> Optional[Dict[str, Any]]:
        """Get cache statistics."""
        return self.cache.stats() if self.cache else None

    def clear_cache(self) -> None:
        """Clear the result cache."""
        if self.cache:
            self.cache.clear()


def _solve_task_worker(task: BatchTask) -> BatchResult:
    """Worker function for process-based parallel solving."""
    # This function runs in a separate process, so no shared cache
    result = solve_subset_sum(
        values=task.values,
        target=task.target,
        mode=task.mode.value,
        strategy=task.strategy.value,
        scale_factor=task.scale_factor,
        time_limit=task.time_limit,
        min_items=task.constraints.min_items if task.constraints else None,
        max_items=task.constraints.max_items if task.constraints else None,
        min_value=task.constraints.min_value if task.constraints else None,
        max_value=task.constraints.max_value if task.constraints else None,
        required_indices=task.constraints.required_indices if task.constraints else None,
        forbidden_indices=task.constraints.forbidden_indices if task.constraints else None,
        tolerance=task.constraints.tolerance if task.constraints else None
    )

    return BatchResult(
        task_id=task.task_id,
        target=task.target,
        result=result,
        cache_hit=False
    )


def solve_parallel_targets(
    values: List[float],
    targets: List[float],
    mode: str = "closest",
    strategy: str = "ortools",
    max_workers: Optional[int] = None,
    use_processes: bool = False,
    enable_caching: bool = True,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    **kwargs
) -> List[BatchResult]:
    """
    Convenience function for parallel solving of multiple targets.

    Args:
        values: List of numbers to choose from
        targets: List of target sums
        mode: Solving mode
        strategy: Solver strategy
        max_workers: Maximum number of parallel workers
        use_processes: Use processes instead of threads
        enable_caching: Enable result caching
        progress_callback: Optional progress callback
        **kwargs: Additional solver parameters

    Returns:
        List of batch results
    """
    config = ParallelConfig(
        max_workers=max_workers,
        use_processes=use_processes,
        enable_caching=enable_caching
    )

    solver = ParallelSolver(config)

    # Build constraints from kwargs
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

    return solver.solve_multiple_targets(
        values=values,
        targets=targets,
        mode=mode,
        strategy=strategy,
        constraints=constraints,
        scale_factor=kwargs.get('scale_factor', 100),
        time_limit=kwargs.get('time_limit', 5.0),
        progress_callback=progress_callback
    )