#!/usr/bin/env python3
"""
CLI: Solve subset-sum on a column exported from Excel and print A1 references.

Usage example:
  python subset_sum_cli.py ^
    --csv "C:\\Users\\wschoenberger\\FuzzySum\\Exports\\col_20250918_130455.csv" ^
    --target 12345.67 ^
    --addr B3 ^
    --mode closest ^
    --time 5 ^
    --scale 100

Output (to STDOUT):
  PICKS: B7,B12,B19
  BEST_SUM: 12345.00
  ERROR: 2.35
  COUNT: 3

Notes:
- --mode closest  -> minimize absolute error |sum - target|
- --mode le       -> choose sum <= target and minimize (target - sum)
- --scale 100     -> scale to ints (e.g., cents). Use 1 for integers.
"""

import argparse
import csv
import re
import sys
from pathlib import Path

# ---- OR-Tools import ----
try:
    from ortools.sat.python import cp_model
except Exception:
    sys.stderr.write("ERROR: ortools not installed. Try: pip install ortools\n")
    sys.exit(2)

# ---- A1 helpers ----
_A1_RE = re.compile(r"^([A-Za-z]+)(\d+)$")

def _col_to_index(col_letters: str) -> int:
    col_letters = col_letters.upper()
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def _index_to_col(idx: int) -> str:
    s = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s.append(chr(ord('A') + rem))
    return ''.join(reversed(s))

def _parse_a1(addr: str):
    m = _A1_RE.match(addr.strip())
    if not m:
        raise ValueError(f"Bad A1 address: {addr}")
    col_letters, row_str = m.groups()
    return _col_to_index(col_letters), int(row_str)

# ---- CSV read ----
def _read_numbers(csv_path: Path):
    vals = []
    # Accept commas and currency symbols; ignore junk rows.
    with csv_path.open(newline='', encoding='utf-8', errors='ignore') as f:
        r = csv.reader(f)
        for row in r:
            if not row:
                continue
            cell = str(row[0]).strip()
            if not cell:
                continue
            cleaned = cell.replace(',', '')
            # Strip common currency symbols
            for sym in ('$','€','£','¥'):
                cleaned = cleaned.replace(sym, '')
            try:
                vals.append(float(cleaned))
            except ValueError:
                # Non-numeric -> ignore
                continue
    return vals

# ---- OR-Tools model ----
def _solve_subset(values_int, target_int, mode="closest", time_limit=5.0):
    m = cp_model.CpModel()
    x = [m.NewBoolVar(f"x{i}") for i in range(len(values_int))]
    s = sum(v * x[i] for i, v in enumerate(values_int))

    if mode == "le":
        m.Add(s <= target_int)
        diff = target_int - s
        m.Minimize(diff)
    else:
        # minimize |s - target|
        abs_max = sum(abs(v) for v in values_int) or 0
        d = m.NewIntVar(0, abs_max, "d")
        m.Add(d >= s - target_int)
        m.Add(d >= target_int - s)
        m.Minimize(d)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(time_limit)
    solver.parameters.num_search_workers = 8

    status = solver.Solve(m)
    picks = [int(solver.Value(v)) for v in x]
    best_sum = sum(v for v, p in zip(values_int, picks) if p)
    return status, picks, best_sum

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True, help="Path to exported one-column CSV")
    ap.add_argument("--target", required=True, type=float, help="Target number (float)")
    ap.add_argument("--addr", required=True, help="A1 start address of the exported range (e.g., B3)")
    ap.add_argument("--mode", choices=["closest", "le"], default="closest")
    ap.add_argument("--time", type=float, default=5.0, help="Max seconds for the solver")
    ap.add_argument("--scale", type=int, default=100, help="Scale factor to convert to ints (100=cents)")
    args = ap.parse_args()

    csv_path = Path(args.csv)
    if not csv_path.exists():
        sys.stderr.write(f"ERROR: CSV not found: {csv_path}\n")
        sys.exit(1)

    values = _read_numbers(csv_path)
    if not values:
        print("PICKS:")
        print("BEST_SUM: 0")
        print("ERROR: NA")
        print("COUNT: 0")
        return

    # Scale to integers for robust solving (avoid float precision).
    scaled_vals = [int(round(v * args.scale)) for v in values]
    scaled_target = int(round(args.target * args.scale))

    _, picks, best_sum_scaled = _solve_subset(
        scaled_vals, scaled_target, args.mode, args.time
    )

    # Map picks back to original A1 references
    start_col_idx, start_row = _parse_a1(args.addr)
    col_letters = _index_to_col(start_col_idx)

    refs = []
    for i, p in enumerate(picks):
        if p:
            refs.append(f"{col_letters}{start_row + i}")

    best_sum = best_sum_scaled / args.scale
    error = abs(args.target - best_sum)

    # Final STDOUT (VBA captures and shows this; no re-import)
    print("PICKS:" + ("" if not refs else " " + ",".join(refs)))
    print(f"BEST_SUM: {best_sum:.10g}")
    print(f"ERROR: {error:.10g}")
    print(f"COUNT: {len(refs)}")

if __name__ == "__main__":
    main()
