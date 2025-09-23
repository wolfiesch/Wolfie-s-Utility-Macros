# FuzzySum - Enhanced Subset Sum Solver

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**FuzzySum** is a powerful, feature-rich command-line tool and Python library for solving subset sum problems with multiple strategies, Excel integration, and beautiful terminal output.

## 🌟 Features

- **Multiple Solving Strategies**: OR-Tools, Dynamic Programming, Greedy, and Hybrid approaches
- **Rich Terminal Output**: Colorful, well-formatted CLI with progress bars and detailed results
- **Excel Integration**: Read from and write to Excel files with cell highlighting
- **Parallel Processing**: Solve multiple targets simultaneously with caching
- **Flexible Constraints**: Min/max items, value ranges, required/forbidden items
- **Configuration Management**: Save and load settings with named profiles
- **Multiple Output Formats**: Terminal, JSON, CSV, Markdown, HTML
- **Batch Processing**: Handle multiple targets from files
- **A1 Notation Support**: Excel-style cell references (A1, B5, etc.)

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd FuzzySum

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

```bash
# Solve for a single target
python subset_sum_cli.py --csv examples/sample_data.csv --target 500 --addr A1

# Use Excel file as input
python subset_sum_cli.py --excel data.xlsx --target 1000 --detailed

# Batch processing with multiple targets
python subset_sum_cli.py --batch examples/batch_targets.txt --csv examples/sample_data.csv

# Export results to JSON
python subset_sum_cli.py --csv data.csv --target 750 --export results.json
```

### Python Library Usage

```python
from fuzzy_sum import solve_subset_sum

# Basic solving
values = [125.50, 234.75, 89.25, 456.80, 178.90]
target = 400.0

result = solve_subset_sum(values, target)

if result.success:
    print(f"Selected: {result.selected_values}")
    print(f"Sum: {result.total_sum}")
    print(f"Error: {result.error}")
```

## 📊 Command Line Interface

### Input Options

- `--csv FILE`: Read data from CSV file
- `--excel FILE`: Read data from Excel file
- `--data FILE`: Auto-detect CSV or Excel format
- `--sheet NAME`: Specify Excel sheet name
- `--addr A1`: Starting cell address (e.g., A1, B5)

### Target Specification

- `--target VALUE`: Single target value
- `--batch FILE`: Multiple targets from file

### Solving Options

- `--mode {closest,le,ge,exact,range}`: Solving mode (default: closest)
- `--strategy {ortools,dynamic,greedy,hybrid}`: Solver strategy (default: ortools)
- `--time SECONDS`: Maximum solve time (default: 5.0)
- `--scale FACTOR`: Float-to-integer scale factor (default: 100)

### Constraints

- `--min-items N`: Minimum number of items in solution
- `--max-items N`: Maximum number of items in solution
- `--min-value V`: Minimum value of items to include
- `--max-value V`: Maximum value of items to include
- `--tolerance T`: Tolerance for range mode
- `--required I1 I2`: Required item indices
- `--forbidden I1 I2`: Forbidden item indices

### Output Options

- `--output {minimal,summary,detailed,verbose}`: Output verbosity
- `--format {terminal,plain,json}`: Output format
- `--export FILE`: Export results (.json, .csv, .md, .html)
- `--width N`: Terminal width
- `--no-progress`: Disable progress indicators

### Examples

```bash
# Find closest sum with constraints
python subset_sum_cli.py --csv data.csv --target 1000 \
  --min-items 3 --max-items 8 --min-value 50

# Use greedy strategy for fast results
python subset_sum_cli.py --excel data.xlsx --target 750 \
  --strategy greedy --time 1.0 --output minimal

# Solve for exact match with verbose output
python subset_sum_cli.py --csv data.csv --target 500 \
  --mode exact --output verbose

# Export detailed results to markdown
python subset_sum_cli.py --data values.csv --target 800 \
  --output detailed --export results.md

# Process batch targets in parallel
python subset_sum_cli.py --batch targets.txt --csv data.csv \
  --export batch_results.json
```

## 🔧 Configuration

FuzzySum supports configuration files for default settings:

```yaml
# .fuzzysum.yaml
solver:
  mode: "closest"
  strategy: "ortools"
  time_limit: 5.0

output:
  level: "summary"
  format: "terminal"
  show_progress: true

profiles:
  fast:
    solver:
      strategy: "greedy"
      time_limit: 1.0
  accurate:
    solver:
      time_limit: 30.0
    output:
      level: "detailed"
```

### Configuration Locations

FuzzySum searches for configuration files in:
1. `~/.fuzzysum.yaml` or `~/.fuzzysum.json`
2. `~/.config/fuzzysum/config.yaml`
3. `./.fuzzysum.yaml` (current directory)

## 📚 Library API

### Core Functions

```python
from fuzzy_sum import solve_subset_sum, SolveMode, SolverStrategy

# Basic solving
result = solve_subset_sum(
    values=[1, 2, 3, 4, 5],
    target=8,
    mode="closest",
    strategy="ortools"
)

# With constraints
result = solve_subset_sum(
    values=[1, 2, 3, 4, 5],
    target=8,
    min_items=2,
    max_items=4,
    min_value=1,
    max_value=10
)
```

### Parallel Solving

```python
from parallel_solver import solve_parallel_targets

values = [1, 2, 3, 4, 5]
targets = [5, 7, 10, 12]

results = solve_parallel_targets(
    values, targets,
    max_workers=4,
    enable_caching=True
)
```

### Excel Integration

```python
from excel_integration import read_excel_data, write_excel_results

# Read data from Excel
values = read_excel_data("data.xlsx", sheet_name="Data", start_address="A1")

# Solve problem
result = solve_subset_sum(values, target=1000)

# Write results back to Excel
write_excel_results(result, "results.xlsx", start_address="A1")
```

### Rich Output Formatting

```python
from cli_formatter import create_formatter

formatter = create_formatter(
    level="detailed",
    format_type="terminal",
    show_progress=True
)

formatter.print_result_summary(result, start_address="A1")
formatter.export_result(result)  # If export path configured
```

## 🧪 Testing

Run the comprehensive test suite:

```bash
# Run all tests
python test_suite.py

# Run performance benchmarks
python test_suite.py --benchmark

# Run specific test class
python -m unittest test_suite.TestSubsetSumSolver -v
```

## 🎯 Solving Modes

### Closest (`closest`)
Find the subset with sum closest to the target value (default mode).

### Less Than or Equal (`le`)
Find the subset with the largest sum that doesn't exceed the target.

### Greater Than or Equal (`ge`)
Find the subset with the smallest sum that meets or exceeds the target.

### Exact (`exact`)
Find a subset that sums exactly to the target value.

### Range (`range`)
Find a subset within a tolerance range of the target.

## ⚡ Solver Strategies

### OR-Tools (`ortools`)
Uses Google's OR-Tools constraint programming solver. Best for optimal solutions on medium-sized problems.

### Dynamic Programming (`dynamic`)
Classic dynamic programming approach. Good for smaller problems with guaranteed optimal solutions.

### Greedy (`greedy`)
Fast approximation algorithm. Best for large problems where speed is more important than optimality.

### Hybrid (`hybrid`)
Tries multiple strategies and returns the best result. Balances quality and performance.

## 📈 Performance Tips

1. **Use appropriate strategy**: OR-Tools for accuracy, Greedy for speed
2. **Set reasonable time limits**: Longer times allow better solutions
3. **Use constraints wisely**: They can speed up solving by reducing search space
4. **Enable caching**: For repeated solving with similar parameters
5. **Consider parallel processing**: For multiple targets or large batches

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- [Google OR-Tools](https://developers.google.com/optimization) for the constraint programming solver
- [Rich](https://github.com/Textualize/rich) for beautiful terminal output
- [Click](https://click.palletsprojects.com/) for the CLI framework
- [OpenPyXL](https://openpyxl.readthedocs.io/) for Excel integration

## 📞 Support

If you encounter any issues or have questions:

1. Check the [examples](examples/) directory for usage examples
2. Run the demo script: `python examples/demo.py`
3. Check the test suite for edge cases: `python test_suite.py`
4. Open an issue on GitHub

---

**Happy subset summing! 🎯**