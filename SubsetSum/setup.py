#!/usr/bin/env python3
"""
Setup script for FuzzySum - Enhanced Subset Sum Solver.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
readme_path = Path(__file__).parent / "README.md"
if readme_path.exists():
    with open(readme_path, "r", encoding="utf-8") as f:
        long_description = f.read()
else:
    long_description = "Enhanced subset sum solver with rich CLI output and Excel integration"

# Read requirements
requirements_path = Path(__file__).parent / "requirements.txt"
if requirements_path.exists():
    with open(requirements_path, "r", encoding="utf-8") as f:
        requirements = [
            line.strip()
            for line in f
            if line.strip() and not line.startswith("#")
        ]
else:
    requirements = [
        "ortools>=9.14.0",
        "pandas>=2.0.0",
        "numpy>=1.20.0",
        "openpyxl>=3.1.0",
        "rich>=13.0.0",
        "click>=8.1.0",
        "colorama>=0.4.6",
        "tabulate>=0.9.0",
        "tqdm>=4.65.0",
        "pyyaml>=6.0",
    ]

setup(
    name="fuzzysum",
    version="1.0.0",
    author="FuzzySum Development Team",
    author_email="fuzzysum@example.com",
    description="Enhanced subset sum solver with rich CLI output and Excel integration",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/example/fuzzysum",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: Financial and Insurance Industry",
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Scientific/Engineering :: Mathematics",
        "Topic :: Office/Business :: Financial",
        "Topic :: Utilities",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
        ],
        "performance": [
            "numba>=0.57.0",
        ],
        "full": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "numba>=0.57.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "fuzzysum=subset_sum_cli:main",
            "fuzzysum-demo=examples.demo:main",
            "fuzzysum-test=test_suite:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": [
            "*.yaml",
            "*.yml",
            "*.json",
            "*.csv",
            "*.txt",
            "*.md",
            "examples/*",
        ],
    },
    project_urls={
        "Bug Reports": "https://github.com/example/fuzzysum/issues",
        "Source": "https://github.com/example/fuzzysum",
        "Documentation": "https://github.com/example/fuzzysum/blob/main/README.md",
    },
    keywords=[
        "subset-sum",
        "optimization",
        "combinatorial",
        "excel",
        "cli",
        "ortools",
        "solver",
        "mathematics",
        "finance",
        "accounting",
    ],
)