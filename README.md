# Excel & CSV Data Auditor

![CI](https://github.com/venkatvatsav2003/excel_auditor/actions/workflows/ci.yml/badge.svg)
![Version](https://img.shields.io/badge/version-2.0.0-blue)
![Language](https://img.shields.io/badge/language-Python%20%2B%20Bash-blue)

An automated data quality auditing tool for CSV files. Detects duplicates, missing values, outliers, type mismatches, and computes a comprehensive data quality score — all with zero external dependencies.

## Features

- **Completeness Analysis** — Missing value counts and percentages per column
- **Duplicate Detection** — Identifies and counts exact duplicate rows
- **Type Inference** — Automatically detects integer, numeric, date, and string columns
- **Statistical Profiling** — Min, max, mean, median, stddev for numeric columns
- **Outlier Detection** — IQR-based (Tukey's method) with configurable threshold
- **String Analysis** — Average, min, and max string length per column
- **Top Values** — Most frequent values for categorical columns
- **Quality Scoring** — 0-100 completeness, duplicate-free, and overall scores
- **HTML Reports** — Self-contained HTML with professional styling
- **JSON Output** — Structured data for pipeline integration
- **Recursive Scanning** — Process all matching files in a directory tree

## Quick Start

```bash
# Single file
./audit.sh data.csv

# Multiple files
./audit.sh file1.csv file2.csv

# Recursive scan
./audit.sh -r -p "*.csv"

# Pipe from stdin
find . -name "*.csv" | ./audit.sh

# JSON output
./audit.sh data.csv --json
```

## Example Output

```
File: inventory.csv
Total rows: 150
Columns: Product_ID, Name, Category, Price, Stock

[Product_ID]
  Missing: 0/150
  Unique values: 148
  Type: integer
  Stats: min=1, max=150, mean=75.5

[Price]
  Missing: 5/150 (3.3%)
  Unique values: 42
  Type: numeric
  Stats: min=0.99, max=999.99, mean=49.95, outliers: 3

Duplicate rows: 3

Quality Scores:
  Completeness:     96.7%
  Duplicate-Free:  98.0%
  Outlier-Free:    95.0%
  Overall:         96.6%
```

## Project Structure

```
excel_auditor/
├── auditor.py              # Python auditing engine
├── audit.sh                # Bash orchestrator
├── config/profiles.yml     # Audit configuration
├── tests/
│   ├── test_auditor.py
│   └── fixtures/
│       ├── clean.csv       # Clean test data
│       └── dirty.csv       # Dirty test data
├── reports/                # Output directory
├── Dockerfile
├── Makefile
└── .github/workflows/
```

## Dependencies

- Python 3.8+ (stdlib only — `csv`, `json`, `math`, `collections`)
- Optional: `pyyaml` for config file support
