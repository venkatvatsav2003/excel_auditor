# Data Auditor

![CI](https://github.com/venkatvatsav2003/excel_auditor/actions/workflows/ci.yml/badge.svg)
![Version](https://img.shields.io/badge/version-3.0.0-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)

**Zero-dependency CSV data quality auditing — duplicates, missing values, type inference, outliers, and quality scoring.**

## Install & Run

```bash
# One-liner (stdlib only, no install needed)
python3 auditor.py data.csv

# Or pip install
pip install data-auditor && data-auditor data.csv

# Clone and run
git clone https://github.com/venkatvatsav2003/excel_auditor.git
cd excel_auditor && pip install -r requirements.txt
./audit.sh data.csv

# Docker
docker run -v $(pwd):/auditor/data ghcr.io/venkatvatsav2003/data-auditor data.csv
```

## Features

- **Zero Dependencies** — uses only Python standard library (`csv`, `json`, `math`, `collections`)
- **Type Inference** — automatically detects integer, numeric, date, and string columns
- **Statistical Profiling** — min, max, mean, median, stddev, quartiles
- **Outlier Detection** — IQR-based (Tukey's method)
- **Quality Scoring** — completeness, duplicate-free, outlier-free scores (0-100)
- **HTML Reports** — self-contained HTML with professional styling
- **Batch Processing** — recursive scanning with glob patterns
- **Custom Rules** — YAML-based validation rules

## Quick Start

```bash
# Single file
./audit.sh data.csv

# Multiple files
./audit.sh file1.csv file2.csv

# Recursive scan
./audit.sh -r -p "*.csv"

# JSON output
./audit.sh data.csv --json

# Output to custom directory
./audit.sh data.csv -o ./reports
```

## Example

```
File: inventory.csv
Rows: 150, Columns: 5
Quality Score: 96.6%

[Price] Missing: 3.3% | Outliers: 3
[Name]  Missing: 0.0% | Unique: 148
Duplicates: 3 rows
```

## Project Structure

```
excel_auditor/
├── auditor.py              # Python engine
├── audit.sh                # Bash launcher
├── pyproject.toml          # pip install
├── docker-compose.yml      # Docker one-command
├── .env.example            # Config template
├── config/profiles.yml     # Audit config
├── tests/fixtures/         # Sample data
├── reports/                # Output directory
├── Dockerfile
└── Makefile
```
