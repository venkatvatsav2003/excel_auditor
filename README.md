# Excel & CSV Data Auditor

A lightweight data quality auditing tool for CSV files. Detects duplicate rows, missing values, and column statistics.

## Ideology
Data quality should be verified before it reaches analytics or ML pipelines. A zero-dependency auditor enables quick validation at any stage of the data pipeline.

## Features
- Duplicate row detection
- Missing value counts per column
- Unique value cardinality
- String length statistics

## Usage
```bash
python auditor.py data.csv
```

## Example
```
File: inventory.csv
Total rows: 150
Columns: Product_ID, Name, Category, Price, Stock

[Product_ID]
  Missing: 0/150
  Unique values: 150
[Name]
  Missing: 2/150
  Unique values: 148
[Price]
  Missing: 5/150
  Unique values: 42

Duplicate rows: 3
```

## Design
See `DESIGN.md` for architecture, data quality dimensions, and design decisions.

## Dependencies
- Python 3 (stdlib only — `csv`, `collections`)
