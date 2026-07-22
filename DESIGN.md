# Excel & CSV Data Auditor — Design Document

## Problem Statement
Data quality issues cost organizations millions annually. Duplicates, missing values, and inconsistent formatting silently corrupt analytics, reporting, and ML training pipelines. Most teams lack a quick, automated way to assess data health before processing.

## Core Ideology
Data quality should be checked before processing, not after. A lightweight, zero-dependency auditor enables teams to validate data at the pipeline's edge without complex infrastructure.

## Architecture

```
CSV/Excel Input
    |
    v
[Parsing Layer] --> csv.DictReader / openpyxl
    |
    v
[Analysis Engine]
    |-- Duplicate Detection (Counter-based)
    |-- Missing Value Analysis (column-wise)
    |-- Uniqueness Analysis (cardinality)
    |-- Length Statistics (avg, min, max)
    |
    v
[Report Generator] --> Console summary (future: HTML/XLSX)
```

## Key Design Decisions
- **Built-in `csv` module only** — avoids pandas dependency for basic audits
- **Column-agnostic** — works with any CSV schema without configuration
- **Deterministic output** — same input always produces same analysis

## Data Quality Dimensions Assessed
1. **Completeness** — missing/null values per column
2. **Uniqueness** — cardinality and duplicate rows
3. **Consistency** — string length distribution as a proxy for format issues

## Threat Model
- **Input**: Malformed CSVs, encoding issues, injection via crafted content
- **Mitigations**: UTF-8 encoding, field stripping, no eval/exec

## Limitations & Future Work
- Numeric outlier detection (IQR) — requires numpy or manual implementation
- Automated report generation (HTML/PDF)
- Schema inference and type detection
- Integration with data pipeline tools (Airflow, dbt)

## Use Cases
- Pre-processing validation for ETL pipelines
- Ad-hoc data exploration for analysts
- CI/CD data quality gates in MLOps workflows
