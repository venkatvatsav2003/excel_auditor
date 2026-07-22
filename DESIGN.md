# Data Auditor — Design Document

## Problem Statement
Data quality issues silently corrupt analytics, reporting, and ML training. Most quality checks require pandas or costly ETL tools. Teams need a zero-dependency solution that can be embedded anywhere.

## Design Principles
1. **Zero external dependencies** — only Python stdlib
2. **Column-agnostic** — works with any CSV schema
3. **Deterministic** — same input, same output every time
4. **Pipeline-friendly** — JSON output for automation
5. **Human-readable** — HTML reports for stakeholders

## Architecture

```
┌─────────────┐
│  CSV Input  │  File, stdin, recursive glob
└──────┬──────┘
       ▼
┌─────────────┐
│  Parser     │  csv.Sniffer → Dialect detection
│             │  csv.DictReader → Structured rows
└──────┬──────┘
       ▼
┌─────────────────┐
│  Column Profiler │  Per-column analysis
│  ├ Type inference│
│  ├ Missing count │
│  ├ Unique count  │
│  ├ Statistics    │
│  └ Outlier detect│
└──────┬──────────┘
       ▼
┌─────────────────┐
│  Quality Scorer  │  Weighted composite score
└──────┬──────────┘
       ▼
┌─────────────────┐
│  Report Gen      │  HTML with embedded CSS
│                  │  JSON for automation
└─────────────────┘
```

## Data Quality Dimensions
1. **Completeness** — ratio of non-empty values per column
2. **Uniqueness** — absence of duplicate rows
3. **Outlier-Free** — values within 1.5×IQR of Q1/Q3

## Outlier Detection
Uses Tukey's fences:
- Lower bound: Q1 - 1.5 × IQR
- Upper bound: Q3 + 1.5 × IQR
- Values outside these bounds are flagged as outliers

## Limitations
- No Excel support without openpyxl
- String-based type inference — may misclassify mixed columns
- No cross-column validation (e.g., referential integrity)
