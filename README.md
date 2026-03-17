# Excel & CSV Data Auditor

This tool is designed to audit Excel and CSV data for common data quality issues such as duplicates, missing values, and data type mismatches. It provides detailed reports in Excel and HTML formats.

## Features

- **Duplicate Detection**: Identifies and highlights duplicate rows.
- **Missing Values**: Detects missing cells and provides counts per column.
- **Type Mismatches**: Identifies columns with inconsistent data types.
- **Outlier Detection**: Detects numeric outliers using the Interquartile Range (IQR) method.
- **Quality Scoring**: Calculates a data quality score for each category.
- **Multiple Formats**: Generates reports in Excel (.xlsx) and HTML formats.
- **CLI & GUI**: Supports both command-line arguments and a simple file picker interface.

## Installation

Ensure you have Python installed and install the required dependencies:

```bash
pip install pandas openpyxl numpy
```

## Usage

### Command Line Interface (CLI)

Run the tool by passing the input file as an argument:

```bash
python tool.py data.xlsx
```

#### Options:

- `input`: Path to the Excel or CSV file to audit.
- `-o`, `--output`: Path to save the audit report.
- `-c`, `--class-col`: Column name to group data by (e.g., 'Class').
- `-f`, `--format`: Output format (`xlsx`, `html`, `both`). Default is `xlsx`.
- `--gui`: Force use of file picker dialog.

#### Example:

```bash
python tool.py inventory.csv -c Category -f both
```

### GUI Mode

If you run the tool without any arguments, it will open a file picker dialog:

```bash
python tool.py
```

## Project Structure

- `tool.py`: Main entry point.
- `src/auditor.py`: Contains the `DataAuditor` class for data analysis.
- `src/reporter.py`: Contains the `ReportGenerator` class for report generation.
- `src/__init__.py`: Makes the `src` directory a Python package.
