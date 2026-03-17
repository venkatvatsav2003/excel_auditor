import pandas as pd
from pathlib import Path
from datetime import datetime
import argparse
import sys
import os
from tkinter import Tk, filedialog

# Add src to sys.path to ensure we can import the modules
sys.path.append(str(Path(__file__).parent / "src"))

from auditor import DataAuditor
from reporter import ReportGenerator

def choose_file():
    """Fallback GUI for file selection if no arguments are provided."""
    try:
        Tk().withdraw()
        return filedialog.askopenfilename(
            title="Select Excel or CSV file",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("All Files", "*.*")
            ]
        )
    except Exception:
        return None

def load_file(filepath):
    """Loads CSV or Excel file into a pandas DataFrame."""
    ext = Path(filepath).suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
    else:
        df = pd.read_excel(filepath)
    
    # Clean column names
    df.columns = df.columns.str.strip()
    return df

def find_class_column(df):
    """Attempts to find a suitable column for grouping (Class/CI Class)."""
    possible_names = [
        "class",
        "ci class",
        "configuration item class",
        "asset class",
        "ci_class",
        "asset_class"
    ]

    normalized_columns = {
        col.lower().replace("_", " ").strip(): col
        for col in df.columns
    }

    for name in possible_names:
        if name in normalized_columns:
            return normalized_columns[name]

    return None

def get_default_output_path(input_path, format="xlsx"):
    """Generates a default output path in the same directory as input or Downloads."""
    input_path = Path(input_path)
    today_date = datetime.now().strftime("%Y-%m-%d")
    output_name = f"{input_path.stem}_Audit_{today_date}.{format}"
    
    downloads_folder = Path.home() / "Downloads"
    if downloads_folder.exists():
        return downloads_folder / output_name
    return input_path.parent / output_name

def main():
    parser = argparse.ArgumentParser(description="Excel and CSV Data Auditor CLI tool.")
    parser.add_argument("input", nargs="?", help="Path to the Excel or CSV file to audit.")
    parser.add_argument("-o", "--output", help="Path to save the audit report.")
    parser.add_argument("-c", "--class-col", help="Column name to group data by (e.g., 'Class').")
    parser.add_argument("-f", "--format", choices=["xlsx", "html", "both"], default="xlsx", 
                        help="Output format of the report (default: xlsx).")
    parser.add_argument("--gui", action="store_true", help="Force use of file picker dialog.")

    args = parser.parse_args()

    # Determine input file
    input_file = args.input
    if args.gui or not input_file:
        input_file = choose_file()
        if not input_file:
            if not args.gui and not args.input:
                parser.print_help()
            return

    # Load data
    try:
        df = load_file(input_file)
    except Exception as e:
        print(f"Error loading file: {e}")
        return

    # Determine class column
    class_col = args.class_col
    if not class_col:
        class_col = find_class_column(df)
        if class_col:
            print(f"Automatically identified class column: {class_col}")
        else:
            print("No class column found for grouping. Proceeding with overall audit.")

    # Run audit
    print(f"Auditing data from {input_file}...")
    auditor = DataAuditor(df, class_col)
    auditor.run_audit()

    # Generate report(s)
    reporter = ReportGenerator(auditor, df)
    
    if args.format in ["xlsx", "both"]:
        out_xlsx = args.output if args.format == "xlsx" else None
        if not out_xlsx or Path(out_xlsx).suffix != ".xlsx":
            out_xlsx = get_default_output_path(input_file, "xlsx")
        
        print(f"Generating Excel report: {out_xlsx}")
        reporter.generate_excel(out_xlsx)
    
    if args.format in ["html", "both"]:
        out_html = args.output if args.format == "html" else None
        if not out_html or Path(out_html).suffix != ".html":
            out_html = get_default_output_path(input_file, "html")
        
        print(f"Generating HTML report: {out_html}")
        reporter.generate_html(out_html)

    print("Audit complete!")

if __name__ == "__main__":
    main()
