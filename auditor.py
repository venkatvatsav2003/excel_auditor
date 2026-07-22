#!/usr/bin/env python3
import csv
import sys
from collections import Counter

def audit(filename):
    with open(filename, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    print(f"File: {filename}")
    print(f"Total rows: {len(rows)}")
    print(f"Columns: {', '.join(reader.fieldnames)}")
    print()

    for col in reader.fieldnames:
        values = [r.get(col, '').strip() for r in rows]
        missing = sum(1 for v in values if not v)
        unique = len(set(v for v in values if v))
        print(f"[{col}]")
        print(f"  Missing: {missing}/{len(rows)}")
        print(f"  Unique values: {unique}")

    full_rows = [tuple(r[col].strip() for col in reader.fieldnames) for r in rows]
    dup_count = sum(c - 1 for r, c in Counter(full_rows).items() if c > 1)
    print(f"\nDuplicate rows: {dup_count}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Usage: python {sys.argv[0]} <file.csv>")
        sys.exit(1)
    audit(sys.argv[1])
