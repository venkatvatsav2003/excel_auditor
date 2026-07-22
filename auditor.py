#!/usr/bin/env python3
import os
import csv
import sys
import json
import yaml
import math
import logging
import argparse
from pathlib import Path
from datetime import datetime
from collections import Counter, defaultdict
from typing import List, Dict, Any

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("auditor")


class ColumnProfile:
    def __init__(self, name: str, values: List[str], dtype: str = "string"):
        self.name = name
        self.raw_values = values
        self.cleaned = [v.strip() for v in values]
        self.non_empty = [v for v in self.cleaned if v]
        self.total = len(values)
        self.missing = self.total - len(self.non_empty)
        self.missing_pct = round(self.missing / self.total * 100, 2) if self.total else 0
        self.unique = len(set(self.non_empty))
        self.unique_pct = round(self.unique / self.total * 100, 2) if self.total else 0
        self.empty_to_null = sum(1 for v in self.cleaned if v.lower() in ('null', 'none', 'na', 'n/a', '-'))
        self.top_values = Counter(self.non_empty).most_common(5)
        self.dtype = self._infer_dtype()
        self.stats = self._compute_stats()
        self.outliers = []

    def _infer_dtype(self) -> str:
        ints = floats = 0
        for v in self.non_empty:
            vc = v.replace(',', '').replace('$', '').replace('%', '').replace('-', '')
            try: int(vc); ints += 1; continue
            except: pass
            try: float(vc); floats += 1; continue
            except: pass
        total_checked = len(self.non_empty) or 1
        if ints / total_checked > 0.8: return "integer"
        if (ints + floats) / total_checked > 0.8: return "numeric"
        if self._detect_date(): return "date"
        return "string"

    def _detect_date(self) -> bool:
        from datetime import datetime
        date_formats = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S"]
        hits = 0
        for v in self.non_empty[:50]:
            for fmt in date_formats:
                try: datetime.strptime(v, fmt); hits += 1; break
                except: pass
        return hits / max(len(self.non_empty), 1) > 0.5

    def _compute_stats(self) -> Dict[str, Any]:
        stats = {}
        if self.dtype in ("integer", "numeric"):
            nums = []
            for v in self.non_empty:
                try: nums.append(float(v.replace(',', '').replace('$', '').replace('%', '')))
                except: pass
            if nums:
                nums.sort()
                stats["min"] = round(min(nums), 2)
                stats["max"] = round(max(nums), 2)
                stats["mean"] = round(sum(nums) / len(nums), 2)
                stats["median"] = round(nums[len(nums)//2], 2)
                stats["stddev"] = round(math.sqrt(sum((x - stats["mean"])**2 for x in nums)/len(nums)), 2)
                q1 = nums[len(nums)//4]
                q3 = nums[3*len(nums)//4]
                iqr = q3 - q1
                lower = q1 - 1.5 * iqr
                upper = q3 + 1.5 * iqr
                self.outliers = [x for x in nums if x < lower or x > upper]
                stats["outliers"] = len(self.outliers)
                stats["outlier_pct"] = round(len(self.outliers)/len(nums)*100, 2) if nums else 0
        elif self.dtype == "string":
            lengths = [len(v) for v in self.non_empty]
            if lengths:
                stats["avg_length"] = round(sum(lengths)/len(lengths), 1)
                stats["min_length"] = min(lengths)
                stats["max_length"] = max(lengths)
        return stats

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "type": self.dtype,
            "total": self.total,
            "missing": self.missing,
            "missing_pct": self.missing_pct,
            "unique": self.unique,
            "unique_pct": self.unique_pct,
            "empty_as_null": self.empty_to_null,
            "top_values": [(v, c) for v, c in self.top_values],
            "stats": self.stats,
            "outlier_count": len(self.outliers),
        }


class DataAuditor:
    def __init__(self, config: dict = None):
        self.config = config or {}
        self.min_completeness = self.config.get("min_completeness", 80)
        self.max_duplicates_pct = self.config.get("max_duplicates_pct", 10)
        self.outlier_threshold = self.config.get("outlier_threshold", 1.5)

    def audit(self, filepath: str) -> dict:
        log.info(f"Auditing: {filepath}")
        ext = Path(filepath).suffix.lower()
        if ext == '.csv':
            return self._audit_csv(filepath)
        elif ext in ('.xlsx', '.xls'):
            log.warning("Excel support requires openpyxl. Use CSV for zero-dependency mode.")
            return {"error": "Excel files require openpyxl"}
        else:
            return {"error": f"Unsupported format: {ext}"}

    def _audit_csv(self, filepath: str) -> dict:
        with open(filepath, newline='', encoding='utf-8', errors='replace') as f:
            sample = f.read(4096)
            f.seek(0)
            dialect = csv.Sniffer().sniff(sample) if self.config.get("auto_detect", True) else csv.excel
            f.seek(0)
            reader = csv.DictReader(f, dialect=dialect)
            rows = list(reader)

        if not rows:
            return {"file": filepath, "error": "Empty file", "rows": 0}

        columns = reader.fieldnames or []
        data = {
            "file": filepath,
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "rows": len(rows),
            "columns": len(columns),
            "column_names": columns,
            "size_bytes": os.path.getsize(filepath),
        }

        col_profiles = {}
        for col in columns:
            values = [r.get(col, '') for r in rows]
            col_profiles[col] = ColumnProfile(col, values)

        profiles = {c: p.to_dict() for c, p in col_profiles.items()}
        data["profiles"] = profiles

        # Duplicate analysis
        full_rows = [tuple(r.get(c, '').strip() for c in columns) for r in rows]
        dup_counts = Counter(full_rows)
        duplicates = {i: c for i, (r, c) in enumerate(dup_counts.items()) if c > 1}
        data["total_duplicate_groups"] = sum(1 for c in duplicates.values())
        data["total_duplicate_rows"] = sum(c - 1 for c in duplicates.values())
        data["duplicate_pct"] = round(data["total_duplicate_rows"] / len(rows) * 100, 2) if rows else 0

        # Quality scoring
        completeness_scores = [p.missing_pct for p in col_profiles.values()]
        avg_completeness = 100 - (sum(completeness_scores) / len(completeness_scores)) if completeness_scores else 0
        dup_score = max(0, 100 - data["duplicate_pct"])
        outlier_count = sum(p.stats.get("outliers", 0) for p in col_profiles.values())
        outlier_score = max(0, 100 - outlier_count)

        data["quality_scores"] = {
            "completeness": round(avg_completeness, 1),
            "duplicate_free": round(dup_score, 1),
            "outlier_free": round(outlier_score, 1),
            "overall": round((avg_completeness + dup_score + outlier_score) / 3, 1),
        }

        data["issues"] = []
        if data["total_duplicate_rows"] > 0:
            data["issues"].append(f"{data['total_duplicate_rows']} duplicate rows ({data['duplicate_pct']}%)")
        for c, p in col_profiles.items():
            if p.missing_pct > (100 - self.min_completeness):
                data["issues"].append(f"Column '{c}': {p.missing_pct}% missing")
            if p.stats.get("outliers", 0) > 0:
                data["issues"].append(f"Column '{c}': {p.stats['outliers']} outliers detected")

        log.info(f"Audit complete: {len(rows)} rows, {len(columns)} columns, score={data['quality_scores']['overall']}")
        return data


def generate_html_report(data: dict, output: str):
    html = """<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Data Audit Report - {file}</title>
<style>
body{{font-family:-apple-system,sans-serif;max-width:1200px;margin:40px auto;padding:20px;background:#0d1117;color:#c9d1d9}}
h1,h2{{color:#58a6ff}}
.summary{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin:20px 0}}
.card{{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:16px}}
.card .num{{font-size:24px;font-weight:700;color:#58a6ff}}
.card .label{{color:#8b949e;font-size:13px}}
.profile{{background:#161b22;border:1px solid #30363d;border-radius:8px;padding:16px;margin:12px 0}}
.profile .stat{{display:inline-block;margin:4px 12px 4px 0;color:#8b949e}}
table{{width:100%;border-collapse:collapse;margin:12px 0}}
th,td{{padding:8px 12px;text-align:left;border-bottom:1px solid #30363d;font-size:14px}}
th{{color:#58a6ff;font-weight:600}}
.issue{{color:#f85149;margin:4px 0}}
</style></head><body>
<h1>Data Quality Audit Report</h1>
<p>File: {file} | {timestamp} | {rows} rows</p>
<div class="summary">
<div class="card"><div class="num">{rows}</div><div class="label">Rows</div></div>
<div class="card"><div class="num">{cols}</div><div class="label">Columns</div></div>
<div class="card"><div class="num">{overall}</div><div class="label">Quality Score</div></div>
<div class="card"><div class="num">{dups}</div><div class="label">Duplicates</div></div>
</div>"""
    qs = data.get("quality_scores", {})
    html = html.format(
        file=data["file"], timestamp=data["timestamp"][:19],
        rows=data["rows"], cols=data["columns"],
        overall=qs.get("overall", "N/A"),
        dups=data.get("total_duplicate_rows", 0)
    )

    html += "<h2>Column Profiles</h2>"
    for cname, prof in data.get("profiles", {}).items():
        html += f'<div class="profile"><h3>{cname} <span style="color:#8b949e;font-weight:400">({prof["type"]})</span></h3>'
        html += f'<div class="stat">Missing: {prof["missing_pct"]}%</div>'
        html += f'<div class="stat">Unique: {prof["unique_pct"]}%</div>'
        html += f'<div class="stat">Top: {", ".join(v[0] for v in prof["top_values"][:3])}</div>'
        if prof.get("stats"):
            s = prof["stats"]
            if "mean" in s: html += f'<div class="stat">Mean: {s["mean"]}</div>'
            if "min" in s: html += f'<div class="stat">Min: {s["min"]}</div>'
            if "max" in s: html += f'<div class="stat">Max: {s["max"]}</div>'
            if "outliers" in s: html += f'<div class="stat">Outliers: {s["outliers"]}</div>'
            if "avg_length" in s: html += f'<div class="stat">Avg length: {s["avg_length"]}</div>'
        if prof.get("outlier_count", 0) > 0:
            html += f'<div class="issue">⚠ {prof["outlier_count"]} outliers detected</div>'
        html += "</div>"

    if data.get("issues"):
        html += "<h2>Issues</h2>"
        for issue in data["issues"]:
            html += f'<div class="issue">⚠ {issue}</div>'

    html += "</body></html>"
    Path(output).write_text(html, encoding='utf-8')
    log.info(f"HTML report: {output}")


def main():
    parser = argparse.ArgumentParser(description="Data Auditor — CSV/Excel Data Quality Analysis")
    parser.add_argument("files", nargs="*", help="Files to audit")
    parser.add_argument("-c", "--config", default="config/profiles.yml", help="Config file")
    parser.add_argument("-o", "--output", help="Output directory (default: reports/)")
    parser.add_argument("-f", "--format", choices=["json", "html", "both"], default="both")
    parser.add_argument("-r", "--recursive", action="store_true", help="Scan directory recursively")
    parser.add_argument("-p", "--pattern", default="*.csv", help="File pattern for recursive (default: *.csv)")
    parser.add_argument("--json", action="store_true", help="JSON output to stdout")
    args = parser.parse_args()

    config = {}
    if Path(args.config).exists():
        config = yaml.safe_load(Path(args.config).read_text()) or {}

    auditor = DataAuditor(config)
    output_dir = args.output or "reports"
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    files = []
    if args.recursive:
        for f in Path(".").rglob(args.pattern):
            files.append(str(f))
    elif args.files:
        files = args.files
    else:
        files = [f for f in sys.stdin.read().splitlines() if f.strip()] if not sys.stdin.isatty() else []
        if not files:
            log.info("No files specified. Use: auditor.py file.csv")
            parser.print_help()
            return

    for filepath in files:
        if not Path(filepath).exists():
            log.warning(f"File not found: {filepath}")
            continue
        data = auditor.audit(filepath)
        if "error" in data:
            log.error(f"{filepath}: {data['error']}")
            continue
        if args.json:
            print(json.dumps(data, indent=2, default=str))
            continue
        if args.format in ("json", "both"):
            report_path = Path(output_dir) / f"{Path(filepath).stem}_audit.json"
            report_path.write_text(json.dumps(data, indent=2, default=str))
            log.info(f"JSON report: {report_path}")
        if args.format in ("html", "both"):
            report_path = Path(output_dir) / f"{Path(filepath).stem}_audit.html"
            generate_html_report(data, str(report_path))


if __name__ == "__main__":
    main()
