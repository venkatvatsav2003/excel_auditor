import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import json

class DataAuditor:
    def __init__(self, df, class_col=None):
        self.df = df
        self.class_col = class_col
        self.results = {}
        self.summary = []

    def run_audit(self):
        """Run all audit checks."""
        if self.class_col:
            for cls, group in self.df.groupby(self.class_col):
                self.results[cls] = self._audit_group(group, cls)
        else:
            self.results["All"] = self._audit_group(self.df, "All")
        
        self._generate_overall_summary()
        return self.results

    def _audit_group(self, group, cls_name):
        group = group.copy()
        
        # 1. Duplicate detection
        duplicates = group.duplicated(keep=False)
        duplicate_count = duplicates.sum()
        
        # 2. Missing values detection
        missing_info = {}
        for col in group.columns:
            missing = group[col].isna() | (group[col].astype(str).str.strip() == "")
            count = missing.sum()
            if count > 0:
                missing_info[col] = int(count)
        
        # 3. Data type mismatches
        type_mismatches = self._detect_type_mismatches(group)
        
        # 4. Outlier detection (for numeric columns)
        outliers = self._detect_outliers(group)
        
        # 5. Incomplete rows
        incomplete_rows = group.isna().any(axis=1).sum()
        
        return {
            "name": cls_name,
            "total_rows": len(group),
            "duplicate_count": int(duplicate_count),
            "missing_values": missing_info,
            "type_mismatches": type_mismatches,
            "outliers": outliers,
            "incomplete_rows": int(incomplete_rows),
            "data_quality_score": self._calculate_score(group, duplicate_count, missing_info, type_mismatches)
        }

    def _detect_type_mismatches(self, df):
        mismatches = {}
        for col in df.columns:
            # Drop NA to check actual data
            non_na = df[col].dropna()
            if non_na.empty:
                continue
            
            # Simple check: are they all of the same basic type?
            types = non_na.apply(type).unique()
            if len(types) > 1:
                mismatches[col] = [t.__name__ for t in types]
        return mismatches

    def _detect_outliers(self, df):
        outliers_info = {}
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            q1 = df[col].quantile(0.25)
            q3 = df[col].quantile(0.75)
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            
            outlier_count = ((df[col] < lower_bound) | (df[col] > upper_bound)).sum()
            if outlier_count > 0:
                outliers_info[col] = int(outlier_count)
        return outliers_info

    def _calculate_score(self, df, duplicates, missing, mismatches):
        total_cells = df.size
        if total_cells == 0:
            return 100
        
        # Simple penalty system
        missing_count = sum(missing.values())
        mismatch_count = len(mismatches)
        
        penalty = (duplicates * 2) + (missing_count * 1) + (mismatch_count * 5)
        score = max(0, 100 - (penalty / total_cells * 100))
        return round(score, 2)

    def _generate_overall_summary(self):
        for res in self.results.values():
            self.summary.append({
                "Category": res["name"],
                "Total Rows": res["total_rows"],
                "Duplicates": res["duplicate_count"],
                "Incomplete Rows": res["incomplete_rows"],
                "Quality Score": res["data_quality_score"]
            })

    def get_summary_df(self):
        return pd.DataFrame(self.summary)
