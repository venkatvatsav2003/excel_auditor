import os
import sys
import csv
import tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from auditor import ColumnProfile, DataAuditor


def test_column_profile_missing():
    profile = ColumnProfile("test", ["a", "", "b", "", "c"])
    assert profile.missing == 2
    assert profile.total == 5


def test_column_profile_unique():
    profile = ColumnProfile("test", ["a", "b", "a", "c", "b"])
    assert profile.unique == 3


def test_data_auditor_csv():
    with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
        f.write("name,age,city\nAlice,30,NYC\nBob,25,LA\nAlice,30,NYC\n")
        fp = f.name
    auditor = DataAuditor()
    result = auditor.audit(fp)
    os.unlink(fp)
    assert result["rows"] == 3
    assert result["total_duplicate_rows"] == 1
    assert result["columns"] == 3
