#!/usr/bin/env python3
"""
diagnostics/forward_fill_candidates.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Report participant keys whose latest snapshot row has blank MOST_RECENT_FIELDS
but an earlier snapshot row has a nonblank value.

Usage
-----
    python diagnostics/forward_fill_candidates.py [output_dir]

If *output_dir* is omitted, runs the Sample input batch into a temp folder
first, then analyzes ``icms_snapshot_history.csv`` in that folder.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import icms_etl_core as core
from diagnostics.ab_compare_deferred_batch import _run_deferred_batch
from diagnostics.ab_compare_master_rebuild import _prepare_output_dir

DEFAULT_OUT = REPO_ROOT / "diagnostics" / "_forward_fill_scan"
REPORT_PATH = REPO_ROOT / "diagnostics" / "forward_fill_candidates.csv"


def _analyze_history(history_path: Path) -> pd.DataFrame:
    if not history_path.exists():
        raise FileNotFoundError(f"Snapshot history not found: {history_path}")

    hist = pd.read_csv(history_path, dtype=str, encoding="utf-8").fillna("")
    if hist.empty or not all(k in hist.columns for k in core.KEY_FIELDS):
        return pd.DataFrame()

    hist = hist.copy()
    hist["_snap"] = pd.to_datetime(hist["snapshot_date"], errors="coerce")
    hist = hist.sort_values(["_snap"], kind="stable")

    rows: List[dict] = []
    for field in core.MOST_RECENT_FIELDS:
        if field not in hist.columns:
            continue

        for key, group in hist.groupby(core.KEY_FIELDS, sort=False):
            latest_row = group.iloc[-1]
            latest_val = core._safe_str(latest_row.get(field, ""))
            if latest_val:
                continue

            earlier = group.iloc[:-1]
            if earlier.empty:
                continue

            earlier_nonblank = earlier[
                earlier[field].map(core._safe_str).ne("")
            ]
            if earlier_nonblank.empty:
                continue

            source_row = earlier_nonblank.iloc[-1]
            rows.append({
                "CHAMP ID": core._safe_str(key[0]),
                "ICMS Project Name": core._safe_str(key[1]),
                "HMIS ID": core._safe_str(latest_row.get("HMIS ID", "")),
                "field_name": field,
                "latest_snapshot_date": core._safe_str(latest_row.get("snapshot_date", "")),
                "latest_source_file_name": core._safe_str(latest_row.get("source_file_name", "")),
                "earlier_snapshot_date": core._safe_str(source_row.get("snapshot_date", "")),
                "earlier_source_file_name": core._safe_str(source_row.get("source_file_name", "")),
                "earlier_value": core._safe_str(source_row.get(field, "")),
            })

    return pd.DataFrame(rows)


def main() -> int:
    output_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_OUT
    if len(sys.argv) == 1:
        _prepare_output_dir(output_dir)
        _run_deferred_batch(output_dir)

    report = _analyze_history(output_dir / "icms_snapshot_history.csv")
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    if report.empty:
        if REPORT_PATH.exists():
            REPORT_PATH.unlink()
        print("No forward-fill candidates found.")
        return 0

    report.to_csv(REPORT_PATH, index=False, encoding="utf-8")
    print(f"Forward-fill candidates: {len(report)}")
    print(f"Unique participant keys: {report[core.KEY_FIELDS].drop_duplicates().shape[0]}")
    print(f"Report written to: {REPORT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
