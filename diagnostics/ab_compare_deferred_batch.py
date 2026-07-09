#!/usr/bin/env python3
"""
diagnostics/ab_compare_deferred_batch.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Compare per-file master rebuild vs deferred end-of-batch rebuild on Sample input.

Usage
-----
    python diagnostics/ab_compare_deferred_batch.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Callable, Dict, List, Tuple

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import icms_etl_core as core
from diagnostics.ab_compare_master_rebuild import (
    HISTORY_FILES,
    MASTER_FILE,
    SAMPLE_INPUT_DIR,
    _compare_history_files,
    _compare_master_files,
    _discover_batch_files,
    _prepare_output_dir,
    _process_file,
    _write_report,
)
from update_icms_snapshot_files import _parse_date_from_filename

OUT_PER_FILE = REPO_ROOT / "diagnostics" / "_deferred_out_per_file"
OUT_DEFERRED = REPO_ROOT / "diagnostics" / "_deferred_out_end_of_batch"
MISMATCH_REPORT = REPO_ROOT / "diagnostics" / "deferred_batch_mismatches.csv"


def _run_per_file_batch(output_dir: Path) -> None:
    _prepare_output_dir(output_dir)
    aliases = core.load_aliases(REPO_ROOT)
    for path, snap in _discover_batch_files():
        _process_file(
            path, snap, output_dir, core.update_client_project_master, aliases
        )


def _run_deferred_batch(output_dir: Path) -> None:
    _prepare_output_dir(output_dir)
    aliases = core.load_aliases(REPO_ROOT)
    for path, snap in _discover_batch_files():
        source_file_name = path.name
        wb, _ = core.open_workbook(path, password="")
        try:
            header_row_idx = core.find_header_row(
                wb, "Participant Level Summary", core.ALL_APPROVED_FIELDS
            )
            df_raw = core.read_sheet(
                wb, "Participant Level Summary", header_row=header_row_idx
            )
            vacant_df = core.read_vacant_slots(wb)
            exit_df = core.read_exit_log(wb)
        finally:
            wb.close()

        vacant_rows = core.build_vacant_slot_rows(vacant_df, snap, source_file_name)
        workbook_columns = [c for c in df_raw.columns if not c.startswith("_")]
        resolved, missing_req, _ = core.resolve_fields(workbook_columns, aliases)
        if missing_req:
            raise ValueError(f"Unresolved required fields in {path.name}: {missing_req}")

        df = core.reduce_to_approved_fields(
            df_raw, resolved, snap, source_file_name
        )
        core.update_snapshot_history(df, snap, output_dir)
        core.record_master_inputs(exit_df, snap, source_file_name, output_dir)
        core.update_dashboard_flatfile(df, output_dir, vacant_rows)

    core.rebuild_client_project_master(output_dir)


def main() -> int:
    print("Deferred batch rebuild comparison")
    print(f"  Sample input : {SAMPLE_INPUT_DIR}")
    print(f"  Per-file out : {OUT_PER_FILE}")
    print(f"  Deferred out : {OUT_DEFERRED}")

    _run_per_file_batch(OUT_PER_FILE)
    _run_deferred_batch(OUT_DEFERRED)

    mismatches: List[dict] = []
    mismatches.extend(_compare_history_files(OUT_PER_FILE, OUT_DEFERRED))
    mismatches.extend(_compare_master_files(OUT_PER_FILE, OUT_DEFERRED))

    if MISMATCH_REPORT.exists() or mismatches:
        pd.DataFrame(mismatches).to_csv(MISMATCH_REPORT, index=False, encoding="utf-8")
    elif MISMATCH_REPORT.exists():
        MISMATCH_REPORT.unlink()

    print(f"\n  History mismatches : {sum(1 for m in mismatches if m['category'].startswith('history_'))}")
    print(f"  Master mismatches  : {sum(1 for m in mismatches if not m['category'].startswith('history_'))}")

    if mismatches:
        print(f"\nFAIL — wrote {MISMATCH_REPORT}")
        return 1

    print("\nPASS — deferred end-of-batch rebuild matches per-file rebuild.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
