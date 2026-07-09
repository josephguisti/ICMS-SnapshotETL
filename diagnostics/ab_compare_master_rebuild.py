#!/usr/bin/env python3
"""
diagnostics/ab_compare_master_rebuild.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Run the Sample input batch through the original and vectorized master-rebuild
implementations, then compare outputs for behavioral equivalence.

Usage
-----
    python diagnostics/ab_compare_master_rebuild.py

Exit code 0 when all comparisons pass; non-zero on any mismatch.
"""

from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path
from typing import Callable, Dict, List, Tuple

import pandas as pd

# Allow imports from repo root when invoked from diagnostics/.
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import icms_etl_core as core
import icms_etl_core_vectorized as core_vec
from update_icms_snapshot_files import _parse_date_from_filename

SAMPLE_INPUT_DIR = REPO_ROOT / "Sample input"
OUT_ORIGINAL = REPO_ROOT / "diagnostics" / "_ab_out_original"
OUT_VECTORIZED = REPO_ROOT / "diagnostics" / "_ab_out_vectorized"
MISMATCH_REPORT = REPO_ROOT / "diagnostics" / "ab_mismatches.csv"

SHEET_NAME = "Participant Level Summary"
HISTORY_FILES = (
    "icms_snapshot_history.csv",
    "icms_exit_log_history.csv",
)
MASTER_FILE = "icms_client_project_master.xlsx"


def _discover_batch_files() -> List[Tuple[Path, str]]:
    """Return Sample input files sorted by parsed snapshot date."""
    if not SAMPLE_INPUT_DIR.is_dir():
        raise FileNotFoundError(f"Sample input folder not found: {SAMPLE_INPUT_DIR}")

    entries: List[Tuple[Path, str]] = []
    for path in sorted(SAMPLE_INPUT_DIR.glob("*.xlsx")):
        snap = _parse_date_from_filename(path.name)
        if snap is None:
            raise ValueError(f"Could not parse snapshot date from: {path.name}")
        entries.append((path, snap))

    entries.sort(key=lambda item: item[1])
    if not entries:
        raise FileNotFoundError(f"No .xlsx files found in {SAMPLE_INPUT_DIR}")
    return entries


def _prepare_output_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for child in path.iterdir():
        if child.is_file():
            child.unlink(missing_ok=True)
        else:
            shutil.rmtree(child, ignore_errors=True)


def _process_file(
    input_path: Path,
    snapshot_date: str,
    output_dir: Path,
    master_updater: Callable[..., Tuple[int, int, int, int, int, int]],
    aliases: Dict[str, str],
) -> float:
    """Ingest one snapshot file; return master-rebuild wall time in seconds."""
    source_file_name = input_path.name
    wb, _was_encrypted = core.open_workbook(input_path, password="")
    try:
        if SHEET_NAME not in core.list_sheet_names(wb):
            raise ValueError(f"Sheet '{SHEET_NAME}' not found in {input_path.name}")

        header_row_idx = core.find_header_row(
            wb, SHEET_NAME, core.ALL_APPROVED_FIELDS
        )
        df_raw = core.read_sheet(wb, SHEET_NAME, header_row=header_row_idx)
        vacant_df = core.read_vacant_slots(wb)
        exit_df = core.read_exit_log(wb)
    finally:
        wb.close()

    if df_raw.empty:
        raise ValueError(f"No participant rows in {input_path.name}")

    vacant_rows = core.build_vacant_slot_rows(
        vacant_df, snapshot_date, source_file_name
    )

    workbook_columns = [c for c in df_raw.columns if not c.startswith("_")]
    resolved, missing_req, _missing_opt = core.resolve_fields(
        workbook_columns, aliases
    )
    if missing_req:
        raise ValueError(
            f"Unresolved required fields in {input_path.name}: {missing_req}"
        )

    df = core.reduce_to_approved_fields(
        df_raw, resolved, snapshot_date, source_file_name
    )

    core.update_snapshot_history(df, snapshot_date, output_dir)

    t0 = time.perf_counter()
    master_updater(df, snapshot_date, source_file_name, output_dir, exit_df)
    master_elapsed = time.perf_counter() - t0

    # Dashboard is not part of this comparison but mirrors a realistic run.
    core.update_dashboard_flatfile(df, output_dir, vacant_rows)

    return master_elapsed


def _run_batch(
    output_dir: Path,
    master_updater: Callable[..., Tuple[int, int, int, int, int, int]],
    label: str,
) -> float:
    """Process all Sample input files; return total master-rebuild time."""
    _prepare_output_dir(output_dir)
    aliases = core.load_aliases(REPO_ROOT)
    batch_files = _discover_batch_files()

    total_master_time = 0.0
    print(f"\n=== {label} ===")
    for idx, (path, snap) in enumerate(batch_files, 1):
        elapsed = _process_file(path, snap, output_dir, master_updater, aliases)
        total_master_time += elapsed
        print(f"  [{idx}/{len(batch_files)}] {path.name} ({snap})  master={elapsed:.3f}s")

    print(f"  Total master-rebuild time: {total_master_time:.3f}s")
    return total_master_time


def _read_master(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    return pd.read_excel(path, engine="openpyxl", dtype=str).fillna("")


def _compare_history_files(
    original_dir: Path,
    vectorized_dir: Path,
) -> List[dict]:
    """Byte-for-byte compare of shared history CSVs."""
    mismatches: List[dict] = []

    for filename in HISTORY_FILES:
        orig_path = original_dir / filename
        vec_path = vectorized_dir / filename

        if not orig_path.exists() and not vec_path.exists():
            continue
        if orig_path.exists() != vec_path.exists():
            mismatches.append({
                "category": "history_missing_file",
                "champ_id": "",
                "icms_project_name": "",
                "hmis_id": "",
                "column": filename,
                "original_value": "exists" if orig_path.exists() else "missing",
                "vectorized_value": "exists" if vec_path.exists() else "missing",
            })
            continue

        orig_bytes = orig_path.read_bytes()
        vec_bytes = vec_path.read_bytes()
        if orig_bytes != vec_bytes:
            mismatches.append({
                "category": "history_content_mismatch",
                "champ_id": "",
                "icms_project_name": "",
                "hmis_id": "",
                "column": filename,
                "original_value": f"{len(orig_bytes)} bytes",
                "vectorized_value": f"{len(vec_bytes)} bytes",
            })

    return mismatches


def _compare_master_files(
    original_dir: Path,
    vectorized_dir: Path,
) -> List[dict]:
    """Row- and cell-level diff of client/project master outputs."""
    mismatches: List[dict] = []
    orig = _read_master(original_dir / MASTER_FILE)
    vec = _read_master(vectorized_dir / MASTER_FILE)

    if orig.empty and vec.empty:
        return mismatches

    key_cols = core.KEY_FIELDS
    for col in key_cols:
        if col not in orig.columns:
            orig[col] = ""
        if col not in vec.columns:
            vec[col] = ""

    orig = orig.set_index(key_cols, drop=False)
    vec = vec.set_index(key_cols, drop=False)

    orig_keys = set(orig.index)
    vec_keys = set(vec.index)

    for key in sorted(orig_keys - vec_keys):
        champ, project = key
        hmis = _safe_row_value(orig.loc[key], "HMIS ID")
        mismatches.append({
            "category": "missing_in_vectorized",
            "champ_id": champ,
            "icms_project_name": project,
            "hmis_id": hmis,
            "column": "",
            "original_value": "row present",
            "vectorized_value": "row absent",
        })

    for key in sorted(vec_keys - orig_keys):
        champ, project = key
        hmis = _safe_row_value(vec.loc[key], "HMIS ID")
        mismatches.append({
            "category": "extra_in_vectorized",
            "champ_id": champ,
            "icms_project_name": project,
            "hmis_id": hmis,
            "column": "",
            "original_value": "row absent",
            "vectorized_value": "row present",
        })

    shared_keys = orig_keys & vec_keys
    all_columns = sorted(set(orig.columns) | set(vec.columns))

    for key in sorted(shared_keys):
        orig_row = orig.loc[key]
        vec_row = vec.loc[key]
        if isinstance(orig_row, pd.DataFrame):
            orig_row = orig_row.iloc[-1]
        if isinstance(vec_row, pd.DataFrame):
            vec_row = vec_row.iloc[-1]

        champ = str(key[0])
        project = str(key[1])
        orig_hmis = _safe_row_value(orig_row, "HMIS ID")
        vec_hmis = _safe_row_value(vec_row, "HMIS ID")

        if orig_hmis != vec_hmis:
            mismatches.append({
                "category": "hmis_id_mismatch",
                "champ_id": champ,
                "icms_project_name": project,
                "hmis_id": f"orig={orig_hmis} vec={vec_hmis}",
                "column": "HMIS ID",
                "original_value": orig_hmis,
                "vectorized_value": vec_hmis,
            })

        for col in all_columns:
            if col in key_cols:
                continue
            orig_val = _safe_row_value(orig_row, col)
            vec_val = _safe_row_value(vec_row, col)
            if orig_val != vec_val:
                mismatches.append({
                    "category": "cell_mismatch",
                    "champ_id": champ,
                    "icms_project_name": project,
                    "hmis_id": orig_hmis,
                    "column": col,
                    "original_value": orig_val,
                    "vectorized_value": vec_val,
                })

    return mismatches


def _safe_row_value(row: pd.Series, column: str) -> str:
    if column not in row.index:
        return ""
    val = row[column]
    if isinstance(val, pd.Series):
        val = val.iloc[0]
    return core._safe_str(val)


def _write_report(mismatches: List[dict]) -> None:
    MISMATCH_REPORT.parent.mkdir(parents=True, exist_ok=True)
    if mismatches:
        pd.DataFrame(mismatches).to_csv(MISMATCH_REPORT, index=False, encoding="utf-8")
    elif MISMATCH_REPORT.exists():
        MISMATCH_REPORT.unlink()


def main() -> int:
    print("ICMS master rebuild A/B comparison")
    print(f"  Sample input : {SAMPLE_INPUT_DIR}")
    print(f"  Original out : {OUT_ORIGINAL}")
    print(f"  Vectorized out: {OUT_VECTORIZED}")

    orig_time = _run_batch(
        OUT_ORIGINAL,
        core.update_client_project_master,
        "Original master rebuild",
    )
    vec_time = _run_batch(
        OUT_VECTORIZED,
        core_vec.update_client_project_master,
        "Vectorized master rebuild",
    )

    print("\n=== Comparisons ===")
    mismatches: List[dict] = []
    mismatches.extend(_compare_history_files(OUT_ORIGINAL, OUT_VECTORIZED))
    mismatches.extend(_compare_master_files(OUT_ORIGINAL, OUT_VECTORIZED))
    _write_report(mismatches)

    history_issues = [
        m for m in mismatches if m["category"].startswith("history_")
    ]
    master_issues = [
        m for m in mismatches if not m["category"].startswith("history_")
    ]

    missing_vec = sum(1 for m in master_issues if m["category"] == "missing_in_vectorized")
    extra_vec = sum(1 for m in master_issues if m["category"] == "extra_in_vectorized")
    cell_mismatches = sum(1 for m in master_issues if m["category"] == "cell_mismatch")
    hmis_mismatches = sum(1 for m in master_issues if m["category"] == "hmis_id_mismatch")

    print(f"  History file mismatches : {len(history_issues)}")
    print(f"  Master rows missing (vec): {missing_vec}")
    print(f"  Master rows extra (vec)  : {extra_vec}")
    print(f"  HMIS ID mismatches       : {hmis_mismatches}")
    print(f"  Cell value mismatches    : {cell_mismatches}")
    print(f"  Original master time     : {orig_time:.3f}s")
    print(f"  Vectorized master time   : {vec_time:.3f}s")

    if mismatches:
        print(f"\nFAIL — {len(mismatches)} issue(s) written to {MISMATCH_REPORT}")
        return 1

    print("\nPASS — original and vectorized outputs are equivalent.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
