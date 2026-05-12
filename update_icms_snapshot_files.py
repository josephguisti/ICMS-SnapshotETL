"""
update_icms_snapshot_files.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CLI runner for the ICMS weekly snapshot ETL.

Usage
-----
python update_icms_snapshot_files.py \\
    --input "ICMS Snapshot-DeIdentified Scrambled.xlsx" \\
    --snapshot-date 2026-05-01 \\
    --sheet "Participant Level Summary"

All ETL logic lives in icms_etl_core.py so the future Tkinter GUI can call
those functions directly without duplicating any code.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import icms_etl_core as core


# ── Argument parsing ──────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="update_icms_snapshot_files",
        description=(
            "Ingest a weekly ICMS Excel snapshot, reduce it to the approved "
            "field set, and update all cumulative output files."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples\n"
            "--------\n"
            '  python update_icms_snapshot_files.py \\\n'
            '      --input "ICMS Snapshot-DeIdentified Scrambled.xlsx" \\\n'
            '      --snapshot-date 2026-05-01 \\\n'
            '      --sheet "Participant Level Summary"\n'
        ),
    )
    parser.add_argument(
        "--input",
        required=True,
        metavar="FILE",
        help="Path to the weekly ICMS Excel snapshot (.xlsx).",
    )
    parser.add_argument(
        "--snapshot-date",
        required=True,
        metavar="YYYY-MM-DD",
        help="Reporting date for this snapshot.",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        metavar="SHEET_NAME",
        help=(
            "Name of the sheet to process.  "
            "If omitted, available sheets are listed and you will be prompted."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        metavar="DIR",
        help="Folder for output files.  Defaults to the current directory.",
    )
    return parser


# ── Interactive sheet selection ───────────────────────────────────────────────

def _select_sheet(wb, provided: str | None) -> str:
    """Return the sheet name to process.

    If ``provided`` is given, validate it exists and return it.
    Otherwise, list available sheets and prompt the user to choose.
    """
    sheets = core.list_sheet_names(wb)

    if provided is not None:
        if provided in sheets:
            return provided
        print(f"\nERROR: Sheet '{provided}' was not found in the workbook.")
        print("Available sheets:")
        for s in sheets:
            print(f"  {s}")
        sys.exit(1)

    print("\nAvailable sheets:")
    for i, s in enumerate(sheets, 1):
        print(f"  {i:3}.  {s}")

    while True:
        choice = input(
            "\nEnter the sheet name or its number: "
        ).strip()

        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(sheets):
                return sheets[idx]
            print(f"  Please enter a number between 1 and {len(sheets)}.")
        elif choice in sheets:
            return choice
        else:
            print("  That is not a valid sheet name or number.  Please try again.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    # ── Resolve paths ─────────────────────────────────────────────────────
    input_path = Path(args.input)
    output_dir = Path(args.output_dir) if args.output_dir else Path.cwd()
    # field_aliases.json lives next to the scripts, not in --output-dir
    script_dir = Path(__file__).parent.resolve()

    # ── Validate snapshot date ─────────────────────────────────────────────
    snapshot_date = args.snapshot_date
    try:
        datetime.strptime(snapshot_date, "%Y-%m-%d")
    except ValueError:
        print(
            f"ERROR: --snapshot-date must be in YYYY-MM-DD format, "
            f"got: '{snapshot_date}'"
        )
        sys.exit(1)

    # ── Ensure output directory exists ─────────────────────────────────────
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        print(f"ERROR: Could not create output directory '{output_dir}': {exc}")
        sys.exit(1)

    run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    source_file_name = input_path.name
    notes: list[str] = []

    # ── Open workbook ──────────────────────────────────────────────────────
    print(f"\nOpening workbook:  {input_path.name}")
    try:
        wb, was_encrypted = core.open_workbook(input_path)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    if was_encrypted:
        print("  (file was password-protected — decrypted in memory, "
              "no copy written to disk)")

    # ── Sheet selection ────────────────────────────────────────────────────
    sheet_name = _select_sheet(wb, args.sheet)
    print(f"  Sheet:  {sheet_name}")

    # ── Read sheet data ────────────────────────────────────────────────────
    print("\nReading sheet data ...")
    try:
        df_raw = core.read_sheet(wb, sheet_name)
    except (ValueError, Exception) as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    finally:
        wb.close()

    input_row_count = len(df_raw)
    print(f"  Rows read (after dropping completely blank rows): {input_row_count}")

    if input_row_count == 0:
        print("ERROR: No data rows found in the selected sheet.  Nothing to process.")
        sys.exit(1)

    # ── Field resolution ───────────────────────────────────────────────────
    print("\nResolving fields ...")
    aliases = core.load_aliases(script_dir)

    # Exclude the internal row-number column from the columns we present
    workbook_columns = [c for c in df_raw.columns if not c.startswith("_")]

    resolved, missing_req, missing_opt, aliases_added = (
        core.resolve_fields_interactive(workbook_columns, aliases, script_dir)
    )

    if missing_req:
        notes.append(f"Missing required fields: {', '.join(missing_req)}")
        print(
            f"\n  WARNING: {len(missing_req)} required field(s) could not be "
            "resolved and will be absent from all outputs:"
        )
        for f in missing_req:
            print(f"    - {f}")

    if missing_opt:
        print(
            f"  Note: {len(missing_opt)} optional field(s) are not present "
            "in this workbook (they will be omitted from outputs)."
        )

    if aliases_added:
        print(
            f"  {aliases_added} new field alias(es) saved to "
            f"{core.ALIASES_FILENAME}."
        )

    resolved_count = len(resolved)
    print(
        f"  Fields resolved: {resolved_count} of "
        f"{len(core.ALL_APPROVED_FIELDS)} approved fields."
    )

    # ── Reduce to approved field set ───────────────────────────────────────
    df = core.reduce_to_approved_fields(
        df_raw, resolved, snapshot_date, source_file_name
    )

    # ── Duplicate detection ────────────────────────────────────────────────
    dup_df = core.find_duplicates(df)
    dup_flag_count = len(dup_df)

    if dup_flag_count > 0:
        unique_dup_pairs = (
            dup_df.groupby(core.KEY_FIELDS, sort=False).ngroups
            if all(k in dup_df.columns for k in core.KEY_FIELDS)
            else 0
        )
        print(
            f"\n  WARNING: {unique_dup_pairs} CHAMP ID + ICMS Project Name "
            f"pair(s) appear more than once ({dup_flag_count} flagged rows).  "
            "All rows will be kept in the dashboard and history files; the "
            "master file will use the last occurrence."
        )
        notes.append(
            f"Duplicate CHAMP ID / ICMS Project Name pairs: {unique_dup_pairs}"
        )
    else:
        unique_dup_pairs = 0

    # ── Write output files ─────────────────────────────────────────────────
    print("\nWriting output files ...")

    # 1. Dashboard flat file
    dashboard_count = 0
    try:
        core.update_dashboard_flatfile(df, output_dir)
        dashboard_count = len(df)
        print(f"  OK  Dashboard flat file      : {dashboard_count} rows")
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        _write_failure_log(
            run_timestamp, snapshot_date, source_file_name,
            sheet_name, input_row_count, str(exc), output_dir,
        )
        sys.exit(1)

    # 2. Snapshot history
    history_count = 0
    try:
        history_count = core.update_snapshot_history(df, snapshot_date, output_dir)
        print(f"  OK  Snapshot history         : {history_count} total rows")
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        _write_failure_log(
            run_timestamp, snapshot_date, source_file_name,
            sheet_name, input_row_count, str(exc), output_dir,
        )
        sys.exit(1)

    # 3. Client/project master
    master_total = new_rows = returning_rows = not_present = 0
    try:
        master_total, new_rows, returning_rows, not_present = (
            core.update_client_project_master(
                df, snapshot_date, source_file_name, output_dir
            )
        )
        print(
            f"  OK  Client/project master    : {master_total} total rows  "
            f"({new_rows} new, {returning_rows} returning, "
            f"{not_present} not in latest snapshot)"
        )
    except RuntimeError as exc:
        print(f"ERROR: {exc}")
        _write_failure_log(
            run_timestamp, snapshot_date, source_file_name,
            sheet_name, input_row_count, str(exc), output_dir,
        )
        sys.exit(1)

    # 4. Duplicate warning file (only when duplicates exist)
    if dup_flag_count > 0:
        try:
            warning_rows_df = core.build_duplicate_warning_rows(
                dup_df, run_timestamp, snapshot_date, source_file_name
            )
            core.append_duplicate_warning(
                warning_rows_df, snapshot_date, output_dir
            )
            print(
                f"  OK  Duplicate warning file   : "
                f"{dup_flag_count} flagged rows written"
            )
        except RuntimeError as exc:
            # Non-fatal — warn but do not abort
            print(f"  WARNING: Could not write duplicate warning file: {exc}")
            notes.append(f"Duplicate warning file write failed: {exc}")

    # 5. Ingest log
    log_row = {
        "run_timestamp": run_timestamp,
        "snapshot_date": snapshot_date,
        "source_file_name": source_file_name,
        "sheet_name": sheet_name,
        "input_row_count": input_row_count,
        "current_dashboard_row_count": dashboard_count,
        "history_total_row_count_after_update": history_count,
        "master_row_count_after_update": master_total,
        "new_client_project_rows": new_rows,
        "returning_client_project_rows": returning_rows,
        "not_present_in_latest_snapshot_count": not_present,
        "required_fields_missing_count": len(missing_req),
        "optional_fields_missing_count": len(missing_opt),
        "field_aliases_added_count": aliases_added,
        "duplicate_client_project_count": dup_flag_count,
        "password_protected_input": was_encrypted,
        "run_status": "success",
        "notes": "; ".join(notes) if notes else "",
    }

    try:
        core.append_ingest_log(log_row, output_dir)
        print(f"  OK  Ingest log               : updated")
    except RuntimeError as exc:
        # Non-fatal — warn but do not abort
        print(f"  WARNING: Could not write ingest log: {exc}")

    # ── Run summary ────────────────────────────────────────────────────────
    _print_summary(
        source_file_name=source_file_name,
        sheet_name=sheet_name,
        snapshot_date=snapshot_date,
        dashboard_count=dashboard_count,
        history_count=history_count,
        master_total=master_total,
        dup_flag_count=dup_flag_count,
        missing_req=missing_req,
        output_dir=output_dir,
    )


def _write_failure_log(
    run_timestamp: str,
    snapshot_date: str,
    source_file_name: str,
    sheet_name: str,
    input_row_count: int,
    error_message: str,
    output_dir: Path,
) -> None:
    """Best-effort attempt to record a failed run in the ingest log."""
    log_row = {
        "run_timestamp": run_timestamp,
        "snapshot_date": snapshot_date,
        "source_file_name": source_file_name,
        "sheet_name": sheet_name,
        "input_row_count": input_row_count,
        "current_dashboard_row_count": "",
        "history_total_row_count_after_update": "",
        "master_row_count_after_update": "",
        "new_client_project_rows": "",
        "returning_client_project_rows": "",
        "not_present_in_latest_snapshot_count": "",
        "required_fields_missing_count": "",
        "optional_fields_missing_count": "",
        "field_aliases_added_count": "",
        "duplicate_client_project_count": "",
        "password_protected_input": "",
        "run_status": "failed",
        "notes": error_message[:500],
    }
    try:
        core.append_ingest_log(log_row, output_dir)
    except Exception:
        pass  # already in an error path; best effort only


def _print_summary(
    *,
    source_file_name: str,
    sheet_name: str,
    snapshot_date: str,
    dashboard_count: int,
    history_count: int,
    master_total: int,
    dup_flag_count: int,
    missing_req: list,
    output_dir: Path,
) -> None:
    sep = "-" * 62
    print(f"\n{sep}")
    print("  RUN SUMMARY")
    print(sep)
    print(f"  Input file       :  {source_file_name}")
    print(f"  Sheet            :  {sheet_name}")
    print(f"  Snapshot date    :  {snapshot_date}")
    print(f"  Dashboard rows   :  {dashboard_count}")
    print(f"  History rows     :  {history_count}")
    print(f"  Master rows      :  {master_total}")
    print(f"  Duplicate flags  :  {dup_flag_count}")
    if missing_req:
        print(f"  Missing required :  {', '.join(missing_req)}")
    print(f"  Output folder    :  {output_dir.resolve()}")
    print(sep)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user.  Exiting.")
        sys.exit(1)
    except Exception as exc:
        print(f"\nUnexpected error: {exc}")
        raise
