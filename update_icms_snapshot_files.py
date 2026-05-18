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
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import icms_etl_core as core


# ── Settings persistence ──────────────────────────────────────────────────────

def load_settings(script_dir: Path) -> dict:
    """Load settings.json from the script directory.

    Returns a dict with keys: ``last_input_dir``, ``last_output_dir``,
    ``last_sheet_name``.  Returns ``{}`` on any error — never raises.
    """
    settings_path = script_dir / "settings.json"
    if not settings_path.exists():
        return {}
    try:
        with open(settings_path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_settings(settings: dict, script_dir: Path) -> None:
    """Persist settings to settings.json in the script directory."""
    settings_path = script_dir / "settings.json"
    try:
        with open(settings_path, "w", encoding="utf-8") as fh:
            json.dump(settings, fh, indent=2, ensure_ascii=False)
    except Exception as exc:
        print(f"WARNING: Could not save settings.json: {exc}")


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
            "  Single file:\n"
            '    python update_icms_snapshot_files.py \\\n'
            '        --input "ICMS Snapshot - 01.28.2026.xlsx" \\\n'
            '        --snapshot-date 2026-01-28 \\\n'
            '        --sheet "Participant Level Summary"\n\n'
            "  Batch:\n"
            '    python update_icms_snapshot_files.py \\\n'
            '        --batch "./snapshots" \\\n'
            '        --output-dir "./output" \\\n'
            '        --sheet "Participant Level Summary"\n'
        ),
    )
    parser.add_argument(
        "--input",
        default=None,
        metavar="FILE",
        help="Path to the weekly ICMS Excel snapshot (.xlsx). Required in single-file mode.",
    )
    parser.add_argument(
        "--snapshot-date",
        default=None,
        metavar="YYYY-MM-DD",
        help="Reporting date for this snapshot. Required in single-file mode.",
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
    parser.add_argument(
        "--batch",
        default=None,
        metavar="DIR",
        help=(
            "Path to a folder containing multiple .xlsx snapshot files "
            "to process in sequence. When provided, --input and --snapshot-date "
            "are ignored."
        ),
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


# ── Snapshot date gap check ───────────────────────────────────────────────────

def _check_snapshot_date_gap(snapshot_date: str, output_dir: Path) -> None:
    """Warn and require confirmation if the incoming snapshot date is more
    than 45 days after the most recent date in the history file.

    Silently skips if the history file is absent or unreadable.
    """
    history_path = output_dir / "icms_snapshot_history.csv"
    if not history_path.exists():
        return
    try:
        hist = pd.read_csv(
            history_path, dtype=str, usecols=["snapshot_date"]
        ).fillna("")
        dates = pd.to_datetime(hist["snapshot_date"], errors="coerce").dropna()
        if dates.empty:
            return
        most_recent = dates.max()
        incoming = pd.to_datetime(snapshot_date)
        gap_days = (incoming - most_recent).days
        if gap_days > 45:
            print(
                f"\n  WARNING: The incoming snapshot date ({snapshot_date}) is "
                f"{gap_days} days after the most recent date in the snapshot "
                f"history ({most_recent.strftime('%Y-%m-%d')})."
            )
            print("  This may indicate a missed snapshot or an incorrect date.")
            while True:
                resp = input(
                    "  Type 'confirm' to proceed anyway, or press Enter to abort: "
                ).strip().lower()
                if resp == "confirm":
                    break
                print("  Aborted by user.")
                sys.exit(0)
    except Exception:
        pass


# ── Batch pipeline helpers ─────────────────────────────────────────────────────

def _parse_date_from_filename(filename: str) -> str | None:
    """Search *filename* for a ``MM.DD.YYYY`` date pattern.

    Returns the date as ``YYYY-MM-DD``, or ``None`` if no match is found.
    """
    m = re.search(r"\d{2}\.\d{2}\.\d{4}", filename)
    if m:
        parts = m.group(0).split(".")
        return f"{parts[2]}-{parts[0]}-{parts[1]}"
    return None


def _confirm_batch_files(
    xlsx_files: list[Path],
) -> list[tuple[Path, str]]:
    """Present a numbered file list, allow inline date correction, and return
    a confirmed ``(path, snapshot_date)`` list sorted chronologically.

    The user must correct all ``UNPARSED`` entries before typing ``confirm``.
    Typing ``abort`` exits the program immediately.
    """
    entries: list[dict] = [
        {"path": f, "date": _parse_date_from_filename(f.name)}
        for f in xlsx_files
    ]

    def _sort_key(e: dict) -> tuple:
        return (e["date"] is None, e["date"] or "")

    while True:
        entries.sort(key=_sort_key)
        print("\nBatch file list (sorted by date, UNPARSED files last):")
        for i, e in enumerate(entries, 1):
            label = e["date"] if e["date"] is not None else "UNPARSED"
            print(f"  {i:3}.  {label:<12}  {e['path'].name}")

        has_unparsed = any(e["date"] is None for e in entries)
        choice = input(
            "\nEnter a file number to correct its date, "
            "'confirm' to proceed, or 'abort' to exit: "
        ).strip()

        if choice.lower() == "abort":
            sys.exit(0)

        if choice.lower() == "confirm":
            if has_unparsed:
                print(
                    "  Cannot confirm: some files still have UNPARSED dates. "
                    "Please correct all dates before confirming."
                )
                continue
            break

        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(entries):
                while True:
                    new_date = input(
                        f"  Corrected date for '{entries[idx]['path'].name}' "
                        "(YYYY-MM-DD): "
                    ).strip()
                    try:
                        datetime.strptime(new_date, "%Y-%m-%d")
                        entries[idx]["date"] = new_date
                        break
                    except ValueError:
                        print("  Invalid format. Please use YYYY-MM-DD.")
            else:
                print(f"  Please enter a number between 1 and {len(entries)}.")
        else:
            print("  Invalid input. Enter a number, 'confirm', or 'abort'.")

    return [(e["path"], e["date"]) for e in entries]


def _run_batch(
    batch_dir: Path,
    output_dir: Path,
    sheet_name_arg: str | None,
    script_dir: Path,
) -> None:
    """Process all .xlsx files in *batch_dir* in chronological order."""
    xlsx_files = sorted(batch_dir.glob("*.xlsx"))
    if not xlsx_files:
        print(f"\nNo .xlsx files found in: {batch_dir}")
        return

    confirmed = _confirm_batch_files(list(xlsx_files))
    total_count = len(confirmed)

    # Aliases accumulate in memory across files and are saved per-file
    aliases = core.load_aliases(script_dir)

    resolved_sheet_name: str | None = sheet_name_arg
    success_count = 0
    fail_count = 0
    master_total = 0

    for file_idx, (input_path, snapshot_date) in enumerate(confirmed, 1):
        print(
            f"\n── File {file_idx} of {total_count}: "
            f"{input_path.name} ({snapshot_date}) ──"
        )
        run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        source_file_name = input_path.name
        notes: list[str] = []

        # ── Open workbook ──────────────────────────────────────────────────
        try:
            wb, was_encrypted = core.open_workbook(input_path)
        except (FileNotFoundError, RuntimeError) as exc:
            print(f"  ERROR: {exc}")
            _write_failure_log(
                run_timestamp, snapshot_date, source_file_name,
                resolved_sheet_name or "", 0, str(exc), output_dir,
            )
            fail_count += 1
            continue

        if was_encrypted:
            print("  (file was password-protected — decrypted in memory)")

        # ── Sheet selection + data reading (always close wb) ───────────────
        try:
            if resolved_sheet_name is None:
                resolved_sheet_name = _select_sheet(wb, None)
            else:
                if resolved_sheet_name not in core.list_sheet_names(wb):
                    raise ValueError(
                        f"Sheet '{resolved_sheet_name}' not found "
                        f"in {input_path.name}"
                    )
            sheet_name = resolved_sheet_name
            print(f"  Sheet: {sheet_name}")

            header_row_idx = core.find_header_row(
                wb, sheet_name, core.ALL_APPROVED_FIELDS
            )
            df_raw = core.read_sheet(wb, sheet_name, header_row=header_row_idx)
            vacant_df = core.read_vacant_slots(wb)
            exit_df = core.read_exit_log(wb)
        except (ValueError, Exception) as exc:
            print(f"  ERROR: {exc}")
            _write_failure_log(
                run_timestamp, snapshot_date, source_file_name,
                resolved_sheet_name or "", 0, str(exc), output_dir,
            )
            fail_count += 1
            continue
        finally:
            wb.close()

        vacant_rows = core.build_vacant_slot_rows(
            vacant_df, snapshot_date, source_file_name
        )
        input_row_count = len(df_raw)
        print(
            f"  Rows read: {input_row_count}  |  "
            f"Vacant slots: {len(vacant_df)}  |  Exit log: {len(exit_df)}"
        )

        if input_row_count == 0:
            print("  ERROR: No data rows found in sheet. Skipping file.")
            _write_failure_log(
                run_timestamp, snapshot_date, source_file_name,
                sheet_name, 0, "No data rows found", output_dir,
            )
            fail_count += 1
            continue

        # ── Field resolution ───────────────────────────────────────────────
        workbook_columns = [c for c in df_raw.columns if not c.startswith("_")]
        resolved, missing_req, missing_opt, aliases_added = (
            core.resolve_fields_interactive(workbook_columns, aliases, script_dir)
        )
        if missing_req:
            notes.append(f"Missing required fields: {', '.join(missing_req)}")

        df = core.reduce_to_approved_fields(
            df_raw, resolved, snapshot_date, source_file_name
        )

        # ── Duplicate detection ────────────────────────────────────────────
        dup_df = core.find_duplicates(df)
        dup_flag_count = len(dup_df)
        if dup_flag_count > 0:
            unique_dup_pairs = (
                dup_df.groupby(core.KEY_FIELDS, sort=False).ngroups
                if all(k in dup_df.columns for k in core.KEY_FIELDS)
                else 0
            )
            notes.append(
                f"Duplicate CHAMP ID / ICMS Project Name pairs: {unique_dup_pairs}"
            )

        # ── Write outputs ──────────────────────────────────────────────────
        vacant_slot_count = len(vacant_rows)
        dashboard_count = 0
        history_count = 0
        master_total_file = 0
        new_rows = returning_rows = not_present = exited_count = unknown_count = 0
        file_failed = False
        error_for_log = ""

        try:
            core.update_dashboard_flatfile(df, output_dir, vacant_rows)
            dashboard_count = len(df) + vacant_slot_count
        except RuntimeError as exc:
            print(f"  ERROR writing dashboard flat file: {exc}")
            error_for_log = str(exc)
            file_failed = True

        if not file_failed:
            try:
                history_count = core.update_snapshot_history(
                    df, snapshot_date, output_dir
                )
            except RuntimeError as exc:
                print(f"  ERROR writing snapshot history: {exc}")
                error_for_log = str(exc)
                file_failed = True

        if not file_failed:
            try:
                (
                    master_total_file, new_rows, returning_rows,
                    not_present, exited_count, unknown_count,
                ) = core.update_client_project_master(
                    df, snapshot_date, source_file_name, output_dir, exit_df
                )
                master_total = master_total_file
            except RuntimeError as exc:
                print(f"  ERROR writing client/project master: {exc}")
                error_for_log = str(exc)
                file_failed = True

        if file_failed:
            _write_failure_log(
                run_timestamp, snapshot_date, source_file_name,
                sheet_name, input_row_count, error_for_log, output_dir,
            )
            fail_count += 1
            continue

        # ── Non-fatal outputs ──────────────────────────────────────────────
        if dup_flag_count > 0:
            try:
                warning_rows_df = core.build_duplicate_warning_rows(
                    dup_df, run_timestamp, snapshot_date, source_file_name
                )
                core.append_duplicate_warning(
                    warning_rows_df, snapshot_date, output_dir
                )
            except RuntimeError as exc:
                notes.append(f"Duplicate warning file write failed: {exc}")

        log_row = {
            "run_timestamp": run_timestamp,
            "snapshot_date": snapshot_date,
            "source_file_name": source_file_name,
            "sheet_name": sheet_name,
            "input_row_count": input_row_count,
            "vacant_slot_rows_injected": vacant_slot_count,
            "exit_log_records_processed": len(exit_df),
            "current_dashboard_row_count": dashboard_count,
            "history_total_row_count_after_update": history_count,
            "master_row_count_after_update": master_total_file,
            "new_client_project_rows": new_rows,
            "returning_client_project_rows": returning_rows,
            "not_present_in_latest_snapshot_count": not_present,
            "master_exited_count": exited_count,
            "master_unknown_count": unknown_count,
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
        except RuntimeError:
            pass

        success_count += 1
        print(f"  Dashboard rows  : {dashboard_count} ({vacant_slot_count} vacant slots)")
        print(f"  History rows    : {history_count} total")
        print(
            f"  Master rows     : {master_total_file} total "
            f"({exited_count} exited, {unknown_count} unknown)"
        )

    # ── Final batch summary ────────────────────────────────────────────────
    border = "\u2550" * 62
    print(f"\n{border}")
    print("  BATCH COMPLETE")
    print(f"  Files processed  : {success_count} of {total_count}")
    print(f"  Files failed     : {fail_count}")
    print(f"  Final master rows: {master_total}")
    print(f"  Output folder    : {output_dir.resolve()}")
    print(border)


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    # ── Resolve paths ─────────────────────────────────────────────────────
    output_dir = Path(args.output_dir) if args.output_dir else Path.cwd()
    # settings.json and field_aliases.json live next to the scripts
    script_dir = Path(__file__).parent.resolve()

    # ── Load and display saved settings ───────────────────────────────────
    settings = load_settings(script_dir)
    if settings:
        print("\nLast used settings:")
        print(f"  Input dir   : {settings.get('last_input_dir', 'not set')}")
        print(f"  Output dir  : {settings.get('last_output_dir', 'not set')}")
        print(f"  Sheet       : {settings.get('last_sheet_name', 'not set')}")

    # ── Batch mode ────────────────────────────────────────────────────────
    if args.batch:
        batch_dir = Path(args.batch)
        if not batch_dir.is_dir():
            print(f"ERROR: --batch path is not a directory: {batch_dir}")
            sys.exit(1)
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            print(f"ERROR: Could not create output directory '{output_dir}': {exc}")
            sys.exit(1)
        _run_batch(batch_dir, output_dir, args.sheet, script_dir)
        save_settings(
            {
                "last_input_dir": str(batch_dir.resolve()),
                "last_output_dir": str(output_dir.resolve()),
                "last_sheet_name": args.sheet or "",
            },
            script_dir,
        )
        return

    # ── Single-file mode: validate required arguments ──────────────────────
    if not args.input:
        print("ERROR: --input is required when not using --batch.")
        sys.exit(1)
    if not args.snapshot_date:
        print("ERROR: --snapshot-date is required when not using --batch.")
        sys.exit(1)

    input_path = Path(args.input)

    # ── Validate snapshot date format ──────────────────────────────────────
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

    # ── Snapshot date gap check ────────────────────────────────────────────
    _check_snapshot_date_gap(snapshot_date, output_dir)

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
    vacant_df: "pd.DataFrame"
    exit_df: "pd.DataFrame"
    try:
        header_row_idx = core.find_header_row(
            wb, sheet_name, core.ALL_APPROVED_FIELDS
        )
        df_raw = core.read_sheet(wb, sheet_name, header_row=header_row_idx)
        vacant_df = core.read_vacant_slots(wb)
        exit_df = core.read_exit_log(wb)
    except (ValueError, Exception) as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)
    finally:
        wb.close()

    vacant_rows = core.build_vacant_slot_rows(vacant_df, snapshot_date, source_file_name)

    input_row_count = len(df_raw)
    print(f"  Rows read (after dropping completely blank rows): {input_row_count}")
    print(f"  Vacant slot records: {len(vacant_df)}")
    print(f"  Exit log records: {len(exit_df)}")

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
    vacant_slot_count = len(vacant_rows)
    dashboard_count = 0
    try:
        core.update_dashboard_flatfile(df, output_dir, vacant_rows)
        dashboard_count = len(df) + vacant_slot_count
        print(
            f"  OK  Dashboard flat file      : {len(df)} participant rows + "
            f"{vacant_slot_count} vacant slot rows = {dashboard_count} total"
        )
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
    master_total = new_rows = returning_rows = not_present = exited_count = unknown_count = 0
    try:
        master_total, new_rows, returning_rows, not_present, exited_count, unknown_count = (
            core.update_client_project_master(
                df, snapshot_date, source_file_name, output_dir, exit_df
            )
        )
        print(
            f"  OK  Client/project master    : {master_total} total rows  "
            f"({new_rows} new, {returning_rows} returning, "
            f"{not_present} not in latest snapshot, "
            f"{exited_count} exited, {unknown_count} unknown)"
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
        "vacant_slot_rows_injected": vacant_slot_count,
        "exit_log_records_processed": len(exit_df),
        "current_dashboard_row_count": dashboard_count,
        "history_total_row_count_after_update": history_count,
        "master_row_count_after_update": master_total,
        "new_client_project_rows": new_rows,
        "returning_client_project_rows": returning_rows,
        "not_present_in_latest_snapshot_count": not_present,
        "master_exited_count": exited_count,
        "master_unknown_count": unknown_count,
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
        vacant_slot_count=vacant_slot_count,
        exit_log_count=len(exit_df),
        history_count=history_count,
        master_total=master_total,
        exited_count=exited_count,
        unknown_count=unknown_count,
        dup_flag_count=dup_flag_count,
        missing_req=missing_req,
        output_dir=output_dir,
    )

    # ── Save settings ──────────────────────────────────────────────────────
    save_settings(
        {
            "last_input_dir": str(input_path.parent.resolve()),
            "last_output_dir": str(output_dir.resolve()),
            "last_sheet_name": sheet_name,
        },
        script_dir,
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
        "vacant_slot_rows_injected": "",
        "exit_log_records_processed": "",
        "current_dashboard_row_count": "",
        "history_total_row_count_after_update": "",
        "master_row_count_after_update": "",
        "new_client_project_rows": "",
        "returning_client_project_rows": "",
        "not_present_in_latest_snapshot_count": "",
        "master_exited_count": "",
        "master_unknown_count": "",
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
    vacant_slot_count: int,
    exit_log_count: int,
    history_count: int,
    master_total: int,
    exited_count: int,
    unknown_count: int,
    dup_flag_count: int,
    missing_req: list,
    output_dir: Path,
) -> None:
    sep = "-" * 62
    print(f"\n{sep}")
    print("  RUN SUMMARY")
    print(sep)
    print(f"  Input file           :  {source_file_name}")
    print(f"  Sheet                :  {sheet_name}")
    print(f"  Snapshot date        :  {snapshot_date}")
    print(f"  Dashboard rows       :  {dashboard_count}")
    print(f"    Vacant slot rows   :  {vacant_slot_count}")
    print(f"  Exit log records     :  {exit_log_count}")
    print(f"  History rows         :  {history_count}")
    print(f"  Master rows          :  {master_total}")
    print(f"    Exited             :  {exited_count}")
    print(f"    Unknown            :  {unknown_count}")
    print(f"  Duplicate flags      :  {dup_flag_count}")
    if missing_req:
        print(f"  Missing required     :  {', '.join(missing_req)}")
    print(f"  Output folder        :  {output_dir.resolve()}")
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
