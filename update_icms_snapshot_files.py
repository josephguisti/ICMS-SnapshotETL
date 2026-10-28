"""
update_icms_snapshot_files.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CLI runner for the ICMS weekly snapshot ETL.

Usage
-----
python update_icms_snapshot_files.py --input "ICMS Snapshot - 10.19.2026.xlsx"
python update_icms_snapshot_files.py --batch "./snapshots" --output-dir "./output"

The GUI is the primary interface. Both call icms_etl_core.run_snapshot_files;
field and sheet mappings live in icms_catalogue.py, so no prompts for field
mappings are ever shown.
"""

from __future__ import annotations

import argparse
import getpass
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import icms_etl_core as core


# ── Settings persistence ──────────────────────────────────────────────────────

def load_settings(script_dir: Path) -> dict:
    """Load settings.json from the script directory.

    Returns ``{}`` on any error — never raises.
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


# Kept for diagnostics/ scripts that import it from here.
_parse_date_from_filename = core.parse_date_from_filename


# ── Argument parsing ──────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="update_icms_snapshot_files",
        description=(
            "Import ICMS Excel snapshots and update all cumulative output files."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples\n"
            "--------\n"
            "  Single file (date defaults to the workbook's Generated On):\n"
            '    python update_icms_snapshot_files.py --input "ICMS Snapshot - 10.19.2026.xlsx"\n\n'
            "  Batch (reprocess files already imported, to repair history):\n"
            '    python update_icms_snapshot_files.py --batch "./snapshots" '
            '--output-dir "./output" --reprocess\n'
        ),
    )
    parser.add_argument("--input", metavar="FILE",
                        help="Snapshot workbook (.xlsx) for single-file mode.")
    parser.add_argument("--snapshot-date", metavar="YYYY-MM-DD",
                        help="Override the snapshot date (single-file mode).")
    parser.add_argument("--sheet", metavar="SHEET_NAME",
                        help="Participant sheet name if it differs from the catalogue names.")
    parser.add_argument("--output-dir", metavar="DIR",
                        help="Folder for output files. Defaults to the current directory.")
    parser.add_argument("--batch", metavar="DIR",
                        help="Folder of .xlsx snapshots to import oldest first.")
    parser.add_argument("--reprocess", action="store_true",
                        help="Reprocess files already imported without asking.")
    parser.add_argument("--no-publish", action="store_true",
                        help="Do not update the current dashboard flat file.")
    return parser


# ── Snapshot date gap check ───────────────────────────────────────────────────

def _check_snapshot_date_gap(snapshot_date: str, output_dir: Path) -> None:
    """Require confirmation if the snapshot is >45 days after the latest in history."""
    latest = core.latest_history_snapshot_date(output_dir)
    if not latest:
        return
    try:
        gap_days = (pd.to_datetime(snapshot_date) - pd.to_datetime(latest)).days
    except Exception:
        return
    if gap_days <= 45:
        return
    print(
        f"\n  WARNING: The snapshot date ({snapshot_date}) is {gap_days} days after "
        f"the most recent snapshot in history ({latest})."
    )
    print("  This may indicate a missed snapshot or an incorrect date.")
    resp = input("  Type 'confirm' to proceed anyway, or press Enter to abort: ").strip().lower()
    if resp != "confirm":
        print("  Aborted by user.")
        sys.exit(0)


# ── File and date confirmation ────────────────────────────────────────────────

def _default_spec(path: Path, password: str) -> core.FileSpec:
    info = core.default_snapshot_date(path, password)
    return core.FileSpec(path, info["date"] or "", info["source"] or "",
                         info["generated_on"], info["filename_date"])


def _confirm_batch_files(specs: list[core.FileSpec]) -> list[core.FileSpec]:
    """List files with their default dates and allow inline correction."""
    while True:
        specs.sort(key=lambda s: (not s.snapshot_date, s.snapshot_date, Path(s.path).name))
        print("\nBatch file list (oldest first):")
        for i, s in enumerate(specs, 1):
            label = s.snapshot_date or "NO DATE"
            note = s.date_source
            if s.generated_on and s.filename_date and s.generated_on != s.filename_date:
                note += f"; filename says {s.filename_date}"
            print(f"  {i:3}.  {label:<12}  {Path(s.path).name}  ({note})")

        choice = input(
            "\nEnter a file number to correct its date, "
            "'confirm' to proceed, or 'abort' to exit: "
        ).strip()
        if choice.lower() == "abort":
            sys.exit(0)
        if choice.lower() == "confirm":
            if any(not s.snapshot_date for s in specs):
                print("  Cannot confirm: enter a date for every file first.")
                continue
            return specs
        if choice.isdigit() and 0 < int(choice) <= len(specs):
            s = specs[int(choice) - 1]
            while True:
                new_date = input(f"  Date for '{Path(s.path).name}' (YYYY-MM-DD): ").strip()
                try:
                    datetime.strptime(new_date, "%Y-%m-%d")
                    s.snapshot_date, s.date_source = new_date, "manual entry"
                    break
                except ValueError:
                    print("  Invalid format. Please use YYYY-MM-DD.")
        else:
            print("  Invalid input. Enter a number, 'confirm', or 'abort'.")


def _password_for(paths: list[Path]) -> str:
    if any(core.is_encrypted(p) for p in paths):
        return getpass.getpass("Workbook password: ")
    return ""


def _print_summary(results: dict) -> None:
    border = "\u2550" * 62
    print(f"\n{border}")
    print("  RUN COMPLETE" if results["success"] else "  RUN FAILED")
    print(f"  Files imported   : {results['batch_success']} of {results['batch_total']}")
    for f in results["batch_failed_files"]:
        print(f"  NOT IMPORTED     : {f['source_file_name']} ({f['snapshot_date']}): {f['error']}")
    if results["skipped_categories"]:
        print(f"  Tabs not imported: {', '.join(results['skipped_categories'])}")
    print("  Dashboard        : "
          + (f"{results['dashboard_count']} rows from {results['dashboard_source']}"
             if results["dashboard_published"] else results["dashboard_note"] or "not updated"))
    print(f"  Master           : {results['master_total']} enrollments "
          f"({results['master_exited_count']} exited, {results['master_unknown_count']} unknown)")
    if results["backup_dir"]:
        print(f"  Backup           : {results['backup_dir']}")
    print(f"  Output folder    : {results['output_dir']}")
    print(border)
    if results["findings_text"]:
        print("\nStructural findings (for the director; nothing to do now):")
        print(results["findings_text"])


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    args = _build_parser().parse_args()
    output_dir = Path(args.output_dir) if args.output_dir else Path.cwd()
    script_dir = Path(__file__).parent.resolve()
    sheet_names = {"participant": args.sheet} if args.sheet else None

    if args.batch:
        batch_dir = Path(args.batch)
        if not batch_dir.is_dir():
            print(f"ERROR: --batch path is not a directory: {batch_dir}")
            sys.exit(1)
        paths = sorted(batch_dir.glob("*.xlsx"))
        if not paths:
            print(f"\nNo .xlsx files found in: {batch_dir}")
            return
        password = _password_for(paths)
        specs = _confirm_batch_files([_default_spec(p, password) for p in paths])
        input_dir = batch_dir
    else:
        if not args.input:
            print("ERROR: --input is required when not using --batch.")
            sys.exit(1)
        path = Path(args.input)
        if not path.exists():
            print(f"ERROR: Input file not found: {path}")
            sys.exit(1)
        password = _password_for([path])
        spec = _default_spec(path, password)
        if args.snapshot_date:
            try:
                datetime.strptime(args.snapshot_date, "%Y-%m-%d")
            except ValueError:
                print(f"ERROR: --snapshot-date must be YYYY-MM-DD, got '{args.snapshot_date}'")
                sys.exit(1)
            if args.snapshot_date != spec.snapshot_date:
                spec.snapshot_date, spec.date_source = args.snapshot_date, "manual entry"
        if not spec.snapshot_date:
            print("ERROR: No date found in Generated On or the filename; pass --snapshot-date.")
            sys.exit(1)
        if spec.generated_on and spec.filename_date and spec.generated_on != spec.filename_date:
            print(f"NOTE: Generated On is {spec.generated_on}; filename says {spec.filename_date}.")
        print(f"Snapshot date: {spec.snapshot_date} ({spec.date_source})")
        output_dir.mkdir(parents=True, exist_ok=True)
        _check_snapshot_date_gap(spec.snapshot_date, output_dir)
        specs = [spec]
        input_dir = path.parent

    output_dir.mkdir(parents=True, exist_ok=True)
    reprocessing = False
    dups = core.already_imported(specs, output_dir)
    if dups:
        print(f"\n{len(dups)} file(s) were already imported into this folder:")
        for d in dups:
            print(f"  {d.snapshot_date}  {Path(d.path).name}")
        if args.reprocess:
            reprocessing = True
        else:
            resp = input("Type 'reprocess' to re-import them (repairs history; outputs "
                         "are backed up first), 'skip' to leave them out, or Enter to abort: "
                         ).strip().lower()
            if resp == "reprocess":
                reprocessing = True
            elif resp == "skip":
                specs = [s for s in specs if s not in dups]
                if not specs:
                    print("Nothing left to import.")
                    return
            else:
                print("Aborted by user.")
                return

    locked = core.check_outputs_writable(output_dir)
    if locked:
        print("ERROR: " + core.format_output_lock_message(locked))
        sys.exit(1)

    results = core.run_snapshot_files(
        specs, output_dir,
        password=password,
        sheet_names=sheet_names,
        publish_dashboard=not args.no_publish,
        reprocessing=reprocessing,
    )
    _print_summary(results)

    settings = load_settings(script_dir)
    settings.update({
        "last_input_dir": str(input_dir.resolve()),
        "last_output_dir": str(output_dir.resolve()),
    })
    save_settings(settings, script_dir)
    if not results["success"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
