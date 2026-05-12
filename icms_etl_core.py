"""
icms_etl_core.py
~~~~~~~~~~~~~~~~
Core ETL logic for the ICMS weekly snapshot pipeline.

All functions are designed to be called by the CLI runner
(update_icms_snapshot_files.py) and, in the future, by a Tkinter GUI
without any modification to this module.
"""

from __future__ import annotations

import io
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pandas as pd

# ── Field catalogue ───────────────────────────────────────────────────────────

METADATA_FIELDS: List[str] = [
    "snapshot_date",
    "reporting_month",
    "source_file_name",
]

REQUIRED_KEY_FIELDS: List[str] = [
    "CHAMP ID",
    "ICMS Project Name",
    "HMIS ID",
    "First Name",
    "Last Name",
    "ICMS Check-In Date",
    "ICMS Provider",
    "PH Building Name",
    "Current ICMS Status",
    "ICMS Case Manager",
    "ICMS Supervisor",
]

REQUIRED_QUEUE_FIELDS: List[str] = [
    "PH Update Still Needed This Month?",
    "5x5 Still Needed This Month?",
    "HAI Still Needed This Month?",
    "Care Plan Pending Action Step Still Needed This Month?",
    "Care Plan Completed Action Step Still Needed This Month?",
    "Eligible Services Still Needed This Month?",
    "Housing Navigation Action Step Needed?",
    "Health Care Action Step Needed?",
    "Key Housing Milestone Data Needed?",
]

REQUIRED_FIELDS: List[str] = REQUIRED_KEY_FIELDS + REQUIRED_QUEUE_FIELDS

OPTIONAL_FIELDS: List[str] = [
    "ICMS Slot #",
    "Medi-Cal CIN",
    "ICMS Enrollment ID",
    "ICMS Billing Rate",
    "Permanently Housed?",
    "Any Required Documentation Still Needed This Month?",
    "Notification of Privacy Practices On Record?",
    "Universal Sharing Consent Status",
    "CalAIM CS Housing Navigation (HN) and Tenancy Sustaining Services (TSS) Verbal Opt-In Status Missing?",
    "Family Billing Rate But Dependent Household Member(s) Not Listed in Household Composition?",
    "Eligible Services Recorded for This Month",
    "Universal Housing Application (UHA) Status",
]

ALL_APPROVED_FIELDS: List[str] = REQUIRED_FIELDS + OPTIONAL_FIELDS

MASTER_TRACKING_FIELDS: List[str] = [
    "first_seen_snapshot_date",
    "last_seen_snapshot_date",
    "last_seen_source_file",
    "present_in_latest_snapshot",
    "days_since_last_seen",
]

KEY_FIELDS: List[str] = ["CHAMP ID", "ICMS Project Name"]

ALIASES_FILENAME = "field_aliases.json"

# Columns written to the duplicate warning file
_WARNING_DATA_COLS: List[str] = [
    "CHAMP ID",
    "ICMS Project Name",
    "ICMS Enrollment ID",
    "ICMS Provider",
    "PH Building Name",
    "ICMS Case Manager",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe_str(val) -> str:
    """Convert a cell value to a clean string; return '' for null/NaN."""
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()


def _strip_internal_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Drop any column whose name starts with '_' before writing outputs."""
    return df.drop(
        columns=[c for c in df.columns if str(c).startswith("_")],
        errors="ignore",
    )


# ── Workbook opening ──────────────────────────────────────────────────────────

def open_workbook(
    path: str | Path,
    password: Optional[str] = None,
) -> Tuple[openpyxl.Workbook, bool]:
    """Open a normal or password-protected Excel workbook.

    Returns
    -------
    (workbook, was_password_protected)

    Password handling
    -----------------
    - If ``password`` is supplied it is used directly (GUI path).
    - If the file is encrypted and ``password`` is None, the user is
      prompted securely via ``getpass`` (CLI path).
    - Decryption is performed entirely in memory via msoffcrypto-tool;
      no decrypted copy is written to disk and the password is not stored.

    Raises
    ------
    FileNotFoundError
        If the input file does not exist.
    RuntimeError
        If the file cannot be opened or decrypted.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    # Attempt to import msoffcrypto (optional dependency)
    try:
        import msoffcrypto  # type: ignore
        _has_msoffcrypto = True
    except ImportError:
        msoffcrypto = None
        _has_msoffcrypto = False

    # ── Detect encryption ─────────────────────────────────────────────────
    is_encrypted = False
    if _has_msoffcrypto:
        try:
            with open(path, "rb") as fh:
                of = msoffcrypto.OfficeFile(fh)
                is_encrypted = of.is_encrypted()
        except Exception as exc:
            raise RuntimeError(f"Could not inspect workbook: {exc}") from exc
    else:
        # Without msoffcrypto, attempt a direct open; openpyxl will raise
        # on encrypted files and we surface a clear install message.
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            return wb, False
        except Exception:
            raise RuntimeError(
                "The workbook could not be opened.  It may be password-protected.\n"
                "Install the decryption library with:\n"
                "    pip install msoffcrypto-tool"
            )

    # ── Plain (unencrypted) file ───────────────────────────────────────────
    if not is_encrypted:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            return wb, False
        except Exception as exc:
            raise RuntimeError(f"Could not open workbook: {exc}") from exc

    # ── Encrypted file ────────────────────────────────────────────────────
    if password is None:
        import getpass
        password = getpass.getpass(f"Password for '{path.name}': ")

    try:
        with open(path, "rb") as fh:
            of = msoffcrypto.OfficeFile(fh)
            of.load_key(password=password)
            buf = io.BytesIO()
            of.decrypt(buf)
        buf.seek(0)
        wb = openpyxl.load_workbook(buf, data_only=True)
        return wb, True
    except Exception as exc:
        raise RuntimeError(
            f"Could not decrypt workbook (wrong password?): {exc}"
        ) from exc


# ── Sheet helpers ─────────────────────────────────────────────────────────────

def list_sheet_names(wb: openpyxl.Workbook) -> List[str]:
    """Return the list of sheet names in the workbook."""
    return list(wb.sheetnames)


def read_sheet(wb: openpyxl.Workbook, sheet_name: str) -> pd.DataFrame:
    """Read a worksheet into a DataFrame.

    - Row 1 is treated as the header row; rows 2+ are data.
    - Adds a ``_source_row_number`` column recording the 1-based Excel row
      number of each data row (useful for duplicate warnings).
    - Strips whitespace from all column headers and string cell values.
    - Drops rows that are completely blank across all data columns.

    Raises
    ------
    ValueError
        If ``sheet_name`` is not present in the workbook.
    """
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[sheet_name]
    all_rows = list(ws.values)
    if not all_rows:
        return pd.DataFrame()

    # Build clean header list
    raw_headers = all_rows[0]
    headers: List[str] = []
    for i, h in enumerate(raw_headers):
        if h is not None and str(h).strip():
            headers.append(str(h).strip())
        else:
            headers.append(f"_unnamed_col_{i}")

    data_rows = all_rows[1:]
    df = pd.DataFrame(data_rows, columns=headers)

    # Excel row numbers: header = row 1, first data row = row 2
    df["_source_row_number"] = range(2, len(df) + 2)

    # Strip string cell values
    for col in df.columns:
        if col == "_source_row_number":
            continue
        df[col] = df[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )

    # Drop completely blank rows (ignoring the internal tracking column)
    data_cols = [c for c in df.columns if c != "_source_row_number"]
    df = df.dropna(subset=data_cols, how="all").reset_index(drop=True)

    return df


# ── Alias management ──────────────────────────────────────────────────────────

def load_aliases(script_dir: Path) -> Dict[str, str]:
    """Load ``field_aliases.json`` from the script directory.

    Returns an empty dict if the file does not exist.
    """
    alias_path = script_dir / ALIASES_FILENAME
    if not alias_path.exists():
        return {}
    try:
        with open(alias_path, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
        return {str(k): str(v) for k, v in raw.items()}
    except Exception as exc:
        print(
            f"WARNING: Could not load {ALIASES_FILENAME}: {exc}. "
            "Proceeding with empty aliases."
        )
        return {}


def save_aliases(aliases: Dict[str, str], script_dir: Path) -> None:
    """Persist the alias dict to ``field_aliases.json`` in the script directory."""
    alias_path = script_dir / ALIASES_FILENAME
    try:
        with open(alias_path, "w", encoding="utf-8") as fh:
            json.dump(aliases, fh, indent=2, ensure_ascii=False)
    except Exception as exc:
        print(f"WARNING: Could not save {ALIASES_FILENAME}: {exc}")


# ── Field resolution ──────────────────────────────────────────────────────────

def resolve_fields(
    df_columns: List[str],
    aliases: Dict[str, str],
) -> Tuple[Dict[str, str], List[str], List[str]]:
    """Resolve canonical field names to actual workbook column names.

    Resolution order for each expected field:
      1. Exact canonical name present in workbook headers.
      2. Saved alias in ``aliases`` whose value is present in workbook headers.
      3. Unresolved — caller must handle (prompt user or accept the gap).

    Parameters
    ----------
    df_columns:
        The stripped column names from the workbook sheet.
    aliases:
        Mapping of canonical field name → known alias column name.

    Returns
    -------
    resolved       : {canonical_name: actual_column_name}
    unresolved_req : required canonical names that could not be resolved
    unresolved_opt : optional canonical names that could not be resolved
    """
    col_set = set(df_columns)
    resolved: Dict[str, str] = {}
    unresolved_req: List[str] = []
    unresolved_opt: List[str] = []

    for field in ALL_APPROVED_FIELDS:
        if field in col_set:
            resolved[field] = field
        elif aliases.get(field) and aliases[field] in col_set:
            resolved[field] = aliases[field]
        else:
            if field in REQUIRED_FIELDS:
                unresolved_req.append(field)
            else:
                unresolved_opt.append(field)

    return resolved, unresolved_req, unresolved_opt


def prompt_for_missing_field(
    canonical: str,
    df_columns: List[str],
    required: bool,
) -> Optional[str]:
    """Interactively prompt the user to map an unresolved field.

    - Prints the missing canonical name and the full list of available columns.
    - Validates the typed name against ``df_columns`` before accepting.
    - For **required** fields: displays a blocking alert and requires the user
      to type ``continue`` before the run proceeds without the field.
    - For **optional** fields: accepts Enter silently with no alert.

    Returns the chosen column name, or ``None`` if the user chose to skip.

    This function is isolated so a future GUI can replace it with a dialog
    without touching any other logic.
    """
    label = "[REQUIRED]" if required else "[OPTIONAL]"
    print(f"\n{label} Field not found in workbook: '{canonical}'")
    print("Available columns:")
    for col in df_columns:
        print(f"    {col}")

    while True:
        response = input(
            f"\nType the column name to use for '{canonical}', "
            "or press Enter to skip: "
        ).strip()

        if response == "":
            if required:
                print(
                    f"\n  !! ALERT: '{canonical}' is a REQUIRED field and "
                    "will be ABSENT from this run's output. !!"
                )
                confirm = input(
                    "  Type  continue  to proceed without this field, "
                    "or press Enter to try again: "
                ).strip()
                if confirm.lower() == "continue":
                    return None
                # User pressed Enter → loop back and ask again
            else:
                return None

        elif response in df_columns:
            return response

        else:
            print(f"  '{response}' is not a recognised column name. Please try again.")


def resolve_fields_interactive(
    df_columns: List[str],
    aliases: Dict[str, str],
    script_dir: Path,
) -> Tuple[Dict[str, str], List[str], List[str], int]:
    """Resolve all approved fields, prompting interactively for missing ones.

    Persists any new aliases discovered during this run to
    ``field_aliases.json``.

    Returns
    -------
    resolved      : {canonical_name: actual_column_name}
    missing_req   : required fields still absent after user interaction
    missing_opt   : optional fields absent (not in workbook; user skipped)
    aliases_added : count of new aliases saved in this run
    """
    resolved, unresolved_req, unresolved_opt = resolve_fields(df_columns, aliases)

    aliases_added = 0
    missing_req: List[str] = []
    missing_opt: List[str] = []

    for field in unresolved_req:
        chosen = prompt_for_missing_field(field, df_columns, required=True)
        if chosen is not None:
            resolved[field] = chosen
            if aliases.get(field) != chosen:
                aliases[field] = chosen
                aliases_added += 1
        else:
            missing_req.append(field)

    for field in unresolved_opt:
        chosen = prompt_for_missing_field(field, df_columns, required=False)
        if chosen is not None:
            resolved[field] = chosen
            if aliases.get(field) != chosen:
                aliases[field] = chosen
                aliases_added += 1
        else:
            missing_opt.append(field)

    if aliases_added > 0:
        save_aliases(aliases, script_dir)

    return resolved, missing_req, missing_opt, aliases_added


# ── Data reduction ────────────────────────────────────────────────────────────

def reduce_to_approved_fields(
    df: pd.DataFrame,
    resolved: Dict[str, str],
    snapshot_date: str,
    source_file_name: str,
) -> pd.DataFrame:
    """Reduce the raw sheet DataFrame to the approved field set.

    - Renames actual workbook column names to their canonical equivalents.
    - Inserts the three metadata columns at the front.
    - Preserves the internal ``_source_row_number`` column for downstream
      duplicate-warning logic; output writers strip it before saving.
    """
    # actual_column → canonical_name
    rename_map = {actual: canonical for canonical, actual in resolved.items()}

    # Select only resolved columns (+ internal row tracker if present)
    cols_to_keep = [c for c in resolved.values() if c in df.columns]
    if "_source_row_number" in df.columns:
        cols_to_keep = cols_to_keep + ["_source_row_number"]

    df_out = df[cols_to_keep].copy()
    df_out = df_out.rename(columns=rename_map)

    # Prepend metadata (inserted at position 0, so reverse order)
    reporting_month = snapshot_date[:7]  # YYYY-MM
    df_out.insert(0, "source_file_name", source_file_name)
    df_out.insert(0, "reporting_month", reporting_month)
    df_out.insert(0, "snapshot_date", snapshot_date)

    return df_out


# ── Duplicate detection ───────────────────────────────────────────────────────

def find_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Return all rows involved in CHAMP ID + ICMS Project Name duplicates.

    Both the first occurrence and every subsequent occurrence are included.
    Returns an empty DataFrame when no duplicates exist or when the key
    columns are absent (e.g., if they were skipped during field resolution).
    """
    missing_keys = [k for k in KEY_FIELDS if k not in df.columns]
    if missing_keys:
        return pd.DataFrame()

    mask = df.duplicated(subset=KEY_FIELDS, keep=False)
    return df[mask].copy()


def build_duplicate_warning_rows(
    dup_df: pd.DataFrame,
    run_timestamp: str,
    snapshot_date: str,
    source_file_name: str,
) -> pd.DataFrame:
    """Format duplicate rows into the warning-file schema.

    Parameters
    ----------
    dup_df:
        Output of ``find_duplicates()`` — still has ``_source_row_number``.
    """
    warning_schema_cols = [
        "run_timestamp",
        "snapshot_date",
        "source_file_name",
        "CHAMP ID",
        "ICMS Project Name",
        "ICMS Enrollment ID",
        "ICMS Provider",
        "PH Building Name",
        "ICMS Case Manager",
        "source_row_number",
    ]

    rows = []
    for _, row in dup_df.iterrows():
        rows.append(
            {
                "run_timestamp": run_timestamp,
                "snapshot_date": snapshot_date,
                "source_file_name": source_file_name,
                "CHAMP ID": _safe_str(row.get("CHAMP ID")),
                "ICMS Project Name": _safe_str(row.get("ICMS Project Name")),
                "ICMS Enrollment ID": _safe_str(row.get("ICMS Enrollment ID")),
                "ICMS Provider": _safe_str(row.get("ICMS Provider")),
                "PH Building Name": _safe_str(row.get("PH Building Name")),
                "ICMS Case Manager": _safe_str(row.get("ICMS Case Manager")),
                "source_row_number": _safe_str(row.get("_source_row_number")),
            }
        )

    return pd.DataFrame(rows, columns=warning_schema_cols)


# ── Output file writers ───────────────────────────────────────────────────────

def update_dashboard_flatfile(df: pd.DataFrame, output_dir: Path) -> None:
    """Overwrite ``icms_current_dashboard_flatfile.xlsx`` completely.

    Raises
    ------
    RuntimeError
        On any I/O failure.
    """
    path = output_dir / "icms_current_dashboard_flatfile.xlsx"
    try:
        _strip_internal_cols(df).to_excel(path, index=False, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(f"Could not write dashboard flat file: {exc}") from exc


def update_snapshot_history(
    df: pd.DataFrame,
    snapshot_date: str,
    output_dir: Path,
) -> int:
    """Append incoming rows to ``icms_snapshot_history.csv``.

    Same-date reload behaviour: all existing rows for ``snapshot_date``
    are removed before the incoming rows are appended, making each
    snapshot date fully idempotent.

    Returns
    -------
    Total row count in the file after the update.

    Raises
    ------
    RuntimeError
        On write failure.
    """
    path = output_dir / "icms_snapshot_history.csv"
    clean_df = _strip_internal_cols(df).copy()
    clean_df = clean_df.astype(str).replace({"nan": "", "None": "", "<NA>": ""})

    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8").fillna("")
            existing = existing[existing["snapshot_date"] != snapshot_date]
        except Exception as exc:
            print(
                f"WARNING: Could not read existing snapshot history: {exc}. "
                "Starting fresh."
            )
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()

    combined = pd.concat([existing, clean_df], ignore_index=True)

    try:
        combined.to_csv(path, index=False, encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Could not write snapshot history: {exc}") from exc

    return len(combined)


def update_client_project_master(
    df: pd.DataFrame,
    snapshot_date: str,
    source_file_name: str,
    output_dir: Path,
) -> Tuple[int, int, int, int]:
    """Update the cumulative client/project master file.

    The master file grain is one row per ``CHAMP ID + ICMS Project Name``.
    When a client has duplicate rows in the incoming snapshot the last
    (bottom-most) row is used to update the master record; all duplicate
    occurrences are logged separately via ``append_duplicate_warning``.

    Parameters
    ----------
    df:
        The reduced and renamed DataFrame from ``reduce_to_approved_fields``.
    snapshot_date:
        The ``YYYY-MM-DD`` snapshot date string.
    source_file_name:
        The basename of the input file.
    output_dir:
        Directory where the master file lives.

    Returns
    -------
    (total_rows, new_rows, returning_rows, not_present_count)

    Raises
    ------
    RuntimeError
        On write failure.
    """
    path = output_dir / "icms_client_project_master.xlsx"

    # Strip internal tracking columns and normalise to string
    incoming = _strip_internal_cols(df).copy()
    incoming = incoming.astype(str).replace({"nan": "", "None": "", "<NA>": ""})

    # Data columns present in this run (excluding metadata)
    current_data_fields = [c for c in incoming.columns if c not in METADATA_FIELDS]

    # ── Load existing master ───────────────────────────────────────────────
    if path.exists():
        try:
            master = pd.read_excel(path, engine="openpyxl", dtype=str).fillna("")
        except Exception as exc:
            print(
                f"WARNING: Could not read existing master: {exc}. Starting fresh."
            )
            master = pd.DataFrame()
    else:
        master = pd.DataFrame()

    # ── Build deduplicated view of incoming snapshot (last row wins per key) ─
    incoming_deduped = (
        incoming.groupby(KEY_FIELDS, sort=False).last().reset_index()
    )
    incoming_key_set: set = set(
        zip(
            incoming_deduped["CHAMP ID"].tolist(),
            incoming_deduped["ICMS Project Name"].tolist(),
        )
    )

    # ── Build master from scratch on first run ─────────────────────────────
    if master.empty:
        new_count = len(incoming_deduped)
        returning_count = 0
        not_present_count = 0

        master_out = incoming_deduped.copy()
        master_out["first_seen_snapshot_date"] = snapshot_date
        master_out["last_seen_snapshot_date"] = snapshot_date
        master_out["last_seen_source_file"] = source_file_name
        master_out["present_in_latest_snapshot"] = "True"
        master_out["days_since_last_seen"] = "0"

        try:
            master_out.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            raise RuntimeError(
                f"Could not write client/project master: {exc}"
            ) from exc

        return len(master_out), new_count, returning_count, not_present_count

    # ── Incremental update ─────────────────────────────────────────────────
    existing_key_set: set = set(
        zip(
            master["CHAMP ID"].tolist(),
            master["ICMS Project Name"].tolist(),
        )
    )

    new_keys = incoming_key_set - existing_key_set
    returning_keys = incoming_key_set & existing_key_set
    absent_keys = existing_key_set - incoming_key_set

    new_count = len(new_keys)
    returning_count = len(returning_keys)
    not_present_count = len(absent_keys)

    # Ensure every current data field has a column in master
    for field in current_data_fields + MASTER_TRACKING_FIELDS:
        if field not in master.columns:
            master[field] = ""

    # Update existing master rows
    for idx in master.index:
        champ = str(master.at[idx, "CHAMP ID"])
        proj = str(master.at[idx, "ICMS Project Name"])
        key = (champ, proj)

        if key in returning_keys:
            # Pull latest known values from the deduped incoming snapshot
            match = incoming_deduped[
                (incoming_deduped["CHAMP ID"] == champ)
                & (incoming_deduped["ICMS Project Name"] == proj)
            ]
            if not match.empty:
                src = match.iloc[0]
                for field in current_data_fields:
                    if field in src.index:
                        master.at[idx, field] = _safe_str(src[field])

            if not master.at[idx, "first_seen_snapshot_date"]:
                master.at[idx, "first_seen_snapshot_date"] = snapshot_date
            master.at[idx, "last_seen_snapshot_date"] = snapshot_date
            master.at[idx, "last_seen_source_file"] = source_file_name
            master.at[idx, "present_in_latest_snapshot"] = "True"
            master.at[idx, "days_since_last_seen"] = "0"

        elif key in absent_keys:
            # Null out fields not present in this run's approved set
            for field in master.columns:
                if (
                    field not in METADATA_FIELDS
                    and field not in MASTER_TRACKING_FIELDS
                    and field not in KEY_FIELDS
                    and field not in current_data_fields
                ):
                    master.at[idx, field] = ""

            master.at[idx, "present_in_latest_snapshot"] = "False"

            # Recalculate days since last seen
            try:
                last_seen_dt = pd.to_datetime(
                    master.at[idx, "last_seen_snapshot_date"], errors="raise"
                )
                snap_dt = pd.to_datetime(snapshot_date)
                days = int((snap_dt - last_seen_dt).days)
                master.at[idx, "days_since_last_seen"] = str(days)
            except Exception:
                master.at[idx, "days_since_last_seen"] = ""

    # Append rows for brand-new client/project combinations
    if new_keys:
        new_rows_list = []
        for key in new_keys:
            champ, proj = key
            match = incoming_deduped[
                (incoming_deduped["CHAMP ID"] == champ)
                & (incoming_deduped["ICMS Project Name"] == proj)
            ]
            if not match.empty:
                new_row = match.iloc[0].to_dict()
                new_row["first_seen_snapshot_date"] = snapshot_date
                new_row["last_seen_snapshot_date"] = snapshot_date
                new_row["last_seen_source_file"] = source_file_name
                new_row["present_in_latest_snapshot"] = "True"
                new_row["days_since_last_seen"] = "0"
                new_rows_list.append(new_row)

        if new_rows_list:
            new_df = pd.DataFrame(new_rows_list)
            # Align columns to master before concatenating
            for col in master.columns:
                if col not in new_df.columns:
                    new_df[col] = ""
            new_df = new_df.reindex(columns=master.columns, fill_value="")
            master = pd.concat([master, new_df], ignore_index=True)

    try:
        master.to_excel(path, index=False, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(
            f"Could not write client/project master: {exc}"
        ) from exc

    return len(master), new_count, returning_count, not_present_count


def append_ingest_log(log_row: Dict, output_dir: Path) -> None:
    """Append a run record to ``icms_ingest_log.csv``.

    Same-date reruns replace the previous row for ``snapshot_date`` so
    each date has at most one log entry.

    Raises
    ------
    RuntimeError
        On write failure.
    """
    path = output_dir / "icms_ingest_log.csv"
    snap_date = str(log_row.get("snapshot_date", ""))

    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8").fillna("")
            existing = existing[existing["snapshot_date"] != snap_date]
        except Exception as exc:
            print(
                f"WARNING: Could not read existing ingest log: {exc}. "
                "Starting fresh."
            )
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()

    new_entry = pd.DataFrame([log_row])
    combined = pd.concat([existing, new_entry], ignore_index=True)

    try:
        combined.to_csv(path, index=False, encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Could not write ingest log: {exc}") from exc


def append_duplicate_warning(
    warning_df: pd.DataFrame,
    snapshot_date: str,
    output_dir: Path,
) -> None:
    """Append duplicate-client rows to ``icms_duplicate_client_project_warning.csv``.

    Same-date reruns replace all rows for ``snapshot_date``.
    No-ops when ``warning_df`` is empty.

    Raises
    ------
    RuntimeError
        On write failure.
    """
    if warning_df.empty:
        return

    path = output_dir / "icms_duplicate_client_project_warning.csv"

    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8").fillna("")
            existing = existing[existing["snapshot_date"] != snapshot_date]
        except Exception as exc:
            print(
                f"WARNING: Could not read existing duplicate warning file: {exc}. "
                "Starting fresh."
            )
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()

    clean_warn = warning_df.astype(str).replace(
        {"nan": "", "None": "", "<NA>": ""}
    )
    combined = pd.concat([existing, clean_warn], ignore_index=True)

    try:
        combined.to_csv(path, index=False, encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(
            f"Could not write duplicate warning file: {exc}"
        ) from exc
