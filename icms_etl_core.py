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
import math
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
    "Date of Birth",
    "First Name",
    "Last Name",
    "HMIS ID",
    "Medi-Cal CIN",
    "ICMS Check-In Date",
    "ICMS Provider",
    "PH Building Name",
    "Current ICMS Status",
    "ICMS Case Manager",
    "ICMS Supervisor",
    "ICMS Project Type",
    "ICMS Slot #",
    "Receiving Pre-Match ICMS?",
    "Checked in 6+ Months and Not Yet Permanently Housed?",
    "Permanent Housing Move-In Date",
    "Housing Voucher Issued Date",
    "Permanent Housing Subsidy",
    "Service Planning Area",
    "Household Size",
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
    "present_in_latest_snapshot",
    "client_status",
]

EXIT_FIELDS: List[str] = [
    "Check-In Date",
    "Check-Out Date",
    "Move-In Date",
    "Move-Out Date",
    "Exit Reason",
    "Exit Destination",
    "Exit Location",
    "Transferred to New Provider Within ICMS Program?",
    "New ICMS Provider",
]

SERVICES_LOG_FIELDS: List[str] = [
    "Row ID",
    "CHAMP ID",
    "DHS Master Agreement #",
    "ICMS Work Order #",
    "ICMS Provider Name",
    "PH Building Name",
    "ICMS Enrollment ID",
    "ICMS Case Manager",
    "Service ID",
    "Service",
    "Place of Service",
    "Service Date",
    "Service Record Created Date",
    "CHAMP User Who Created Record",
    "Organization of CHAMP User Who Created Record",
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


def find_header_row(
    wb: openpyxl.Workbook,
    sheet_name: str,
    candidate_fields: List[str],
    max_scan: int = 5,
) -> int:
    """Return the 0-based header row index that best matches candidate fields.

    Scans the first ``max_scan`` rows and returns the row with the highest
    overlap against ``candidate_fields``.  Falls back to 0 when no overlap is
    detected so existing single-row header behaviour is preserved.
    """
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[sheet_name]
    rows = list(ws.values)
    if not rows:
        return 0

    candidate_set = {str(f).strip() for f in candidate_fields if str(f).strip()}
    scan_limit = min(max_scan, len(rows))
    best_idx = 0
    best_score = -1

    for idx in range(scan_limit):
        row = rows[idx] or ()
        values = {
            str(v).strip()
            for v in row
            if v is not None and str(v).strip()
        }
        score = len(values & candidate_set)
        if score > best_score:
            best_idx = idx
            best_score = score

    if best_score <= 0:
        return 0
    return best_idx


def read_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    header_row: int = 0,
    *,
    stats: Optional[Dict[str, object]] = None,
) -> pd.DataFrame:
    """Read a worksheet into a DataFrame.

    - ``header_row`` determines which row is treated as the header row.
      Rows after that become data.
    - Adds a ``_source_row_number`` column recording the 1-based Excel row
      number of each data row (useful for duplicate warnings).
    - Strips whitespace from all column headers and string cell values.
    - Drops rows that are completely blank across all data columns.
    - Optionally writes lightweight diagnostics into ``stats``.

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
    if header_row < 0 or header_row >= len(all_rows):
        raise ValueError(
            f"Header row index {header_row} is out of range for sheet "
            f"'{sheet_name}' (rows: {len(all_rows)})."
        )

    # Build clean header list
    raw_headers = all_rows[header_row]
    headers: List[str] = []
    for i, h in enumerate(raw_headers):
        if h is not None and str(h).strip():
            headers.append(str(h).strip())
        else:
            headers.append(f"_unnamed_col_{i}")

    data_rows = all_rows[header_row + 1:]
    df = pd.DataFrame(data_rows, columns=headers)

    # Excel row numbers are 1-based; first data row follows header_row.
    first_data_row_num = header_row + 2
    df["_source_row_number"] = range(first_data_row_num, len(df) + first_data_row_num)

    if stats is not None:
        stats["header_row_detected-diag"] = header_row
        stats["participant_column_count-diag"] = len(headers)

    # Strip string cell values; normalise whole-number floats to integer
    # strings so that CHAMP IDs stored as numeric cells (e.g. 100300.0) match
    # those stored as text cells ("100300") across different source files.
    def _norm(x):
        if isinstance(x, str):
            return x.strip()
        if isinstance(x, float):
            if math.isnan(x) or math.isinf(x):
                return None
            if x == int(x):
                return str(int(x))
        return x

    key_float_norm_count = 0

    def _norm_key_field(x):
        nonlocal key_float_norm_count
        if isinstance(x, str):
            return x.strip()
        if isinstance(x, float):
            if math.isnan(x) or math.isinf(x):
                return None
            if x == int(x):
                key_float_norm_count += 1
                return str(int(x))
        return x

    for col in df.columns:
        if col == "_source_row_number":
            continue
        if col in KEY_FIELDS:
            df[col] = df[col].apply(_norm_key_field)
        else:
            df[col] = df[col].apply(_norm)

    # Drop completely blank rows (ignoring the internal tracking column)
    data_cols = [c for c in df.columns if c != "_source_row_number"]
    df = df.dropna(subset=data_cols, how="all").reset_index(drop=True)

    if stats is not None:
        stats["key_field_float_normalizations-diag"] = key_float_norm_count

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
    *,
    stats: Optional[Dict[str, object]] = None,
) -> pd.DataFrame:
    """Reduce the raw sheet DataFrame to the approved field set.

    - Renames actual workbook column names to their canonical equivalents.
    - Inserts the three metadata columns at the front.
    - Preserves the internal ``_source_row_number`` column for downstream
      duplicate-warning logic; output writers strip it before saving.
    - Drops rows where either key field (``CHAMP ID`` or
      ``ICMS Project Name``) is blank.
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

    dropped_blank_key_rows = 0
    if all(k in df_out.columns for k in KEY_FIELDS):
        key_df = (
            df_out[KEY_FIELDS]
            .fillna("")
            .astype(str)
            .apply(lambda s: s.str.strip())
        )
        keep_mask = key_df["CHAMP ID"].ne("") & key_df["ICMS Project Name"].ne("")
        dropped_blank_key_rows = int((~keep_mask).sum())
        if dropped_blank_key_rows > 0:
            df_out = df_out.loc[keep_mask].reset_index(drop=True)

    if stats is not None:
        stats["dropped_blank_key_rows-diag"] = dropped_blank_key_rows

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


# ── Vacant slots and exit log readers ────────────────────────────────────────

def read_vacant_slots(wb: openpyxl.Workbook) -> pd.DataFrame:
    """Read the ``Vacant Slots`` sheet from an already-open workbook.

    Retains only: ``ICMS Project Name``, ``ICMS Provider``, ``PH Building Name``,
    ``ICMS Project Type``, ``# Days Vacant``.

    Returns an empty DataFrame with a printed warning if the sheet is missing.
    """
    sheet_name = "Vacant Slots"
    if sheet_name not in wb.sheetnames:
        print(
            f"WARNING: Sheet '{sheet_name}' not found in workbook. "
            "Vacant slot data will be skipped."
        )
        return pd.DataFrame()

    df = read_sheet(wb, sheet_name)
    keep_cols = [
        "ICMS Project Name",
        "ICMS Provider",
        "PH Building Name",
        "ICMS Project Type",
        "# Days Vacant",
    ]
    available = [c for c in keep_cols if c in df.columns]
    return df[available].copy()


def build_vacant_slot_rows(
    vacant_df: pd.DataFrame,
    snapshot_date: str,
    source_file_name: str,
) -> pd.DataFrame:
    """Convert a vacant slots DataFrame into synthetic flat-file rows.

    Adds metadata columns, sets ``is_vacant_slot = True``, and fills every
    approved participant field not present in the Vacant Slots sheet with
    ``None``.  Returns one row per vacant slot record.
    """
    if vacant_df.empty:
        return pd.DataFrame()

    out = vacant_df.copy()
    out["snapshot_date"] = snapshot_date
    out["reporting_month"] = snapshot_date[:7]
    out["source_file_name"] = source_file_name
    out["is_vacant_slot"] = True

    for field in ALL_APPROVED_FIELDS:
        if field not in out.columns:
            out[field] = None

    return out


def read_exit_log(wb: openpyxl.Workbook) -> pd.DataFrame:
    """Read the ``Exit Log`` sheet from an already-open workbook.

    Retains these columns where present: ``CHAMP ID``, ``ICMS Project Name``,
    ``ICMS Provider``, ``PH Building Name``, ``ICMS Project Type``,
    ``ICMS Enrollment ID``, ``ICMS Case Manager``, ``First Name``,
    ``Last Name``, ``Date of Birth``, ``HMIS ID``, ``Medi-Cal CIN``,
    ``Check-In Date``, ``Check-Out Date``, ``Move-In Date``, ``Move-Out Date``,
    ``Exit Reason``, ``Exit Destination``, ``Exit Location``,
    ``Transferred to New Provider Within ICMS Program?``, ``New ICMS Provider``.

    Returns an empty DataFrame with a printed warning if the sheet is missing.
    """
    sheet_name = "Exit Log"
    if sheet_name not in wb.sheetnames:
        print(
            f"WARNING: Sheet '{sheet_name}' not found in workbook. "
            "Exit log data will be skipped."
        )
        return pd.DataFrame()

    df = read_sheet(wb, sheet_name)
    keep_cols = [
        "CHAMP ID",
        "ICMS Project Name",
        "ICMS Provider",
        "PH Building Name",
        "ICMS Project Type",
        "ICMS Enrollment ID",
        "ICMS Case Manager",
        "First Name",
        "Last Name",
        "Date of Birth",
        "HMIS ID",
        "Medi-Cal CIN",
        "Check-In Date",
        "Check-Out Date",
        "Move-In Date",
        "Move-Out Date",
        "Exit Reason",
        "Exit Destination",
        "Exit Location",
        "Transferred to New Provider Within ICMS Program?",
        "New ICMS Provider",
    ]
    available = [c for c in keep_cols if c in df.columns]
    return df[available].copy()


def read_services_log(
    wb: openpyxl.Workbook,
    sheet_name: str = "Services Log",
) -> pd.DataFrame:
    """Read the Services Log sheet from an already-open workbook.

    Retains known services-log columns where present.
    Returns an empty DataFrame with a warning if the sheet is missing.
    """
    if sheet_name not in wb.sheetnames:
        print(
            f"WARNING: Sheet '{sheet_name}' not found in workbook. "
            "Services log data will be skipped."
        )
        return pd.DataFrame()

    df = read_sheet(wb, sheet_name)
    available = [c for c in SERVICES_LOG_FIELDS if c in df.columns]
    return df[available].copy()


# ── Output file writers ───────────────────────────────────────────────────────

def update_dashboard_flatfile(
    df: pd.DataFrame,
    output_dir: Path,
    vacant_rows: pd.DataFrame,
) -> None:
    """Overwrite ``icms_current_dashboard_flatfile.xlsx`` completely.

    All participant rows receive ``is_vacant_slot = False``.  If
    ``vacant_rows`` is non-empty it is concatenated (columns are aligned;
    any missing columns are filled with ``None``).

    Raises
    ------
    RuntimeError
        On any I/O failure.
    """
    path = output_dir / "icms_current_dashboard_flatfile.xlsx"
    out = _strip_internal_cols(df).copy()
    out["is_vacant_slot"] = False

    if not vacant_rows.empty:
        combined = pd.concat([out, vacant_rows], ignore_index=True, sort=False)
    else:
        combined = out

    try:
        combined.to_excel(path, index=False, engine="openpyxl")
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


def update_exit_log_history(
    exit_df: pd.DataFrame,
    snapshot_date: str,
    source_file_name: str,
    output_dir: Path,
) -> int:
    """Append exit-log rows to ``icms_exit_log_history.csv``.

    Same file/date reruns replace previously written rows for the same
    ``snapshot_date`` + ``source_file_name`` combination.

    Returns total row count in the file after update.
    """
    path = output_dir / "icms_exit_log_history.csv"

    if exit_df.empty:
        if path.exists():
            try:
                return len(pd.read_csv(path, dtype=str, encoding="utf-8"))
            except Exception:
                return 0
        return 0

    clean_exit = _strip_internal_cols(exit_df).copy()
    clean_exit = clean_exit.astype(str).replace({"nan": "", "None": "", "<NA>": ""})
    clean_exit.insert(0, "source_file_name", source_file_name)
    clean_exit.insert(0, "reporting_month", snapshot_date[:7])
    clean_exit.insert(0, "snapshot_date", snapshot_date)

    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8").fillna("")
            if (
                "snapshot_date" in existing.columns
                and "source_file_name" in existing.columns
            ):
                existing = existing[
                    ~(
                        (existing["snapshot_date"] == snapshot_date)
                        & (existing["source_file_name"] == source_file_name)
                    )
                ]
        except Exception as exc:
            print(
                f"WARNING: Could not read existing exit log history: {exc}. "
                "Starting fresh."
            )
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()

    combined = pd.concat([existing, clean_exit], ignore_index=True, sort=False)
    try:
        combined.to_csv(path, index=False, encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Could not write exit log history: {exc}") from exc

    return len(combined)


def _build_services_event_key_df(df: pd.DataFrame) -> pd.DataFrame:
    """Return deterministic service-event keys for deduplication/comparison."""
    work = df.copy()
    for col in SERVICES_LOG_FIELDS:
        if col not in work.columns:
            work[col] = ""
    for col in SERVICES_LOG_FIELDS:
        work[col] = work[col].fillna("").astype(str).str.strip()

    service_id = work["Service ID"]
    has_service_id = service_id.ne("")
    fallback = (
        work["CHAMP ID"] + "|" +
        work["ICMS Enrollment ID"] + "|" +
        work["Service"] + "|" +
        work["Place of Service"] + "|" +
        work["Service Date"] + "|" +
        work["Service Record Created Date"]
    )
    event_key = service_id.where(has_service_id, "__fallback__|" + fallback)
    return pd.DataFrame(
        {
            "service_event_key": event_key,
            "has_service_id": has_service_id.astype(bool),
        }
    )


def summarize_services_overlap(
    services_df: pd.DataFrame,
    output_dir: Path,
) -> Dict[str, float]:
    """Summarise incoming services-table overlap against cumulative history."""
    summary: Dict[str, float] = {
        "incoming_event_key_count": 0,
        "existing_overlap_rate": 0.0,
    }
    if services_df.empty:
        return summary

    incoming_keys_df = _build_services_event_key_df(services_df)
    incoming_keys = set(incoming_keys_df["service_event_key"].astype(str))
    summary["incoming_event_key_count"] = int(len(incoming_keys))
    if not incoming_keys:
        return summary

    path = output_dir / "icms_services_log_history.csv"
    if not path.exists():
        return summary
    try:
        existing = pd.read_csv(path, dtype=str, encoding="utf-8").fillna("")
    except Exception:
        return summary
    if existing.empty:
        return summary
    existing_keys = set(_build_services_event_key_df(existing)["service_event_key"].astype(str))
    if not existing_keys:
        return summary

    overlap = len(incoming_keys & existing_keys)
    summary["existing_overlap_rate"] = float(overlap / len(incoming_keys))
    return summary


def update_services_log_history(
    services_df: pd.DataFrame,
    snapshot_date: str,
    source_file_name: str,
    output_dir: Path,
) -> Dict[str, float]:
    """Append/update deduplicated Services Log history.

    Cumulative history is deduplicated deterministically by a service-event key.
    The chosen representative row per event key is the earliest
    ``snapshot_date`` then ``source_file_name`` (stable sort), making results
    order-independent for historical backfills.

    Returns metrics used by GUI/logging.
    """
    path = output_dir / "icms_services_log_history.csv"
    metrics: Dict[str, float] = {
        "services_log_records_processed": 0,
        "services_log_history_total_rows": 0,
        "services_log_duplicate_records_dropped-diag": 0,
        "services_log_blank_service_id_count-diag": 0,
    }

    if services_df.empty:
        if path.exists():
            try:
                metrics["services_log_history_total_rows"] = len(
                    pd.read_csv(path, dtype=str, encoding="utf-8")
                )
            except Exception:
                metrics["services_log_history_total_rows"] = 0
        return metrics

    incoming = _strip_internal_cols(services_df).copy()
    incoming = incoming.astype(str).replace({"nan": "", "None": "", "<NA>": ""})
    incoming.insert(0, "source_file_name", source_file_name)
    incoming.insert(0, "reporting_month", snapshot_date[:7])
    incoming.insert(0, "snapshot_date", snapshot_date)

    metrics["services_log_records_processed"] = int(len(incoming))
    if "Service ID" in incoming.columns:
        metrics["services_log_blank_service_id_count-diag"] = int(
            incoming["Service ID"].fillna("").astype(str).str.strip().eq("").sum()
        )

    if path.exists():
        try:
            existing = pd.read_csv(path, dtype=str, encoding="utf-8").fillna("")
        except Exception as exc:
            print(
                f"WARNING: Could not read existing services log history: {exc}. "
                "Starting fresh."
            )
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()

    combined = pd.concat([existing, incoming], ignore_index=True, sort=False)
    if combined.empty:
        return metrics

    key_df = _build_services_event_key_df(combined)
    combined = combined.copy()
    combined["_service_event_key"] = key_df["service_event_key"]
    combined["_snap"] = pd.to_datetime(combined.get("snapshot_date", ""), errors="coerce")
    combined["_src"] = combined.get("source_file_name", "").fillna("").astype(str)

    before = len(combined)
    combined = combined.sort_values(["_snap", "_src"], kind="stable")
    deduped = combined.groupby("_service_event_key", sort=False).first().reset_index(drop=True)
    dropped = max(0, before - len(deduped))
    metrics["services_log_duplicate_records_dropped-diag"] = int(dropped)

    deduped = deduped.drop(columns=["_service_event_key", "_snap", "_src"], errors="ignore")
    try:
        deduped.to_csv(path, index=False, encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(f"Could not write services log history: {exc}") from exc

    metrics["services_log_history_total_rows"] = int(len(deduped))
    return metrics


def _build_master_from_histories(output_dir: Path) -> pd.DataFrame:
    """Rebuild client/project master deterministically from history files."""
    history_path = output_dir / "icms_snapshot_history.csv"
    exit_history_path = output_dir / "icms_exit_log_history.csv"

    if history_path.exists():
        try:
            hist = pd.read_csv(history_path, dtype=str, encoding="utf-8").fillna("")
        except Exception as exc:
            raise RuntimeError(f"Could not read snapshot history: {exc}") from exc
    else:
        hist = pd.DataFrame()

    if exit_history_path.exists():
        try:
            exit_hist = pd.read_csv(
                exit_history_path, dtype=str, encoding="utf-8"
            ).fillna("")
        except Exception as exc:
            raise RuntimeError(f"Could not read exit log history: {exc}") from exc
    else:
        exit_hist = pd.DataFrame()

    if hist.empty:
        # Exit-only histories can still seed a master.
        if exit_hist.empty:
            return pd.DataFrame(
                columns=METADATA_FIELDS + ALL_APPROVED_FIELDS
                + MASTER_TRACKING_FIELDS + EXIT_FIELDS
            )

    all_master_cols = METADATA_FIELDS + ALL_APPROVED_FIELDS + MASTER_TRACKING_FIELDS + EXIT_FIELDS

    # Participant history (canonical latest known values by max snapshot_date)
    if not hist.empty and all(k in hist.columns for k in KEY_FIELDS):
        hist = hist.copy()
        hist["_snap"] = pd.to_datetime(hist["snapshot_date"], errors="coerce")
        hist = hist.sort_values(["_snap"], kind="stable")
        latest_participant = hist.groupby(KEY_FIELDS, sort=False).last().reset_index()
        latest_participant = latest_participant.drop(columns=["_snap"], errors="ignore")

        latest_snap = hist["_snap"].max()
        if pd.isna(latest_snap):
            latest_keys: set = set()
        else:
            latest_rows = hist[hist["_snap"] == latest_snap]
            latest_keys = set(
                zip(
                    latest_rows["CHAMP ID"].astype(str),
                    latest_rows["ICMS Project Name"].astype(str),
                )
            )
    else:
        latest_participant = pd.DataFrame(columns=KEY_FIELDS)
        latest_keys = set()

    # Exit history (latest exit record by max parsed date fallback snapshot date)
    if not exit_hist.empty and all(k in exit_hist.columns for k in KEY_FIELDS):
        exit_hist = exit_hist.copy()
        exit_hist["_snap"] = pd.to_datetime(exit_hist.get("snapshot_date", ""), errors="coerce")
        exit_hist["_checkout"] = pd.to_datetime(exit_hist.get("Check-Out Date", ""), errors="coerce")
        exit_hist["_moveout"] = pd.to_datetime(exit_hist.get("Move-Out Date", ""), errors="coerce")
        exit_hist["_event"] = exit_hist["_checkout"].fillna(exit_hist["_moveout"]).fillna(exit_hist["_snap"])
        exit_hist = exit_hist.sort_values(["_event", "_snap"], kind="stable")
        latest_exit = exit_hist.groupby(KEY_FIELDS, sort=False).last().reset_index()
    else:
        latest_exit = pd.DataFrame(columns=KEY_FIELDS)

    master = latest_participant.copy()
    for col in all_master_cols:
        if col not in master.columns:
            master[col] = ""
    master = master.reindex(columns=all_master_cols, fill_value="")

    if not latest_exit.empty:
        exit_map = {}
        for _, r in latest_exit.iterrows():
            key = (_safe_str(r.get("CHAMP ID", "")), _safe_str(r.get("ICMS Project Name", "")))
            exit_map[key] = r

        if master.empty:
            master = pd.DataFrame(columns=all_master_cols)

        # Update participant-derived rows with exit data where present.
        for idx in master.index:
            key = (_safe_str(master.at[idx, "CHAMP ID"]), _safe_str(master.at[idx, "ICMS Project Name"]))
            row = exit_map.get(key)
            if row is None:
                continue
            for f in EXIT_FIELDS:
                if f in latest_exit.columns:
                    master.at[idx, f] = _safe_str(row.get(f, ""))
            # Backfill approved fields from exit history when participant history is empty for them.
            for f in ALL_APPROVED_FIELDS:
                if f in latest_exit.columns and _safe_str(master.at[idx, f]) == "":
                    master.at[idx, f] = _safe_str(row.get(f, ""))

        # Add exit-only keys absent from participant history.
        participant_keys = set(
            zip(master["CHAMP ID"].astype(str), master["ICMS Project Name"].astype(str))
        ) if not master.empty else set()
        for key, row in exit_map.items():
            if key in participant_keys:
                continue
            new_row = {c: "" for c in all_master_cols}
            new_row["CHAMP ID"] = key[0]
            new_row["ICMS Project Name"] = key[1]
            for f in ALL_APPROVED_FIELDS:
                if f in latest_exit.columns:
                    new_row[f] = _safe_str(row.get(f, ""))
            for f in EXIT_FIELDS:
                if f in latest_exit.columns:
                    new_row[f] = _safe_str(row.get(f, ""))
            # Preserve useful metadata from exit record when available.
            if "snapshot_date" in latest_exit.columns:
                new_row["snapshot_date"] = _safe_str(row.get("snapshot_date", ""))
            if "reporting_month" in latest_exit.columns:
                new_row["reporting_month"] = _safe_str(row.get("reporting_month", ""))
            if "source_file_name" in latest_exit.columns:
                new_row["source_file_name"] = _safe_str(row.get("source_file_name", ""))
            master = pd.concat([master, pd.DataFrame([new_row])], ignore_index=True)

    # Derive deterministic tracking fields.
    exit_keys = set()
    if not latest_exit.empty:
        exit_keys = set(
            zip(
                latest_exit["CHAMP ID"].astype(str),
                latest_exit["ICMS Project Name"].astype(str),
            )
        )

    if master.empty:
        return master

    for idx in master.index:
        key = (
            _safe_str(master.at[idx, "CHAMP ID"]),
            _safe_str(master.at[idx, "ICMS Project Name"]),
        )
        if key in latest_keys:
            master.at[idx, "present_in_latest_snapshot"] = "True"
            master.at[idx, "client_status"] = "Active"
        else:
            master.at[idx, "present_in_latest_snapshot"] = "False"
            if key in exit_keys:
                master.at[idx, "client_status"] = "Exited"
            else:
                master.at[idx, "client_status"] = "Unknown"

    return master.reindex(columns=all_master_cols, fill_value="")


def upsert_exit_log_to_master(
    exit_df: pd.DataFrame,
    master: pd.DataFrame,
    active_key_set: set,
) -> pd.DataFrame:
    """Merge exit log records into the master DataFrame.

    Key: ``CHAMP ID + ICMS Project Name``.

    Active keys take precedence — if a key appears in both the exit log and
    ``active_key_set`` (the current Participant Level Summary), the master row
    is not modified by the exit log.

    For each exit record whose key is **not** in ``active_key_set``:

    * **Key exists in master** → update all EXIT_FIELDS, set
      ``client_status = "Exited"`` and ``present_in_latest_snapshot = "False"``.
    * **Key absent from master** → append a new row with all available exit
      log columns populated, ``client_status = "Exited"``, and
      ``present_in_latest_snapshot = "False"``; all other columns set to ``""``.

    Returns the updated master DataFrame.
    """
    if exit_df.empty:
        return master

    # Ensure required columns exist in master before updates
    for col in EXIT_FIELDS + ["client_status", "present_in_latest_snapshot"]:
        if col not in master.columns:
            master[col] = ""

    new_exit_rows: list = []

    for _, row in exit_df.iterrows():
        champ = _safe_str(row.get("CHAMP ID", ""))
        proj = _safe_str(row.get("ICMS Project Name", ""))
        key = (champ, proj)

        if key in active_key_set:
            continue

        mask = (
            (master["CHAMP ID"].astype(str) == champ)
            & (master["ICMS Project Name"].astype(str) == proj)
        )

        if mask.any():
            idx = master[mask].index[0]
            for field in EXIT_FIELDS:
                if field in row.index:
                    master.at[idx, field] = _safe_str(row[field])
            master.at[idx, "client_status"] = "Exited"
            master.at[idx, "present_in_latest_snapshot"] = "False"
        else:
            new_row: Dict[str, str] = {col: "" for col in master.columns}
            for col in exit_df.columns:
                if col in master.columns:
                    new_row[col] = _safe_str(row[col])
            new_row["client_status"] = "Exited"
            new_row["present_in_latest_snapshot"] = "False"
            new_exit_rows.append(new_row)

    if new_exit_rows:
        new_df = pd.DataFrame(new_exit_rows, columns=master.columns)
        master = pd.concat([master, new_df], ignore_index=True)

    return master


def update_client_project_master(
    df: pd.DataFrame,
    snapshot_date: str,
    source_file_name: str,
    output_dir: Path,
    exit_df: pd.DataFrame,
) -> Tuple[int, int, int, int, int, int]:
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
    exit_df:
        Exit log DataFrame from ``read_exit_log()``.

    Returns
    -------
    (total_rows, new_rows, returning_rows, not_present_count,
     exited_count, unknown_count)

    Raises
    ------
    RuntimeError
        On write failure.
    """
    path = output_dir / "icms_client_project_master.xlsx"

    incoming = _strip_internal_cols(df).copy()
    incoming = incoming.astype(str).replace({"nan": "", "None": "", "<NA>": ""})

    # Keep exit history cumulative so master rebuild remains order-independent.
    update_exit_log_history(exit_df, snapshot_date, source_file_name, output_dir)

    incoming_deduped = incoming.groupby(KEY_FIELDS, sort=False).last().reset_index()
    incoming_key_set = set(
        zip(
            incoming_deduped["CHAMP ID"].astype(str),
            incoming_deduped["ICMS Project Name"].astype(str),
        )
    ) if all(k in incoming_deduped.columns for k in KEY_FIELDS) else set()

    if path.exists():
        try:
            existing_master = pd.read_excel(path, engine="openpyxl", dtype=str).fillna("")
        except Exception:
            existing_master = pd.DataFrame()
    else:
        existing_master = pd.DataFrame()

    if (
        not existing_master.empty
        and all(k in existing_master.columns for k in KEY_FIELDS)
    ):
        existing_key_set = set(
            zip(
                existing_master["CHAMP ID"].astype(str),
                existing_master["ICMS Project Name"].astype(str),
            )
        )
    else:
        existing_key_set = set()

    new_count = len(incoming_key_set - existing_key_set)
    returning_count = len(incoming_key_set & existing_key_set)
    not_present_count = len(existing_key_set - incoming_key_set)

    master = _build_master_from_histories(output_dir)
    exited_count = int((master["client_status"] == "Exited").sum()) if not master.empty else 0
    unknown_count = int((master["client_status"] == "Unknown").sum()) if not master.empty else 0

    try:
        master.to_excel(path, index=False, engine="openpyxl")
    except Exception as exc:
        raise RuntimeError(
            f"Could not write client/project master: {exc}"
        ) from exc

    return (
        len(master), new_count, returning_count, not_present_count,
        exited_count, unknown_count,
    )


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
