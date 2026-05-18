# ICMS Snapshot ETL — Command Line User Manual

This manual describes the **command line** runner (`update_icms_snapshot_files.py`) for the ICMS weekly snapshot ETL. Operational behavior applies to the whole pipeline (`icms_etl_core.py`); for architecture and GUI-specific features, see `architecture planning.md` and run `python icms_gui.py` directly.

---

## 1. What the tool does

The ETL:

1. Opens a weekly DHS/HSH ICMS Excel workbook (`.xlsx`), including optional password protection.
2. Reads the **participant** sheet you choose (typically **Participant Level Summary**), detecting the header row automatically when vendors use multi-row headers.
3. Optionally reads **`Vacant Slots`** and **`Exit Log`** sheets (fixed names in the CLI; sheet names cannot be overridden from the CLI). Missing sheets produce warnings and are treated as empty.
4. Cleans headers and text, drops completely blank rows, resolves column names via aliases and prompts, and keeps only the **approved field set** plus metadata.
5. Writes or updates dashboard, history, master, logs, and alias files under your `--output-dir`.
6. Persists **`field_aliases.json`** beside the scripts and **`settings.json`** after successful runs.

Reusable logic: `icms_etl_core.py`. CLI entry point: `update_icms_snapshot_files.py`.

---

## 2. Prerequisites

- **Python 3** supported by your organization.
- **Packages** — install dependencies (recommended):

  ```text
  pip install -r requirements.txt
  ```

  Minimum for the CLI: `pandas`, `openpyxl`, `msoffcrypto-tool`. The GUI adds `tkcalendar` and `keyring`; they are not required for the CLI.

- **Files** next to each other (or adjust `PYTHONPATH`):

  - `icms_etl_core.py`
  - `update_icms_snapshot_files.py`

---

## 3. How to run (command line)

### Single-file mode

Open PowerShell or a terminal, `cd` to the folder containing the scripts, then run:

```powershell
python update_icms_snapshot_files.py `
  --input "ICMS Snapshot-DeIdentified Scrambled.xlsx" `
  --snapshot-date 2026-05-01 `
  --sheet "Participant Level Summary" `
  --output-dir "C:\Path\To\Tableau\Data"
```

On macOS or Linux:

```bash
python update_icms_snapshot_files.py \
  --input "ICMS Snapshot-DeIdentified Scrambled.xlsx" \
  --snapshot-date 2026-05-01 \
  --sheet "Participant Level Summary" \
  --output-dir "/path/to/tableau/data"
```

If you omit `--output-dir`, outputs go to the **current working directory**.

### Batch mode (folder of workbooks)

```bash
python update_icms_snapshot_files.py \
  --batch "./snapshots" \
  --output-dir "./output" \
  --sheet "Participant Level Summary"
```

In batch mode **`--input` and `--snapshot-date` are ignored**. The CLI lists every `.xlsx` in the batch folder, infers **`MM.DD.YYYY`** dates from filenames, lets you fix any **`UNPARSED`** entries interactively, then processes files **oldest → newest**, continuing past individual file failures where possible.

### Minimal single-file invocation

```bash
python update_icms_snapshot_files.py --input "weekly.xlsx" --snapshot-date 2026-05-01
```

Built-in help:

```bash
python update_icms_snapshot_files.py --help
```

---

## 4. Command line arguments

| Argument | When required | Description |
|----------|----------------|-------------|
| `--input` | Single-file mode | Path to one weekly `.xlsx` snapshot. |
| `--snapshot-date` | Single-file mode | **`YYYY-MM-DD`**. Snapshot date used in metadata, history grain, and master logic. |
| `--sheet` | No | Participant sheet name. If omitted, sheets are listed and you **prompt** for name or number. In batch mode, if omitted you are prompted once on the first workbook; thereafter the **same name** must exist in later files or that file fails for that workbook. |
| `--output-dir` | No | Output folder (created if missing). Defaults to **`cwd`** if omitted. |
| `--batch` | Mutually exclusive with single-file paths | Directory containing multiple `.xlsx` files (see §3). Implies **`--input` / `--snapshot-date` not used**. |

Single-file validation: **`--input`** and **`--snapshot-date`** are required unless **`--batch`** is set.

---

## 5. Saved settings (`settings.json`)

On startup the CLI reads **`settings.json`** in the script directory (if present) and prints **last used** input directory, output directory, and participant sheet name. After each **successful single-file run** or after **batch completes**, settings are overwritten with:

- `last_input_dir` — folder of the workbook (single file) or the batch folder  
- `last_output_dir`  
- `last_sheet_name` — resolved participant sheet (`--sheet` value or interactive choice)

This file is advisory for humans; startup does **not** auto-fill CLI arguments beyond the printed reminder.

---

## 6. Snapshot date gap warning (single-file)

Before processing, if **`icms_snapshot_history.csv`** already exists under `--output-dir`, the CLI compares your **`--snapshot-date`** to the newest date in history. If the incoming date is **more than 45 days** later, it warns and asks you to type **`confirm`** to proceed or press **Enter** to abort (exit 0).

Batch mode does not run this check per file automatically in the same way; rely on diligence when correcting dates during batch confirmation.

---

## 7. Password-protected workbooks

1. The tool tries normal open first.  
2. If encrypted, it prompts for a password (**hidden input** where supported).  
3. Decryption is **in memory** only — no decrypted file on disk, no credential storage.

If encryption is detected but **`msoffcrypto-tool`** is missing, installation instructions are printed and the program exits.

---

## 8. Participant sheet and headers

- The participant sheet contains the tabular census data requiring **field mapping** across runs.
- The CLI **does not assume row 1 is the header**. It scans the first several rows (**`find_header_row`**) and picks the row that best matches approved canonical field labels, so vendor “title” rows above the column headers no longer force dozens of bogus alias prompts.
- **`Vacant Slots`** / **`Exit Log`** are read with a **single** header row (first row after openpyxl’s used range). Columns are narrowed in code; unresolved names are not prompted on those sheets.

If `--sheet` is wrong or missing from a workbook, the program lists sheets and exits with a non‑zero status (or skips the file in batch with a logged failure).

---

## 9. Field resolution and `field_aliases.json`

For the **participant** sheet only:

1. **Exact** match after trimming headers.  
2. **`field_aliases.json`** in the **script directory** (not `--output-dir`), mapping canonical names to workbook columns.  
3. **Interactive terminal prompt** if still unresolved.

Prompt rules:

- You must supply a workbook column name that **exactly** matches after header stripping (or press Enter).  
- A valid mapping is saved into **`field_aliases.json`** for later runs.  
- **Required** field skipped (**Enter**) — you must acknowledge with **`continue`** to proceed without that column; missing required fields appear in ingest notes.  
- **Optional** fields — skip silently; column omitted from outputs for that run.

The full approved catalogue (required, optional, master tracking fields, exit fields) is summarized in **`architecture planning.md`**. **`Medi-Cal CIN`** and **`ICMS Slot #`** are among the **required** participant fields — do not treat them as optional when verifying prompt behavior.

---

## 10. Data cleaning

- Leading/trailing spaces removed from **headers** and **string** cells.  
- Rows that are entirely blank across data columns dropped.  
- Other values are passed through according to spreadsheet types (minimal numeric/date coercion unless already strings that get trimmed).

---

## 11. Output files (written to `--output-dir`)

Use a **stable** output folder so Tableau and downstream jobs keep working.

| File | Description |
|------|-------------|
| `icms_current_dashboard_flatfile.xlsx` | **Participant rows for one snapshot** merged with synthetic **vacant slot** placeholder rows (`is_vacant_slot`). **Participant** rows carry `is_vacant_slot = False`; vacant rows use `True` and null/absent participant fields. **Overwritten** every time a run completes this step successfully. **`--batch` mode** writes the dashboard once per successful file in **old→new order**, so the file normally ends on the **latest snapshot date among successes** — if the last chronological file fails, the dashboard stays at the prior successful snapshot. (**`icms_gui.py`** adds an optional *skip dashboard* flag for historic loads.) |
| `icms_snapshot_history.csv` | **Participant rows only.** Cumulative CSV; grain **`CHAMP ID + ICMS Project Name + snapshot_date`**. Reruns for the **same** `snapshot_date` replace rows for that date. UTF‑8. |
| `icms_exit_log_history.csv` | Cumulative exit rows keyed by **`snapshot_date` + `source_file_name`** semantics in core (replacing rows for reruns with the same pairing). Supports rebuilding the master. |
| `icms_client_project_master.xlsx` | One row per **`CHAMP ID + ICMS Project Name`**. Rebuilt deterministically each run from **snapshot history + exit log history**, so ingest **order within the cumulative files** does not corrupt “latest participant” or status logic. Tracking fields: **`present_in_latest_snapshot`**, **`client_status`** (**`Active`** / **`Exited`** / **`Unknown`**). |
| `icms_ingest_log.csv` | One row per **`snapshot_date`** run (successful rows updated; failures append with `run_status=failed`). UTF‑8. |
| `icms_duplicate_client_project_warning.csv` | When duplicate **`CHAMP ID + ICMS Project Name`** pairs appear within one snapshot; same‑date reruns replace prior warnings for that snapshot date. UTF‑8. |

**Script directory**

| File | Description |
|------|-------------|
| `field_aliases.json` | Column alias mappings. |
| `settings.json` | Last CLI/GUI dirs and sheet hints (CLI prints on launch). |

**Duplicate-ingest warning UI** (optional check before rerun) exists in **`icms_gui.py`** only — the CLI does not block on “already imported” filename/date pairs.

---

## 12. Metadata columns added to row outputs

- `snapshot_date` — from **`--snapshot-date`** or batch confirmation  
- `reporting_month` — `YYYY-MM` from `snapshot_date`  
- `source_file_name` — base name of workbook  
- **`is_vacant_slot`** — **`False`** on participant rows written to dashboard; **`True`** on vacant-slot synthetic rows *(dashboard output only)*

---

## 13. Run summary

After single-file success, the CLI prints totals including:

- Input file  
- Participant sheet  
- Snapshot date  
- Dashboard row counts — **participant + vacant slot injections**  
- Exit log records read from workbook  
- History row total after update  
- Master totals — **`exited`** and **`unknown`** client_status counts  
- Duplicate flag rows  
- Missing required fields (if any)  
- Output folder path  

Batch mode ends with **`BATCH COMPLETE`**: successes, failures, and final master row estimate.

---

## 14. Exit codes and failures

- **Recoverable skips** inside batch advance to the next file; failures append to **`icms_ingest_log`** with **`run_status: failed`** when possible (**`_write_failure_log`**).  
- Hard errors (**bad date**, **missing `--input`** in single mode, decrypt failure before progress, unreadable workbook, etc.) print a message and exit **non‑zero**.

---

## 15. Suggested workflows

### Weekly incremental (single workbook)

1. Place the `.xlsx` in a known path.  
2. Use one consistent **`--output-dir`**.  
3. Run **`--input`**, **`--snapshot-date`**, and usually **`--sheet "Participant Level Summary"`**.  
4. Review stdout summary and **`icms_ingest_log.csv`**.  
5. If DHS renames participant columns once, approve the prompts; **`field_aliases.json`** records them.

### Historical backfill (many files)

1. Prefer **`--batch`** with **`MM.DD.YYYY`** in filenames, or ingest one file per run while watching order.  
2. The **master** remains consistent when **`icms_snapshot_history.csv`** and **`icms_exit_log_history.csv`** are complete — order of **historic** ingestion matters only for incremental dashboard freshness, not logical master rebuild.  
3. **Re-ingesting** past dates updates history replacements for those dates automatically.

---

## 16. Ingest log schema (successful runs)

Representative columns include: **`run_timestamp`**, **`snapshot_date`**, **`source_file_name`**, **`sheet_name`**, **`input_row_count`**, **`vacant_slot_rows_injected`**, **`exit_log_records_processed`**, **`current_dashboard_row_count`**, **`history_total_row_count_after_update`**, **`master_row_count_after_update`**, new/returning/not-present counts, **`master_exited_count`**, **`master_unknown_count`**, missing required/optional counts, alias additions, duplicate count, **`password_protected_input`**, **`run_status`**, **`notes`**. Failures populate the same structure with blanks where metrics are unavailable.

---

## 17. Related entry points

- **`architecture planning.md`** — full behavioral spec, GUI batch/dashboard publish toggle, duplicate-warning behavior.  
- **`icms_gui.py`** — Tkinter front end calling **`icms_etl_core`** (editable vacant/exit sheet names, dashboard publish checkbox, duplicate warnings).

---

*Document version: 2.0 — Command line ETL user manual aligned with vacant slots, exit log history, deterministic master rebuild, batch mode, and header-row detection.*
