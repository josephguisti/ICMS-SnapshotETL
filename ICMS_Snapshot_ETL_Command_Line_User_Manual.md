# ICMS Snapshot ETL — Command Line User Manual

This manual describes the **command line** version of the ICMS weekly snapshot ETL. It is intended for technical staff who run `update_icms_snapshot_files.py` from a terminal for development, testing, or scheduled batch runs. A future graphical interface will call the same core module; behavior described here applies to the pipeline, not only to the command line entry point.

---

## 1. What the tool does

The ETL:

1. Opens a weekly DHS/HSH ICMS Excel workbook (`.xlsx`), including optional password protection.
2. Reads the **Participant Level Summary** sheet (or another sheet you choose).
3. Cleans headers and text, drops completely blank rows, and keeps only the **approved field set** plus three metadata columns.
4. Writes or updates Tableau-oriented outputs in a folder you specify.
5. Maintains a **field alias** file next to the scripts so column renames are remembered across runs.

All reusable logic lives in `icms_etl_core.py`. The command line runner is `update_icms_snapshot_files.py`.

---

## 2. Prerequisites

- **Python 3** (version your organization supports; the tool uses standard library plus third-party packages).
- **Packages** (install with pip):

  ```text
  pip install pandas openpyxl msoffcrypto-tool
  ```

  - `pandas` — data frames and CSV I/O.
  - `openpyxl` — Excel read/write engine used for all `.xlsx` operations.
  - `msoffcrypto-tool` — decrypt password-protected workbooks **in memory** only. If the file is encrypted and this package is missing, the program exits with an install message.

- **Scripts** in the same folder (or on `PYTHONPATH`):

  - `icms_etl_core.py`
  - `update_icms_snapshot_files.py`

---

## 3. How to run (command line)

Open PowerShell, Command Prompt, or a terminal, change to the folder that contains the scripts and your workbook, then run:

```bash
python update_icms_snapshot_files.py ^
  --input "ICMS Snapshot-DeIdentified Scrambled.xlsx" ^
  --snapshot-date 2026-05-01 ^
  --sheet "Participant Level Summary" ^
  --output-dir "C:\Path\To\Tableau\Data"
```

On macOS or Linux, use backslashes for line continuation or put everything on one line:

```bash
python update_icms_snapshot_files.py \
  --input "ICMS Snapshot-DeIdentified Scrambled.xlsx" \
  --snapshot-date 2026-05-01 \
  --sheet "Participant Level Summary" \
  --output-dir "/path/to/tableau/data"
```

If you omit line breaks, a minimal successful command looks like:

```bash
python update_icms_snapshot_files.py --input "weekly.xlsx" --snapshot-date 2026-05-01
```

---

## 4. Command line arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--input` | Yes | Path to the weekly ICMS Excel snapshot (`.xlsx`). |
| `--snapshot-date` | Yes | Snapshot date in **`YYYY-MM-DD`** format. Used in metadata, history grain, and master tracking. |
| `--sheet` | No | Exact worksheet name. If omitted, the program lists sheet names and **prompts** you to enter a sheet name or its list number. |
| `--output-dir` | No | Folder where outputs are written. Defaults to the **current working directory** if omitted. |

Built-in help:

```bash
python update_icms_snapshot_files.py --help
```

---

## 5. Password-protected workbooks

1. The tool first tries to open the file normally.
2. If the file is encrypted, it prompts for a password using a **hidden** prompt (no echo on screen).
3. The file is decrypted **in memory** only. No decrypted copy is written to disk, and the password is not stored.

If encryption is detected but `msoffcrypto-tool` is not installed, the program stops and tells you to install it with pip.

---

## 6. Sheet selection

- Row **1** is treated as column headers; row **2** onward is data.
- If `--sheet` is wrong or missing from the workbook, the program prints available sheets and exits with a non-zero status.
- If `--sheet` is omitted, you interactively choose by **sheet name** or **number** from the printed list.

---

## 7. Field resolution and `field_aliases.json`

The ETL expects a fixed set of **canonical** column names (required and optional). For each expected field it:

1. Looks for an **exact** match in the workbook headers (after stripping whitespace from headers).
2. If not found, checks **`field_aliases.json`** in the **same directory as the Python scripts** (not `--output-dir`). That file maps canonical name → actual column name from a previous run.
3. If still not found, it **prompts** you to type the workbook column name or press Enter to skip.

**Prompt rules**

- You must type a name that **exactly** matches a workbook column (after the tool’s header stripping).
- If you supply a valid name, it is used for this run and saved to `field_aliases.json` for future runs.
- **Required field** skipped (Enter): you see an alert; you must type **`continue`** to proceed without that column. The run continues; missing required fields are noted in the ingest log.
- **Optional field** skipped (Enter): no extra alert; that column is omitted from outputs for this run.

`field_aliases.json` is created or updated automatically when new mappings are saved.

**Client identifiers and demographics**

`HMIS ID`, `First Name`, and `Last Name` are **required** fields: if they are missing from the workbook headers, the tool prompts for a column mapping, and skipping one after Enter requires typing **`continue`** before the run proceeds without that column (see prompt rules above). `ICMS Slot #` and `Medi-Cal CIN` remain **optional** and can be omitted per run. If a production extract renames any of these headers, map them once via `field_aliases.json` or the interactive prompt.

---

## 8. Data cleaning (what changes in the data)

- Leading and trailing spaces are removed from **column headers** and from **string** cell values.
- Rows that are **entirely blank** across data columns are dropped.
- Other cell values are **not** transformed (dates and numbers stay as interpreted by pandas/openpyxl unless they are strings that get trimmed).

---

## 9. Output files (written to `--output-dir`)

Use a **stable** output folder path so Tableau connections do not break.

| File | Description |
|------|-------------|
| `icms_current_dashboard_flatfile.xlsx` | **Latest snapshot only.** One row per source row for the run’s snapshot date. Grain for dashboard use: **CHAMP ID + ICMS Project Name** may repeat if the source has duplicates; duplicates are also logged separately. **Overwritten** every run. |
| `icms_snapshot_history.csv` | **Cumulative history.** Grain: **CHAMP ID + ICMS Project Name + snapshot_date**. If you reload the **same** `snapshot_date`, all rows for that date are **removed** and replaced with the new run’s rows. UTF-8. |
| `icms_client_project_master.xlsx` | **Cumulative master** per **CHAMP ID + ICMS Project Name**. Includes tracking fields (`first_seen_snapshot_date`, `last_seen_snapshot_date`, `last_seen_source_file`, `present_in_latest_snapshot`, `days_since_last_seen`) plus latest known values for approved fields. Clients not in the latest snapshot remain in the file with `present_in_latest_snapshot` false and recalculated `days_since_last_seen`. |
| `icms_ingest_log.csv` | One row per **snapshot_date** (same date reruns **replace** the prior log row for that date). UTF-8. |
| `icms_duplicate_client_project_warning.csv` | Written when duplicate **CHAMP ID + ICMS Project Name** keys appear **within the incoming snapshot**. Same **snapshot_date** reruns **replace** prior warning rows for that date. UTF-8. |

**Script directory (not output folder)**

| File | Description |
|------|-------------|
| `field_aliases.json` | Persists column alias mappings. Lives next to `icms_etl_core.py` / `update_icms_snapshot_files.py`. |

---

## 10. Metadata columns added to row outputs

These are prepended to the reduced dataset:

- `snapshot_date` — from `--snapshot-date`
- `reporting_month` — `YYYY-MM` derived from `snapshot_date`
- `source_file_name` — base file name of `--input`

---

## 11. After a successful run

The command line prints a **run summary**, including:

- Input file name  
- Selected sheet  
- Snapshot date  
- Current dashboard row count  
- Snapshot history row count after update  
- Master row count after update  
- Duplicate flag count  
- Missing required fields (if any)  
- Output folder path  

---

## 12. Exit codes and errors

- On **unrecoverable** errors (missing file, bad date format, write failure, decrypt failure, and similar), the program prints a clear message and exits with a **non-zero** status.
- Some secondary writes (for example duplicate warning or ingest log) may warn without aborting the whole run if the primary outputs already succeeded; read the console output and the log files.

---

## 13. Suggested workflow for analysts

1. Place the new weekly `.xlsx` in a known location.
2. Choose or create an **output directory** used consistently for Tableau.
3. Run the command line with `--input`, `--snapshot-date`, and usually `--sheet "Participant Level Summary"`.
4. Confirm the run summary row counts and open `icms_ingest_log.csv` for the latest `snapshot_date` if you need run metrics or notes.
5. If DHS renames a column, answer the prompt once; confirm `field_aliases.json` updated next to the scripts.

---

## 14. Quick reference — ingest log columns

The ingest log row includes (among others): run timestamp, snapshot date, source file, sheet name, input row count, dashboard row count, history total after update, master row count after update, new vs returning client/project counts, not-present count, missing required/optional field counts, aliases added count, duplicate count, whether the input was password protected, run status, and notes.

---

## 15. Support and scope

- **v1 scope:** command line runner plus core ETL only; no bundled GUI or executable in this manual.
- For behavior details not covered here, refer to inline comments in `icms_etl_core.py` and `update_icms_snapshot_files.py`.

---

*Document version: 1.0 — Command line ETL user manual.*
