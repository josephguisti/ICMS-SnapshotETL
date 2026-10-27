# ICMS Snapshot ETL — Architecture and Behavior Summary

## What this application is

A Python ETL tool for DHS/HSH ICMS snapshot workbooks. The production path is the Tkinter GUI. It reads each workbook's supported tabs through a director-maintained catalogue, keeps cumulative histories, and rebuilds a deterministic enrollment-level master from those histories.

Codebase:
- `icms_catalogue.py` — the only definition of output fields, accepted vendor header names, supported tabs, and reviewed exclusions.
- `icms_etl_core.py` — all ETL logic. `run_snapshot_files` is the single entry point for every run.
- `icms_gui.py` — desktop GUI. Collects choices, shows progress and results; contains no import logic.
- `update_icms_snapshot_files.py` — CLI over the same entry point (development/backup use).
- `tests/` — focused regression tests (`python -m pytest tests -p no:cacheprovider`).

---

## Operating cadence

Import a snapshot **weekly**, and additionally import a snapshot **generated on the last day of each month**.

- A snapshot's Exit Log appears to cover the current service month. The app's exit history keeps every exit it has observed across imports, so exits survive after they roll off the vendor's tab.
- A month-end snapshot only contains records entered before it was generated. Late or backdated exits can still be missing; those enrollments stay **Unknown**. Nobody needs to intervene, and the app never invents exit dates.
- An exit that disappears from a later Exit Log is **not** treated as reversed. An enrollment that reappears in a participant snapshot becomes Active by the status rule below; absence alone proves nothing.

---

## GUI workflow

### Start and run
- Launch with `python icms_gui.py` or the EXE.
- `Browse…` for one workbook (weekly), `Browse files…` for several (backfill or repair).
- Snapshot dates fill in automatically (see "Snapshot dates"). Override any date by typing over it.
- Choose the output folder and run. No field-mapping questions are ever asked.

### Already-imported files
Before a run, the app checks the output folder for files already imported (same file name + snapshot date, from snapshot history and successful ingest-log rows). If any are found, one dialog offers:
- **Reprocess all selected** — re-import them; their stored rows are replaced with freshly read values. This is how historical data is repaired.
- **Skip already-imported** — shown only when some selected files are new.
- **Cancel**.

### Dashboard publish toggle
- Single file: the toggle decides whether this file becomes the current dashboard flatfile.
- Batch: with the toggle on, only the newest **successful** file is published, and only if it is not older than a snapshot already in history. With the toggle off, the dashboard is left as is.

### Results panel
Shows per-file outcomes, dashboard source, master counts, backup location, any tab that was not imported, and the grouped **structural findings**. "Copy findings for director" copies the findings as plain text. If any tab was not imported the status reads "Run complete with gaps" instead of "Run complete".

---

## Input model

Each workbook (`.xlsx`, optionally password-protected; decrypted in memory) supplies up to four data tabs. Nothing else is imported.

| Category | Accepted tab names | If missing |
|---|---|---|
| Participant | `Participant Level Summary` (or the name typed in Setup) | File fails; nothing from it is written. |
| Vacant | `Vacant Slots` | Reported (from 2026-04-08 on); file continues. |
| Exit | `Exit Log` | Reported; file continues. |
| Services | `Service Log`, `Services Log` (both always accepted) | Reported; file continues. |

`Generated On` is read for the snapshot date only.

Tab names are matched after trimming, collapsing whitespace, and ignoring case. If a workbook contains two candidate tabs for one category that differ, neither is imported and the gap is reported; identical copies are harmless.

### Header resolution (per workbook, per tab)
- The header row is detected within the first five rows by overlap with accepted names.
- Each header is matched to an output field if it equals the output name or one of the field's accepted vendor names, after the same whitespace/case normalization. No fuzzy matching.
- Two accepted names for one field may coexist. The first non-blank value in catalogue order is kept. If they disagree, a finding is recorded. If they disagree on an **essential** identity field, the tab is not imported (participant: the file fails) rather than risk misattributing records.
- Headers outside the catalogue and its reviewed-exclusion list are reported as **unfamiliar** and never exported.
- A missing non-essential field becomes blank and is reported once per run, except fields marked retired, absent by design, or not yet introduced for that snapshot's date.
- Essential fields: participant `CHAMP ID`, `ICMS Project Name`, `ICMS Enrollment ID`; exit the same three; vacant `ICMS Project Name`; services `CHAMP ID`. A participant tab missing one of these fails that file. Other tabs missing one are skipped for that file.
- Identifiers (`CHAMP ID`, `ICMS Enrollment ID`) are normalized as text: `100300.0`, `100300`, and `"100300 "` all become `100300`.
- Participant rows with a blank `CHAMP ID` or `ICMS Project Name` are dropped and counted. Rows with a blank enrollment ID are kept as their own blank-enrollment key and reported.

### Snapshot dates
- Default: the date in the workbook's `Generated On` tab, found by locating the "Generated On" label (not a fixed cell).
- Fallback: `MM.DD.YYYY` in the filename.
- When both exist and disagree, the Setup panel shows both next to that file's date. The selected date and its origin (`Generated On`, `filename`, or `manual entry`) are written to the ingest log, along with both candidate dates.
- Already-stored records keep the dates they were imported with unless that file is reprocessed.

---

## Output files

Field names and column order for every output come from `icms_catalogue.py`. Metadata columns `snapshot_date`, `reporting_month` (`YYYY-MM`), and `source_file_name` lead every history. Columns found in older stored files but no longer in the catalogue are kept at the end.

### `icms_current_dashboard_flatfile.xlsx`
**Grain:** one row per ICMS slot for the published snapshot (participants plus vacant placeholders).
- Columns: metadata, every participant output field, `is_vacant_slot`, `# Days Vacant`.
- Vacant rows carry the vacant fields and blank participant fields.
- Written only when publishing (see toggle rules). Date columns are real Excel dates.

### `icms_snapshot_history.csv`
**Grain:** one participant row per enrollment per `snapshot_date`.
- Reprocessing a date replaces every row for that `snapshot_date`.
- Sorted by snapshot date then source file, so import order never changes the file.

### `icms_exit_log_history.csv`
**Grain:** one exit row per enrollment per imported file.
- Reprocessing a file replaces its rows (same `source_file_name` + `snapshot_date`). Rows from other files are untouched.
- A file whose Exit Log is empty or not imported leaves history unchanged.

### `icms_services_log_history.csv`
**Grain:** one row per service event.
- Event key: `Service ID` when present; otherwise `CHAMP ID | ICMS Enrollment ID | Service | Place of Service | Service Date | Service Record Created Date`. Keys are computed after header resolution.
- If a fallback row has a blank `Service`, it borrows the single non-blank `Service` seen for the same remaining components. This lets damaged rows (written when the vendor renamed the column) merge with their repaired copies instead of duplicating.
- Reprocessing a file first removes rows previously stored from that file + date. Then per event key, each column takes the earliest snapshot's non-blank value. A corrected reimport therefore replaces its own old values, and a blank never hides a value another snapshot supplied.
- Blank `PH Building Name` is stored as `Scattered Site`.

### `icms_client_project_master.xlsx`
**Grain:** one row per **`CHAMP ID` + `ICMS Project Name` + `ICMS Enrollment ID`**. The filename is unchanged.

Rebuilt from the full snapshot and exit histories on every run:
- Participant values come from that enrollment's latest participant observation.
- Exit columns come from that enrollment's latest exit observation (latest snapshot, then latest check-out/move-out date). Blank participant fields are filled from the exit row.
- Enrollments seen only in exit history are included.
- `present_in_latest_snapshot` is True when the enrollment appears in the latest imported participant snapshot.
- `client_status`, per enrollment:
  - `Active` — in the latest participant snapshot.
  - `Exited` — not in it, with a matching exit recorded.
  - `Unknown` — previously observed, not in it, no matching exit.
- A new enrollment for the same client and project is a separate row and never overwrites an older enrollment's values or status. Service events never establish enrollment or exit.
- If a stored history lacks `ICMS Enrollment ID` entirely, the rebuild stops with a message instead of falling back to client + project.

### `icms_ingest_log.csv`
One row per `snapshot_date` (replaced on rerun), including failed attempts (`run_status = failed`). It records:
- counts;
- the date and its origin, plus the Generated On and filename dates;
- the tabs used;
- `skipped_categories`;
- `structural_findings`, the grouped findings as one line;
- whether the dashboard was published.

### `icms_duplicate_client_project_warning.csv`
Participant rows that share all three enrollment-key fields within one snapshot. Reprocessing a date replaces that date's rows.

### `backups/`
Before a run that reprocesses already-imported files, and before the first run that upgrades outputs written by an earlier release, the app copies every existing output file to `backups/<timestamp> <reason>/` inside the output folder. Nothing else needs to be managed; old backup folders can be deleted when no longer wanted.

### `icms_output_format.txt`
Records the output format version so the one-time upgrade backup happens once.

### `settings.json`
Stored beside the scripts. Last-used folders and tab names.

---

## Failure and transaction rules

- Everything is read and validated before anything is written.
- All writes for one file happen as a unit. If any write fails, every output file is restored to its state before that file.
- A failed file is logged, listed in Results, and never published. Other files in a batch continue.
- Files run oldest first (snapshot date, then filename) whatever order they were selected in. The master is rebuilt once at the end of a batch.
- Stored histories that cannot be read stop the run; they are never silently started over.

---

## Structural findings

Findings are maintenance information for the director, not questions for the analyst. They are grouped across a batch (one line per finding with the number of affected files) and appear in the Results panel, the run log, and the ingest log.

| Finding | Meaning |
|---|---|
| File not imported: essential column missing / conflicting identity columns | Participant identity could not be read safely. |
| Tab not imported | A non-participant tab lacked an essential column. |
| Supported tab not found / several candidate tabs differ | A data category is absent for that file. |
| Unfamiliar tab / unfamiliar column | Not in the catalogue or the reviewed exclusions; not imported. |
| Expected column missing | A current catalogue field was absent; left blank. |
| Two accepted column names both present and disagree | First in catalogue order kept. |
| Participant rows with blank ICMS Enrollment ID | Kept, reported. |
| Snapshot date differs from Generated On or filename | Shows all three dates. |
| Generated On date not found | Filename or manual date used. |

The reviewed exclusions seeded from the 02.18, 04.30, 05.19, 07.07, and 10.19.2026 samples mean those workbooks produce no unfamiliar-header findings.

---

## CLI (`update_icms_snapshot_files.py`)

```bash
python update_icms_snapshot_files.py --input "ICMS Snapshot - 10.19.2026.xlsx"
python update_icms_snapshot_files.py --batch "./snapshots" --output-dir "./output" --reprocess
```

Same pipeline and outputs as the GUI. Dates default from `Generated On`; `--snapshot-date` overrides in single-file mode, and batch mode lists the dates for correction before running. Options: `--sheet` (participant tab name), `--no-publish`, `--reprocess`. Prompts once for a password if any file is encrypted.

---

## Build Windows EXE

Prerequisites: dependencies from `requirements.txt` and `pyinstaller`.

Bump `__version__` in `version.py`, then from the repo root:

```bash
python build_exe.py
```

Output: `dist/ICMS Snapshot ETL v<version>.exe` (one-file, windowed, built from `ICMS Snapshot ETL.spec`).
