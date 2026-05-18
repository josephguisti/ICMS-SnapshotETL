# ICMS Snapshot ETL — Architecture and Behavior Summary

## What this application is

A Python ETL tool for weekly DHS/HSH ICMS snapshot workbooks. It supports both CLI and Tkinter GUI workflows, normalizes varying source workbooks, and maintains cumulative historical outputs plus a deterministic master file.

Codebase:
- `icms_etl_core.py` — reusable ETL logic.
- `update_icms_snapshot_files.py` — CLI runner.
- `icms_gui.py` — desktop GUI runner.

---

## GUI-first operating workflow

For ongoing use, the intended day-to-day path is the GUI (`icms_gui.py`), not the CLI.

### Start and run
- Launch with `python icms_gui.py`.
- Choose one file (`Browse...`) for routine weekly updates, or many files (`Browse files...`) for historical backfill.
- Pick/confirm snapshot date(s), confirm participant/vacant/exit sheet names, and choose output folder.

### Typical weekly run (single file)
- Keep the dashboard publish toggle enabled.
- Run one latest workbook.
- Result: dashboard flatfile, snapshot history, exit-log history, master, warnings/logs updated.

### Historical backfill (multi-file batch)
- Select all historical files together, then confirm/correct parsed dates.
- Let the app process oldest to newest.
- Recommended: enable dashboard publish only for the newest intended current file.
- Result: histories and deterministic master are built progressively even when files vary.

### Safety features in GUI
- Duplicate-import warning before run (`source_file_name + snapshot_date`) with include/exclude options.
- Missing `Vacant Slots` / `Exit Log` sheets are warnings, not hard failures.
- Alias prompts are modal dialogs; saved aliases reduce repeated mapping work.

### When to use CLI
- Keep CLI for automation/scripting or terminal-only environments.
- Functional behavior aligns with the same core ETL logic, but UX controls (duplicate exclusion dialogs, publish toggle) are GUI conveniences.

---

## Input model

Each run processes one workbook (`.xlsx`), optionally encrypted.

Primary sheet categories:
1. **Participant sheet** (user-configurable; typically `Participant Level Summary`)
2. **Vacant slots sheet** (user-configurable; typically `Vacant Slots`)
3. **Exit log sheet** (user-configurable; typically `Exit Log`)

Header handling:
- Participant sheet supports multi-row vendor headers through automatic header-row detection (`find_header_row`).
- Vacant/exit sheets are read as standard single-header rows.

---

## Output files

All outputs are written to one selected output folder.

### `icms_current_dashboard_flatfile.xlsx`
**Grain:** one row per ICMS slot (participant or vacant placeholder).

- Participant rows: approved fields + metadata (`snapshot_date`, `reporting_month`, `source_file_name`), `is_vacant_slot=False`.
- Vacant rows: available vacant-sheet fields + metadata, participant fields null, `is_vacant_slot=True`.
- Overwritten when dashboard publish is enabled for a run.

### `icms_snapshot_history.csv`
**Grain:** one participant enrollment row per `snapshot_date`.

- Participant rows only (no vacant placeholders).
- Same-date reload is idempotent (existing same `snapshot_date` rows replaced).

### `icms_exit_log_history.csv`
**Grain:** exit-log rows with `snapshot_date` + `source_file_name`.

- Cumulative exit-history store.
- Same source/date reruns replace previously written rows for that source/date pair.

### `icms_client_project_master.xlsx`
**Grain:** one row per `CHAMP ID + ICMS Project Name`.

Current behavior is **deterministic and order-independent**:
- Rebuilt from full `icms_snapshot_history.csv` (+ `icms_exit_log_history.csv`) each run.
- Latest participant values are chosen by max `snapshot_date`.
- `present_in_latest_snapshot` is based on membership in the global latest snapshot.
- `client_status` derived as:
  - `Active` if present in latest participant snapshot
  - `Exited` if absent from latest participant snapshot and exit history exists
  - `Unknown` otherwise
- Exit fields are populated from latest relevant exit-history row.

### `icms_ingest_log.csv`
One row per run (`snapshot_date` replaced on rerun). Tracks key counts, status, and notes.

### `icms_duplicate_client_project_warning.csv`
Rows involved in same-snapshot duplicate `CHAMP ID + ICMS Project Name` keys.

### `field_aliases.json`
Stored beside scripts. Persists canonical field → workbook column mappings.

### `settings.json`
Stored beside scripts. Persists recent UI/CLI settings (input/output dirs, sheet names).

---

## Approved field catalogue

### Metadata fields
`snapshot_date`, `reporting_month`, `source_file_name`

### Required fields
`CHAMP ID`, `ICMS Project Name`, `Date of Birth`, `First Name`, `Last Name`, `HMIS ID`, `Medi-Cal CIN`, `ICMS Check-In Date`, `ICMS Provider`, `PH Building Name`, `Current ICMS Status`, `ICMS Case Manager`, `ICMS Supervisor`, `ICMS Project Type`, `ICMS Slot #`, `Receiving Pre-Match ICMS?`, `Checked in 6+ Months and Not Yet Permanently Housed?`, `Permanent Housing Move-In Date`, `Housing Voucher Issued Date`, `Permanent Housing Subsidy`, `Service Planning Area`, `Household Size`, `PH Update Still Needed This Month?`, `5x5 Still Needed This Month?`, `HAI Still Needed This Month?`, `Care Plan Pending Action Step Still Needed This Month?`, `Care Plan Completed Action Step Still Needed This Month?`, `Eligible Services Still Needed This Month?`, `Housing Navigation Action Step Needed?`, `Health Care Action Step Needed?`, `Key Housing Milestone Data Needed?`

### Optional fields
`ICMS Enrollment ID`, `ICMS Billing Rate`, `Permanently Housed?`, `Any Required Documentation Still Needed This Month?`, `Notification of Privacy Practices On Record?`, `Universal Sharing Consent Status`, `CalAIM CS Housing Navigation (HN) and Tenancy Sustaining Services (TSS) Verbal Opt-In Status Missing?`, `Family Billing Rate But Dependent Household Member(s) Not Listed in Household Composition?`, `Eligible Services Recorded for This Month`, `Universal Housing Application (UHA) Status`

### Master tracking fields
`present_in_latest_snapshot`, `client_status`

### Exit fields
`Check-In Date`, `Check-Out Date`, `Move-In Date`, `Move-Out Date`, `Exit Reason`, `Exit Destination`, `Exit Location`, `Transferred to New Provider Within ICMS Program?`, `New ICMS Provider`

---

## Field resolution and aliases

Resolution order:
1. exact canonical column name,
2. saved alias from `field_aliases.json`,
3. user prompt for unresolved fields.

Required unresolved fields require explicit user acknowledgment to continue.
Optional unresolved fields may be skipped.

GUI prompts are modal dialogs; CLI prompts are terminal-interactive.

---

## CLI behavior (`update_icms_snapshot_files.py`)

### Single-file mode
```bash
python update_icms_snapshot_files.py \
  --input "ICMS Snapshot - 01.28.2026.xlsx" \
  --snapshot-date 2026-01-28 \
  --sheet "Participant Level Summary"
```

### Batch mode
```bash
python update_icms_snapshot_files.py \
  --batch "./snapshots" \
  --output-dir "./output"
```

CLI batch:
- scans folder for `.xlsx`,
- parses date from filename (`MM.DD.YYYY`),
- confirms/corrects dates interactively,
- processes oldest→newest,
- continues when individual files fail.

Additional CLI features:
- startup settings display,
- snapshot-date-gap warning (>45 days),
- settings persistence after runs.

---

## GUI behavior (`icms_gui.py`)

Single-window app with 3 panels (Setup → Run → Results), background worker thread, queue-based UI updates.

### Setup panel
- Single-file picker (`Browse…`) and multi-file picker (`Browse files…`).
- Batch file summary display.
- Snapshot date input (disabled in multi-file mode; dates come from batch confirmation).
- Editable sheet names (participant, vacant, exit).
- Password input with optional keyring save/load.
- Output folder picker.
- Run behavior toggle: publish current dashboard file on/off.
- Inline validation.

### Batch confirmation
- Modal list of selected files with parsed dates.
- Date correction in `YYYY-MM-DD`.
- Requires valid date for every file.
- Sorts processing oldest→newest.

### Duplicate-import warning / exclusion
Before ingest, GUI checks selected output folder history for already integrated tuples:
- match key: `source_file_name + snapshot_date`.
- sources: `icms_snapshot_history.csv` (primary), successful `icms_ingest_log.csv` rows (secondary).

Single-file:
- warns and asks user to cancel or continue.

Batch:
- lists duplicates and allows:
  - exclude flagged files (recommended),
  - include all,
  - cancel batch.
- handles “all excluded” safely by stopping run.

### Run panel
- 7-step progress indicator.
- live log stream.
- resilient per-file behavior in batch (continue on file failure).
- missing vacant/exit sheets are warnings, not hard failures.

### Batch dashboard publish rule
- In batch mode, dashboard flatfile publishes only for newest snapshot in confirmed batch.
- Earlier files in batch update history/master but do not overwrite dashboard file.

### Results panel
- Single-file: detailed run summary.
- Batch: files processed/succeeded/failed, dashboard source file, per-file outcomes.
- Includes skipped-duplicate reporting when applicable.
- Clickable output folder link.

### UX style
Warm off-white palette, muted accent color, flat controls, generous spacing, fixed-width design with vertical scrolling support in Setup.