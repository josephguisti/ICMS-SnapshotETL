# Maintainer Guide

This guide is written for a successor who can use Cursor or another LLM but may not know the codebase. The **human-authored intended behavior** is documented in **`architecture planning.md`** at the repository root (title: *ICMS Snapshot ETL — Architecture and Behavior Summary*). This file turns that summary into a maintenance map and flags places where the **summary and the code diverge**.

---

## 1. What this app does

**Plain-language overview** (aligned with `architecture planning.md`, verified against the code):

- **Problem it solves:** Weekly DHS/HSH-style ICMS snapshot Excel workbooks vary slightly in layout and column naming. The app normalizes them into stable outputs for reporting and cumulative history.
- **Input:** One or more `.xlsx` files per run. Files may be **password-protected**; decryption is in memory via optional `msoffcrypto-tool` (`icms_etl_core.open_workbook`).
- **Processing:** Reads a **participant** sheet (multi-row vendor headers supported via `find_header_row`), optional **vacant**, **exit**, and **services** sheets, maps columns to a **canonical field catalogue** (`REQUIRED_FIELDS`, `OPTIONAL_FIELDS`, `MOST_RECENT_FIELDS`, etc. in `icms_etl_core.py`), drops rows with blank `CHAMP ID` or `ICMS Project Name` (`reduce_to_approved_fields`), detects duplicate keys (`find_duplicates`), updates **dashboard flatfile**, **CSV histories**, **client/project master**, **ingest log**, and **duplicate warning** CSV. Unresolved optional fields are backfilled with null so output columns stay stable when older workbooks lack newer vendor headers. Services events are **deduplicated** in `update_services_log_history` using `_build_services_event_key_df`.
- **Outputs:** Excel and CSV files in a user-chosen **output folder** (see §7). `field_aliases.json` and `settings.json` are stored next to the **GUI script directory** (`Path(__file__).parent` in `icms_gui.ICMSApp`), i.e. the folder containing `icms_gui.py` when run from source.
- **Who outputs are for:** Operational staff and downstream analytics; the architecture summary describes a **Tableau-oriented** “current dashboard” flatfile. The repo does **not** contain Tableau workbooks or connection strings—only file outputs you point Tableau at.

---

## 2. How the app is organized

| Path | Purpose | Why a maintainer might need it |
|------|---------|--------------------------------|
| `architecture planning.md` | Authoritative *intended* behavior and field catalogue | Start here for “what should happen”; compare to code via §11. |
| `icms_etl_core.py` | All shared ETL: constants, read/write, dedupe, master rebuild | Almost any data rule or output change touches this file. |
| `icms_gui.py` | Tkinter GUI: setup, batch UI, worker thread, dialogs, settings | UX, batch orchestration, duplicate/service warnings, ingest log shape. |
| `update_icms_snapshot_files.py` | CLI entry (`argparse`), settings load/save, batch/single runners | Legacy/dev path; **not feature-parity with GUI** (see §11). |
| `field_aliases.json` | Canonical field name → workbook column name (optional; created at runtime) | Survives vendor column renames; lives beside `icms_gui.py` when using GUI. |
| `settings.json` | Last-used dirs, sheet names, publish toggle, etc. | Persisted by GUI via `load_settings` / `save_settings` from `update_icms_snapshot_files.py`. |
| `requirements.txt` | Runtime deps: `openpyxl`, `pandas`, `msoffcrypto-tool`, `tkcalendar`, `keyring` | Version bumps can affect Excel I/O, dates, or optional GUI features. |
| `version.py` | Application version (`__version__`); used by GUI title, PyInstaller spec, and `build_exe.py` | Bump before each release EXE build. |
| `build_exe.py` | Builds versioned EXE via PyInstaller spec | Preferred packaging entry point. |
| `ICMS Snapshot ETL.spec` | PyInstaller spec (one-file windowed EXE from `icms_gui.py`; reads `version.py` for output name) | Packaging changes. |
| `build/` | PyInstaller build artifacts (often git-ignored locally) | Regenerated on build; do not hand-edit. |
| `.gitignore` | Ignores venv, `Outputs/`, several standard output filenames | Prevents committing generated data; extend if new outputs should stay private. |
| `Sample input/`, `Sample output/` (if present) | Example workbooks / outputs | Manual regression; not wired to automated tests. |

**Not in repo:** There is **no** root `README.md` at the time of this writing.

---

## 3. End-to-end workflow

1. **Start:** `python icms_gui.py` (see `icms_gui.py` module docstring). PyInstaller build per `architecture planning.md`: `python build_exe.py` → `dist/ICMS Snapshot ETL v<version>.exe`.
2. **Setup panel (`ICMSApp`):** Pick one file (`Browse…`) or multiple (`Browse files…`). When **more than one** file is selected, the UI switches to **batch mode** (`len(self._batch_files) > 1`): per-file snapshot dates are edited in **inline rows** on the Setup panel (`_render_batch_date_rows`, `_collect_batch_dates`). Enter participant / vacant / exit / services **sheet names**, optional **password** (optional **keyring** save), output folder, and **“publish current dashboard file”** toggle.
3. **Run:** `_on_run` validates, checks **duplicate imports** (`_load_integrated_file_keys`, `_find_duplicate_candidates`) and **services-table duplicate** heuristics (`_services_table_already_integrated`), then starts `_etl_worker` on a **background thread**.
4. **Per file (`_process_one_file`):** `core.open_workbook` → list sheets → **participant:** `find_header_row` + `read_sheet` → **vacant/exit:** `read_sheet` + column subsets (GUI) → **services:** `read_services_log` → `_gui_resolve_fields` / `reduce_to_approved_fields` → duplicates → `update_dashboard_flatfile` (if enabled) → `update_snapshot_history` → `update_client_project_master` (which calls `update_exit_log_history` inside) → step 7: `summarize_services_overlap`, `update_services_log_history`, optional `append_duplicate_warning`, `append_ingest_log`.
5. **Passwords:** GUI passes `password` into `open_workbook` so CLI `getpass` is not used. Optional keyring integration in GUI (`_KR_SERVICE`, `_KR_USER`).
6. **Logs:** Live log in Run panel via queue; structured row in `icms_ingest_log.csv` via `append_ingest_log`.
7. **Success:** Results panel shows counts, paths, batch summary; `run_status` success in ingest log.

**GUI vs CLI (`architecture planning.md` §“CLI status”):** The summary is correct that **production intent is GUI-first** and the CLI is for dev/backward compatibility. The repo **does** ship both `icms_gui.py` and `update_icms_snapshot_files.py`. The CLI is **not** a full substitute for the GUI pipeline (see §11).

---

## 4. Input file assumptions

| Input assumption | Where it is handled in code | What breaks if it changes | How to update it |
|------------------|----------------------------|---------------------------|------------------|
| File is `.xlsx` read with `openpyxl` | `open_workbook`, `read_sheet` | Other formats fail | New format would need a new reader (large change). |
| Encrypted workbooks need `msoffcrypto-tool` | `open_workbook` | Open fails with install hint | Ensure `requirements.txt` + venv include `msoffcrypto-tool`. |
| Participant sheet name is user-supplied and must exist | `icms_gui._process_one_file` (raises if missing) | Hard error | User edits Setup field or `settings.json` key `last_sheet_name` (GUI load/save). |
| Participant headers may span multiple rows | `find_header_row`, then `read_sheet(..., header_row=...)` | Wrong columns / empty data | Adjust `find_header_row` logic or approved-field list used for scoring. |
| Vacant / exit / services sheet names are user-configured (**GUI**) | `icms_gui._process_one_file` | Missing sheets → warning, skipped data | Edit Setup defaults or persisted `last_*_sheet_name` in `settings.json`. |
| CLI always looks for sheets named **`Vacant Slots`** and **`Exit Log`** | `read_vacant_slots`, `read_exit_log` in `icms_etl_core.py` | CLI silently skips data if names differ | Rename sheets in source **or** extend CLI to pass sheet names (currently not exposed). |
| Vacant sheet columns (subset) | GUI: `vacant_keep_cols`; core `read_vacant_slots`: same list | Vacant rows missing fields | Edit the keep list in **both** places if CLI must stay aligned, or refactor to one helper (future change). |
| Exit sheet columns (subset) | GUI: `exit_keep_cols`; core `read_exit_log`: same list | Missing exit fields in master | Same as vacant. |
| Services log columns | `SERVICES_LOG_FIELDS`, `read_services_log` | Missing keys for dedupe / Tableau | Extend `SERVICES_LOG_FIELDS` and any logic using those names. |
| `CHAMP ID` / `ICMS Project Name` identify rows | `KEY_FIELDS`, `reduce_to_approved_fields`, `find_duplicates`, master logic | Wrong grain, wrong dedupe | Change with extreme care (§12). |
| Numeric Excel cells for keys normalized to integer strings | `read_sheet` (`_norm_key_field` for `KEY_FIELDS`) | Mismatched keys across files | Adjust normalization in `read_sheet`. |
| Rows blank on all data columns dropped | `read_sheet` | Sparse rows vanish | Change `dropna` logic if business requires keeping them. |
| Required / optional canonical fields | `REQUIRED_FIELDS`, `OPTIONAL_FIELDS`, `ALL_APPROVED_FIELDS` | Validation / column expectations shift | Edit lists in `icms_etl_core.py` and alias prompts; update `architecture planning.md` if catalogue is authoritative. |
| Architecture doc merges required fields into one list; code splits | `REQUIRED_KEY_FIELDS` + `REQUIRED_QUEUE_FIELDS` → `REQUIRED_FIELDS` | None if union unchanged | Maintain union consistency when editing. |

---

## 5. Field resolution and aliases

| Behavior | Code location | Data/file dependency | Maintenance notes |
|----------|---------------|----------------------|-------------------|
| Resolution order: exact header → saved alias → unresolved | `resolve_fields` in `icms_etl_core.py` | `field_aliases.json` via `load_aliases` / `save_aliases` | Same order as `architecture planning.md` §“Field resolution and aliases”. |
| CLI prompts for unresolved required and optional | `resolve_fields_interactive` → `prompt_for_missing_field` | Terminal stdin | Optional fields: Enter skips. |
| GUI prompts only for **unresolved required**; optional skipped silently | `_gui_resolve_fields`, `_FieldAliasDialog` | `field_aliases.json` | Doc said optional may be skipped—GUI never asks for optional columns; they are simply absent if not auto-resolved. |
| Saving aliases | `save_aliases(script_dir)` | `field_aliases.json` next to `icms_gui.py` (GUI `self._script_dir`) | Stale aliases map to wrong columns—delete or edit JSON carefully. |
| Required field skipped with user confirm | `_FieldAliasDialog._on_skip` | None | Produces `missing_req` notes in ingest log. |

---

## 6. Business rules and transformation logic

| Rule / transformation | Code location | Depends on | Affects | Maintenance notes |
|----------------------|---------------|------------|---------|-------------------|
| `reporting_month` = first 7 chars of `snapshot_date` | `reduce_to_approved_fields`, exit/services writers | `snapshot_date` string `YYYY-MM-DD` | All outputs carrying `reporting_month` | Wrong date string format breaks month. |
| Drop participant rows with blank key fields | `reduce_to_approved_fields` | `CHAMP ID`, `ICMS Project Name` | Row counts, all downstream | Changing keys changes everything downstream. |
| Same-date snapshot history replace | `update_snapshot_history` | `snapshot_date` column in CSV | `icms_snapshot_history.csv` | Idempotent per date. |
| Exit history replace per file+date | `update_exit_log_history` | `snapshot_date`, `source_file_name` | `icms_exit_log_history.csv` | Rerun same file/date replaces rows. |
| Services event key: `Service ID` if nonblank else fallback composite | `_build_services_event_key_df` | `SERVICES_LOG_FIELDS` columns | `icms_services_log_history.csv` dedupe | Changing key definition changes “same event” identity. |
| Services dedupe: earliest `snapshot_date` then `source_file_name` wins | `update_services_log_history` (sort + `groupby().first()`) | History file | Cumulative services CSV | Order-independent **within current key definition**. |
| Duplicate CHAMP+Project in one ingest | `find_duplicates`, `build_duplicate_warning_rows`, `append_duplicate_warning` | Participant data | `icms_duplicate_client_project_warning.csv` | Same-date replace in `append_duplicate_warning`. |
| Master rebuilt from histories each run | `update_client_project_master` → `_build_master_from_histories` | `icms_snapshot_history.csv`, `icms_exit_log_history.csv` | `icms_client_project_master.xlsx` | **Not** incremental merge from prior master file alone. |
| Latest participant row per key | `_build_master_from_histories` (`groupby(KEY_FIELDS).last()` after sort by `_snap`) | Snapshot dates in history | Master participant columns | “Last” = latest `snapshot_date` in sorted history. |
| `present_in_latest_snapshot` / `client_status` | `_build_master_from_histories` | Global max snapshot in participant history; exit keys | Master tracking columns | Logic matches architecture summary (Active / Exited / Unknown). |
| Exit fields on master | `_build_master_from_histories` + exit_map merge | Exit history columns | `EXIT_FIELDS` on master | Uses latest exit row by `_event` ordering. |
| Dashboard flatfile overwrite | `update_dashboard_flatfile` | Participant `df`, `vacant_rows` | `icms_current_dashboard_flatfile.xlsx` | Concatenates vacant synthetic rows (`build_vacant_slot_rows`). |
| Ingest log one row per `snapshot_date` | `append_ingest_log` | `snapshot_date` in log row | `icms_ingest_log.csv` | Replaces prior row for same date. |
| `upsert_exit_log_to_master` | `icms_etl_core.py` (defined only) | — | — | **Not called** by GUI or CLI; master path uses `_build_master_from_histories` only. See §11. |

---

## 7. Output files and output structure

| Output | Format | Grain | Created by | Important columns | Downstream use | How to modify |
|--------|--------|-------|------------|-------------------|----------------|---------------|
| `icms_current_dashboard_flatfile.xlsx` | Excel | One row per participant row + vacant synthetic rows | `update_dashboard_flatfile` | Approved fields + metadata + `is_vacant_slot` | Tableau / “current” reporting | Change writer; keep column names stable for Tableau. |
| `icms_snapshot_history.csv` | CSV | One row per participant row per `snapshot_date` | `update_snapshot_history` | Catalogue + metadata | Cumulative history; feeds master rebuild | Adjust `reduce_to_approved_fields` output columns. |
| `icms_exit_log_history.csv` | CSV | One row per exit sheet row per ingest | `update_exit_log_history` | Exit columns + metadata | Master exit merge | Column list from GUI/core exit readers. |
| `icms_services_log_history.csv` | CSV | One row per deduplicated service event | `update_services_log_history` | `SERVICES_LOG_FIELDS` + metadata | Cumulative services | Dedupe key in `_build_services_event_key_df`. |
| `icms_client_project_master.xlsx` | Excel | One row per `CHAMP ID` + `ICMS Project Name` | `update_client_project_master` | Full master column set in `_build_master_from_histories` | Deterministic roster | Rebuild logic in `_build_master_from_histories`. |
| `icms_ingest_log.csv` | CSV | One row per `snapshot_date` (latest wins) | `append_ingest_log` | Many operational + `-diag` columns (GUI) | Audit / ops | Extend `log_row` dict in `_process_one_file`; CLI row shape differs (fewer services columns). |
| `icms_duplicate_client_project_warning.csv` | CSV | Rows implicated in duplicates | `append_duplicate_warning` | See `build_duplicate_warning_rows` | QA | Schema in `build_duplicate_warning_rows`. |
| `field_aliases.json` | JSON | N/A | `save_aliases` | Canonical → column | Less manual mapping | Hand-edit or clear bad mappings. |
| `settings.json` | JSON | N/A | `save_settings` in GUI flow | UI defaults | Convenience | Keys written by GUI save handler in `icms_gui.py`. |

**Name check:** Output filenames match `architecture planning.md` for the files listed there.

---

## 8. Maintenance surface area

| Maintenance area | Common reason for change | Files/functions to inspect | Risk level | Notes |
|------------------|-------------------------|----------------------------|------------|-------|
| Input workbook format / Excel quirks | Vendor export change | `read_sheet`, `open_workbook` | High | Float vs string keys already handled—new quirks need evidence. |
| Sheet names | ICMS renames tabs | `icms_gui.py` Setup; CLI `read_vacant_slots` / `read_exit_log` | **High for CLI** | CLI hardcodes vacant/exit names. |
| Column names / mapping | Header rename | `resolve_fields`, `field_aliases.json`, `_gui_resolve_fields` | Medium | Prefer aliases before code changes. |
| Required / optional catalogue | Policy change | `REQUIRED_FIELDS`, `OPTIONAL_FIELDS` | High | Touches validation, GUI prompts, outputs. |
| Field aliases | Drift / new vendor | `field_aliases.json`, `load_aliases` / `save_aliases` | Medium | JSON syntax errors break load. |
| Password / encryption | Security tooling change | `open_workbook`, GUI password + keyring | Medium | Missing `msoffcrypto-tool` blocks encrypted files. |
| Validation | New mandatory checks | `reduce_to_approved_fields`, GUI `_validate` | Medium | Distinguish “warn” vs “fail”. |
| Business rules / transforms | Reporting change | `reduce_to_approved_fields`, `_build_master_from_histories`, vacant/exit builders | High | Many implicit assumptions (§6). |
| Participant dedupe (same file) | QA policy | `find_duplicates`, warning append | Medium | Does not remove rows from main outputs—warns only. |
| Services dedupe / event key | Double-counting in services history | `_build_services_event_key_df`, `update_services_log_history` | **High** | Small logic change redefines “same event”. |
| Client/project master | Status / exit display | `_build_master_from_histories`, `update_client_project_master` | **High** | Full rebuild—ordering and joins matter. |
| Output columns | Tableau field change | Writers in `icms_etl_core.py`, `reduce_to_approved_fields` | **High** | Silent Tableau breaks if columns renamed. |
| Output naming | File convention | Paths inside `update_*` functions | Medium | Search string literals ending in `.csv` / `.xlsx`. |
| Tableau compatibility | Calc / join change | (downstream) + stable column names here | Medium | Repo has no `.twb`; validate externally. |
| Dashboard publish behavior | Weekly vs backfill | `_run_batch_worker`, `_process_one_file` | **High** | Batch behavior has a documented quirk (§11). |
| GUI copy / steps | UX | `icms_gui.py` strings, `_STEP_LABELS` | Low | Step labels do not list every sub-step (e.g. services in step 7). |
| File pickers | OS dialog change | `filedialog` usage in `icms_gui.py` | Low | |
| Batch ordering | Chronology | `_run_batch_worker` sorts by date string | High | Lexicographic sort works for `YYYY-MM-DD`. |
| Duplicate-import warnings | Ops policy | `_BatchDuplicateDialog`, `_load_integrated_file_keys` | Medium | Matches `source_file_name` + `snapshot_date`. |
| Missing sheet behavior | Vendor drops sheet | `icms_gui._process_one_file` branches | Low | Warnings + notes; participant sheet still required. |
| Error messages | Clarity | `messagebox`, `log()`, exceptions | Low | |
| Logging / ingest diagnostics | New KPI | `log_row` construction in `icms_gui.py` | Low | Many `*-diag` columns—Tableau may ignore. |
| Sample / test data | Regression | `Sample input/` etc. | Low | No automated `pytest` in repo. |
| Dependencies | Security / compat | `requirements.txt` | Medium | Pin versions if production needs stability. |
| Packaging | EXE distribution | `version.py`, `build_exe.py`, `ICMS Snapshot ETL.spec` | Medium | Bump `__version__` before release; test EXE on clean VM; see §14 for path behavior. |

---

## 9. LLM-assisted maintenance playbook

Use **`architecture planning.md`** as the behavioral contract, and the files below as the implementation. If the LLM “fixes” a mismatch without human approval, stop and use §10 / §11.

| Task | Point the LLM at | Ask the LLM to do this | Human check before accepting changes |
|------|------------------|------------------------|--------------------------------------|
| Change an input column name (vendor rename) | `architecture planning.md` (Approved field catalogue); `icms_etl_core.resolve_fields`, `field_aliases.json` | Explain current resolution; propose alias JSON update or code change **only** if alias is insufficient | Run GUI on real workbook; confirm outputs have non-empty mapped columns. |
| Add / remove / reclassify required or optional field | `icms_etl_core.REQUIRED_*`, `OPTIONAL_FIELDS`, `MOST_RECENT_FIELDS`, `ALL_APPROVED_FIELDS`; `reduce_to_approved_fields`; `architecture planning.md` | Make minimal constant + downstream writer changes; list all writers affected | Tableau + master + CSV column alignment; update architecture doc if catalogue is official. |
| Change participant sheet | `icms_gui.py` (sheet `StringVar` / settings); `_process_one_file` | Trace sheet name from UI to `read_sheet` | Confirm `find_header_row` still finds headers. |
| Support multiple possible sheet names | `icms_gui.py`, optionally new helper in `icms_etl_core.py` | Try first existing sheet from a list | Test with workbooks that have only one of the names. |
| Change vacant/exit **GUI** sheet handling | `icms_gui._process_one_file` (`vacant_keep_cols`, `exit_keep_cols`) | Align column lists with vendor export | Compare to `read_vacant_slots` / `read_exit_log` if CLI must match. |
| Fix CLI vacant/exit names only | `icms_etl_core.read_vacant_slots`, `read_exit_log`; `update_icms_snapshot_files.py` | Parameterize sheet names with argparse flags (narrow addition) | Regression on existing CLI scripts. |
| Change password behavior | `icms_etl_core.open_workbook`; GUI password fields + keyring | Document and implement without writing decrypted files to disk | Test encrypted sample workbook. |
| Update date parsing (filename batch) | `update_icms_snapshot_files._parse_date_from_filename`; GUI uses same helper | Adjust regex / format; keep `YYYY-MM-DD` output | Batch order and duplicate detection. |
| Update reporting month logic | `reduce_to_approved_fields` (`snapshot_date[:7]`), services/exit inserts | Change derivation consistently | Downstream month filters in Tableau. |
| Change participant dedupe / duplicate warnings | `find_duplicates`, `build_duplicate_warning_rows`, `append_duplicate_warning` | Preserve warning file schema or migrate intentionally | QA review of warning CSV. |
| Change services event key | `_build_services_event_key_df`, `summarize_services_overlap`, `architecture planning.md` §services | Prove two rows merge or split under new key | Re-run same workbook twice; compare `icms_services_log_history.csv` row counts. |
| Change master / `client_status` logic | `_build_master_from_histories`, `update_client_project_master` | Diff behavior vs prior master for fixture CSVs | Spot-check Active/Exited/Unknown against business rules. |
| Add output column | `reduce_to_approved_fields` or relevant writer (`update_*`) | Add column once at source of truth | Tableau workbook refresh. |
| Rename output file | Each `output_dir / "..."` literal in `icms_etl_core.py` + any GUI text | Global rename with care for ingest/history readers | Any external scripts reading old names. |
| Change batch dashboard publish rules | `icms_gui._run_batch_worker`, `_etl_worker`; `architecture planning.md` | **Read §11 first**—today’s code may not match the doc’s “toggle” story | Explicit sign-off from stakeholder. |
| Change GUI text / steps | `icms_gui.py` labels, `_STEP_LABELS` | Copy-only diff | N/A |
| Improve error messages | `icms_gui._process_one_file` exception paths; `open_workbook` | Clear, non-leaky messages | Show sample to end user. |
| Add logging field to ingest | `icms_gui.py` `log_row` dict | Add column with safe string conversion | CSV consumers tolerate new column. |
| Packaging / PyInstaller | `version.py`, `build_exe.py`, `ICMS Snapshot ETL.spec`, `architecture planning.md` §Build | Bump version, rebuild; document command | Run EXE from §13. |
| Troubleshoot failed run | Last `icms_ingest_log.csv` row; GUI log; `architecture planning.md` workflow | Trace exception to function | Reproduce on sample workbook. |
| Tableau mismatch | Output writers + field catalogue docs | Map Tableau field to code column | Validate in Tableau, not only in code. |

---

## 10. Safe prompting patterns for future maintainers

Copy/paste and replace the bracketed parts.

### Before changing code

> Read **`architecture planning.md`** (sections on GUI workflow and outputs) and the implementation in **`[icms_etl_core.py | icms_gui.py | path]`** for **`[behavior]`**. Summarize current behavior as bullet steps and list all functions involved. **Do not edit any files.**

### Small targeted change

> Make the **smallest possible change** to **`[file]`** so that **`[observable outcome]`**. Do **not** refactor unrelated code, rename variables broadly, or change other outputs. If you touch more than **`[N]`** functions, stop and explain why.

### Debugging

> Given input **`[file.xlsx]`** and output **`[unexpected symptom]`**, trace the data path from **`open_workbook` / `read_sheet`** through **`reduce_to_approved_fields`** and **`[writer]`** in **`icms_etl_core.py`** and the GUI worker in **`icms_gui.py`**. Identify the first function where results diverge from **`architecture planning.md`**. Propose a fix with a minimal diff.

### Input file changed

> Compare **`architecture planning.md`** input assumptions to the actual headers/sheets in **`[new workbook path]`**. List differences, then list exact code edits needed in **`icms_etl_core.py`** / **`icms_gui.py`** (constants, column lists, `SERVICES_LOG_FIELDS`, etc.). **Do not assume** the architecture doc is wrong without labeling it a doc update.

### Output / dashboard request changed

> Tableau (or reporting) needs **`[column/logic change]`**. Trace which writers populate affected fields (**`update_dashboard_flatfile`**, **`update_snapshot_history`**, **`_build_master_from_histories`**, **`update_services_log_history`**). Propose the **smallest** set of edits and list downstream files that must stay column-compatible.

### Architecture / code mismatch review

> Compare **`architecture planning.md`** to the current code in **`icms_etl_core.py`**, **`icms_gui.py`**, and **`update_icms_snapshot_files.py`**. List each claim in the doc that is **not implemented**, **partially implemented**, or **implemented differently**—with file/function evidence. Do **not** “fix” the doc or code in this pass; only report.

### Post-change review

> Review the diff you just made for **`[feature]`**. Check for regressions: **`KEY_FIELDS` handling**, **services dedupe keys**, **master rebuild inputs**, **GUI thread safety** (queue usage), and **CSV schemas**. List risks and any tests or manual steps I should run.

---

## 11. Architecture / code mismatches to review

| Topic | Architecture summary says | Code appears to do | Risk | Recommendation |
|-------|---------------------------|--------------------|------|------------------|
| Batch date confirmation UI | Modal **batch confirmation** dialog listing files and dates (`_BatchConfirmDialog` described in spirit) | **`_BatchConfirmDialog` is never instantiated.** Batch uses **inline per-file date fields** on the Setup panel (`_collect_batch_dates`). | Medium | Treat the modal class as **dead code** unless you wire it up; update `architecture planning.md` to describe inline batch dates **or** restore the modal workflow. |
| Batch “publish dashboard” toggle | User can choose to publish only for newest file (**recommended**); batch rule publishes only newest | **`_run_batch_worker` always sets `publish_latest_only=True`**, so **only the chronologically last file** in the sorted batch gets `publish_dashboard=True`. The **Setup toggle is ignored for batch** (single-file still honors it). | **High** | Decide intended UX; if toggle should apply, thread `publish_dashboard` into batch logic; if not, document that batch always publishes the latest date’s successful file only. |
| CLI feature parity | CLI documented for batch/single with same core | **`update_icms_snapshot_files.py` never calls `read_services_log` or `update_services_log_history`.** No `icms_services_log_history.csv` from CLI. | **High** | Use GUI for full pipeline **or** extend CLI to mirror `_process_one_file` services steps. |
| CLI vacant / exit sheet names | Implied general “workbook” model | **`read_vacant_slots` / `read_exit_log` hardcode `"Vacant Slots"` and `"Exit Log"`** | **High** | Parameterize CLI or document exact required sheet titles. |
| Optional field prompting (GUI) | Optional unresolved fields may be skipped | GUI **`_gui_resolve_fields`** only loops **`unresolved_req`**; optional unresolved columns are **never prompted** | Low–medium | If optional mapping matters, add optional dialogs or document “map via exact header or alias only.” |
| Master merge helper | (summary describes master behavior at high level) | **`upsert_exit_log_to_master` exists but is unused** by GUI/CLI | Low | Remove if dead, or integrate if it was meant to replace part of `_build_master_from_histories`—needs human decision. |
| Ingest log schema | Summary lists high-level tracking | **GUI** `log_row` includes many extra `*-diag` keys vs **CLI** `log_row` | Low | Any external parser of `icms_ingest_log.csv` must tolerate optional columns. |

---

## 12. What not to touch casually

- **`_build_services_event_key_df` / `update_services_log_history`:** Redefines identity of service events; mistakes **silently** merge or split history.
- **`KEY_FIELDS` and `reduce_to_approved_fields` blank-key drop:** Feeds every aggregate; changing keys without a migration plan breaks continuity with past CSVs.
- **`_build_master_from_histories`:** Full rebuild from CSVs; subtle changes to “latest” row selection or `client_status` logic propagate to every master row.
- **`read_sheet` normalization for `KEY_FIELDS`:** Cross-file CHAMP ID matching depends on numeric/string normalization.
- **Output column names** consumed by Tableau: renaming is a breaking change unless Tableau is updated in lockstep.
- **Batch publish logic (`_run_batch_worker`):** Easy to produce a dashboard file that **does not** match operator expectations relative to the Setup toggle (see §11).
- **`open_workbook` / encryption path:** Security-sensitive; avoid writing temporary decrypted paths.
- **PyInstaller one-file EXE:** Combined with `Path(__file__).parent` for settings/aliases, runtime paths for frozen apps can surprise users—validate on a clean machine (§14).

---

## 13. Testing and validation checklist

**Automated tests:** None found (`test_*.py` absent). Validation is **manual**.

After any change:

- [ ] `python icms_gui.py` launches without import errors.
- [ ] Single-file: browse workbook, set date, run → success in Results + `run_status=success` in `icms_ingest_log.csv`.
- [ ] Multi-file: select 2+ files, set each date (`YYYY-MM-DD`), confirm sort order processes oldest → newest (`_run_batch_worker` sorts by date string).
- [ ] Encrypted sample: confirm `msoffcrypto-tool` installed and password path works.
- [ ] Wrong participant sheet name → clear error (not silent).
- [ ] Missing optional sheets → warnings in log / notes, run continues (`icms_gui._process_one_file`).
- [ ] Outputs present: dashboard (if published), `icms_snapshot_history.csv`, `icms_exit_log_history.csv`, `icms_services_log_history.csv`, `icms_client_project_master.xlsx`, `icms_ingest_log.csv`, duplicate warning only if duplicates.
- [ ] Row counts plausible vs raw participant count (account for blank-key drops).
- [ ] Re-run **same** `source_file_name` + `snapshot_date` → snapshot history replaced for that date; exit history rows for that pair replaced; ingest log row for date replaced.
- [ ] Services: rerun same services data → deduped row count stable (no unbounded growth).
- [ ] Client/project master: spot-check `client_status` and `present_in_latest_snapshot` against known clients.
- [ ] If Tableau is in scope: refresh extracts pointing at the output folder.
- [ ] Optional: `python update_icms_snapshot_files.py --help` still runs; understand CLI **does not** update services history (§11).
- [ ] EXE: rebuild per `architecture planning.md`; smoke-test on a machine without Python.

---

## 14. Known unknowns / documentation gaps

- **Tableau:** No workbook or data-source definitions in-repo; all “dashboard depends on X column” knowledge may live outside git.
- **Business meaning** of specific queue flags and ICMS policy for “required” fields: only partially encoded in constant names—ultimate authority may be program staff, not code.
- **Frozen EXE paths:** Whether `settings.json` / `field_aliases.json` land where operators expect when using the one-file EXE depends on PyInstaller runtime layout—**verify empirically** before wide rollout.
- **`upsert_exit_log_to_master`:** Purpose vs current `_build_master_from_histories` design is unclear without author context.
- **Services “already integrated” heuristic:** `_services_table_already_integrated` uses overlap threshold **0.98** and/or same file+date in history—false positives/negatives possible; no automated spec.

---

## 15. Recommended future documentation improvements

- Short **operator runbook** (non-coder): screenshots for weekly vs backfill, what folder to pick, what “already integrated” means.
- **Column dictionary** for each output CSV/Excel (maintained alongside `architecture planning.md`).
- **Tableau dependency map** (which calculated fields reference which output columns).
- **CLI/GUI parity decision:** Either document CLI as “participant-only pipeline” or bring it level with GUI.
- **Minimal pytest** with tiny synthetic workbooks checked into `tests/fixtures/` (de-identified).

---

## Quick reference: architecture document location

The planning source titled **“ICMS Snapshot ETL — Architecture and Behavior Summary”** is in **`architecture planning.md`** (repository root), not under `docs/`.
