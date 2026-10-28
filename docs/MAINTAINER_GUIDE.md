# Maintainer Guide

This guide is for a successor who can use Cursor or another LLM but may not know the codebase. The intended behavior is defined in **`architecture planning.md`** at the repository root. Field names and vendor mappings are defined only in **`icms_catalogue.py`**. This guide maps behavior to code and says how to keep it working.

---

## 1. Director maintenance instruction

When a run's Results panel lists **structural findings**, the analyst copies them ("Copy findings for director") and sends them to you. Nothing is blocked while you decide. For each finding:

1. **Unfamiliar column** — HSH added or renamed a header.
   - If it is a rename of a field we export, add the new header to that field's `sources=(...)` in `icms_catalogue.py`. Do not rename `output`; that breaks downstream readers.
   - If we do not want it, add it to `KNOWN_EXCLUDED_HEADERS` for that tab.
   - If it should become a new export, add a `FieldSpec` (append at the end so column order stays stable) and get sign-off first.
2. **Expected column missing** — HSH dropped or renamed a field.
   - If renamed, handle it as above.
   - If intentionally gone, set `retired=True`. For a field that never appears on that tab, set `report_missing=False`.
3. **Supported tab not found / Unfamiliar tab** — add the new tab name to `SHEETS` in `icms_catalogue.py`, or to `KNOWN_EXCLUDED_SHEETS` if it is not a data source.
4. Run `python -m pytest tests -p no:cacheprovider`, bump `version.py`, run `python build_exe.py`, and give the analyst the new EXE.
5. If the analyst needs the affected snapshots corrected, have them select those files and choose **Reprocess all selected**. Outputs are backed up automatically.

Analysts never edit mappings, and there is no local alias file.

---

## 2. How the app is organized

| Path | Purpose |
|------|---------|
| `README.md` | Front door: how to run, which document to trust. |
| `architecture planning.md` | Behavior contract: workflow, inputs, outputs, status rules, findings. |
| `icms_catalogue.py` | Output fields (`FieldSpec`), accepted vendor headers, essential/retired flags, `expected_since` dates, supported tabs (`SHEETS`), reviewed exclusions. |
| `icms_etl_core.py` | All ETL logic. Entry point `run_snapshot_files`; one file via `process_snapshot_file`. |
| `icms_gui.py` | Tkinter GUI: Setup → Run → Results. Calls `core.run_snapshot_files` on a worker thread. |
| `update_icms_snapshot_files.py` | CLI over the same entry point; also hosts `load_settings` / `save_settings`. |
| `tests/` | `test_import_rules.py` (synthetic workbooks, fast) and `test_reference_fixtures.py` (July/October fixtures, ~1.5 min). |
| `Sample input/`, `Deidentified Sample Output/` | Deidentified reference fixtures. Tests copy them to temp folders and never modify them. |
| `diagnostics/` | Older one-off A/B scripts. They run against compatibility helpers at the bottom of `icms_etl_core.py`; not part of the pipeline. |
| `icms_etl_core_vectorized.py` | Legacy copy used only by `diagnostics/ab_compare_master_rebuild.py`. |
| `version.py`, `build_exe.py`, `ICMS Snapshot ETL.spec` | Versioned one-file EXE build. |

---

## 3. Run path

`run_snapshot_files(files, output_dir, ...)`:
1. Probes existing outputs for write access (`check_outputs_writable` / `ensure_outputs_writable`) so a file open in Excel fails before any snapshot is processed. The GUI offers Retry/Cancel; the CLI prints the locked names and exits.
2. Sorts files oldest first, then backs up outputs if reprocessing or upgrading the format (`backup_outputs`, `output_format_is_current`).
3. Calls `process_snapshot_file` for each file. That function:
   - opens the workbook (`open_workbook`) and reads `Generated On` (`read_generated_on`);
   - picks tabs (`resolve_sheets`) and reads each through the catalogue (`read_category`);
   - validates essential fields and builds the participant frame (`build_participant_frame`);
   - writes everything inside `_OutputTransaction`, which restores every output if any write fails. The writers are `update_snapshot_history`, `update_exit_log_history`, `update_services_log_history`, the master rebuild (single file only), the dashboard (single file only), `append_duplicate_warning`, and `append_ingest_log`.
4. Rebuilds the master once for a batch (`rebuild_client_project_master` → `_build_master_from_histories`).
5. Publishes the newest successful file for a batch, if allowed.
6. Groups findings (`group_findings`, `format_findings_text`) and returns one result dict that the GUI and CLI display.

Snapshot-date defaults for the Setup panel come from `default_snapshot_date(path, password)`. The GUI's already-imported check is `already_imported(specs, output_dir)`.

---

## 4. Business rules and where they live

| Rule | Code |
|------|------|
| Header/tab normalization: trim, collapse whitespace, casefold | `icms_catalogue.normalize_header` |
| Coexisting accepted headers: first non-blank in catalogue order; conflicts recorded; identity conflicts block | `read_category` |
| Identifier normalization | `normalize_identifier` |
| Missing-field reporting suppressed for retired / by-design / pre-`expected_since` fields | `_field_expected` |
| Participant rows without CHAMP ID or project dropped | `build_participant_frame` |
| Snapshot history: replace by `snapshot_date` | `update_snapshot_history` |
| Exit history: replace by `source_file_name` + `snapshot_date` | `update_exit_log_history` |
| Services key, blank-Service adoption, source replacement, earliest non-blank per column | `_build_services_event_key_df`, `update_services_log_history` |
| Master grain, latest-observation selection, exit-only rows, Active/Exited/Unknown | `_build_master_from_histories` |
| Duplicate enrollment rows within a snapshot | `find_duplicates` (three-field key) |
| Batch publishing rule | end of `run_snapshot_files` |
| Ingest log fields | `log_row` dict in `process_snapshot_file` |

---

## 5. Where each output is written

| Output | Writer |
|--------|--------|
| `icms_current_dashboard_flatfile.xlsx` | `update_dashboard_flatfile` (columns `FLATFILE_COLUMNS`) |
| `icms_snapshot_history.csv` | `update_snapshot_history` |
| `icms_exit_log_history.csv` | `update_exit_log_history` |
| `icms_services_log_history.csv` | `update_services_log_history` |
| `icms_client_project_master.xlsx` | `rebuild_client_project_master` (columns `MASTER_COLUMNS`) |
| `icms_ingest_log.csv` | `append_ingest_log` |
| `icms_duplicate_client_project_warning.csv` | `append_duplicate_warning` |
| `backups/…` | `backup_outputs` |
| `icms_output_format.txt` | `write_format_marker` (bump `OUTPUT_FORMAT_VERSION` when stored outputs change shape) |
| `settings.json` | `save_settings` |

Excel date columns are every catalogue field with `is_date=True`, plus `snapshot_date` (`DATE_FIELDS`).

---

## 6. What not to touch casually

- **`FieldSpec.output` names**: these are the exported column names. Renaming one breaks every downstream reader. Add a source instead.
- **Catalogue order**: this is the output column order. Append new fields at the end.
- **`KEY_FIELDS` (CHAMP ID + project + enrollment)** and `_build_master_from_histories`: changing the grain or the "latest" selection changes every master row.
- **`_build_services_event_key_df`**: redefines what counts as the same service event. Mistakes silently merge or split history.
- **Replacement keys** in the three history writers: these are what make reprocessing a repair rather than a duplication.
- **`_OutputTransaction`**: guarantees a failed file leaves outputs unchanged.
- **`open_workbook`**: decryption stays in memory; never write decrypted copies.

---

## 7. Testing and validation

Automated:

```bash
python -m pytest tests -p no:cacheprovider
```

(`-p no:cacheprovider` avoids an OneDrive permission error on `.pytest_cache`.) The fixture tests check the acceptance figures: the October flatfile has 1,409 rows; the October exits keep their dates; exit replay repairs only the replayed rows; CHAMP 9038 and 51546 enrollments stay separate; the 41 July-only enrollments come out as 31 Exited and 10 Unknown; and reruns do not inflate counts.

Manual, after a release build:
- [ ] The EXE launches. Pick one workbook: the date fills from Generated On and a run succeeds.
- [ ] Pick several workbooks including one already imported: the reprocess dialog appears, the run succeeds, and a backup folder is created.
- [ ] An encrypted workbook opens with the saved password.
- [ ] The Results panel lists findings, and "Copy findings for director" pastes readable text.

---

## 8. LLM-assisted maintenance prompts

**Vendor header changed**

> Read `icms_catalogue.py` and the findings below. For each unfamiliar or missing column, propose the smallest catalogue edit (a source alias, an exclusion, `retired`, or `report_missing=False`). Do not change `output` names or `icms_etl_core.py`. Findings: [paste]

**Behavior question**

> Read `architecture planning.md` and `icms_etl_core.py`. Explain how [behavior] works, citing functions. Do not edit files.

**Small change**

> Make the smallest change to [file] so that [outcome]. Add or adjust one focused test in `tests/`. Do not refactor unrelated code.

**Post-change review**

> Review this diff for regressions in: enrollment key handling, services event keys, history replacement keys, master status rules, column order, and the per-file transaction. List risks and the tests that cover them.

---

## 9. Known limitations

- **Fixtures with blank project names:** the scrambled 02.18, 04.30, and 05.19.2026 samples have `ICMS Project Name` blank on every row, so the app correctly refuses them. Production files are not expected to look like this.
- **Unrepaired historical exits:** exits stored by earlier releases with blank dates are only repaired when their source workbook is reprocessed. In the deidentified exit history, 92 rows from July 7, 2026 onward have blank dates. The July 7 and October 19 fixtures repair 12 of them; the other 80 need their original workbooks.
- **Stored participant/master/services histories from production** were not available. Enrollment-ID completeness was confirmed only for the deidentified exit history (all 250 rows) and the deidentified snapshot history. A production history file without the `ICMS Enrollment ID` column stops the master rebuild with a clear message rather than merging enrollments.
- **Generated On on encrypted files:** a file can only be read once the password is entered. If a file is selected before the password, its date falls back to the filename; re-select it after entering the password.
- **Tableau:** no workbook definitions are in the repo. New columns are appended, and existing columns keep their names and order.
