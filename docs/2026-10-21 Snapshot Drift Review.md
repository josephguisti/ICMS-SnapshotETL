# ICMS Snapshot ETL — Snapshot Drift Review (2026-10-21)

Review of whether the ETL app (v1.0.1) still handles the current ICMS Snapshot layout, using:

- HSH change log: https://homeless.lacounty.gov/icms-snapshot-updates/ (entries through OCT 1 2026)
- Input: `Sample input/ICMS Snapshot - 023-ICMS-PH - 10.19.2026 - HFHPI2026 Scrambled.xlsx`
- Output: `Deidentified Sample Output/icms_exit_log_history Scrambled 10-19-26.csv`
- Output: `Deidentified Sample Output/icms_current_dashboard_flatfile Scrambled 10-19-26.xlsx`
- Older sample snapshots (02.18, 04.30, 05.19, 07.07) for header-drift comparison
- Internal docs: `architecture planning.md`, `docs/MAINTAINER_GUIDE.md`; code: `icms_etl_core.py`, `icms_gui.py`

All analysis was done by reading headers and values with the app's own `read_sheet` and comparing to the hardcoded field lists. No PHI-bearing production outputs were opened.

---

## 1. Headline findings

| # | Finding | Severity |
|---|---------|----------|
| 1 | **Services Log is no longer being ingested.** The tab was renamed `Services Log` → `Service Log` in the 10.19 snapshot. The app matches sheet names exactly (settings: `last_services_sheet_name = "Services Log"`), so the 10.19 run skipped services with only a warning. `icms_services_log_history.csv` stopped growing. | High |
| 2 | **Four REQUIRED participant fields are now blank in every output** because the vendor renamed/retired them: `Service Planning Area` → `Service Planning Area (SPA)`, `Permanent Housing Subsidy` → `Current Rental Subsidy`, `Receiving Pre-Match ICMS?` (retired), `Key Housing Milestone Data Needed?` (retired). The GUI prompts for these on every run and they are being skipped. | High |
| 3 | **Exit dates have been blank in the exit-log history since the 07.07.2026 snapshot.** The Exit Log tab renamed `Check-In Date`/`Check-Out Date`/`Move-In Date`/`Move-Out Date` → `ICMS Check-In Date`/`ICMS Check-Out Date`/`Permanent Housing Move-In Date`/`Permanent Housing Move-Out Date`. The app keeps columns by exact name, so 92 of 250 history rows (all rows from 2026-07-07 onward) have no dates. **This is why the October exits look missing — they are present (all 11) but undated.** | High |
| 4 | **Services columns have been partially blank since 07.07.2026.** `Service` → `ICMS Service`, `ICMS Provider Name` → `ICMS Provider`, `CHAMP User Who Created Record` → `CHAMP User Who Created Service Record`, `Organization of CHAMP User Who Created Record` → `...Service Record`, `DHS Master Agreement #` → `HSH Master Agreement #` (and dropped entirely in 10.19). Only `Service ID` kept the dedupe key intact. | High |
| 5 | **Vacant slots lost `# Days Vacant`** (renamed `Days Vacant` in 10.19). The column disappeared from the flatfile entirely (60 columns now; it was present before), which is a silent Tableau-breaking change. | Medium |
| 6 | **Three optional fields silently went blank** in 10.19: `Notification of Privacy Practices On Record?` → `...Status on Record?`, `CalAIM ... Opt-In Status Missing?` → `...Status Needed?`, `Family Billing Rate But Dependent Household Member(s) Not Listed...` → `...Eligible Dependent Household Member(s) Not Recorded...`. Optional fields are never prompted in the GUI, so nothing flagged this. | Medium |
| 7 | **~20 new participant indicators added since the catalogue was written are not surfaced** (Target Rental Subsidy, TLS / Shallow Subsidy Month, PHA Sharing Addendum Status, income/AMI, Medi-Cal MCP, Nominated for Graduation, Fiscal Agent, Care Plan Focus, etc.). See §4. | Medium (business decision) |
| 8 | Everything the app *was* able to match is working as documented: key normalization, blank-key drop, same-date replace, exit history replace-per-file+date, master rebuild, vacant synthetic rows, idempotent ingest log. | — |

---

## 2. Is the app doing what the internal documentation says it should?

**Mostly yes, for the parts of the workbook whose headers have not changed.** Verified against `architecture planning.md` §Output files and `MAINTAINER_GUIDE.md` §6:

| Documented behavior | Observed in 10-19 outputs | Status |
|---|---|---|
| Flatfile grain = one row per slot (participant + vacant), `is_vacant_slot` flag | 1,409 rows = 1,354 participant + 55 vacant; matches Provider Level Summary (Total ICMS Slots 1,409; Active Households 1,354; Vacant 55) | OK |
| Flatfile columns = metadata + approved catalogue + `is_vacant_slot`, null-backfilled | 60 columns, all 55 catalogue fields present, unresolved ones blank | OK (but see §1 #2, #6) |
| Dashboard overwritten per run; single snapshot_date | All rows `2026-10-19`, one `source_file_name` | OK |
| Exit history grain = exit rows per `snapshot_date`+`source_file_name`; rerun replaces | 250 rows across 36 snapshot dates (Jan 28 → Oct 19); each file+date appears once (06.08 appears twice because two differently-named files were ingested for that date — correct per the replace key) | OK |
| Exit history is accretive | Grew 165 → 250 between the two sample CSVs; all older rows retained unchanged | OK |
| Master: Active / Exited / Unknown from histories | Sample master is stale (built 07-20: 1,280 Active / 59 Exited / 32 Unknown); not re-exported for 10-19 so not re-verified | Not verified (stale sample) |
| Services history deduped by `Service ID` | Not re-verified — sample services history not provided; ingest log shows services processed through 07-20 | Not verified (and now broken per §1 #1) |
| Missing optional sheets are warnings, not failures | The 10.19 run completed despite `Services Log` being absent | OK (but the warning was the only signal) |
| Key normalization numeric→string | CHAMP IDs match across exit sheet, flatfile, and history (e.g. `53469`, `179512`) | OK |

**Deviations from documented intent**

- `architecture planning.md` says "Optional unresolved fields may be skipped" and the schema stays stable. True, but the practical effect is that **vendor renames of optional fields go completely unnoticed** — there is no log line, ingest-log anomaly, or UI warning when a field that was populated last week is unresolved this week. The ingest log's `missing_optional_fields-diag` records it, but nobody reads that routinely.
- The doc says the GUI "prompts for unresolved required fields"; it does, but a skip is **not remembered**, so the operator is asked four times on every run and has to click Skip → OK each time. This encourages reflexive skipping rather than mapping.
- Exit/vacant/services column lists are hardcoded exact-match lists (duplicated in `icms_gui.py` and `icms_etl_core.py`) and are **not covered by the alias mechanism** at all. Only participant fields can be aliased. The maintainer guide (§4) notes the duplication but not that renames can't be fixed without a code change.
- `MAINTAINER_GUIDE.md` §11 flags that the CLI hardcodes sheet names; the GUI now has the same practical problem because the configured sheet name (`Services Log`) no longer exists.

---

## 3. Are all exits being accretively added to the exit log?

**Yes — every exit on the 10.19 Exit Log tab is in the exported history.** The provisional "no" is explained by the blank date columns.

Exit Log tab in the 10.19 snapshot: 11 rows. Exit-history CSV rows with `snapshot_date = 2026-10-19`: 11 rows. CHAMP ID match: 11/11.

| CHAMP ID | HMIS ID | Enrollment | Project | ICMS Check-Out (sheet) | Exit Reason | In history? | Check-Out Date in history |
|---|---|---|---|---|---|---|---|
| 53469 | 9F8F965B9 | 533409 | SJC Cobb Apartments (MHSA) | 2026-10-15 | Transferred to Another HFH Project | Yes (10-19) | **blank** |
| 179512 | EA0DA100A | 326796 | SJC La Veranda (D7) | 2026-10-01 | Deceased | Yes (10-05, 10-13, 10-19) | **blank** |
| 151672 | EC49382D7 | 269770 | SJC Chesterfield Sr. DMH (D7) | 2026-10-16 | Deceased | Yes (10-19) | **blank** |
| 8582 | 04C48C820 | 283072 | SJC Firmin Court (D7) | 2026-10-07 | Non-Payment of Rent | Yes (10-13, 10-19) | **blank** |
| 12499 | 10BCC4146 | 159965 | SJC Scattered 1 (D7) | 2026-10-16 | Moved to Another Housing Opportunity | Yes (10-19) | **blank** |
| 88803 | 6A9035D31 | 188875 | SJC Metro at Buckingham Sr. BACKSTOP | 2026-10-01 | Non-Payment of Rent | Yes (10-05, 10-13, 10-19) | **blank** |
| 83527 | 7AC162DB2 | 145238 | SJC Immanuel Senior Housing (DHS) | 2026-10-16 | Change in Level of Care | Yes (10-19) | **blank** |
| 150145 | 8FC7EB22B | 401553 | SJC Immanuel Senior Housing (DHS) | 2026-10-16 | Client Left Voluntarily | Yes (10-19) | **blank** |
| 163007 | 6F4C170BD | 349347 | SJC Metro at Buckingham Sr. BACKSTOP | 2026-10-16 | Client Left Voluntarily | Yes (10-19) | **blank** |
| 148202 | EE4FB175B | 499526 | SJC Metro at Buckingham Sr. BACKSTOP | 2026-10-16 | Deceased | Yes (10-19) | **blank** |
| 101944 | 8F7260344 | 174483 | SJC Westmont Vistas (D7) | 2026-10-16 | Moved to Another Housing Opportunity | Yes (10-19) | **blank** |

None of the 11 appear as participant rows in the 10-19 flatfile, which is correct (they have been removed from the Participant Level Summary).

### Why they look missing

Exit-history fill rates, before vs. after the 07.07.2026 header change:

| Column | Rows before 2026-07-07 (n=158) | Rows from 2026-07-07 on (n=92) |
|---|---|---|
| Check-In Date | 158 | **0** |
| Check-Out Date | 158 | **0** |
| Move-In Date | 109 | **0** |
| Move-Out Date | 95 | **0** |
| Exit Reason / Destination | 156 | 92 |
| Date of Birth | 77 | 92 |

Any downstream view that filters or sorts exits by `Check-Out Date` shows nothing after June. The `reporting_month` column (derived from snapshot date) is the only usable date for post-July exits, and it is a *snapshot* month, not an *exit* month.

Secondary effect in the master: `_build_master_from_histories` orders exit rows by `Check-Out Date` → `Move-Out Date` → `snapshot_date`. With the first two blank, it silently falls back to snapshot date. Status derivation (Active/Exited/Unknown) is unaffected, but the master's exit date columns are blank for every exit since July.

### Other observations on exit accretion

- **Within-month exit reversals exist and are preserved.** Three enrollments appeared on an Exit Log and then disappeared from later snapshots in the same month (CHAMP 6507 in April, 74077 in May, 234798 in August). The history correctly keeps the earlier rows (it is append-per-snapshot), but any consumer counting unique exits per month will over-count by these. The master handles it correctly only if the client reappears in the latest participant snapshot (Active wins).
- **Each exit appears once per snapshot it was on** (e.g. 179512 appears 3 times in October). That is the documented grain, but "unique exits per month" has to be computed downstream as `distinct(CHAMP ID, ICMS Enrollment ID)`. October: 16 rows → 11 unique exits.
- **Structural gap, not an app bug:** the Exit Log tab lists only exits whose check-out falls in the *current* service month (verified: in all pre-July rows, check-out month == reporting month, 0 exceptions, and no enrollment ever appears under two reporting months). Exits that occur after the last ingested snapshot of a month, or are entered into CHAMP late for a prior month, will never appear in any ingested Exit Log. The master's `client_status = Unknown` (absent from latest snapshot, no exit record) is the existing detector for this; the stale sample master shows 32 such rows. Recommend ingesting the last snapshot generated each month (the 30th/31st) and periodically reviewing Unknowns.

---

## 4. New or renamed columns the flatfile is not surfacing

### 4a. Participant Level Summary — header drift between the 07.07 and 10.19 snapshots

**Removed (app fields now permanently blank):**

| Catalogue field (app) | Class | What happened | Change-log reference |
|---|---|---|---|
| `Receiving Pre-Match ICMS?` | Required | Retired — pre-match pathway discontinued | JUL 14 2026 |
| `Key Housing Milestone Data Needed?` | Required | Retired | — |
| `Service Planning Area` | Required | Renamed → `Service Planning Area (SPA)` | — |
| `Permanent Housing Subsidy` | Required | Renamed → `Current Permanent Housing Subsidy` → `Current Rental Subsidy` | AUG 27 / AUG 31 2026 |
| `Notification of Privacy Practices On Record?` | Optional | Renamed → `Notification of Privacy Practices Status on Record?` | — |
| `CalAIM CS ... Verbal Opt-In Status Missing?` | Optional | Renamed → `CalAIM CS ... Verbal Opt-In Status Needed?` | — |
| `Family Billing Rate But Dependent Household Member(s) Not Listed in Household Composition?` | Optional | Renamed → `Family Billing Rate But Eligible Dependent Household Member(s) Not Recorded in Household Composition?` | JUN 12 2026 |
| `TLS?` | not in catalogue | Removed; superseded by Current/Target Rental Subsidy + Fiscal Agent | — |
| `Main Income Source (Amount)` | not in catalogue | Replaced by two income indicators | AUG 4 2026 |
| `Geographic Preference(s) for Housing` | not in catalogue | Removed | — |

**Added in 10.19 (none surfaced):**

| New header | Change-log reference | Notes |
|---|---|---|
| `Current Rental Subsidy` | AUG 31 2026 | Successor to required field `Permanent Housing Subsidy`; 1,354/1,354 populated |
| `Target Rental Subsidy` | AUG 27 2026 | 5/1,354 populated |
| `TLS / Shallow Subsidy Month` | SEP 23 2026 | 195/1,354 populated; integer months |
| `TLS / Shallow Subsidy Fiscal Agent` | JUL 15 2026 | 205 populated (all `HOM`) |
| `Public Housing Authority (PHA) Sharing Addendum Status` | JUL 23 2026 | 607 populated |
| `Total Monthly Household Income (Primary Income Source)` | AUG 4 2026 | 655 populated; text like `$221 (GR)` |
| `Household Income as Percentage of Area Median Income (AMI)` | AUG 4 2026 | 654 populated; decimal fraction |
| `Medi-Cal Managed Care Plan (MCP)` | SEP 22 2026 | renamed from `Medi-Cal Managed Care Plan`; new non-Medi-Cal values possible |
| `Service Planning Area (SPA)` | — | successor to required field |
| `Recommended Care Plan Focus(es) Based on Recent Assessment(s)` | — | new |
| `Current Care Plan Focus(es) According to CHAMP Record` | — | new |
| `In-Person Encounters Recorded For This Month` | — | case change from `...for This Month` |
| `Notification of Privacy Practices Status on Record?` | — | successor to optional field |
| `CalAIM CS ... Verbal Opt-In Status Needed?` | — | successor to optional field |
| `Family Billing Rate But Eligible Dependent ... Not Recorded ...?` | JUN 12 2026 | successor to optional field |

**Present since before the catalogue was written but never surfaced** (so not drift, but relevant to "should we add them"):
`Matched Via DMH?`, `Citizenship Status`, `U.S. Veteran?`, `Primary Language`, `Gender Identity`, `Race and Ethnicity`, `Address`, `Dependents`, `Minors`, `Medi-Cal Status`, `Housing Subsidy Application Date`, `Permanent Housing Move-Out Date`, `Mobility/Accessibility Needs`, `Disabilities`, `Living with HIV/AIDS?`, `Prior Criminal Justice System Involvement?`, five `5x5 ... Score` columns, three `HAI ... Subtotal` columns, `Emergency Room or Inpatient Visit Within Past 2 Weeks?`, `Emergency Room Visits Within Past 6 Months`, `Inpatient Visits Within Past 6 Months`, `Nominated for Potential ICMS Graduation?` (MAY 14 2026), `Unsuccessful Outreach Attempts Recorded For This Month`, `Property Manager`, `Interim Housing Provider`, `Brilliant Corners Housing Coordinator`, `DMH Mental Health Services Provider`, `Caregiver`.

Also: `Household Size` is still mapped, but per the JUN 12 2026 change-log entry its meaning changed (blank unless the Household Composition profile was updated within the past year and not on the check-in date). Any dashboard logic treating blank as "unknown" vs "single" should be re-checked.

### 4b. Exit Log tab

| App keep-column | 10.19 header | Status |
|---|---|---|
| `Check-In Date` | `ICMS Check-In Date` | **not captured since 07.07** |
| `Check-Out Date` | `ICMS Check-Out Date` | **not captured since 07.07** |
| `Move-In Date` | `Permanent Housing Move-In Date` | **not captured since 07.07** |
| `Move-Out Date` | `Permanent Housing Move-Out Date` | **not captured since 07.07** |
| `Medi-Cal CIN` | (never on this tab) | always blank |
| — | `Current Rental Subsidy`, `Target Rental Subsidy`, `TLS / Shallow Subsidy Month`, `TLS / Shallow Subsidy Fiscal Agent` | new (AUG–SEP 2026), not captured |
| — | `ICMS Slot #`, `ICMS Slot ID`, `ICMS Project ID`, `Property Manager`, demographics, `Household Size`/`Dependents`/`Minors` | present since Feb, not captured |

### 4c. Vacant Slots tab

| App keep-column | 10.19 header | Status |
|---|---|---|
| `# Days Vacant` | `Days Vacant` | **column dropped from flatfile in 10.19** |
| — | `ICMS Slot #`, `ICMS Slot ID`, `ICMS Slot Activation Date`, `ICMS Slot End Date`, `Slot Vacancy Status` (`Never Filled`/turnover), `Default ICMS Billing Rate`, `CHAMP ID That Last Occupied Slot`, `Slot to Be Filled in Partnership With DMH?` | APR 8 2026 feature set; not captured |

The 07.07 vacant tab also used `Slot #`, `Slot ID`, `Slot Activation Date`, `Slot End Date`; 10.19 prefixes these with `ICMS `. Expect further churn here.

### 4d. Service(s) Log tab

| App `SERVICES_LOG_FIELDS` | 07.07 header | 10.19 header | Status |
|---|---|---|---|
| sheet name `Services Log` | `Services Log` | **`Service Log`** | **sheet skipped in 10.19** |
| `DHS Master Agreement #` | `HSH Master Agreement #` | (removed) | blank since 07.07 |
| `ICMS Work Order #` | same | (removed) | blank in 10.19 |
| `ICMS Provider Name` | `ICMS Provider` | `ICMS Provider` | blank since 07.07 |
| `Service` | `ICMS Service` | `ICMS Service` | **blank since 07.07** — the service type itself |
| `CHAMP User Who Created Record` | `...Created Service Record` | same | blank since 07.07 |
| `Organization of CHAMP User Who Created Record` | `...Created Service Record` | same | blank since 07.07 |
| `Service ID`, `CHAMP ID`, `ICMS Enrollment ID`, `Place of Service`, `Service Date`, `Service Record Created Date`, `PH Building Name`, `ICMS Case Manager` | same | same | OK |

Dedupe still works because `Service ID` is populated, but the fallback composite key (which includes `Service`) has been weaker since July. Also note the JUN 22 2026 change-log item: services whose `Service Date` is after `Service Record Created Date` are now flagged by HSH as data-entry errors; both columns are captured so this check could be reproduced downstream.

### 4e. Tabs the app does not read at all

`UHA Log` (MAR 27 2026), `Hospitalization Log` (MAR 3 2025), `Action Step Log` (APR 29 2024), `Training Log` (MAR 17 2026; renamed from `Trainings Log` in 10.19), `Billing Rate Appeal Log` (JUL 15 2026), and the bundled summaries (`Provider`/`Building`/`Scattered Site`/`Case Manager Level Summary`). Whether to ingest any of these is a business decision; the Provider Level Summary now includes `Total ICMS Exits` (OCT 1 2026), which would be a cheap reconciliation check against the exit-log count.

The `Generated On` tab carries the snapshot generation timestamp (`10/19/2026 9:01 am`). The app currently derives snapshot date from the filename; this tab could be used to validate or auto-fill it.

---

## 5. Other things that need to change or be updated

### Code (in priority order)

1. **Services sheet name.** Accept a list of candidate names (`Service Log`, `Services Log`) in `icms_gui._process_one_file` rather than a single exact name, or at minimum surface a loud warning in the Results panel when a configured sheet is missing. Same treatment for vacant/exit (`Vacant Slots`, `Exit Log` are unchanged today but are equally brittle).
2. **Header synonyms for exit, vacant, and services keep-columns.** Replace the exact-match keep lists (duplicated in `icms_gui.py` ~L1757–1800 and `icms_etl_core.read_vacant_slots`/`read_exit_log`/`SERVICES_LOG_FIELDS`) with canonical → `[accepted header names]` mappings and rename to the canonical name on read. Minimum synonym set:
   - Exit: `Check-In Date` ← `ICMS Check-In Date`; `Check-Out Date` ← `ICMS Check-Out Date`; `Move-In Date` ← `Permanent Housing Move-In Date`; `Move-Out Date` ← `Permanent Housing Move-Out Date`
   - Vacant: `# Days Vacant` ← `Days Vacant`
   - Services: `Service` ← `ICMS Service`; `ICMS Provider Name` ← `ICMS Provider`; `CHAMP User Who Created Record` ← `CHAMP User Who Created Service Record`; `Organization of CHAMP User Who Created Record` ← `Organization of CHAMP User Who Created Service Record`; `DHS Master Agreement #` ← `HSH Master Agreement #`
   Keeping the canonical output names unchanged preserves CSV/Tableau compatibility and lets the master's `_event` ordering work again.
3. **Participant catalogue update** (`REQUIRED_KEY_FIELDS`, `REQUIRED_QUEUE_FIELDS`, `OPTIONAL_FIELDS` in `icms_etl_core.py`):
   - Add built-in aliases (or seed `field_aliases.json`) so `Service Planning Area` ← `Service Planning Area (SPA)`, `Permanent Housing Subsidy` ← `Current Rental Subsidy`, and the three renamed optional fields resolve without prompting.
   - Demote or remove `Receiving Pre-Match ICMS?` and `Key Housing Milestone Data Needed?` from required (they no longer exist); keep them as optional so historical CSV columns stay stable.
   - Decide which of the new indicators to add as optional fields. Strong candidates given the change-log emphasis: `Target Rental Subsidy`, `TLS / Shallow Subsidy Month`, `TLS / Shallow Subsidy Fiscal Agent`, `Public Housing Authority (PHA) Sharing Addendum Status`, `Total Monthly Household Income (Primary Income Source)`, `Household Income as Percentage of Area Median Income (AMI)`, `Medi-Cal Status`, `Medi-Cal Managed Care Plan (MCP)`, `Nominated for Potential ICMS Graduation?`, `Dependents`, `Minors`, `Permanent Housing Move-Out Date`, `In-Person Encounters Recorded For This Month`, `Unsuccessful Outreach Attempts Recorded For This Month`.
   - Adding fields is schema-additive for the CSV histories and the flatfile (new columns appended); Tableau extracts need a refresh but nothing breaks.
4. **Drift detection.** Add an ingest-time check that compares this run's resolved/unresolved field set and sheet list to the previous successful ingest-log row and raises a visible warning (Results panel + `data_anomalies-diag`) when a previously-populated field becomes unresolved or a previously-present sheet disappears. This would have caught items #1–#6 in §1 at the 07.07 and 10.19 runs.
5. **Remember required-field skips** per header fingerprint, or let the alias dialog offer "mark as retired", so operators aren't trained to click Skip four times per run.
6. **Exit uniqueness helper.** Consider adding `icms_exit_log_unique.csv` (or a column in the master) keyed on `CHAMP ID + ICMS Enrollment ID` with first-seen/last-seen snapshot and a `reversed` flag when the enrollment vanishes from a later same-month Exit Log while the client reappears in the participant sheet. This resolves the over-counting noted in §3 without changing the existing history grain.
7. **Snapshot date from `Generated On`.** Low effort; removes a class of manual date errors (the history already has one date with two filename variants).
8. **Row ID.** Every input sheet and both outputs carry a `Row ID` UUID column as the first column. The app code does not produce it (`reduce_to_approved_fields` would drop it), so it is almost certainly an artifact of the scrambling step; confirm, and if the real outputs also contain it, trace where it is added.

### Documentation

- `architecture planning.md` — Approved field catalogue: mark `Receiving Pre-Match ICMS?` and `Key Housing Milestone Data Needed?` as retired-by-vendor; add header-synonym tables for exit/vacant/services; note the `Service Log` tab rename; note that the Exit Log tab is a current-service-month window.
- `MAINTAINER_GUIDE.md` §4/§11 — add "vendor header renames in exit/vacant/services are **not** alias-able and require a code change" and record the 07.07 and 10.19 drift events as worked examples.
- Keep a dated copy of the HSH change log in `docs/` with each release so future reviews can diff against it.

### Sample outputs in the repo

`icms_ingest_log.csv`, `icms_snapshot_history Scrambled.csv`, and `icms_client_project_master Scrambled.xlsx` in `Deidentified Sample Output/` are from the 07-20 run, not 10-19. A matching 10-19 set (including `icms_services_log_history`) would let the next review verify master status and services dedupe directly.

---

## 6. Evidence appendix

- Participant Level Summary 10.19: 97 vendor columns (+ `Row ID`). App catalogue: 55 fields. Exact-match resolved: 48 (7 unresolved: 4 required, 3 optional). Unmapped vendor columns: 49.
- Flatfile 10-19: 1,409 × 60. All-blank participant columns: `Receiving Pre-Match ICMS?`, `Permanent Housing Subsidy`, `Service Planning Area`, `Key Housing Milestone Data Needed?`, `Notification of Privacy Practices On Record?`, `CalAIM ... Opt-In Status Missing?`, `Family Billing Rate ... Not Listed ...?`. Vacant rows: 55, `PH Building Name` 35/55 (scattered-site slots are legitimately blank), no `# Days Vacant` column.
- Exit history 10-19: 250 × 24; 36 snapshot dates; reporting-month row counts Jan 4 / Feb 16 / Mar 28 / Apr 67 / May 41 / Jun 2 / Jul 7 / Aug 39 / Sep 30 / Oct 16; unique October enrollments 11.
- Sheet lists: 02.18 had 10 tabs (no `Vacant Slots`, `UHA Log`, `Trainings Log`); 04.30–07.07 had 13 (`Services Log`, `Trainings Log`); 10.19 has 14 (`Service Log`, `Training Log`, + `Billing Rate Appeal Log`).
- Ingest log (through 07-20) confirms services were processed on every run up to and including 07-20 (e.g. 2,586 records on 07-20), so the services break is new with the 10.19 sheet rename; the column blanking began 07.07.
