# ICMS ETL update — Cursor implementation prompt
Date: 2026-10-27

Implement the changes below in the existing ICMS ETL repository. These requirements consolidate the director's decisions after reviewing the 2026-10-21 Snapshot Drift Review. This is an implementation request, not another exploratory audit. Inspect the actual code and docs, make the changes, run meaningful checks, and report the result. Resolve straightforward implementation details yourself. Ask only when a genuinely ambiguous business meaning or missing essential input prevents a sound decision.

## Goal and operating constraints

Restore imports across historical and current snapshot layouts, repair accumulated data through the existing batch workflow, add the approved fields, and make the master enrollment-specific.

The analyst runs the app; the director maintains field definitions and mappings in the repository. Normal runs must not require analysts to interpret unfamiliar columns, approve mappings, acknowledge retired fields repeatedly, or maintain local mapping files.

Prefer a small change to the existing architecture. Reuse the GUI, batch processing, Results panel, and ingest log. Do not introduce new databases, reconciliation applications, or detailed-log datasets. Keep the production workflow GUI-first.

## Read before editing

Locate these files in the repo; paths may differ slightly:

- `architecture planning.md` (a supplied copy is named `architecture planning(1).md`).
- `2026-10-21 Snapshot Drift Review.md`.
- `docs/MAINTAINER_GUIDE.md`.
- `icms_etl_core.py`, `icms_gui.py`, and relevant tests/configuration.
- July and October sample snapshots:
  - `ICMS Snapshot - 023-HFH-ICMS - 07.07.2026 - HFHPI2026 Scrambled.xlsx`
  - `ICMS Snapshot - 023-ICMS-PH - 10.19.2026 - HFHPI2026 Scrambled.xlsx`
- Sample exports:
  - `icms_current_dashboard_flatfile Scrambled 10-19-26.xlsx`
  - `icms_exit_log_history Scrambled 10-19-26.csv`

The director also reviewed https://homeless.lacounty.gov/icms-snapshot-updates/ . Consult it for ambiguous meanings rather than guessing from names. Distinguish findings independently verified from the provided files from assertions requiring other repository samples or production evidence.

The samples are partially scrambled. CHAMP and HMIS IDs remain traceable. Some demographic/date values are intentionally invalid, and `Row ID` appears to have been added during scrambling. Do not redesign validation around those artifacts or make Row ID a business key.

## 1. Centralize source-name resolution and preserve export compatibility

Maintain a version-controlled catalogue of output fields and accepted historical source names, shared by the core and GUI. Resolve names independently for every workbook and sheet category. Reading a new snapshot must never overwrite mappings needed for older snapshots in the same batch.

Retain existing output column names unless this prompt expressly changes them. New approved fields use their current HSH names. Preserve established column presence and sensible ordering even when a source field is absent. In particular, restore `# Days Vacant` to the flatfile's stable schema.

Normalize harmless formatting differences, including surrounding/repeated whitespace and case, without fuzzy semantic matching. Do not merge two measures merely because their labels sound alike.

Use mappings at the appropriate sheet/output scope: for example, participant `Permanent Housing Move-Out Date` and existing exit-history `Move-Out Date` can have distinct output names.

Required known mappings include:

| Sheet category | Existing output field | Additional accepted source header(s) |
| --- | --- | --- |
| Participant | Service Planning Area | Service Planning Area (SPA) |
| Participant | Permanent Housing Subsidy | Current Permanent Housing Subsidy; Current Rental Subsidy |
| Participant | Notification of Privacy Practices On Record? | Notification of Privacy Practices Status on Record? |
| Participant | CalAIM CS Housing Navigation (HN) and Tenancy Sustaining Services (TSS) Verbal Opt-In Status Missing? | CalAIM CS Housing Navigation (HN) and Tenancy Sustaining Services (TSS) Verbal Opt-In Status Needed? |
| Participant | Family Billing Rate But Dependent Household Member(s) Not Listed in Household Composition? | Family Billing Rate But Eligible Dependent Household Member(s) Not Recorded in Household Composition? |
| Exit | Check-In Date | ICMS Check-In Date |
| Exit | Check-Out Date | ICMS Check-Out Date |
| Exit | Move-In Date | Permanent Housing Move-In Date |
| Exit | Move-Out Date | Permanent Housing Move-Out Date |
| Vacant | # Days Vacant | Days Vacant |
| Services | Service | ICMS Service |
| Services | ICMS Provider Name | ICMS Provider |
| Services | CHAMP User Who Created Record | CHAMP User Who Created Service Record |
| Services | Organization of CHAMP User Who Created Record | Organization of CHAMP User Who Created Service Record |
| Services | DHS Master Agreement # | HSH Master Agreement # |

Retain the original headers as accepted inputs too. For the newly approved `Medi-Cal Managed Care Plan (MCP)`, also accept historical `Medi-Cal Managed Care Plan`. Accept the historical capitalization of `In-Person Encounters Recorded for This Month`.

Accept both `Service Log` and `Services Log` as worksheet names automatically, regardless of a stale saved setting. Reuse the same sheet-resolution approach for existing supported sheet categories; do not require an analyst to discover a vendor rename. Avoid silently selecting an arbitrary sheet if multiple candidates are materially different.

If two recognized source headers coexist, combine agreeing/complementary values safely. Conflicting populated values must not be silently discarded. Use a documented deterministic policy and record a concise anomaly; nonessential field conflicts should not trigger mapping dialogs. Identity ambiguity must not result in merged or misattributed records.

## 2. Remove routine analyst mapping decisions

Known retired participant fields remain in the output schema and accept values from older snapshots:

- `Receiving Pre-Match ICMS?`
- `Key Housing Milestone Data Needed?`

When unavailable in a newer source, leave them blank without repeated prompts or recurring warnings. Missing legacy agreement/work-order service fields are likewise not fatal when unavailable by design.

Unexpectedly missing nonessential reporting fields: preserve the column, use nulls for that snapshot, complete processing, and report the issue once in the Results summary/log. Never carry an older snapshot's value into a new snapshot merely to conceal missing source information.

A newly discovered, unapproved column does NOT automatically enter the exported data schema. Report it for director review. Preserve the original input workbook unchanged; it is the source for later batch reprocessing.

Missing information essential to identify/process records reliably may fail the affected file. Identify the actual missing requirement. Do not wipe or partially replace that file's existing history/output contributions. Continue other valid files in a batch, and never publish a failed file as a successful current snapshot.

Repository mappings are authoritative. Inspect the existing `field_aliases.json` behavior: preserve useful approved mappings, but remove the routine workflow in which an analyst's local dialog choices become an undocumented competing schema. Any supported override must have explicit, documented precedence and must not silently supersede the released catalogue.

## 3. Add the approved optional fields

Add these fields to the current dashboard flatfile and participant snapshot history. Carry them through the master where participant fields are already carried through. Add corresponding approved fields to exit history when its source tab supplies them; do not copy current participant values into historical exit rows or create unrelated service fields.

1. Target Rental Subsidy
2. TLS / Shallow Subsidy Month
3. TLS / Shallow Subsidy Fiscal Agent
4. Public Housing Authority (PHA) Sharing Addendum Status
5. Total Monthly Household Income (Primary Income Source)
6. Household Income as Percentage of Area Median Income (AMI)
7. Medi-Cal Status
8. Medi-Cal Managed Care Plan (MCP)
9. Nominated for Potential ICMS Graduation?
10. Dependents
11. Minors
12. Permanent Housing Move-Out Date
13. In-Person Encounters Recorded For This Month
14. Unsuccessful Outreach Attempts Recorded For This Month

Preserve existing current-subsidy information through its approved alias mapping. Do not introduce a competing participant `Current Rental Subsidy` output column alongside `Permanent Housing Subsidy`.

Older snapshots lacking these optional fields produce nulls and do not require acknowledgment. Preserve source meaning and types through the existing storage/export conventions. For example, AMI is a fraction; household income is supplied as text such as an amount plus primary source. Do not invent parsing or calculation rules. A blank household size must not be interpreted automatically as one person.

No dashboard calculations, task rules, or new displays are authorized by the addition of these fields.

## 4. Change the master to client + project + enrollment

Change `icms_client_project_master.xlsx` from one row per CHAMP ID + project to:

**One row per CHAMP ID + ICMS Project Name + ICMS Enrollment ID.**

Keep the filename. Multiple enrollments in the same project must remain separate rows with their own dates, values, and status. Normalize identifiers as identifiers, not floating-point measures.

Apply this key to master construction, participant-value selection, exit matching, latest-snapshot presence, grouping, and relevant duplicate diagnostics. Do not change snapshot-history grain to one lifetime row per enrollment: snapshot history must still retain separate observations by snapshot date.

Rebuild the master deterministically from the accumulated histories. Include observed enrollments supported by participant and/or exit history rather than losing an exit-only enrollment simply because a participant observation was never imported.

For each specific enrollment:

- Active: that enrollment appears in the latest successfully imported participant snapshot.
- Exited: it does not appear there and has a matching recorded exit.
- Unknown: previously observed but not in the latest participant snapshot, with no matching recorded exit.

A new active enrollment must not overwrite an older enrollment's Exited/Unknown status, dates, or other historical values. Choose participant values from that enrollment's latest observation and exit details from its appropriate latest recorded exit/correction. A service event alone does not establish current enrollment or an exit.

Check completeness of enrollment IDs in the full repository histories. Do not silently fall back to the old client/project grouping when IDs are missing. If identifiers can be recovered unambiguously from archived sources, use them. If not, report the actual limitation and preserve ambiguous data without falsely merging enrollments; raise a targeted question only if necessary.

Do not add `icms_exit_log_unique.csv` or another deduplicated exit export. The revised master and existing histories are the agreed output package.

## 5. Make existing batch reprocessing repair data

Historical repair uses the existing multi-file batch function. Do not create a separate repair application or require manual CSV editing.

Reprocessing already-imported snapshots must update corrected values, preserve unrelated accumulated history, and avoid duplicates:

- Participant history retains its documented same-date replacement behavior.
- Exit history retains replacement by source filename + snapshot date.
- Services remain deduplicated by Service ID, with the existing composite fallback when necessary.
- Rebuild the master from the full resulting histories.
- Respect the dashboard-publish option. In a batch, publish only the newest successful intended snapshot, not intermediate historical snapshots.

Existing duplicate-import detection must allow an intentional rerun of all selected files through a single batch choice. Do not silently exclude already-imported files when the purpose is repairing them.

Services require particular care. The documented earliest-snapshot/source representative rule must not freeze damaged rows permanently. When that representative source is reprocessed, its corrected values must replace its prior imported values. Resolve aliases BEFORE calculating fallback keys. Verify that repairing previously blank service names does not create duplicate events under old versus repaired fallback keys. Preserve deterministic behavior when files are selected in a different order.

Use the simplest reliable approach compatible with the existing stores. If a particular damaged record cannot be reconstructed without a missing source workbook, report that limitation instead of claiming it was repaired or inventing values.

Create an automatic backup of existing output files before the one-time rebuild/reprocessing that changes existing histories, using a simple documented approach. Do not turn this into an analyst-managed backup routine.

## 6. Summarize unexpected structural changes without interrupting runs

Use the existing Results panel and ingest log for a concise summary of unfamiliar headers, unexpectedly missing fields, and supported sheets that could not be imported. Make maintenance details easy to copy for the director.

For a batch, group equivalent findings and identify affected files; do not display the same modal for every file/field.

Distinguish intentional exclusions and recognized historical/retired fields from unfamiliar changes. Seed the known intentionally excluded fields/sheets from reviewed samples, so the first release does not repeatedly announce dozens of already known exclusions. Do not infer semantic changes from column names or treat legitimately blank source values as missing headers.

Known mappings work without warnings. A newly absent expected field or a new unrecognized header is maintenance information, not a demand for analyst approval. An output missing services must not be presented as fully complete without noting the skipped category.

No extra recurring review checklist, email automation, or mandatory reconciliation process is requested.

## 7. Default snapshot dates from the workbook

For each workbook independently, default snapshot date to the usable date in `Generated On`. Fall back to filename parsing if unavailable/unreadable.

Preserve the existing manual date override. If workbook and filename dates disagree, display both in the existing single-file/batch date-confirmation interface. Do not add a separate dialog sequence for each file. Record the selected date and its origin in the existing log where practical. Do not silently change dates of records already stored.

Examples in supplied workbooks:

- July: `07/07/2026 8:13 am`
- October: `10/19/2026 9:01 am`

Locate the actual metadata value rather than relying on fixed cell coordinates; scrambled samples contain an extra Row ID column.

## 8. Scope and operating guidance

Continue importing only Participant Level Summary, Vacant Slots, Exit Log, and Service/Services Log as data sources.

Do not expand ingestion to UHA, Hospitalization, Action Step, Training, Billing Rate Appeal, or the aggregate summary datasets. Reading `Generated On` for date metadata is in scope. Do not add unrelated vacant-slot fields or participant fields beyond the approved additions simply because they exist.

Document this operating cadence: import weekly and additionally import a snapshot generated on the last day of each month. Snapshot Exit Logs appear to cover the current service month; accumulated app exit history retains observed exits across imports. A month-end snapshot only includes records entered before generation, so late/backdated exits can still be absent. Retain Unknown without requiring intervention or fabricating exit dates.

Do not label an exit as definitively reversed solely because it disappears from a later Exit Log. A reappearing active enrollment can take Active status under the rules above; absence alone is not proof of a correction.

## 9. Acceptance checks

Use isolated temporary output directories and avoid modifying the deidentified reference fixtures. Add focused regression tests for the changes; do not build a broad test framework just for this release.

At minimum verify:

- July and October files process together without mapping dialogs for known differences. Selection order and repeated import produce the same final histories/master.
- October flatfile has 1,354 participant rows plus 55 vacant rows, total 1,409, if the fixtures remain unchanged. Existing columns remain; approved additions appear. Renamed participant fields populate from the correct source.
- October Exit Log has 11 records and every one is retained with check-in/check-out/move-in/move-out dates matching the source where present. Null source dates remain null.
- The supplied old export has 250 rows, including 92 rows from July 7 onward whose four date fields are blank. Replaying available July/October sources repairs their corresponding records without removing unrelated dates. Do not claim all 92 were repaired unless their source snapshots were actually available and replayed.
- Service/Services Log import works in both vintages, including service type, creator, and provider aliases. A previously damaged representative service row is repaired; reruns do not inflate event counts.
- `# Days Vacant` is present and populated from `Days Vacant` for October vacant records.
- Retired and unavailable optional fields do not block old or new snapshots. A deliberately unfamiliar header is reported and excluded from exports; no local analyst mapping is requested.
- New approved fields match source values and are null where not supplied.
- The master separates multiple enrollments in the same client/project, matches exits to the correct enrollment, and preserves old rows when a new enrollment becomes active.
- At least one exit-only enrollment, a genuinely missing essential identity field, a date-source conflict, and a known field absent by design are handled as specified.

Useful independently checked examples:

- CHAMP 9038: July enrollment 20415; October enrollment 529134 in the same project, checked in August 28. No exit-history row for the old enrollment in the supplied CSV. The old enrollment must not disappear into the new Active row.
- CHAMP 51546: July enrollment 106708; October enrollment 531738 in the same project, checked in September 16. Same separation requirement.
- Comparing July participants to October participants: 41 July enrollments are absent in October; 31 have exit-history matches and 10 do not. Eight of those 10 clients are absent entirely from the October participant list; two have the new enrollments above. This demonstrates unaccounted-for enrollments, NOT proof that a month-end omission caused them.
- All 250 supplied exit-history rows have enrollment IDs. That does not establish completeness of every other stored historical file.

The full participant-history/master/services-history production outputs were not attached to this conversation. Verify against repository versions where available and clearly identify what remains unverified.

## 10. Finish and report

Update architecture and maintainer documentation to match the implemented behavior, particularly the master key, shared aliases, director-owned mapping decisions, missing-field behavior, date precedence, batch correction behavior, and weekly/month-end workflow. Remove obsolete guidance that exit/service/vacancy renames require duplicated hardcoded edits once the shared mechanism exists.

Keep documentation concise. Include a short director maintenance instruction: review the structural-change details and original snapshot, update approved mappings/fields in the repo using Cursor, test historical/current examples, release the updated app, and reprocess affected archives.

Bump the app version when preparing the updated release. Run the existing build procedure if this environment supports it; otherwise state precisely what was validated and provide the existing Windows build command without claiming an executable was built.

Return a concise implementation summary, checks/results, any real unresolved limitations, and the analyst's exact steps for using the updated release and rerunning archived snapshots. Do not ask for approval on routine mapping changes already specified here.

