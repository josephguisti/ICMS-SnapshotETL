# Change Log

What changed in each version of the ICMS Snapshot ETL app, newest first. The version number appears in the app's window title and in the program's file name (for example, `ICMS Snapshot ETL v1.1.3.exe`).

When a new version is released, add an entry at the top.

---

## Version 1.1.4 — November 24, 2026

- **The services history now includes a client name.** The Service Log tab does not include first or last names. The app fills `Client Name` from the same snapshot's participant row for that CHAMP ID and enrollment ID, or from the Exit Log when the participant row has no name. Names are written as `First Last`. If either part is missing, the blank part is marked as `[BLANK FIRST]` or `[BLANK LAST]`. If the enrollment can't be found, or the participant list has conflicting names, the field is left blank rather than guessed.
- Services already saved by earlier versions keep a blank client name until a later snapshot includes the same service, or the older snapshot files are reprocessed.
- The first run with this version saves a backup copy of the existing output files before writing anything.

## Version 1.1.3 — November 16, 2026

- **The services history now shows which project each service belongs to.** The vendor's Service Log tab does not include a project name. The app now looks up each service's client and enrollment in the same snapshot's participant list, or its Exit Log, and fills in the project from there. If the enrollment can't be found, or it is listed under more than one project, the project is left blank rather than guessed.
- Services already saved by earlier versions keep a blank project until a later snapshot includes the same service, or the older snapshot files are reprocessed.
- The first run with this version saves a backup copy of the existing output files before writing anything.

## Version 1.1.2 — October 28, 2026

- **The client/project master now includes Gender Identity and Race and Ethnicity.** Each enrollment shows the values from its most recent participant snapshot.
- The first run with this version saves a backup copy of the existing output files before the new columns appear.

## Version 1.1.1 — October 28, 2026

- **The app now checks that it can write its output files before it starts.** Previously, if an output file (such as the dashboard flatfile) was open in Excel, a long batch would run to the end and then fail. Now the app warns you right away so you can close the file and try again.
- The program file is about half its previous size (about 37 MB instead of 67 MB), so it downloads and opens faster.

## Version 1.1.0 — October 27, 2026

A major update to keep up with changes the vendor made to the snapshot workbook.

- **No more field-mapping questions.** The app now recognizes the vendor's old and new column names on its own. Analysts are no longer asked to match up fields during a run.
- **Renamed tabs and columns are handled.** The Service Log tab (previously "Services Log") and renamed columns on the Exit Log, Service Log, and Vacant Slots tabs are read correctly again. Before this, some exit dates and service details had been saved blank since early July 2026.
- **Unexpected changes are reported, not asked about.** When the vendor adds, renames, or drops a column or tab, the Results screen lists it with a "Copy findings for director" button. The run is not interrupted.
- **One row per enrollment in the master.** The client/project master now has one row per client, project, and enrollment. A client who leaves a project and later re-enrolls has two rows instead of one overwriting the other.
- **All exits are kept.** Enrollments that only ever appeared on an Exit Log are now included in the master.
- **Fixing old data is built in.** Re-running snapshots that were already imported replaces their saved rows with freshly read values. This is how past blanks are repaired. Output files are backed up automatically first.
- **Snapshot dates fill in automatically** from the workbook's "Generated On" tab. If that disagrees with the date in the file name, both are shown.
- **Each file is saved all-or-nothing.** If something goes wrong partway through a file, every output is restored to how it was before that file.
- **Batch runs respect the "publish dashboard" setting.** Only the newest file in a batch becomes the current dashboard.
- **14 approved fields were added**, including Target Rental Subsidy, TLS / Shallow Subsidy Month and Fiscal Agent, household income and AMI, Medi-Cal status and plan, graduation nomination, dependents and minors, Permanent Housing Move-Out Date, and this month's in-person encounters and unsuccessful outreach attempts.
- **"# Days Vacant" is back** on the vacant-slot rows after the vendor renamed it.
- Excel date columns are saved as real dates, so date filters work in Excel. *This change was made in September 2026 but first reached users in this version.*

## Version 1.0.1 — July 21, 2026

- **The app remembers your folders.** File pickers now open in the folder you last used for input workbooks and for outputs.

## Version 1.0.0 — July 21, 2026

- **Versions are now numbered.** The version shows in the window title and the program file name, so it's easy to tell which copy you're running.

## Before version numbers — May to July 2026

- **July 9:** Large batches run much faster. The master file is now rebuilt once at the end of a batch instead of after every file.
- **May 21:** Services with no building listed are recorded as "Scattered Site." Error messages are clearer when a workbook can't be opened, and the Results screen shows why a file in a batch failed.
- **May 20:** The app began keeping a cumulative history of services from the Services Log tab.
- **May 18:** The desktop app (with windows and buttons) was added, including importing several workbooks at once and safeguards against importing the same file twice.
- **May 12:** First release. A command-line tool that turned snapshot workbooks into the current dashboard file, cumulative histories, and the client/project master.
