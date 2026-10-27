# ICMS Snapshot ETL

A desktop Python app that turns DHS/HSH ICMS snapshot workbooks into stable files for reporting and cumulative history. The day-to-day path is the Tkinter GUI.

It reads the participant, vacant, exit, and service tabs of one or more `.xlsx` files (password-protected files are decrypted in memory) through a fixed field catalogue. It then writes a current dashboard flatfile, cumulative CSV histories, and an enrollment-level client/project master.

## Run

From this folder, with the packages in `requirements.txt` installed:

```bash
python icms_gui.py
```

Pick one workbook for the weekly update, or several for a backfill or repair. Snapshot dates fill in from each workbook's `Generated On` tab. Choose an output folder and run. Import weekly, plus the snapshot generated on the last day of each month.

Tests:

```bash
python -m pytest tests -p no:cacheprovider
```

To build the Windows executable, bump `__version__` in `version.py`, then:

```bash
python build_exe.py
```

The executable is written to `dist/ICMS Snapshot ETL v<version>.exe`.

## Which document to trust

- `architecture planning.md` is the behavior contract: the workflow, inputs, outputs, master status rules, and structural findings.
- `icms_catalogue.py` is the only definition of output field names, accepted vendor header names, and supported tabs.
- [docs/MAINTAINER_GUIDE.md](docs/MAINTAINER_GUIDE.md) maps that contract to the code. It also has the director's instruction for handling structural findings (§1) and the known limitations (§9).

Historical records:

- [docs/2026-10-21 Snapshot Drift Review.md](docs/2026-10-21%20Snapshot%20Drift%20Review.md) records the vendor drift that the 2026-10-27 update (v1.1.0) addressed.
- `CONFIG_EXTRACTION_PLAN.md` and `ODR_ETL_PORT_PLAN.md` are plans, not descriptions of the running app.
