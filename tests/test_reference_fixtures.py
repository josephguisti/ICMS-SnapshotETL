"""Acceptance checks against the deidentified July and October fixtures.

Fixtures are copied into a temporary folder under their production file
names (the stored deidentified exit history refers to those names); the
originals are never modified.
"""

import shutil
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

import icms_catalogue as catalogue
import icms_etl_core as core
from conftest import SAMPLE_DIR, SAMPLE_OUTPUT_DIR, quiet

JULY_SAMPLE = "ICMS Snapshot - 023-HFH-ICMS - 07.07.2026 - HFHPI2026 Scrambled.xlsx"
OCT_SAMPLE = "ICMS Snapshot - 023-ICMS-PH - 10.19.2026 - HFHPI2026 Scrambled.xlsx"
JULY = "ICMS Snapshot - 023-HFH-ICMS - 07.07.2026 - HFHPI2026.xlsx"
OCT = "ICMS Snapshot - 023-ICMS-PH - 10.19.2026 - HFHPI2026.xlsx"
OLD_EXIT_HISTORY = SAMPLE_OUTPUT_DIR / "icms_exit_log_history Scrambled 10-19-26.csv"
EXIT_DATES = ["Check-In Date", "Check-Out Date", "Move-In Date", "Move-Out Date"]
# The 2026-10-27 additions. Gender and race were approved later and are
# present on both fixture vintages, so they are not part of this set.
NEW_FIELDS = [
    f for f in core.ALL_APPROVED_FIELDS
    if f not in ("Gender Identity", "Race and Ethnicity")
][-14:]

pytestmark = pytest.mark.skipif(
    not (SAMPLE_DIR / OCT_SAMPLE).exists() or not OLD_EXIT_HISTORY.exists(),
    reason="deidentified fixtures not present",
)


def _raw(path: Path, sheet: str):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()
    return pd.DataFrame(rows[1:], columns=[str(h).strip() if h else "" for h in rows[0]])


def _date(v):
    p = pd.to_datetime(v, errors="coerce", format="mixed")
    return None if pd.isna(p) else p.date()


@pytest.fixture(scope="module")
def run_dir(tmp_path_factory):
    base = tmp_path_factory.mktemp("fixtures")
    inputs, out = base / "in", base / "out"
    inputs.mkdir()
    out.mkdir()
    shutil.copy2(SAMPLE_DIR / JULY_SAMPLE, inputs / JULY)
    shutil.copy2(SAMPLE_DIR / OCT_SAMPLE, inputs / OCT)
    shutil.copy2(OLD_EXIT_HISTORY, out / core.EXIT_HISTORY_FILENAME)

    files = []
    for name in (OCT, JULY):  # deliberately newest first
        d = core.default_snapshot_date(inputs / name)
        files.append(core.FileSpec(inputs / name, d["date"], d["source"],
                                   d["generated_on"], d["filename_date"]))
    results = core.run_snapshot_files(files, out, log=quiet)
    return inputs, out, results


@pytest.fixture(scope="module")
def outputs(run_dir):
    _, out, _ = run_dir
    read = lambda n: pd.read_csv(out / n, dtype=str).fillna("")  # noqa: E731
    return {
        "hist": read(core.SNAPSHOT_HISTORY_FILENAME),
        "exit": read(core.EXIT_HISTORY_FILENAME),
        "svc": read(core.SERVICES_HISTORY_FILENAME),
        "master": pd.read_excel(out / core.MASTER_FILENAME, dtype=str).fillna(""),
        "flat": pd.read_excel(out / core.DASHBOARD_FILENAME, dtype=str).fillna(""),
    }


def test_every_sample_header_is_catalogued_or_known_excluded():
    for path in sorted(SAMPLE_DIR.glob("*.xlsx")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        chosen, findings = core.resolve_sheets(wb, "2026-10-19")
        assert not [f for f in findings if f["kind"] == "unfamiliar_sheet"], path.name
        for category, sheet in chosen.items():
            if not sheet:
                continue
            read = core.read_category(wb, sheet, category, "2026-10-19")
            assert not [f for f in read.findings if f["kind"] == "unfamiliar_header"], \
                (path.name, category)
        wb.close()


def test_july_and_october_import_without_mapping_prompts(run_dir):
    _, _, res = run_dir
    assert res["success"] and res["batch_success"] == 2
    assert not res["skipped_categories"]
    by_date = {r["snapshot_date"]: r for r in res["batch_results"]}
    assert set(by_date) == {"2026-07-07", "2026-10-19"}
    assert all(r["snapshot_date_source"] == "Generated On" for r in by_date.values())
    assert not [g for g in res["finding_groups"] if g["kind"] == "unfamiliar_header"]


def test_october_flatfile_counts_and_columns(outputs, run_dir):
    inputs, _, _ = run_dir
    flat = outputs["flat"]
    assert len(flat) == 1409
    assert (flat["is_vacant_slot"] == "False").sum() == 1354
    assert (flat["is_vacant_slot"] == "True").sum() == 55
    assert list(flat.columns) == core.FLATFILE_COLUMNS
    vacant = flat[flat["is_vacant_slot"] == "True"]
    raw_vacant = _raw(inputs / OCT, "Vacant Slots")
    assert vacant["# Days Vacant"].ne("").all()
    assert sorted(vacant["# Days Vacant"]) == sorted(raw_vacant["Days Vacant"].map(lambda v: str(int(v))))


def test_renamed_and_new_fields_match_october_source(outputs, run_dir):
    inputs, _, _ = run_dir
    raw = _raw(inputs / OCT, "Participant Level Summary")
    raw["key"] = raw["CHAMP ID"].map(core.normalize_identifier) + "|" + \
        raw["ICMS Enrollment ID"].map(core.normalize_identifier)
    hist = outputs["hist"][outputs["hist"]["snapshot_date"] == "2026-10-19"].copy()
    hist["key"] = hist["CHAMP ID"] + "|" + hist["ICMS Enrollment ID"]
    merged = hist.merge(raw, on="key", suffixes=("", "__raw"))
    assert len(merged) == 1354

    def same(out_col, raw_col):
        got = merged[out_col]
        want = merged[raw_col if raw_col != out_col else f"{raw_col}__raw"]
        if out_col in core.DATE_FIELDS:
            return (got.map(_date) == want.map(_date)) | (got.eq("") & want.isna())
        return got == want.map(lambda v: core._safe_str(core._norm_cell(v)))

    assert same("Permanent Housing Subsidy", "Current Rental Subsidy").all()
    for field in NEW_FIELDS:
        raw_name = next(n for n in catalogue.PARTICIPANT_FIELDS[core.ALL_APPROVED_FIELDS.index(field)]
                        .accepted if n in raw.columns)
        assert same(field, raw_name).all(), field


def test_new_fields_are_null_where_july_does_not_supply_them(outputs, run_dir):
    inputs, _, _ = run_dir
    july_headers = {catalogue.normalize_header(h) for h in _raw(inputs / JULY, "Participant Level Summary").columns}
    hist = outputs["hist"][outputs["hist"]["snapshot_date"] == "2026-07-07"]
    assert len(hist) == 1095
    for field in NEW_FIELDS:
        spec = catalogue.PARTICIPANT_FIELDS[core.ALL_APPROVED_FIELDS.index(field)]
        if not {catalogue.normalize_header(n) for n in spec.accepted} & july_headers:
            assert hist[field].eq("").all(), field


def test_october_exits_retained_with_source_dates(outputs, run_dir):
    inputs, _, _ = run_dir
    raw = _raw(inputs / OCT, "Exit Log")
    got = outputs["exit"][outputs["exit"]["source_file_name"] == OCT]
    assert len(raw) == 11 and len(got) == 11
    raw_by_enr = raw.set_index(raw["ICMS Enrollment ID"].map(core.normalize_identifier))
    for _, row in got.iterrows():
        src = raw_by_enr.loc[row["ICMS Enrollment ID"]]
        for out_col, raw_col in zip(EXIT_DATES, ["ICMS Check-In Date", "ICMS Check-Out Date",
                                                 "Permanent Housing Move-In Date",
                                                 "Permanent Housing Move-Out Date"]):
            assert _date(row[out_col]) == _date(src[raw_col]), (row["ICMS Enrollment ID"], out_col)


def test_replay_repairs_only_replayed_exit_rows(outputs):
    old = pd.read_csv(OLD_EXIT_HISTORY, dtype=str).fillna("")
    assert len(old) == 250
    blank = old[EXIT_DATES].eq("").all(axis=1)
    from_july = old["snapshot_date"] >= "2026-07-07"
    assert int((blank & from_july).sum()) == 92

    new = outputs["exit"]
    assert len(new) == 250
    replayed = new["source_file_name"].isin([JULY, OCT])
    order = ["snapshot_date", "source_file_name", "CHAMP ID", "ICMS Enrollment ID", "Row ID"]
    unrelated_old = (old[~old["source_file_name"].isin([JULY, OCT])]
                     .sort_values(order).reset_index(drop=True))
    unrelated_new = new[~replayed].sort_values(order).reset_index(drop=True)
    for col in EXIT_DATES:
        assert list(unrelated_old[col].map(_date)) == list(unrelated_new[col].map(_date)), col

    still_blank = new[EXIT_DATES].eq("").all(axis=1) & (new["snapshot_date"] >= "2026-07-07")
    repaired = int((blank & from_july).sum()) - int(still_blank.sum())
    assert repaired == int(old["source_file_name"].isin([JULY, OCT]).sum()) == 12
    assert not new.loc[replayed, "Check-In Date"].eq("").any()


def test_master_gender_and_race_follow_the_latest_participant_row(outputs, run_dir):
    inputs, _, _ = run_dir
    master = outputs["master"]
    assert {"Gender Identity", "Race and Ethnicity"} <= set(master.columns)
    withheld = {"Citizenship Status", "U.S. Veteran?", "Primary Language", "Address"}
    assert not withheld & set(master.columns)

    hist = outputs["hist"]
    seen = set(zip(hist["CHAMP ID"], hist["ICMS Project Name"], hist["ICMS Enrollment ID"]))
    sourced = master[[k in seen for k in zip(master["CHAMP ID"], master["ICMS Project Name"],
                                             master["ICMS Enrollment ID"])]].copy()
    assert len(sourced) == len(seen)

    for name, day in ((OCT, "2026-10-19"), (JULY, "2026-07-07")):
        raw = _raw(inputs / name, "Participant Level Summary")
        raw["key"] = (
            raw["CHAMP ID"].map(core.normalize_identifier) + "|"
            + raw["ICMS Project Name"].map(core._safe_str) + "|"
            + raw["ICMS Enrollment ID"].map(core.normalize_identifier)
        )
        part = sourced[sourced["snapshot_date"].str.startswith(day)].copy()
        part["key"] = part["CHAMP ID"] + "|" + part["ICMS Project Name"] + "|" + part["ICMS Enrollment ID"]
        merged = part.merge(raw, on="key", suffixes=("", "__raw"))
        assert len(merged) == len(part) > 0, day
        for col in ("Gender Identity", "Race and Ethnicity"):
            got = merged[col]
            want = merged[f"{col}__raw"].map(lambda v: core._safe_str(core._norm_cell(v)))
            assert (got == want).all(), (day, col)


def test_new_enrollments_in_same_project_stay_separate(outputs):
    m = outputs["master"]
    for champ, old_enr, new_enr in (("9038", "20415", "529134"), ("51546", "106708", "531738")):
        rows = m[m["CHAMP ID"] == champ].set_index("ICMS Enrollment ID")
        assert {old_enr, new_enr} <= set(rows.index), champ
        assert rows.loc[new_enr, "client_status"] == "Active"
        assert rows.loc[old_enr, "client_status"] in {"Unknown", "Exited"}
        assert rows.loc[old_enr, "snapshot_date"].startswith("2026-07-07")
        assert rows.loc[old_enr, "ICMS Project Name"] == rows.loc[new_enr, "ICMS Project Name"]
    assert not m.duplicated(core.KEY_FIELDS).any()


def test_july_enrollments_absent_in_october_are_exited_or_unknown(outputs):
    hist, m = outputs["hist"], outputs["master"]
    keys = lambda df: set(zip(df["CHAMP ID"], df["ICMS Project Name"], df["ICMS Enrollment ID"]))  # noqa: E731
    gone = keys(hist[hist["snapshot_date"] == "2026-07-07"]) - keys(hist[hist["snapshot_date"] == "2026-10-19"])
    assert len(gone) == 41
    status = m.set_index(core.KEY_FIELDS)["client_status"]
    counts = pd.Series([status[k] for k in gone]).value_counts().to_dict()
    assert counts == {"Exited": 31, "Unknown": 10}


def test_exit_only_enrollments_are_in_master(outputs):
    m = outputs["master"]
    participant_keys = set(zip(outputs["hist"]["CHAMP ID"], outputs["hist"]["ICMS Project Name"],
                               outputs["hist"]["ICMS Enrollment ID"]))
    exit_only = m[[k not in participant_keys for k in zip(m["CHAMP ID"], m["ICMS Project Name"],
                                                          m["ICMS Enrollment ID"])]]
    assert len(exit_only) > 0
    assert set(exit_only["client_status"]) == {"Exited"}


def test_services_both_vintages_use_aliases(outputs):
    svc = outputs["svc"]
    for name in (JULY, OCT):
        part = svc[svc["source_file_name"] == name]
        assert len(part) > 0, name
        for col in ("Service", "ICMS Provider Name", "CHAMP User Who Created Record",
                    "Organization of CHAMP User Who Created Record"):
            assert part[col].ne("").mean() > 0.95, (name, col)


def _unique_names(df: pd.DataFrame) -> dict:
    if "First Name" not in df.columns or "Last Name" not in df.columns:
        return {}
    work = pd.DataFrame({
        "champ": df["CHAMP ID"].map(core.normalize_identifier),
        "enrollment": df["ICMS Enrollment ID"].map(core.normalize_identifier),
        "name": [core.format_client_name(f, l) for f, l in zip(df["First Name"], df["Last Name"])],
    })
    work = work[work["champ"].ne("") & work["enrollment"].ne("")]
    if work.empty:
        return {}
    grouped = work.drop_duplicates().groupby(["champ", "enrollment"])["name"]
    return {k: v for k, v in grouped.first()[grouped.nunique() == 1].items()}


def _ambiguous_name_keys(df: pd.DataFrame) -> set:
    if "First Name" not in df.columns or "Last Name" not in df.columns:
        return set()
    work = pd.DataFrame({
        "champ": df["CHAMP ID"].map(core.normalize_identifier),
        "enrollment": df["ICMS Enrollment ID"].map(core.normalize_identifier),
        "name": [core.format_client_name(f, l) for f, l in zip(df["First Name"], df["Last Name"])],
    })
    work = work[work["champ"].ne("") & work["enrollment"].ne("")]
    if work.empty:
        return set()
    grouped = work.drop_duplicates().groupby(["champ", "enrollment"])["name"]
    return set(grouped.nunique()[grouped.nunique() > 1].index)


def test_services_client_name_matches_same_snapshot_enrollment(outputs, run_dir):
    inputs, _, _ = run_dir
    svc = outputs["svc"]
    assert "Client Name" in svc.columns

    for out_name in (JULY, OCT):
        part = _raw(inputs / out_name, "Participant Level Summary")
        exit_df = _raw(inputs / out_name, "Exit Log")
        part_names = _unique_names(part)
        exit_names = _unique_names(exit_df)
        ambiguous = _ambiguous_name_keys(part)
        rows = svc[svc["source_file_name"] == out_name]
        assert len(rows) > 0, out_name
        matched = 0
        for _, row in rows.iterrows():
            key = (row["CHAMP ID"], row["ICMS Enrollment ID"])
            if not key[0] or not key[1] or key in ambiguous:
                expected = ""
            else:
                part_name = part_names.get(key)
                if part_name is not None and part_name != core._BLANK_CLIENT_NAME:
                    expected = part_name
                elif key in exit_names:
                    expected = exit_names[key]
                elif part_name is not None:
                    expected = part_name
                else:
                    expected = ""
            assert row["Client Name"] == expected, (out_name, key, row["Service ID"])
            if expected:
                matched += 1
        assert matched > 0, out_name


def test_rerun_does_not_inflate_counts(run_dir, outputs):
    inputs, out, _ = run_dir
    d = core.default_snapshot_date(inputs / OCT)
    res = core.run_snapshot_files(
        [core.FileSpec(inputs / OCT, d["date"], d["source"])], out,
        reprocessing=True, log=quiet,
    )
    assert res["success"] and res["backup_dir"]
    for name, key in ((core.SNAPSHOT_HISTORY_FILENAME, "hist"), (core.EXIT_HISTORY_FILENAME, "exit"),
                      (core.SERVICES_HISTORY_FILENAME, "svc")):
        assert len(pd.read_csv(out / name, dtype=str)) == len(outputs[key]), name
    assert len(pd.read_excel(out / core.MASTER_FILENAME)) == len(outputs["master"])
