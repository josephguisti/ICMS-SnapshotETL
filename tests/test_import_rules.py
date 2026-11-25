"""Synthetic-workbook regression tests for the 2026-10-27 import rules."""

from pathlib import Path

import pandas as pd

import icms_catalogue as catalogue
import icms_etl_core as core
from conftest import (
    NAMED_EXIT_HEADERS,
    NAMED_PARTICIPANT_HEADERS,
    make_workbook,
    run,
    spec,
)


def _csv(out: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(out / name, dtype=str).fillna("")


def _kinds(results, kind):
    return [g for g in results["finding_groups"] if g["kind"] == kind]


# ── Catalogue ────────────────────────────────────────────────────────────────

def test_catalogue_names_are_unambiguous_within_each_tab():
    for category, specs in catalogue.FIELDS_BY_CATEGORY.items():
        seen = {}
        for s in specs:
            for name in s.accepted:
                key = catalogue.normalize_header(name)
                assert key not in seen, f"{category}: '{name}' used by {seen.get(key)} and {s.output}"
                seen[key] = s.output
        excluded = {catalogue.normalize_header(h) for h in catalogue.KNOWN_EXCLUDED_HEADERS[category]}
        assert not excluded & set(seen), f"{category}: header both accepted and excluded"


def test_aliases_whitespace_and_case_resolve_without_prompts(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx",
        [(1, "Proj A", 10, "SJC", "HACLA", "SPA 6", 0.25)],
        participant_headers=["champ  id", "ICMS Project Name", "ICMS Enrollment ID",
                             "ICMS Provider", "Current Rental Subsidy",
                             "Service Planning Area (SPA)",
                             "Household Income as Percentage of Area Median Income (AMI)"],
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    hist = _csv(out_dir, core.SNAPSHOT_HISTORY_FILENAME)
    row = hist.iloc[0]
    assert row["CHAMP ID"] == "1"
    assert row["Permanent Housing Subsidy"] == "HACLA"
    assert row["Service Planning Area"] == "SPA 6"
    assert row["Household Income as Percentage of Area Median Income (AMI)"] == "0.25"
    assert "Current Rental Subsidy" not in hist.columns
    assert list(hist.columns[:3]) == core.METADATA_FIELDS
    assert list(hist.columns[3:3 + len(core.ALL_APPROVED_FIELDS)]) == core.ALL_APPROVED_FIELDS


def test_unfamiliar_header_is_reported_and_not_exported(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx",
        [(1, "Proj A", 10, "SJC", "surprise")],
        participant_headers=PARTICIPANT_HEADERS_PLUS("Brand New Vendor Column"),
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    found = _kinds(res, "unfamiliar_header")
    assert [g["item"] for g in found] == ["Brand New Vendor Column"]
    assert "Brand New Vendor Column" in res["findings_text"]
    assert "Brand New Vendor Column" not in _csv(out_dir, core.SNAPSHOT_HISTORY_FILENAME).columns
    assert "Brand New Vendor Column" not in pd.read_excel(out_dir / core.DASHBOARD_FILENAME).columns


def PARTICIPANT_HEADERS_PLUS(*extra):
    return ["CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID", "ICMS Provider", *extra]


def test_retired_and_by_design_fields_are_silent(tmp_path, out_dir):
    wb = make_workbook(tmp_path / "s.xlsx", [(1, "Proj A", 10, "SJC")],
                       exit_rows=[(2, "Proj A", 20, "2026-10-01")])
    res = run([spec(wb, "2026-10-19")], out_dir)
    missing = {g["item"] for g in _kinds(res, "missing_field")}
    assert "Receiving Pre-Match ICMS?" not in missing
    assert "Key Housing Milestone Data Needed?" not in missing
    assert "Medi-Cal CIN" not in {g["item"] for g in _kinds(res, "missing_field")
                                  if g["category"] == "Exit Log"}
    # Ordinary missing fields are reported once, as a group, and left blank.
    assert "Household Size" in missing
    hist = _csv(out_dir, core.SNAPSHOT_HISTORY_FILENAME)
    assert hist["Household Size"].eq("").all()


def test_expected_since_suppresses_reports_for_older_snapshots(tmp_path, out_dir):
    wb = make_workbook(tmp_path / "s.xlsx", [(1, "Proj A", 10, "SJC")],
                       generated_on="02/18/2026 7:00 am")
    res = run([spec(wb, "2026-02-18")], out_dir)
    missing = {g["item"] for g in _kinds(res, "missing_field")}
    assert "Target Rental Subsidy" not in missing
    assert "Most Recent HMIS DV Assessment" not in missing
    assert "Household Size" in missing


def test_coexisting_headers_keep_first_and_record_conflict(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx",
        [(1, "Proj A", 10, "SJC", "HACLA", "LACDA"), (2, "Proj A", 11, "SJC", None, "HUD")],
        participant_headers=PARTICIPANT_HEADERS_PLUS("Current Permanent Housing Subsidy",
                                                     "Current Rental Subsidy"),
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    hist = _csv(out_dir, core.SNAPSHOT_HISTORY_FILENAME).set_index("CHAMP ID")
    assert hist.loc["1", "Permanent Housing Subsidy"] == "HACLA"
    assert hist.loc["2", "Permanent Housing Subsidy"] == "HUD"
    conflicts = _kinds(res, "value_conflict")
    assert conflicts and "1 row(s) disagree" in conflicts[0]["details"][0]


def test_identity_conflict_and_missing_essential_fail_only_that_file(tmp_path, out_dir):
    good = make_workbook(tmp_path / "good 10.12.2026.xlsx", [(1, "Proj A", 10, "SJC")],
                         generated_on="10/12/2026 9:00 am")
    run([spec(good, "2026-10-12")], out_dir)
    before = {n: (out_dir / n).read_bytes() for n in core.OUTPUT_FILENAMES
              if (out_dir / n).exists() and n != core.INGEST_LOG_FILENAME}

    conflict = make_workbook(
        tmp_path / "conflict 10.19.2026.xlsx", [(1, "Proj A", 10, 99)],
        participant_headers=["CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID",
                             "icms enrollment id"],
    )
    no_enrollment = make_workbook(
        tmp_path / "noenr 10.20.2026.xlsx", [(1, "Proj A", "SJC")],
        participant_headers=["CHAMP ID", "ICMS Project Name", "ICMS Provider"],
        generated_on="10/20/2026 9:00 am",
    )
    res = run([spec(conflict, "2026-10-19"), spec(no_enrollment, "2026-10-20")], out_dir,
              reprocessing=False)
    assert res["batch_failed"] == 2 and not res["success"]
    assert {g["kind"] for g in res["finding_groups"]} >= {"essential_conflict", "essential_missing"}
    after = {n: (out_dir / n).read_bytes() for n in before}
    assert after == before, "failed files must not change stored outputs or the dashboard"
    log = _csv(out_dir, core.INGEST_LOG_FILENAME)
    assert set(log.loc[log.snapshot_date != "2026-10-12", "run_status"]) == {"failed"}


def test_failed_file_does_not_block_rest_of_batch(tmp_path, out_dir):
    bad = make_workbook(tmp_path / "bad 10.12.2026.xlsx", [(1, "Proj A", "x")],
                        participant_headers=["CHAMP ID", "ICMS Project Name", "ICMS Provider"])
    good = make_workbook(tmp_path / "good 10.19.2026.xlsx", [(1, "Proj A", 10, "SJC")])
    res = run([spec(bad, "2026-10-12"), spec(good, "2026-10-19")], out_dir)
    assert res["success"] and res["batch_success"] == 1 and res["batch_failed"] == 1
    assert res["dashboard_source"] == good.name


# ── Sheets and dates ─────────────────────────────────────────────────────────

def test_service_and_services_log_both_accepted_regardless_of_setting(tmp_path, out_dir):
    svc = [(1, 10, "S1", "Case Management", "Office", "2026-10-01", "2026-10-02")]
    a = make_workbook(tmp_path / "a 10.12.2026.xlsx", [(1, "Proj A", 10, "SJC")],
                      services_sheet="Services Log", services_rows=svc,
                      generated_on="10/12/2026 9:00 am")
    b = make_workbook(tmp_path / "b 10.19.2026.xlsx", [(1, "Proj A", 10, "SJC")],
                      services_sheet="Service Log", services_rows=svc)
    res = run([spec(a, "2026-10-12"), spec(b, "2026-10-19")], out_dir,
              sheet_names={"services": "Services Log"})
    assert res["batch_success"] == 2 and not res["skipped_categories"]
    svc_hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME)
    assert len(svc_hist) == 1 and svc_hist.loc[0, "Service"] == "Case Management"


def test_differing_candidate_tabs_are_not_guessed_and_marked_incomplete(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx", [(1, "Proj A", 10, "SJC")],
        services_sheet="Service Log",
        services_rows=[(1, 10, "S1", "CM", "Office", "2026-10-01", "2026-10-02")],
        extra_sheets={"Services Log": [["CHAMP ID", "Service ID"], [2, "S9"]]},
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    assert res["skipped_categories"] == ["Service Log"]
    assert _kinds(res, "ambiguous_sheet")
    assert not (out_dir / core.SERVICES_HISTORY_FILENAME).exists()


def test_generated_on_is_found_by_content_and_preferred(tmp_path):
    wb = make_workbook(tmp_path / "ICMS Snapshot - 10.18.2026.xlsx", [(1, "Proj A", 10, "SJC")],
                       generated_on="10/19/2026 9:01 am")
    info = core.default_snapshot_date(wb)
    assert info == {"date": "2026-10-19", "source": "Generated On", "generated_on": "2026-10-19",
                    "filename_date": "2026-10-18", "mismatch": True}


def test_filename_fallback_and_date_mismatch_logged(tmp_path, out_dir):
    no_meta = make_workbook(tmp_path / "ICMS Snapshot - 10.12.2026.xlsx",
                            [(1, "Proj A", 10, "SJC")], generated_on=None)
    info = core.default_snapshot_date(no_meta)
    assert (info["date"], info["source"]) == ("2026-10-12", "filename")

    mism = make_workbook(tmp_path / "ICMS Snapshot - 10.18.2026.xlsx", [(1, "Proj A", 10, "SJC")])
    d = core.default_snapshot_date(mism)
    res = run([core.FileSpec(mism, d["date"], d["source"], d["generated_on"], d["filename_date"])],
              out_dir)
    assert _kinds(res, "date_mismatch")
    log = _csv(out_dir, core.INGEST_LOG_FILENAME).iloc[-1]
    assert (log["snapshot_date"], log["snapshot_date_source"]) == ("2026-10-19", "Generated On")
    assert (log["generated_on_date"], log["filename_date"]) == ("2026-10-19", "2026-10-18")


# ── Master grain and status ──────────────────────────────────────────────────

def test_master_keeps_gender_and_race_from_the_participant_sheet(tmp_path, out_dir):
    headers = PARTICIPANT_HEADERS_PLUS(
        "Gender Identity", "Race and Ethnicity", "Citizenship Status", "Primary Language",
    )
    july = make_workbook(
        tmp_path / "july.xlsx",
        [
            (1, "Proj A", 10, "SJC", "Female", "White", "US Citizen", "English"),
            (2, "Proj A", 20, "SJC", "Male", "Black", "US Citizen", "Spanish"),
        ],
        participant_headers=headers,
        generated_on="07/07/2026 8:00 am",
    )
    october = make_workbook(
        tmp_path / "oct.xlsx",
        [(1, "Proj A", 10, "SJC", "Non-Binary", "Asian", "US Citizen", "English")],
        participant_headers=headers,
        exit_headers=["CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID",
                      "ICMS Check-Out Date", "Gender Identity", "Race and Ethnicity"],
        exit_rows=[(2, "Proj A", 20, "2026-09-01", "Should Not Win", "Should Not Win")],
    )
    res = run([spec(july, "2026-07-07"), spec(october, "2026-10-19")], out_dir)
    assert res["success"]
    reported = {g["item"] for g in res["finding_groups"]}
    assert "Citizenship Status" not in reported
    assert "Primary Language" not in reported

    m = pd.read_excel(out_dir / core.MASTER_FILENAME, dtype=str).fillna("")
    assert list(m.columns) == core.MASTER_COLUMNS
    tracking = core.MASTER_COLUMNS.index("present_in_latest_snapshot")
    assert core.MASTER_COLUMNS[tracking - 2:tracking] == ["Gender Identity", "Race and Ethnicity"]

    active = m[m["ICMS Enrollment ID"] == "10"].iloc[0]
    assert active["Gender Identity"] == "Non-Binary"
    assert active["Race and Ethnicity"] == "Asian"
    exited = m[m["ICMS Enrollment ID"] == "20"].iloc[0]
    assert exited["client_status"] == "Exited"
    assert exited["Gender Identity"] == "Male"
    assert exited["Race and Ethnicity"] == "Black"

    withheld = ("Citizenship Status", "Primary Language", "U.S. Veteran?", "Address",
                "Disabilities", "Living with HIV/AIDS?")
    hist = _csv(out_dir, core.SNAPSHOT_HISTORY_FILENAME)
    flat = pd.read_excel(out_dir / core.DASHBOARD_FILENAME, dtype=str).fillna("")
    for col in withheld:
        assert col not in m.columns
        assert col not in hist.columns
        assert col not in flat.columns


def test_master_is_per_enrollment_with_exit_only_and_unknown(tmp_path, out_dir):
    july = make_workbook(
        tmp_path / "july 07.07.2026.xlsx",
        [(9038, "Proj A", 20415, "SJC"), (77, "Proj B", 700, "SJC"), (88, "Proj B", 800, "SJC")],
        generated_on="07/07/2026 8:13 am",
    )
    october = make_workbook(
        tmp_path / "oct 10.19.2026.xlsx",
        [(9038, "Proj A", 529134, "SJC"), (77, "Proj B", 700, "SJC")],
        exit_rows=[(88, "Proj B", 800, "2026-09-01"), (55, "Proj C", 555, "2026-10-01")],
    )
    run([spec(october, "2026-10-19"), spec(july, "2026-07-07")], out_dir)
    m = pd.read_excel(out_dir / core.MASTER_FILENAME, dtype=str).fillna("")
    status = {(r["CHAMP ID"], r["ICMS Enrollment ID"]): r["client_status"] for _, r in m.iterrows()}
    assert status == {
        ("9038", "20415"): "Unknown",
        ("9038", "529134"): "Active",
        ("77", "700"): "Active",
        ("88", "800"): "Exited",
        ("55", "555"): "Exited",
    }
    old = m[(m["CHAMP ID"] == "9038") & (m["ICMS Enrollment ID"] == "20415")].iloc[0]
    assert old["snapshot_date"].startswith("2026-07-07")
    exited = m[m["CHAMP ID"] == "88"].iloc[0]
    assert exited["Check-Out Date"].startswith("2026-09-01")


def test_two_enrollments_same_project_are_not_duplicates(tmp_path, out_dir):
    wb = make_workbook(tmp_path / "s.xlsx", [(1, "Proj A", 10, "SJC"), (1, "Proj A", 11, "SJC")])
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["batch_results"][0]["dup_flag_count"] == 0


# ── Repair, idempotence, ordering, publishing ────────────────────────────────

def _three_files(tmp_path):
    files = []
    for i, (d, gen) in enumerate([("2026-09-01", "09/01/2026"), ("2026-09-15", "09/15/2026"),
                                  ("2026-10-01", "10/01/2026")]):
        files.append(spec(make_workbook(
            tmp_path / f"f{i} {d[5:7]}.{d[8:]}.2026.xlsx",
            [(1, "Proj A", 10, "SJC"), (2 + i, "Proj B", 20 + i, "SJC")],
            exit_rows=[(9, "Proj A", 90, f"{d}")] if i else [],
            services_rows=[(1, 10, "", "CM", "Office", "2026-08-01", "2026-08-02"),
                           (1, 10, f"S{i}", "Visit", "Home", d, d)],
            generated_on=f"{gen} 9:00 am",
        ), d))
    return files


def _snapshot_outputs(out: Path):
    data = {n: _csv(out, n) for n in (core.SNAPSHOT_HISTORY_FILENAME, core.EXIT_HISTORY_FILENAME,
                                      core.SERVICES_HISTORY_FILENAME)}
    data["master"] = pd.read_excel(out / core.MASTER_FILENAME, dtype=str).fillna("")
    return data


def test_selection_order_and_repeat_import_give_identical_outputs(tmp_path):
    files = _three_files(tmp_path)
    a, b = tmp_path / "a", tmp_path / "b"
    run(files, a)
    run(list(reversed(files)), b)
    run(files[1:2], b, reprocessing=True)
    run(files, b, reprocessing=True)
    left, right = _snapshot_outputs(a), _snapshot_outputs(b)
    for name in left:
        pd.testing.assert_frame_equal(left[name], right[name], check_like=False, obj=name)
    assert len(left[core.SERVICES_HISTORY_FILENAME]) == 4


def test_reprocessing_backs_up_outputs(tmp_path, out_dir):
    files = _three_files(tmp_path)
    run(files[:1], out_dir)
    res = run(files[:1], out_dir, reprocessing=True)
    assert res["backup_dir"]
    backed = {p.name for p in Path(res["backup_dir"]).iterdir()}
    assert core.SNAPSHOT_HISTORY_FILENAME in backed


def test_old_format_outputs_are_backed_up_once(tmp_path, out_dir):
    files = _three_files(tmp_path)
    pd.DataFrame({"snapshot_date": ["2026-01-01"], "CHAMP ID": ["5"], "ICMS Project Name": ["P"],
                  "ICMS Enrollment ID": ["50"]}).to_csv(out_dir / core.SNAPSHOT_HISTORY_FILENAME,
                                                        index=False)
    first = run(files[:1], out_dir)
    second = run(files[1:2], out_dir)
    assert "format upgrade" in first["backup_dir"] and second["backup_dir"] == ""


def test_damaged_services_rows_are_repaired_without_duplicates(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "july 07.07.2026.xlsx", [(1, "Proj A", 10, "SJC")],
        services_rows=[(1, 10, "", "Case Management", "Office", "2026-07-01", "2026-07-02"),
                       (1, 10, "S5", "Visit", "Home", "2026-07-03", "2026-07-03")],
        generated_on="07/07/2026 8:13 am",
    )
    other = make_workbook(
        tmp_path / "july 07.14.2026.xlsx", [(1, "Proj A", 10, "SJC")],
        services_rows=[(1, 10, "", "Case Management", "Office", "2026-07-01", "2026-07-02")],
        generated_on="07/14/2026 8:13 am",
    )
    # History as an older release wrote it: Service blank (missed "ICMS Service"
    # rename) for both the representative and a later observation.
    damaged = pd.DataFrame([
        {"snapshot_date": "2026-07-07", "reporting_month": "2026-07", "source_file_name": wb.name,
         "CHAMP ID": "1", "ICMS Enrollment ID": "10", "Service ID": "", "Service": "",
         "Place of Service": "Office", "Service Date": "2026-07-01",
         "Service Record Created Date": "2026-07-02", "PH Building Name": "Scattered Site"},
        {"snapshot_date": "2026-07-07", "reporting_month": "2026-07", "source_file_name": wb.name,
         "CHAMP ID": "1", "ICMS Enrollment ID": "10", "Service ID": "S5", "Service": "",
         "Place of Service": "Home", "Service Date": "2026-07-03",
         "Service Record Created Date": "2026-07-03", "PH Building Name": "Scattered Site"},
    ])
    damaged.to_csv(out_dir / core.SERVICES_HISTORY_FILENAME, index=False)

    run([spec(other, "2026-07-14")], out_dir)
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME)
    assert len(hist) == 2, "repaired name must not create a second fallback event"

    run([spec(wb, "2026-07-07")], out_dir, reprocessing=True)
    run([spec(wb, "2026-07-07"), spec(other, "2026-07-14")], out_dir, reprocessing=True)
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME).set_index("Service Date")
    assert len(hist) == 2
    assert hist.loc["2026-07-01", "Service"] == "Case Management"
    assert hist.loc["2026-07-03", "Service"] == "Visit"
    assert hist.loc["2026-07-01", "snapshot_date"] == "2026-07-07"


def test_publish_toggle_and_newest_only(tmp_path, out_dir):
    files = _three_files(tmp_path)
    res = run(files[:2], out_dir, publish_dashboard=False)
    assert not (out_dir / core.DASHBOARD_FILENAME).exists() and not res["dashboard_published"]

    res = run(files, out_dir, reprocessing=True)
    assert res["dashboard_source"] == Path(files[2].path).name
    flat = pd.read_excel(out_dir / core.DASHBOARD_FILENAME, dtype=str)
    assert set(flat["snapshot_date"].str[:10]) == {"2026-10-01"}
    assert list(flat.columns) == core.FLATFILE_COLUMNS

    older = run(files[:2], out_dir, reprocessing=True)
    assert not older["dashboard_published"] and "older" in older["dashboard_note"]
    flat = pd.read_excel(out_dir / core.DASHBOARD_FILENAME, dtype=str)
    assert set(flat["snapshot_date"].str[:10]) == {"2026-10-01"}


# ── Services project name ────────────────────────────────────────────────────

def test_services_project_name_comes_from_same_snapshot(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx",
        [(1, "Proj A", 10, "SJC"), (1, "Proj B", 11, "SJC")],
        exit_rows=[(2, "Proj C", 20, "2026-10-10")],
        services_rows=[(1, 10, "S1", "Visit", "Home", "2026-10-01", "2026-10-01"),
                       (1, 11, "S2", "Visit", "Home", "2026-10-02", "2026-10-02"),
                       (2, 20, "S3", "Visit", "Home", "2026-10-03", "2026-10-03"),
                       (3, 30, "S4", "Visit", "Home", "2026-10-04", "2026-10-04")],
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    assert "ICMS Project Name" not in {g["item"] for g in _kinds(res, "missing_field")
                                       if g["category"] == "Service Log"}
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME).set_index("Service ID")
    assert hist.loc["S1", "ICMS Project Name"] == "Proj A"
    assert hist.loc["S2", "ICMS Project Name"] == "Proj B"
    assert hist.loc["S3", "ICMS Project Name"] == "Proj C"
    assert hist.loc["S4", "ICMS Project Name"] == ""
    cols = list(hist.reset_index().columns)
    assert cols.index("ICMS Project Name") == cols.index("CHAMP ID") + 1


# ── Services client name ─────────────────────────────────────────────────────

def test_services_client_name_comes_from_same_snapshot(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx",
        [
            (1, "Proj A", 10, "SJC", "Ada", "Lovelace"),
            (1, "Proj B", 11, "SJC", "Grace", "Hopper"),
            (4, "Proj D", 40, "SJC", "OnlyFirst", None),
            (5, "Proj E", 50, "SJC", None, "OnlyLast"),
            (6, "Proj F", 60, "SJC", None, None),
            (7, "Proj G", 70, "SJC", "Part", "One"),
            (7, "Proj G", 70, "SJC", "Part", "Two"),
        ],
        participant_headers=NAMED_PARTICIPANT_HEADERS,
        exit_headers=NAMED_EXIT_HEADERS,
        exit_rows=[
            (2, "Proj C", 20, "2026-10-10", "Exit", "Client"),
            (6, "Proj F", 60, "2026-10-11", "Exit", "Fallback"),
            (7, "Proj G", 70, "2026-10-12", "Exit", "Ignored"),
            (8, "Proj H", 80, "2026-10-13", "Exit", "Only"),
            (8, "Proj H", 80, "2026-10-13", "Other", "Name"),
        ],
        services_rows=[
            (1, 10, "S1", "Visit", "Home", "2026-10-01", "2026-10-01"),
            (1, 11, "S2", "Visit", "Home", "2026-10-02", "2026-10-02"),
            (2, 20, "S3", "Visit", "Home", "2026-10-03", "2026-10-03"),
            (3, 30, "S4", "Visit", "Home", "2026-10-04", "2026-10-04"),
            (4, 40, "S5", "Visit", "Home", "2026-10-05", "2026-10-05"),
            (5, 50, "S6", "Visit", "Home", "2026-10-06", "2026-10-06"),
            (6, 60, "S7", "Visit", "Home", "2026-10-07", "2026-10-07"),
            (7, 70, "S8", "Visit", "Home", "2026-10-08", "2026-10-08"),
            (8, 80, "S9", "Visit", "Home", "2026-10-09", "2026-10-09"),
            (1, "", "S10", "Visit", "Home", "2026-10-10", "2026-10-10"),
        ],
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    assert "Client Name" not in {g["item"] for g in _kinds(res, "missing_field")
                                 if g["category"] == "Service Log"}
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME).set_index("Service ID")
    assert hist.loc["S1", "Client Name"] == "Ada Lovelace"
    assert hist.loc["S2", "Client Name"] == "Grace Hopper"
    assert hist.loc["S3", "Client Name"] == "Exit Client"
    assert hist.loc["S4", "Client Name"] == ""
    assert hist.loc["S5", "Client Name"] == "OnlyFirst [BLANK LAST]"
    assert hist.loc["S6", "Client Name"] == "[BLANK FIRST] OnlyLast"
    assert hist.loc["S7", "Client Name"] == "Exit Fallback"
    assert hist.loc["S8", "Client Name"] == ""
    assert hist.loc["S9", "Client Name"] == ""
    assert hist.loc["S10", "Client Name"] == ""
    cols = list(hist.reset_index().columns)
    assert cols[-1] == "Client Name" or cols.index("Client Name") > cols.index(
        "Organization of CHAMP User Who Created Record"
    )
    assert "Client Name" in cols
    assert len(hist) == 10


def test_services_client_name_participant_wins_over_exit(tmp_path, out_dir):
    wb = make_workbook(
        tmp_path / "s.xlsx",
        [(1, "Proj A", 10, "SJC", "Part", "Name")],
        participant_headers=NAMED_PARTICIPANT_HEADERS,
        exit_headers=NAMED_EXIT_HEADERS,
        exit_rows=[(1, "Proj A", 10, "2026-10-10", "Exit", "Name")],
        services_rows=[(1, 10, "S1", "Visit", "Home", "2026-10-01", "2026-10-01")],
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME)
    assert hist.loc[0, "Client Name"] == "Part Name"


def test_empty_services_run_rewrites_history_with_client_name_column(tmp_path, out_dir):
    # Seed an older-format services history without Client Name.
    old = pd.DataFrame([{
        "snapshot_date": "2026-10-01",
        "reporting_month": "2026-10",
        "source_file_name": "prior.xlsx",
        "CHAMP ID": "1",
        "ICMS Enrollment ID": "10",
        "Service ID": "S1",
        "Service": "Visit",
        "Place of Service": "Home",
        "Service Date": "2026-10-01",
        "Service Record Created Date": "2026-10-01",
        "PH Building Name": "Scattered Site",
    }])
    old.to_csv(out_dir / core.SERVICES_HISTORY_FILENAME, index=False)
    # A workbook with no Service Log tab still advances the format and rewrites.
    wb = make_workbook(
        tmp_path / "no-services.xlsx",
        [(1, "Proj A", 10, "SJC", "Ada", "Lovelace")],
        participant_headers=NAMED_PARTICIPANT_HEADERS,
        services_sheet=None,
    )
    res = run([spec(wb, "2026-10-19")], out_dir)
    assert res["success"]
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME)
    assert "Client Name" in hist.columns
    assert len(hist) == 1
    assert hist.loc[0, "Service ID"] == "S1"
    assert hist.loc[0, "Client Name"] == ""


def test_services_client_name_backfills_without_duplicating_events(tmp_path, out_dir):
    # Older history has the event but no client name yet (format upgrade / prior blank).
    old = pd.DataFrame([{
        "snapshot_date": "2026-10-01",
        "reporting_month": "2026-10",
        "source_file_name": "prior.xlsx",
        "CHAMP ID": "1",
        "ICMS Enrollment ID": "10",
        "Service ID": "S1",
        "Service": "Visit",
        "Place of Service": "Home",
        "Service Date": "2026-10-01",
        "Service Record Created Date": "2026-10-01",
        "PH Building Name": "Scattered Site",
        "Client Name": "",
    }])
    old.to_csv(out_dir / core.SERVICES_HISTORY_FILENAME, index=False)
    named = make_workbook(
        tmp_path / "b.xlsx",
        [(1, "Proj A", 10, "SJC", "Ada", "Lovelace")],
        participant_headers=NAMED_PARTICIPANT_HEADERS,
        services_rows=[(1, 10, "S1", "Visit", "Home", "2026-10-01", "2026-10-01")],
        generated_on="10/19/2026 9:00 am",
    )
    run([spec(named, "2026-10-19")], out_dir)
    hist = _csv(out_dir, core.SERVICES_HISTORY_FILENAME)
    assert len(hist) == 1
    assert hist.loc[0, "Client Name"] == "Ada Lovelace"
    assert hist.loc[0, "snapshot_date"] == "2026-10-01"
