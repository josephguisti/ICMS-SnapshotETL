import sys
from pathlib import Path

import openpyxl
import pytest

REPO = Path(__file__).resolve().parents[1]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

import icms_etl_core as core  # noqa: E402

SAMPLE_DIR = REPO / "Sample input"
SAMPLE_OUTPUT_DIR = REPO / "Deidentified Sample Output"

PARTICIPANT_HEADERS = ["CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID", "ICMS Provider"]
# Optional named-participant layout for Client Name enrichment tests.
NAMED_PARTICIPANT_HEADERS = [
    "CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID", "ICMS Provider",
    "First Name", "Last Name",
]
NAMED_EXIT_HEADERS = [
    "CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID", "ICMS Check-Out Date",
    "First Name", "Last Name",
]


def quiet(*_args, **_kwargs):
    pass


def make_workbook(
    path: Path,
    participant_rows,
    *,
    participant_headers=PARTICIPANT_HEADERS,
    exit_headers=None,
    exit_rows=(),
    services_sheet="Service Log",
    services_headers=None,
    services_rows=(),
    vacant_rows=None,
    generated_on="10/19/2026 9:01 am",
    extra_sheets=None,
):
    """Write a minimal snapshot workbook with the supported tabs."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Participant Level Summary"
    ws.append(list(participant_headers))
    for r in participant_rows:
        ws.append(list(r))

    ws = wb.create_sheet("Exit Log")
    ws.append(list(exit_headers or ["CHAMP ID", "ICMS Project Name", "ICMS Enrollment ID",
                                    "ICMS Check-Out Date"]))
    for r in exit_rows:
        ws.append(list(r))

    if services_sheet:
        ws = wb.create_sheet(services_sheet)
        ws.append(list(services_headers or ["CHAMP ID", "ICMS Enrollment ID", "Service ID",
                                            "ICMS Service", "Place of Service", "Service Date",
                                            "Service Record Created Date"]))
        for r in services_rows:
            ws.append(list(r))

    ws = wb.create_sheet("Vacant Slots")
    ws.append(["ICMS Project Name", "ICMS Provider", "Days Vacant"])
    for r in vacant_rows or [("Proj A", "SJC", 12)]:
        ws.append(list(r))

    if generated_on is not None:
        ws = wb.create_sheet("Generated On")
        ws.append(["Row ID", "Generated On"])
        ws.append(["abc-123", generated_on])

    for name, rows in (extra_sheets or {}).items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(list(r))
    wb.save(path)
    return path


def spec(path: Path, date: str) -> core.FileSpec:
    return core.FileSpec(Path(path), date, "test")


def run(files, out, **kwargs):
    kwargs.setdefault("log", quiet)
    return core.run_snapshot_files(files, out, **kwargs)


@pytest.fixture
def out_dir(tmp_path):
    d = tmp_path / "out"
    d.mkdir()
    return d
