"""
icms_catalogue.py
~~~~~~~~~~~~~~~~~
Director-maintained catalogue of approved output fields, accepted source
header names, and supported worksheets. This file is the only place mappings
live; the core, GUI, and CLI all read it. Analysts never edit mappings.

Editing rules
-------------
- ``output`` is the column name written to every export. Renaming it is a
  breaking change for downstream readers; add an accepted source instead.
- ``sources`` lists additional accepted workbook headers, oldest vendor names
  included. The output name is always accepted first. When more than one
  accepted header is present in the same sheet, the earliest in
  ``(output,) + sources`` wins on conflict and the conflict is reported.
- Headers are matched after trimming, collapsing internal whitespace, and
  ignoring case. Nothing else is fuzzy-matched.
- ``expected_since`` (YYYY-MM-DD, from the HSH change log or the first sample
  that carried the header) suppresses "missing field" reports for snapshots
  generated before HSH introduced the field.
- ``retired`` fields stay in the schema for older snapshots; their absence is
  never reported.
- ``report_missing=False`` marks fields that are absent by design.
- ``KNOWN_EXCLUDED_HEADERS`` lists workbook headers that were reviewed and
  intentionally not exported. Headers outside the catalogue and this list are
  reported as unfamiliar for director review and are never exported.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


@dataclass(frozen=True)
class FieldSpec:
    output: str
    sources: Tuple[str, ...] = ()
    essential: bool = False
    expected_since: Optional[str] = None
    retired: bool = False
    report_missing: bool = True
    is_date: bool = False

    @property
    def accepted(self) -> Tuple[str, ...]:
        return (self.output,) + tuple(self.sources)


@dataclass(frozen=True)
class SheetSpec:
    category: str
    names: Tuple[str, ...]
    required: bool = False
    expected_since: Optional[str] = None


def normalize_header(value: object) -> str:
    """Trim, collapse whitespace, and casefold a header or sheet name."""
    if value is None:
        return ""
    return " ".join(str(value).split()).casefold()


F = FieldSpec

# ── Participant Level Summary ────────────────────────────────────────────────
# Order here is the output column order.

PARTICIPANT_FIELDS: List[FieldSpec] = [
    F("CHAMP ID", essential=True),
    F("ICMS Project Name", essential=True),
    F("Date of Birth", is_date=True),
    F("First Name"),
    F("Last Name"),
    F("HMIS ID"),
    F("Medi-Cal CIN"),
    F("ICMS Check-In Date", is_date=True),
    F("ICMS Provider"),
    F("PH Building Name"),
    F("Current ICMS Status"),
    F("ICMS Case Manager"),
    F("ICMS Supervisor"),
    F("ICMS Project Type"),
    F("ICMS Slot #"),
    F("Receiving Pre-Match ICMS?", retired=True),
    F("Checked in 6+ Months and Not Yet Permanently Housed?"),
    F("Permanent Housing Move-In Date", is_date=True),
    F("Housing Voucher Issued Date", is_date=True),
    F("Permanent Housing Subsidy",
      sources=("Current Permanent Housing Subsidy", "Current Rental Subsidy")),
    F("Service Planning Area", sources=("Service Planning Area (SPA)",)),
    F("Household Size"),
    F("PH Update Still Needed This Month?"),
    F("5x5 Still Needed This Month?"),
    F("HAI Still Needed This Month?"),
    F("Care Plan Pending Action Step Still Needed This Month?"),
    F("Care Plan Completed Action Step Still Needed This Month?"),
    F("Eligible Services Still Needed This Month?"),
    F("Housing Navigation Action Step Needed?"),
    F("Health Care Action Step Needed?"),
    F("Key Housing Milestone Data Needed?", retired=True),
    F("ICMS Enrollment ID", essential=True),
    F("ICMS Billing Rate"),
    F("Permanently Housed?"),
    F("Any Required Documentation Still Needed This Month?"),
    F("Notification of Privacy Practices On Record?",
      sources=("Notification of Privacy Practices Status on Record?",)),
    F("Universal Sharing Consent Status"),
    F("CalAIM CS Housing Navigation (HN) and Tenancy Sustaining Services (TSS) "
      "Verbal Opt-In Status Missing?",
      sources=("CalAIM CS Housing Navigation (HN) and Tenancy Sustaining Services "
               "(TSS) Verbal Opt-In Status Needed?",)),
    F("Family Billing Rate But Dependent Household Member(s) Not Listed in "
      "Household Composition?",
      sources=("Family Billing Rate But Eligible Dependent Household Member(s) "
               "Not Recorded in Household Composition?",)),
    F("Eligible Services Recorded for This Month"),
    F("Universal Housing Application (UHA) Status", expected_since="2026-03-27"),
    F("Most Recent Demographic Profile Update", expected_since="2026-03-19", is_date=True),
    F("Most Recent Household Composition Profile Update", expected_since="2026-03-26", is_date=True),
    F("Most Recent Medical Profile Update", expected_since="2026-03-27", is_date=True),
    F("Most Recent Legal Profile Update", expected_since="2026-03-27", is_date=True),
    F("Most Recent PH Update", is_date=True),
    F("Most Recent HMIS Universal Data Assessment", expected_since="2026-04-20", is_date=True),
    F("Most Recent HMIS Financial Assessment", expected_since="2026-04-20", is_date=True),
    F("Most Recent HMIS Barriers Assessment", expected_since="2026-04-20", is_date=True),
    F("Most Recent HMIS DV Assessment", expected_since="2026-04-20", is_date=True),
    F("Most Recent 5x5 Assessment", is_date=True),
    F("Most Recent Housing Acuity Index (HAI) Assessment", is_date=True),
    F("Most Recent Pending Action Step in the Care Plan", is_date=True),
    F("Most Recent Completed Action Step in the Care Plan", is_date=True),
    F("Most Recent Health Care Accompaniment", is_date=True),
    # Approved additions, 2026-10-27 release.
    F("Target Rental Subsidy", expected_since="2026-08-27"),
    F("TLS / Shallow Subsidy Month", expected_since="2026-09-23"),
    F("TLS / Shallow Subsidy Fiscal Agent", expected_since="2026-07-15"),
    F("Public Housing Authority (PHA) Sharing Addendum Status", expected_since="2026-07-23"),
    F("Total Monthly Household Income (Primary Income Source)", expected_since="2026-08-04"),
    F("Household Income as Percentage of Area Median Income (AMI)", expected_since="2026-08-04"),
    F("Medi-Cal Status"),
    F("Medi-Cal Managed Care Plan (MCP)", sources=("Medi-Cal Managed Care Plan",)),
    F("Nominated for Potential ICMS Graduation?", expected_since="2026-05-14"),
    F("Dependents"),
    F("Minors"),
    F("Permanent Housing Move-Out Date", is_date=True),
    # Older vintages capitalize "for"; case-insensitive matching covers it.
    F("In-Person Encounters Recorded For This Month"),
    F("Unsuccessful Outreach Attempts Recorded For This Month"),
]

# ── Exit Log ─────────────────────────────────────────────────────────────────

EXIT_FIELDS: List[FieldSpec] = [
    F("CHAMP ID", essential=True),
    F("ICMS Project Name", essential=True),
    F("ICMS Provider"),
    F("PH Building Name"),
    F("ICMS Project Type"),
    F("ICMS Enrollment ID", essential=True),
    F("ICMS Case Manager"),
    F("First Name"),
    F("Last Name"),
    F("Date of Birth", is_date=True),
    F("HMIS ID"),
    F("Medi-Cal CIN", report_missing=False),
    F("Check-In Date", sources=("ICMS Check-In Date",), is_date=True),
    F("Check-Out Date", sources=("ICMS Check-Out Date",), is_date=True),
    F("Move-In Date", sources=("Permanent Housing Move-In Date",), is_date=True),
    F("Move-Out Date", sources=("Permanent Housing Move-Out Date",), is_date=True),
    F("Exit Reason"),
    F("Exit Destination"),
    F("Exit Location"),
    F("Transferred to New Provider Within ICMS Program?", expected_since="2026-05-19"),
    F("New ICMS Provider"),
    # Approved additions, 2026-10-27 release (only those the Exit Log supplies).
    F("Target Rental Subsidy", expected_since="2026-08-27"),
    F("TLS / Shallow Subsidy Month", expected_since="2026-09-23"),
    F("TLS / Shallow Subsidy Fiscal Agent", expected_since="2026-07-15"),
    F("Dependents"),
    F("Minors"),
]

# ── Vacant Slots ─────────────────────────────────────────────────────────────

VACANT_FIELDS: List[FieldSpec] = [
    F("ICMS Project Name", essential=True),
    F("ICMS Provider"),
    F("PH Building Name"),
    F("ICMS Project Type"),
    F("# Days Vacant", sources=("Days Vacant",)),
]

# ── Service Log ──────────────────────────────────────────────────────────────

SERVICES_FIELDS: List[FieldSpec] = [
    F("Row ID", report_missing=False),
    F("CHAMP ID", essential=True),
    F("DHS Master Agreement #", sources=("HSH Master Agreement #",), report_missing=False),
    F("ICMS Work Order #", report_missing=False),
    F("ICMS Provider Name", sources=("ICMS Provider",)),
    F("PH Building Name"),
    F("ICMS Enrollment ID"),
    F("ICMS Case Manager"),
    F("Service ID"),
    F("Service", sources=("ICMS Service",)),
    F("Place of Service"),
    F("Service Date", is_date=True),
    F("Service Record Created Date", is_date=True),
    F("CHAMP User Who Created Record",
      sources=("CHAMP User Who Created Service Record",)),
    F("Organization of CHAMP User Who Created Record",
      sources=("Organization of CHAMP User Who Created Service Record",)),
]

FIELDS_BY_CATEGORY: Dict[str, List[FieldSpec]] = {
    "participant": PARTICIPANT_FIELDS,
    "exit": EXIT_FIELDS,
    "vacant": VACANT_FIELDS,
    "services": SERVICES_FIELDS,
}

# ── Worksheets ───────────────────────────────────────────────────────────────

SHEETS: Dict[str, SheetSpec] = {
    "participant": SheetSpec("participant", ("Participant Level Summary",), required=True),
    "vacant": SheetSpec("vacant", ("Vacant Slots",), expected_since="2026-04-08"),
    "exit": SheetSpec("exit", ("Exit Log",)),
    "services": SheetSpec("services", ("Service Log", "Services Log")),
}

CATEGORY_LABELS: Dict[str, str] = {
    "participant": "Participant Level Summary",
    "vacant": "Vacant Slots",
    "exit": "Exit Log",
    "services": "Service Log",
}

METADATA_SHEET = "Generated On"

# Reviewed tabs that are intentionally not imported.
KNOWN_EXCLUDED_SHEETS: Tuple[str, ...] = (
    "Generated On",
    "Provider Level Summary",
    "Building Level Summary",
    "Scattered Site Level Summary",
    "Case Manager Level Summary",
    "UHA Log",
    "Hospitalization Log",
    "Action Step Log",
    "Training Log",
    "Trainings Log",
    "Billing Rate Appeal Log",
)

# Reviewed headers that are intentionally not exported, seeded from the
# 02.18, 04.30, 05.19, 07.07, and 10.19.2026 sample snapshots.
KNOWN_EXCLUDED_HEADERS: Dict[str, Tuple[str, ...]] = {
    "participant": (
        "Row ID",
        "Matched Via DMH?",
        "TLS?",
        "Citizenship Status",
        "U.S. Veteran?",
        "Primary Language",
        "Gender Identity",
        "Race and Ethnicity",
        "Address",
        "Main Income Source (Amount)",
        "Housing Subsidy Application Date",
        "Geographic Preference(s) for Housing",
        "Mobility/Accessibility Needs",
        "Disabilities",
        "Living with HIV/AIDS?",
        "Prior Criminal Justice System Involvement?",
        "5x5 ADLs/IADLs Score",
        "5x5 Life Skills Score",
        "5x5 Physical Health Score",
        "5x5 Mental Health Score",
        "5x5 Substance Use Score",
        "HAI Housing Subtotal",
        "HAI Income Health Support Subtotal",
        "HAI Parent and Child Services Subtotal",
        "Emergency Room or Inpatient Visit Within Past 2 Weeks?",
        "Emergency Room Visits Within Past 6 Months",
        "Inpatient Visits Within Past 6 Months",
        "Recommended Care Plan Focus(es) Based on Recent Assessment(s)",
        "Current Care Plan Focus(es) According to CHAMP Record",
        "Property Manager",
        "Interim Housing Provider",
        "Brilliant Corners Housing Coordinator",
        "DMH Mental Health Services Provider",
        "Caregiver",
    ),
    "exit": (
        "Row ID",
        "DHS Master Agreement #",
        "HSH Master Agreement #",
        "ICMS Work Order #",
        "ICMS Project ID",
        "ICMS Slot #",
        "ICMS Slot ID",
        "Property Manager",
        "Receiving Pre-Match ICMS?",
        "Housing Subsidy",
        "Current Permanent Housing Subsidy",
        "Current Rental Subsidy",
        "Citizenship Status",
        "U.S. Veteran?",
        "Primary Language",
        "Gender Identity",
        "Race and Ethnicity",
        "Household Size",
    ),
    "vacant": (
        "Row ID",
        "Slot #",
        "ICMS Slot #",
        "DHS Master Agreement #",
        "ICMS Work Order #",
        "Property Manager",
        "ICMS Project ID",
        "Slot Description",
        "Slot Connected to HSH MRSA Unit?",
        "Pre-Match ICMS Slot?",
        "Slot to Be Filled in Partnership With DMH?",
        "Slot Activation Date",
        "ICMS Slot Activation Date",
        "Slot End Date",
        "ICMS Slot End Date",
        "Slot ID",
        "ICMS Slot ID",
        "Default ICMS Billing Rate",
        "Slot Operation Cost",
        "Slot Vacancy Status",
        "CHAMP ID That Last Occupied Slot",
    ),
    "services": (),
}
