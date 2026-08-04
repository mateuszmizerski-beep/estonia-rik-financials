"""Translate normalized RIK statement rows into the existing Estonia extractor model."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata

from estonia_extractor import EstonianReport, load_terms
from rik_xml_client import (
    AnnualReportAvailability,
    CompanyDocument,
    RikError,
    RikXmlClient,
    StatementLine,
)


BALANCE_TYPES = ("14", "01", "1", "12", "08", "8", "22")
INCOME_TYPES = (
    "35",
    "36",
    "15",
    "16",
    "02",
    "2",
    "03",
    "3",
    "23",
    "37",
    "40",
    "41",
    "28",
    "29",
    "30",
)
CASH_FLOW_TYPES = ("18", "19", "04", "4", "13", "32", "33")
FIXED_ASSET_TYPES = ("05", "5", "06", "6")


@dataclass(frozen=True)
class StructuredFetchResult:
    reports: list[EstonianReport]
    lines: list[StatementLine]
    warnings: list[str]


def _norm(value: str) -> str:
    value = unicodedata.normalize("NFKD", value.casefold())
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


ALIASES: dict[str, tuple[str, ...]] = {
    "Revenue": (
        "Müügitulu",
        "Müügitulu kliendilepingutest",
        "Revenue",
        "Sales revenue",
        "Net sales",
    ),
    "Other income": (
        "Muud äritulud",
        "Muu äritulu",
        "Other operating income",
        "Other business income",
    ),
    "COGS": (
        "Kaubad, toore, materjal ja teenused",
        "Müüdud toodangu (kaupade, teenuste) kulu",
        "Müügikulud",
        "Raw materials and consumables used",
        "Goods, raw materials and services",
        "Cost of goods sold",
        "Cost of sales",
    ),
    "Reported EBIT": (
        "Ärikasum (kahjum)",
        "Ärikasum",
        "Kasum äritegevusest",
        "Operating profit (loss)",
        "Operating profit",
        "Profit (loss) from operating activities",
    ),
    "D&A": (
        "Põhivarade kulum ja väärtuse langus",
        "Põhivara kulum ja amortisatsioon",
        "Depreciation, amortisation and impairment loss",
        "Depreciation, amortization and impairment loss",
        "Depreciation and amortisation",
        "Depreciation and amortization",
    ),
    "Fixed assets": (
        "Kokku põhivarad",
        "Kokku põhivara",
        "Põhivara kokku",
        "Põhivarad",
        "Non-current assets",
        "Total non-current assets",
        "Fixed assets",
    ),
    "Current assets": (
        "Kokku käibevarad",
        "Kokku käibevara",
        "Käibevara kokku",
        "Käibevarad",
        "Current assets",
        "Total current assets",
    ),
    "Stocks / inventories": (
        "Varud",
        "Kokku varud",
        "Inventories",
        "Stocks",
    ),
    "Trade debtors / receivables": (
        "Nõuded ostjate vastu",
        "Nõuded ostjate vastu ja muud nõuded",
        "Nõuded klientide vastu",
        "Trade receivables",
        "Trade and other receivables",
        "Receivables from customers",
    ),
    "Trade creditors / payables": (
        "Võlad tarnijatele",
        "Võlad tarnijatele ja ettemaksed",
        "Trade payables",
        "Trade and other payables",
    ),
    "Cash and cash equivalents": (
        "Raha",
        "Raha ja raha ekvivalendid",
        "Raha ja pangakontod",
        "Cash",
        "Cash and cash equivalents",
        "Cash and bank",
    ),
    "Debt LT": (
        "Pikaajalised laenukohustised",
        "Pikaajalised laenud ja võlakohustised",
        "Long-term borrowings",
        "Long-term loans and borrowings",
        "Non-current borrowings",
        "Non-current loans",
    ),
    "Debt ST": (
        "Lühiajalised laenukohustised",
        "Lühiajalised laenud ja võlakohustised",
        "Short-term borrowings",
        "Short-term loans and borrowings",
        "Current borrowings",
        "Current loans",
    ),
    "CAPEX": (
        "Tasutud materiaalsete ja immateriaalsete põhivarade soetamisel",
        "Tasutud materiaalse põhivara soetamisel",
        "Põhivara soetamine",
        "Purchase of property, plant and equipment and intangible assets",
        "Acquisition of property, plant and equipment and intangible assets",
        "Purchase of fixed assets",
    ),
}

NORMALIZED_ALIASES = {
    item: {_norm(alias) for alias in aliases} for item, aliases in ALIASES.items()
}

# Stable row identifiers for the most common XBRL income-statement layout.
ROW_NUMBER_HINTS: dict[tuple[str, str], str] = {
    ("35", "10"): "Revenue",
    ("35", "40"): "Other income",
    ("35", "60"): "COGS",
    ("35", "90"): "D&A",
    ("35", "110"): "Reported EBIT",
    ("15", "10"): "Revenue",
    ("15", "40"): "Other income",
    ("15", "60"): "COGS",
    ("15", "90"): "D&A",
    ("15", "110"): "Reported EBIT",
}


def canonical_item(line: StatementLine) -> str | None:
    report_type = line.report_type.lstrip("0") or "0"
    hinted = ROW_NUMBER_HINTS.get((line.report_type, line.row_number)) or ROW_NUMBER_HINTS.get(
        (report_type, line.row_number)
    )
    normalized_name = _norm(line.row_name)
    if hinted and normalized_name in NORMALIZED_ALIASES[hinted]:
        return hinted
    for item, aliases in NORMALIZED_ALIASES.items():
        if normalized_name in aliases:
            return item
    return None


def _value_year(line: StatementLine) -> int | None:
    if line.report_type in FIXED_ASSET_TYPES:
        return None
    code = line.column_code.upper()
    if code == "A1":
        return line.fiscal_year
    if code == "A2":
        return line.fiscal_year - 1
    return None


def _source(line: StatementLine) -> str:
    return (
        f"RIK XML | statement {line.report_type} {line.report_name} | "
        f"row {line.row_number} {line.row_name} | column {line.column_code} {line.column_name} | "
        f"retrieved {line.retrieved_at.isoformat()}"
    )


def report_from_lines(
    registry_code: str,
    company_name: str,
    fiscal_year: int,
    lines: list[StatementLine],
) -> EstonianReport:
    dated_line = next((line for line in lines if line.period_end is not None), None)
    report = EstonianReport(
        source_path=Path(f"RIK_API_{registry_code}_{fiscal_year}.xml"),
        payload_name=f"RIK structured statements FY{fiscal_year}",
        lines=[],
        terms=load_terms(None),
        period_start=dated_line.period_start if dated_line else None,
        period_end=dated_line.period_end if dated_line else None,
        company=company_name,
    )
    if report.period_end is None:
        report.period_end = next(
            (line.period_end for line in lines if line.period_end is not None), None
        )
    for line in lines:
        if line.value is None:
            continue
        item = canonical_item(line)
        value_year = _value_year(line)
        if item is None or value_year is None:
            continue
        if report.get_value(value_year, item) is not None:
            continue
        report.set_value(value_year, item, line.value, _source(line))
    return report


def _pick_one(available: set[str], priorities: tuple[str, ...]) -> str | None:
    normalized = {value.lstrip("0") or "0": value for value in available}
    for priority in priorities:
        if priority in available:
            return priority
        stripped = priority.lstrip("0") or "0"
        if stripped in normalized:
            return normalized[stripped]
    return None


def preferred_report_types(
    availability: list[AnnualReportAvailability], fiscal_year: int
) -> list[str]:
    available = {item.report_type for item in availability if item.fiscal_year == fiscal_year}
    selected: list[str] = []
    for priorities in (BALANCE_TYPES, INCOME_TYPES, CASH_FLOW_TYPES, FIXED_ASSET_TYPES):
        report_type = _pick_one(available, priorities)
        if report_type and report_type not in selected:
            selected.append(report_type)
    return selected


def source_report_years(
    availability: list[AnnualReportAvailability], target_years: list[int]
) -> list[int]:
    """Return filings needed to fill targets using the next filing's comparatives."""
    available_years = {item.fiscal_year for item in availability}
    selected = set(target_years)
    selected.update(year + 1 for year in target_years if year + 1 in available_years)
    return sorted(selected, reverse=True)


def fetch_structured_reports(
    client: RikXmlClient,
    registry_code: str,
    company_name: str,
    availability: list[AnnualReportAvailability],
    fiscal_years: list[int],
) -> StructuredFetchResult:
    reports: list[EstonianReport] = []
    all_lines: list[StatementLine] = []
    warnings: list[str] = []
    for year in sorted(set(fiscal_years), reverse=True):
        year_lines: list[StatementLine] = []
        report_types = preferred_report_types(availability, year)
        if not report_types:
            warnings.append(f"FY{year}: RIK has no transferred structured statements.")
            continue
        for report_type in report_types:
            try:
                lines = client.get_annual_report_lines(
                    registry_code, report_type, year, language="eng"
                )
            except RikError as exc:
                warnings.append(f"FY{year} statement {report_type}: {exc}")
                continue
            if not lines:
                warnings.append(f"FY{year} statement {report_type}: no rows returned.")
                continue
            year_lines.extend(lines)
            all_lines.extend(lines)
        if year_lines:
            reports.append(report_from_lines(registry_code, company_name, year, year_lines))
    return StructuredFetchResult(reports=reports, lines=all_lines, warnings=warnings)


def merge_missing(primary: EstonianReport, fallback: EstonianReport) -> EstonianReport:
    """Fill only gaps in an API report from a parsed document report."""
    if primary.period_start is None:
        primary.period_start = fallback.period_start
    if primary.period_end is None:
        primary.period_end = fallback.period_end
    if not primary.company:
        primary.company = fallback.company
    for year, values in fallback.values.items():
        for item, value in values.items():
            if primary.get_value(year, item) is None:
                primary.set_value(year, item, value, fallback.get_source(year, item) or "Annual report")
    for year, segmentations in fallback.segments.items():
        for segment_by, records in segmentations.items():
            destination = primary.segments.setdefault(year, {}).setdefault(segment_by, {})
            for label, record in records.items():
                if label not in destination:
                    destination[label] = copy(record)
    return primary


def needs_document_fallback(report: EstonianReport, fiscal_year: int) -> bool:
    core_items = (
        "Revenue",
        "Reported EBIT",
        "Fixed assets",
        "Current assets",
        "Cash and cash equivalents",
        "FTEs",
    )
    return any(report.get_value(fiscal_year, item) is None for item in core_items) or not report.segments.get(
        fiscal_year
    )


def select_fallback_document(
    documents: list[CompanyDocument], fiscal_year: int
) -> CompanyDocument | None:
    candidates = [
        document
        for document in documents
        if document.fiscal_year == fiscal_year
        and document.document_type in {"A", "D"}
        and document.validity in {None, "", "K"}
        and document.report_kind in {None, "", "A", "P"}
    ]
    candidates.sort(
        key=lambda document: (
            0 if document.document_type == "A" else 1,
            -(document.size_bytes or 0),
        )
    )
    return candidates[0] if candidates else None
