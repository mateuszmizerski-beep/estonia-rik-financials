"""Translate normalized RIK statement rows into the existing Estonia extractor model."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import date
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
UNCONSOLIDATED_TYPES = {"22", "23", "28", "29", "30", "32", "33", "37", "40", "41"}


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
        "Other revenue",
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
        "Depreciation and impairment of fixed assets",
    ),
    "Fixed assets": (
        "Kokku põhivarad",
        "Kokku põhivara",
        "Põhivara kokku",
        "Põhivarad",
        "Non-current assets",
        "Total non-current assets",
        "Non-current assets total",
        "Fixed assets",
    ),
    "Current assets": (
        "Kokku käibevarad",
        "Kokku käibevara",
        "Käibevara kokku",
        "Käibevarad",
        "Current assets",
        "Total current assets",
        "Current assets total",
    ),
    "Stocks / inventories": (
        "Varud",
        "Kokku varud",
        "Inventories",
        "Inventories total",
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
        "Long-term loan liabilities total",
    ),
    "Debt ST": (
        "Lühiajalised laenukohustised",
        "Lühiajalised laenud ja võlakohustised",
        "Short-term borrowings",
        "Short-term loans and borrowings",
        "Current borrowings",
        "Current loans",
        "Loan liabilities total",
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

# Stable RIK row identifiers, guarded by the expected normalized line label.  RIK
# uses the same row scheme across consolidated, standalone and unconsolidated
# variants, but different schemes for income-statement layout 1 and layout 2.
ROW_NUMBER_HINTS: dict[tuple[str, str], str] = {}
for report_type in BALANCE_TYPES:
    ROW_NUMBER_HINTS.update(
        {
            (report_type, "10"): "Cash and cash equivalents",
            (report_type, "70"): "Stocks / inventories",
            (report_type, "100"): "Current assets",
            (report_type, "200"): "Fixed assets",
            (report_type, "310"): "Debt ST",
            (report_type, "410"): "Debt LT",
        }
    )
for report_type in ("35", "15", "02", "2", "23", "37", "40", "28", "30"):
    ROW_NUMBER_HINTS.update(
        {
            (report_type, "10"): "Revenue",
            (report_type, "40"): "Other income",
            (report_type, "110"): "COGS",
            (report_type, "140"): "D&A",
            (report_type, "170"): "Reported EBIT",
        }
    )
for report_type in ("36", "16", "03", "3", "41", "29"):
    ROW_NUMBER_HINTS.update(
        {
            (report_type, "10"): "Revenue",
            (report_type, "25"): "COGS",
            (report_type, "40"): "Other income",
            (report_type, "170"): "Reported EBIT",
        }
    )
for report_type in CASH_FLOW_TYPES:
    ROW_NUMBER_HINTS[(report_type, "281")] = "CAPEX"


def guarded_row_hint(line: StatementLine) -> str | None:
    report_type = line.report_type.lstrip("0") or "0"
    hinted = ROW_NUMBER_HINTS.get((line.report_type, line.row_number)) or ROW_NUMBER_HINTS.get(
        (report_type, line.row_number)
    )
    normalized_name = _norm(line.row_name)
    if hinted and normalized_name in NORMALIZED_ALIASES[hinted]:
        return hinted
    return None


def canonical_item(line: StatementLine) -> str | None:
    hinted = guarded_row_hint(line)
    if hinted:
        return hinted
    normalized_name = _norm(line.row_name)
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


def structured_accounting_basis(lines: list[StatementLine]) -> str:
    report_types = {line.report_type.lstrip("0") or "0" for line in lines}
    report_names = {_norm(line.report_name) for line in lines}
    if report_types and report_types.issubset(UNCONSOLIDATED_TYPES) or any(
        "konsolideerimata" in name or "unconsolidated" in name for name in report_names
    ):
        return "unconsolidated"
    if any("konsolideeritud" in name or "consolidated" in name for name in report_names):
        return "consolidated"
    return "reported"


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
        accounting_basis=structured_accounting_basis(lines),
    )
    if report.period_end is None:
        report.period_end = next(
            (line.period_end for line in lines if line.period_end is not None), None
        )
    ordered_lines = sorted(lines, key=lambda line: 0 if guarded_row_hint(line) else 1)
    for line in ordered_lines:
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
    availability: list[AnnualReportAvailability],
    target_years: list[int],
    documents: list[CompanyDocument] | None = None,
) -> list[int]:
    """Return filings needed to fill targets using the next filing's comparatives."""
    available_years = {item.fiscal_year for item in availability}
    if documents:
        available_years.update(
            document.fiscal_year
            for document in documents
            if document.document_type == "X" and document.fiscal_year is not None
        )
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
    if primary.accounting_basis == "unknown":
        primary.accounting_basis = fallback.accounting_basis
    for year, values in fallback.values.items():
        has_detailed_da = all(
            values.get(item) is not None
            for item in ("Total depreciation", "Total amortisation")
        )
        primary_da_source = primary.get_source(year, "D&A") or ""
        if has_detailed_da and not primary_da_source.startswith("RIK XBRL |"):
            # The statement-level D&A row can differ from the audited fixed-asset
            # notes. When both note components are disclosed, keep the more
            # granular pair and let the workbook calculate their combined total.
            primary.values.setdefault(year, {}).pop("D&A", None)
            primary.sources.setdefault(year, {}).pop("D&A", None)
        for item, value in values.items():
            if item == "D&A" and has_detailed_da:
                continue
            if primary.get_value(year, item) is None:
                primary.set_value(year, item, value, fallback.get_source(year, item) or "Annual report")
    for year, segmentations in fallback.segments.items():
        for segment_by, records in segmentations.items():
            destination = primary.segments.setdefault(year, {}).setdefault(segment_by, {})
            if destination and any(
                (record.source or "").startswith("RIK XBRL |")
                for record in destination.values()
            ):
                continue
            for label, record in records.items():
                if label not in destination:
                    destination[label] = copy(record)
    return primary


WORKBOOK_FINANCIAL_ITEMS = {
    "Revenue",
    "Other income",
    "COGS",
    "Reported EBIT",
    "Total depreciation",
    "Total amortisation",
    "D&A",
    "Fixed assets",
    "Current assets",
    "Stocks / inventories",
    "Trade debtors / receivables",
    "Trade creditors / payables",
    "Cash and cash equivalents",
    "Debt LT",
    "Debt ST",
    "Investments in tangible assets",
    "Investments in intangible assets",
    "CAPEX",
    "FTEs",
    "Goodwill amortisation",
    "Management EBITDA",
}


def replace_unconsolidated_with_consolidated(
    primary: EstonianReport, fallback: EstonianReport
) -> dict[int, list[str]]:
    """Replace a complete workbook block instead of mixing accounting perimeters."""
    if primary.accounting_basis != "unconsolidated" or fallback.accounting_basis != "consolidated":
        return {}

    replaced: dict[int, list[str]] = {}
    for year, fallback_values in fallback.values.items():
        anchors = {
            item
            for item in ("Revenue", "Reported EBIT", "Fixed assets", "Current assets")
            if fallback_values.get(item) is not None
        }
        if len(anchors) < 3:
            continue
        primary_values = primary.values.setdefault(year, {})
        primary_sources = primary.sources.setdefault(year, {})
        for item in WORKBOOK_FINANCIAL_ITEMS:
            primary_values.pop(item, None)
            primary_sources.pop(item, None)
        for item, value in fallback_values.items():
            if item not in WORKBOOK_FINANCIAL_ITEMS:
                continue
            primary.set_value(year, item, value, fallback.get_source(year, item) or "Annual report")
        primary.segments[year] = {
            segment_by: {label: copy(record) for label, record in records.items()}
            for segment_by, records in fallback.segments.get(year, {}).items()
        }
        replaced[year] = sorted(
            item for item in fallback_values if item in WORKBOOK_FINANCIAL_ITEMS
        )
    if replaced:
        primary.accounting_basis = "consolidated"
    return replaced


def needs_document_fallback(report: EstonianReport, fiscal_year: int) -> bool:
    core_items = (
        "Revenue",
        "Reported EBIT",
        "Fixed assets",
        "Current assets",
        "Cash and cash equivalents",
        "FTEs",
    )
    return (
        report.accounting_basis == "unconsolidated"
        or any(report.get_value(fiscal_year, item) is None for item in core_items)
        or not report.segments.get(fiscal_year)
    )


def fallback_document_candidates(
    documents: list[CompanyDocument], fiscal_year: int
) -> list[CompanyDocument]:
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
    return candidates


def select_fallback_document(
    documents: list[CompanyDocument], fiscal_year: int
) -> CompanyDocument | None:
    candidates = fallback_document_candidates(documents, fiscal_year)
    return candidates[0] if candidates else None


def select_xbrl_document(
    documents: list[CompanyDocument], fiscal_year: int
) -> CompanyDocument | None:
    """Select the latest valid full annual-report XBRL package for a fiscal year."""
    candidates = [
        document
        for document in documents
        if document.fiscal_year == fiscal_year
        and document.document_type == "X"
        and document.validity in {None, "", "K"}
        and document.report_kind in {None, "", "A", "P"}
    ]
    candidates.sort(
        key=lambda document: (
            document.status_date or date.min,
            document.size_bytes or 0,
            document.document_id,
        ),
        reverse=True,
    )
    return candidates[0] if candidates else None


def select_annual_report_pdf(
    documents: list[CompanyDocument], fiscal_year: int
) -> CompanyDocument | None:
    """Select the complete, valid annual-report PDF for a fiscal year."""
    candidates = [
        document
        for document in documents
        if document.fiscal_year == fiscal_year
        and document.document_type == "A"
        and document.validity in {None, "", "K"}
        and document.report_kind in {None, "", "A", "P"}
    ]
    candidates.sort(key=lambda document: -(document.size_bytes or 0))
    return candidates[0] if candidates else None
