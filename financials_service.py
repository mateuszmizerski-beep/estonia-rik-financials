"""Application service for document fallback and template workbook generation."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import tempfile

from openpyxl import load_workbook

import estonia_extractor as extractor
from api_adapter import merge_missing, needs_document_fallback, select_fallback_document
from rik_xml_client import CompanyDocument, RikError, RikXmlClient
from workbook_preservation import restore_extended_validations


@dataclass(frozen=True)
class GeneratedWorkbook:
    content: bytes
    filename: str
    messages: list[str]


@dataclass(frozen=True)
class FallbackAdditions:
    value_items: dict[int, tuple[str, ...]]
    segment_counts: dict[int, dict[str, int]]

    @property
    def has_additions(self) -> bool:
        return bool(self.value_items or self.segment_counts)


def safe_company_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "", value or "").strip()
    value = re.sub(r"\s+", " ", value)
    return value[:120] or "Estonia Company"


def fallback_additions(
    primary: extractor.EstonianReport | None,
    fallback: extractor.EstonianReport,
) -> FallbackAdditions:
    value_items: dict[int, tuple[str, ...]] = {}
    segment_counts: dict[int, dict[str, int]] = {}
    for value_year, values in fallback.values.items():
        added = sorted(
            item
            for item in values
            if primary is None or primary.get_value(value_year, item) is None
        )
        if added:
            value_items[value_year] = tuple(added)
    for value_year, segmentations in fallback.segments.items():
        for segment_by, records in segmentations.items():
            existing = (
                primary.segments.get(value_year, {}).get(segment_by, {})
                if primary is not None
                else {}
            )
            added_count = sum(label not in existing for label in records)
            if added_count:
                segment_counts.setdefault(value_year, {})[segment_by] = added_count
    return FallbackAdditions(value_items, segment_counts)


def format_fallback_additions(report_year: int, additions: FallbackAdditions) -> str:
    details: list[str] = []
    for value_year in sorted(additions.value_items, reverse=True):
        details.append(
            f"FY{value_year} items: {', '.join(additions.value_items[value_year])}"
        )
    for value_year in sorted(additions.segment_counts, reverse=True):
        dimensions = ", ".join(
            f"{segment_by} ({count})"
            for segment_by, count in sorted(additions.segment_counts[value_year].items())
        )
        details.append(f"FY{value_year} segmentations: {dimensions}")
    if not details:
        return (
            f"AR{report_year}: the source document was parsed, but it contained no "
            "additional mapped items or segmentations."
        )
    return f"AR{report_year}: source document additions — {'; '.join(details)}."


def add_document_fallbacks(
    client: RikXmlClient,
    reports: list[extractor.EstonianReport],
    documents: list[CompanyDocument],
    fiscal_years: list[int],
    *,
    company_name: str,
) -> tuple[list[extractor.EstonianReport], list[str]]:
    reports_by_year = {report.year: report for report in reports if report.period_end is not None}
    warnings: list[str] = []
    with tempfile.TemporaryDirectory() as temporary_directory:
        temporary_path = Path(temporary_directory)
        for year in sorted(set(fiscal_years), reverse=True):
            primary = reports_by_year.get(year)
            if primary is not None and not needs_document_fallback(primary, year):
                continue
            document = select_fallback_document(documents, year)
            if document is None:
                warnings.append(f"FY{year}: no PDF or BDOC fallback document was available.")
                continue
            try:
                downloaded = client.download_document(document)
                suffix = document.extension or Path(downloaded.filename).suffix or ".bin"
                source_path = temporary_path / f"{document.document_id}_{year}{suffix}"
                source_path.write_bytes(downloaded.content)
                parsed_reports = extractor.parse_reports_from_input(
                    source_path, extractor.load_terms(None)
                )
                fallback = next(
                    (candidate for candidate in parsed_reports if candidate.year == year),
                    parsed_reports[0],
                )
                fallback.source_path = Path(downloaded.filename)
                if not fallback.company:
                    fallback.company = company_name
            except (RikError, OSError, ValueError) as exc:
                warnings.append(f"FY{year}: document fallback could not be parsed ({exc}).")
                continue
            additions = fallback_additions(primary, fallback)
            if primary is None:
                reports.append(fallback)
                reports_by_year[year] = fallback
            else:
                merge_missing(primary, fallback)
            warnings.append(format_fallback_additions(year, additions))
    return reports, warnings


def generate_workbook(
    reports: list[extractor.EstonianReport],
    fiscal_years: list[int],
    template_path: Path,
    output_path: Path,
    *,
    company_name: str,
    fill_segments: bool = True,
    fill_cogs: bool = True,
    fill_goodwill_amortisation: bool = False,
) -> GeneratedWorkbook:
    if not reports:
        raise ValueError("No structured or document report data was available to fill the workbook.")
    selected_years = sorted(set(fiscal_years), reverse=True)
    jobs, messages = extractor.build_fill_jobs(
        reports,
        years=max(len(selected_years), 1),
        fill_comparative=True,
        target_years=selected_years,
    )
    jobs = [job for job in jobs if job.year in selected_years]
    if not jobs:
        raise ValueError("None of the selected fiscal years could be prepared for the workbook.")

    workbook = load_workbook(template_path)
    if extractor.FINANCIALS_SHEET not in workbook.sheetnames:
        raise ValueError("The embedded template does not contain the Financials sheet.")
    worksheet = workbook[extractor.FINANCIALS_SHEET]

    for job in jobs:
        messages.append(
            f"=== Filling FY{job.year} from {job.report.payload_name} values FY{job.value_year} ==="
        )
        messages.extend(
            extractor.fill_period(
                worksheet,
                job,
                overwrite=True,
                add_comments=True,
                fill_cogs=fill_cogs,
                fill_goodwill_amortisation=fill_goodwill_amortisation,
                dry_run=False,
            )
        )
    messages.extend(extractor.update_cagr_formulas(worksheet, jobs, dry_run=False))
    messages.extend(extractor.group_unused_year_columns(worksheet, jobs, dry_run=False))
    if fill_segments:
        messages.extend(
            extractor.fill_segments_sheet(
                workbook,
                jobs,
                overwrite=True,
                add_comments=True,
                dry_run=False,
            )
        )

    extractor.save_workbook(workbook, output_path)
    restore_extended_validations(template_path, output_path)
    filename = f"{safe_company_name(company_name)} Financials.xlsx"
    return GeneratedWorkbook(
        content=output_path.read_bytes(),
        filename=filename,
        messages=messages,
    )
