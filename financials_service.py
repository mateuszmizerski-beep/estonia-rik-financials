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
    source_details: list[str]


def safe_company_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "", value or "").strip()
    value = re.sub(r"\s+", " ", value)
    return value[:120] or "Estonia Company"


def workbook_source_details(
    jobs: list[extractor.FillJob],
    *,
    fill_segments: bool,
    fill_cogs: bool,
    fill_goodwill_amortisation: bool,
) -> list[str]:
    enabled_flags = {
        "fill_cogs": fill_cogs,
        "fill_goodwill_amortisation": fill_goodwill_amortisation,
    }
    mapped_items: list[str] = []
    for mapping in extractor.WORKBOOK_MAPPINGS:
        if mapping.optional_flag and not enabled_flags.get(mapping.optional_flag, False):
            continue
        if mapping.item not in mapped_items:
            mapped_items.append(mapping.item)

    details: list[str] = []
    for job in sorted(jobs, key=lambda item: item.year, reverse=True):
        source_year = job.report.year
        source_period = (
            "comparative"
            if source_year == job.year + 1
            else "current"
            if source_year == job.year
            else "available figures"
        )
        fallback_items = [
            item
            for item in mapped_items
            if job.report.get_value(job.value_year, item) is not None
            and not (job.report.get_source(job.value_year, item) or "").startswith("RIK XML |")
        ]
        parts = [f"FY{job.year} ← AR{source_year} {source_period}"]
        parts.append(
            f"document fallback items: {', '.join(fallback_items)}"
            if fallback_items
            else "structured XML only"
        )
        if fill_segments:
            segment_counts: list[str] = []
            for segment_by, records in sorted(job.report.segments.get(job.value_year, {}).items()):
                count = sum(
                    not (record.source or "").startswith("RIK XML |")
                    for record in records.values()
                )
                if count:
                    segment_counts.append(f"{segment_by} ({count})")
            if segment_counts:
                parts.append(f"document segmentations: {', '.join(segment_counts)}")
        details.append(" — ".join(parts) + ".")
    return details


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
            if primary is None:
                reports.append(fallback)
                reports_by_year[year] = fallback
            else:
                merge_missing(primary, fallback)
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
    source_details = workbook_source_details(
        jobs,
        fill_segments=fill_segments,
        fill_cogs=fill_cogs,
        fill_goodwill_amortisation=fill_goodwill_amortisation,
    )
    return GeneratedWorkbook(
        content=output_path.read_bytes(),
        filename=filename,
        messages=messages,
        source_details=source_details,
    )
