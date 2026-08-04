from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import load_workbook

from api_adapter import preferred_report_types, report_from_lines, source_report_years
from estonia_extractor import EstonianReport, build_fill_jobs, load_terms
from financials_service import (
    build_annual_report_pdf_bundle,
    generate_workbook,
    workbook_source_details,
)
from rik_xml_client import (
    AnnualReportAvailability,
    CompanyDocument,
    DownloadedDocument,
    StatementLine,
)
from workbook_preservation import count_extended_validations


PROJECT = Path(__file__).parents[1]
TEMPLATE = PROJECT / "assets" / "gainpro_template_eur.xlsx"


def availability(report_type: str, year: int = 2024) -> AnnualReportAvailability:
    return AnnualReportAvailability(
        report_type,
        f"Report {report_type}",
        year,
        date(year, 1, 1),
        date(year, 12, 31),
    )


def line(
    report_type: str,
    row_number: str,
    row_name: str,
    column: str,
    value: str,
) -> StatementLine:
    return StatementLine(
        registry_code="70000310",
        fiscal_year=2024,
        report_type=report_type,
        report_name=f"Report {report_type}",
        period_start=date(2024, 1, 1),
        period_end=date(2024, 12, 31),
        row_number=row_number,
        row_name=row_name,
        column_code=column,
        column_name="Current" if column == "A1" else "Previous",
        value=Decimal(value),
        retrieved_at=datetime(2025, 1, 2, tzinfo=timezone.utc),
    )


class AdapterTests(unittest.TestCase):
    def test_consolidated_statements_are_preferred(self) -> None:
        items = [
            availability("14"),
            availability("22"),
            availability("35"),
            availability("40"),
            availability("18"),
            availability("32"),
        ]
        self.assertEqual(preferred_report_types(items, 2024), ["14", "35", "18"])

    def test_income_layout_two_and_comparative_mapping(self) -> None:
        report = report_from_lines(
            "70000310",
            "Fixture Company",
            2024,
            [
                line("36", "10", "Revenue", "A1", "2000000"),
                line("36", "10", "Revenue", "A2", "1500000"),
                line("36", "110", "Operating profit", "A1", "100000"),
            ],
        )
        self.assertEqual(report.get_value(2024, "Revenue"), Decimal("2000000"))
        self.assertEqual(report.get_value(2023, "Revenue"), Decimal("1500000"))
        self.assertIn("statement 36", report.get_source(2024, "Revenue") or "")

    def test_source_report_years_include_next_filing(self) -> None:
        items = [availability("35", 2025), availability("35", 2024)]
        self.assertEqual(source_report_years(items, [2024]), [2025, 2024])

    def test_next_annual_report_comparatives_take_priority(self) -> None:
        reports: list[EstonianReport] = []
        for report_year, values in (
            (2025, {2025: "100", 2024: "91"}),
            (2024, {2024: "80", 2023: "71"}),
            (2023, {2023: "60"}),
        ):
            report = EstonianReport(
                Path(f"AR{report_year}.xml"),
                f"AR{report_year}",
                [],
                load_terms(None),
                period_start=date(report_year, 1, 1),
                period_end=date(report_year, 12, 31),
                company="Fixture Company",
            )
            for value_year, value in values.items():
                report.set_value(
                    value_year,
                    "Revenue",
                    Decimal(value),
                    f"RIK XML | AR{report_year} "
                    f"{'current' if value_year == report_year else 'comparative'}",
                )
            reports.append(report)

        jobs, _ = build_fill_jobs(
            reports,
            years=3,
            fill_comparative=True,
            target_years=[2025, 2024, 2023],
        )
        by_year = {job.year: job for job in jobs}
        self.assertEqual(by_year[2025].report.get_value(2025, "Revenue"), Decimal("100"))
        self.assertEqual(by_year[2024].report.get_value(2024, "Revenue"), Decimal("91"))
        self.assertEqual(by_year[2023].report.get_value(2023, "Revenue"), Decimal("71"))
        self.assertIn("AR2025 comparative", by_year[2024].report.get_source(2024, "Revenue") or "")
        self.assertIn("AR2024 comparative", by_year[2023].report.get_source(2023, "Revenue") or "")

        by_year[2024].report.set_value(2024, "FTEs", Decimal("42"), "AR2025 PDF")
        details = workbook_source_details(
            jobs,
            fill_segments=True,
            fill_cogs=True,
            fill_goodwill_amortisation=True,
        )
        detail_2024 = next(detail for detail in details if detail.startswith("FY2024"))
        self.assertIn("FY2024 ← AR2025 comparative", detail_2024)
        self.assertIn("document fallback items: FTEs", detail_2024)


class WorkbookTests(unittest.TestCase):
    def make_report(self, days: int) -> EstonianReport:
        report = EstonianReport(
            Path("fixture.xml"),
            "RIK structured fixture",
            [],
            load_terms(None),
            period_start=date(2024, 1, 1),
            period_end=date.fromordinal(date(2024, 1, 1).toordinal() + days - 1),
            company="Fixture Company",
        )
        for year, revenue in ((2024, "910000"), (2023, "3650000")):
            report.set_value(year, "Revenue", Decimal(revenue), f"RIK row 10 A{1 if year == 2024 else 2}")
            report.set_value(year, "Reported EBIT", Decimal("91000"), "RIK row 110")
            report.set_value(year, "D&A", Decimal("9000"), "RIK row 90")
            report.set_value(year, "Fixed assets", Decimal("1000000"), "RIK balance row 190")
            report.set_value(year, "Current assets", Decimal("500000"), "RIK balance row 100")
            report.set_value(year, "Cash and cash equivalents", Decimal("100000"), "RIK balance row 20")
        return report

    def test_annualisation_boundaries(self) -> None:
        self.assertFalse(self.make_report(90).requires_annualisation)
        self.assertTrue(self.make_report(91).requires_annualisation)
        self.assertFalse(self.make_report(365).requires_annualisation)
        self.assertFalse(self.make_report(366).requires_annualisation)

    def test_workbook_preserves_template_and_annualises_flows_only(self) -> None:
        report = self.make_report(91)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "generated.xlsx"
            generate_workbook(
                [report],
                [2024, 2023],
                TEMPLATE,
                output,
                company_name="Fixture Company",
            )
            workbook = load_workbook(output, data_only=False)
            worksheet = workbook["Financials"]
            self.assertEqual(worksheet["P136"].value, 0.91)
            self.assertEqual(worksheet["P61"].value, "=P136/$P$133")
            self.assertEqual(worksheet["P78"].value, 1)
            self.assertEqual(worksheet["U2"].value, "CAGR 2023-2024")
            self.assertTrue(worksheet.column_dimensions["R"].hidden)
            self.assertEqual(workbook.calculation.calcMode, "auto")
            workbook.close()
            self.assertEqual(count_extended_validations(output), 2)

    def test_complete_annual_report_pdfs_are_bundled_by_year(self) -> None:
        document = CompanyDocument(
            document_id="pdf-2024",
            registry_code="70000310",
            document_type="A",
            document_name="Annual report PDF",
            size_bytes=15,
            status_date=date(2025, 5, 1),
            validity="K",
            report_kind="A",
            fiscal_year=2024,
            url="https://example.invalid/ar2024.pdf",
            retrieved_at=datetime(2025, 5, 1, tzinfo=timezone.utc),
        )

        class DownloadClient:
            def download_document(self, selected):
                self.assert_document(selected)
                return DownloadedDocument(b"%PDF-1.4\n%%EOF", "source.pdf", "application/pdf")

            @staticmethod
            def assert_document(selected):
                if selected.document_id != "pdf-2024":
                    raise AssertionError("Unexpected PDF document")

        bundle = build_annual_report_pdf_bundle(
            DownloadClient(),
            [document],
            [2024],
            company_name="Fixture Company",
        )
        self.assertEqual(bundle.included_years, (2024,))
        with ZipFile(BytesIO(bundle.content)) as archive:
            self.assertEqual(archive.namelist(), ["Fixture Company AR 2024.pdf"])
            self.assertTrue(archive.read(archive.namelist()[0]).startswith(b"%PDF-"))


if __name__ == "__main__":
    unittest.main()
