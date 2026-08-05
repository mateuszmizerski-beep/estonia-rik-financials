from __future__ import annotations

import os
from pathlib import Path
import tempfile
from functools import lru_cache

from openpyxl import load_workbook
import streamlit as st

from api_adapter import fetch_structured_reports, source_report_years
from financials_service import (
    add_document_fallbacks,
    build_annual_report_pdf_bundle,
    generate_workbook,
)
from rik_xml_client import RikError, RikXmlClient


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "assets" / "gainpro_template_eur.xlsx"
MAX_YEARS = 6
# Parser revision: Exmet regression (D&A source priority and segment rest buckets).


@lru_cache(maxsize=1)
def template_years() -> set[int]:
    workbook = load_workbook(TEMPLATE_PATH, read_only=True, data_only=False)
    try:
        worksheet = workbook["Financials"]
        return {
            int(cell.value)
            for cell in worksheet[2]
            if isinstance(cell.value, (int, float)) and float(cell.value).is_integer()
        }
    finally:
        workbook.close()


def document_format(document_type: str) -> str:
    return {"A": "PDF", "D": "BDOC/DDOC", "X": "XBRL", "P": "Articles"}.get(
        document_type, document_type
    )


def credentials() -> tuple[str, str]:
    username = os.getenv("RIK_USERNAME", "")
    password = os.getenv("RIK_PASSWORD", "")
    try:
        username = username or str(st.secrets.get("RIK_USERNAME", ""))
        password = password or str(st.secrets.get("RIK_PASSWORD", ""))
        rik_section = st.secrets.get("rik", {})
        if rik_section:
            username = username or str(rik_section.get("username", ""))
            password = password or str(rik_section.get("password", ""))
    except (FileNotFoundError, KeyError):
        pass
    return username.strip(), password


def make_client() -> RikXmlClient:
    username, password = credentials()
    if not username or not password:
        raise RikError(
            "RIK credentials are not configured. Add RIK_USERNAME and RIK_PASSWORD to Streamlit Secrets."
        )
    return RikXmlClient(username, password)


def availability_rows(items):
    return [
        {
            "Year": item.fiscal_year,
            "Statement code": item.report_type,
            "Statement": item.report_name,
            "Period start": item.period_start,
            "Period end": item.period_end,
        }
        for item in items
    ]


def statement_rows(items):
    return [
        {
            "Year": item.fiscal_year,
            "Statement": f"{item.report_type} — {item.report_name}",
            "Row": item.row_number,
            "Line item": item.row_name,
            "Column": item.column_code,
            "Column name": item.column_name,
            "Value": None if item.value is None else str(item.value),
        }
        for item in items
    ]


def document_rows(items):
    return [
        {
            "Year": item.fiscal_year,
            "Format": document_format(item.document_type),
            "Name": item.document_name,
            "Size (MB)": round((item.size_bytes or 0) / 1024 / 1024, 2),
            "Valid": item.validity or "",
        }
        for item in items
    ]


def load_structured(registry_code: str, company_name: str, availability, years):
    client = make_client()
    report_years = source_report_years(availability, list(years))
    result = fetch_structured_reports(
        client,
        registry_code,
        company_name,
        availability,
        report_years,
    )
    st.session_state["structured_key"] = (registry_code, tuple(sorted(years)))
    st.session_state["structured_result"] = result
    return result


st.set_page_config(page_title="Gain Estonia Financials", page_icon="🇪🇪", layout="wide")
st.title("Gain Estonia Financials Extractor")
st.write(
    "Enter an Estonian registry code to retrieve structured annual-report statements "
    "from RIK and generate a completed EUR financials workbook."
)
st.info(
    "Consolidated RIK XML is preferred for the core statements. If RIK exposes only "
    "unconsolidated XML, the complete consolidated annual report replaces that financial "
    "block; PDF or BDOC reports also supply detailed disclosures such as FTEs and segments."
)

with st.form("company_lookup"):
    registry_code = st.text_input(
        "Estonian registry code",
        value=st.session_state.get("registry_code", ""),
        max_chars=8,
        autocomplete="off",
        placeholder=None,
    ).strip()
    lookup = st.form_submit_button("Find annual reports", type="primary")

if lookup:
    if not (registry_code.isdigit() and len(registry_code) == 8):
        st.error("Enter a valid eight-digit Estonian registry code.")
    else:
        try:
            with st.spinner("Retrieving company and annual-report information from RIK..."):
                client = make_client()
                company = client.get_company_summary(registry_code)
                availability = client.list_annual_reports(registry_code)
                documents = client.list_company_documents(registry_code)
            st.session_state["registry_code"] = registry_code
            st.session_state["company"] = company
            st.session_state["availability"] = availability
            st.session_state["documents"] = documents
            st.session_state.pop("structured_result", None)
            st.session_state.pop("generated_workbook", None)
            st.session_state.pop("annual_report_pdf_bundle", None)
        except RikError as exc:
            st.error(str(exc))

company = st.session_state.get("company")
availability = st.session_state.get("availability", [])
documents = st.session_state.get("documents", [])

if company is not None:
    st.subheader(company.name)
    details = [company.registry_code]
    if company.legal_form:
        details.append(company.legal_form)
    if company.status:
        details.append(company.status)
    st.caption(" · ".join(details))
    if company.address:
        st.caption(company.address)

    if availability:
        with st.expander("Available structured statements", expanded=False):
            st.dataframe(
                availability_rows(availability),
                hide_index=True,
                use_container_width=True,
            )
    else:
        st.warning(
            "RIK did not return transferred structured statements. Available annual-report "
            "documents can still be used as a fallback."
        )

    annual_documents = [item for item in documents if item.document_type in {"A", "D", "X"}]
    available_years = sorted(
        (
            {item.fiscal_year for item in availability}
            | {item.fiscal_year for item in annual_documents if item.fiscal_year is not None}
        )
        & template_years(),
        reverse=True,
    )
    if not available_years:
        st.warning("No annual-report years match the years available in the EUR template.")
    else:
        default_years = available_years[:MAX_YEARS]
        selected_years = st.multiselect(
            "Fiscal years to include",
            options=available_years,
            default=default_years,
            max_selections=MAX_YEARS,
        )
        with st.expander("Advanced settings"):
            use_document_fallback = st.checkbox(
                "Use PDF/BDOC reports to supplement missing XML values and segmentations",
                value=True,
            )
            fill_goodwill_amortisation = st.checkbox(
                "Fill goodwill amortisation adjustment when disclosed",
                value=True,
            )

        preview_col, generate_col = st.columns(2)
        preview_clicked = preview_col.button(
            "Preview structured data", disabled=not selected_years, use_container_width=True
        )
        generate_clicked = generate_col.button(
            "Generate Excel workbook",
            type="primary",
            disabled=not selected_years,
            use_container_width=True,
        )

        if preview_clicked:
            try:
                with st.spinner("Retrieving selected statements..."):
                    load_structured(company.registry_code, company.name, availability, selected_years)
            except RikError as exc:
                st.error(str(exc))

        if generate_clicked:
            try:
                with st.spinner("Retrieving statements and creating the workbook..."):
                    key = (company.registry_code, tuple(sorted(selected_years)))
                    if st.session_state.get("structured_key") == key:
                        structured = st.session_state["structured_result"]
                    else:
                        structured = load_structured(
                            company.registry_code, company.name, availability, selected_years
                        )
                    reports = list(structured.reports)
                    generation_warnings = list(structured.warnings)
                    client = make_client()
                    downloaded_annual_reports = {}
                    if use_document_fallback:
                        reports, fallback_warnings = add_document_fallbacks(
                            client,
                            reports,
                            documents,
                            source_report_years(availability, list(selected_years)),
                            company_name=company.name,
                            downloaded_documents=downloaded_annual_reports,
                        )
                        generation_warnings.extend(fallback_warnings)
                    with tempfile.TemporaryDirectory() as temporary_directory:
                        output_path = Path(temporary_directory) / "Filled_Financials.xlsx"
                        generated = generate_workbook(
                            reports,
                            selected_years,
                            TEMPLATE_PATH,
                            output_path,
                            company_name=company.name,
                            fill_goodwill_amortisation=fill_goodwill_amortisation,
                        )
                    annual_report_pdf_bundle = build_annual_report_pdf_bundle(
                        client,
                        documents,
                        list(selected_years),
                        company_name=company.name,
                        cached_documents=downloaded_annual_reports,
                    )
                    generation_warnings.extend(annual_report_pdf_bundle.warnings)
                st.session_state["generated_workbook"] = generated
                st.session_state["annual_report_pdf_bundle"] = annual_report_pdf_bundle
                st.session_state["generation_warnings"] = generation_warnings
            except (RikError, ValueError, OSError) as exc:
                st.error(f"The workbook could not be created: {exc}")

        structured = st.session_state.get("structured_result")
        if structured is not None:
            for warning in structured.warnings:
                st.warning(warning)
            if structured.lines:
                with st.expander("Structured statement preview", expanded=True):
                    st.dataframe(
                        statement_rows(structured.lines),
                        hide_index=True,
                        use_container_width=True,
                        height=420,
                    )

        generated = st.session_state.get("generated_workbook")
        if generated is not None:
            for warning in st.session_state.get("generation_warnings", []):
                st.warning(warning)
            st.success("Your filled Excel workbook is ready.")
            with st.expander("Generation details", expanded=True):
                for detail in generated.source_details:
                    st.markdown(f"- {detail}")
            excel_download_col, pdf_download_col = st.columns(2)
            excel_download_col.download_button(
                "Download filled Excel workbook",
                data=generated.content,
                file_name=generated.filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            annual_report_pdf_bundle = st.session_state.get("annual_report_pdf_bundle")
            if annual_report_pdf_bundle and annual_report_pdf_bundle.included_years:
                pdf_download_col.download_button(
                    "Download annual report PDFs (.zip)",
                    data=annual_report_pdf_bundle.content,
                    file_name=annual_report_pdf_bundle.filename,
                    mime="application/zip",
                    use_container_width=True,
                )
                st.caption(
                    "Complete RIK annual-report PDFs included for: "
                    + ", ".join(str(year) for year in annual_report_pdf_bundle.included_years)
                    + "."
                )

    if annual_documents:
        with st.expander("Source documents", expanded=False):
            st.dataframe(document_rows(annual_documents), hide_index=True, use_container_width=True)
            selected_document = st.selectbox(
                "Select a source document",
                options=annual_documents,
                format_func=lambda item: (
                    f"FY{item.fiscal_year or '—'} · "
                    f"{document_format(item.document_type)} "
                    f"· {item.document_name}"
                ),
            )
            if st.button("Fetch selected source document"):
                try:
                    with st.spinner("Downloading the source document from RIK..."):
                        downloaded = make_client().download_document(selected_document)
                    st.session_state["downloaded_document"] = downloaded
                except RikError as exc:
                    st.error(str(exc))
            downloaded = st.session_state.get("downloaded_document")
            if downloaded is not None:
                st.download_button(
                    "Download source document",
                    data=downloaded.content,
                    file_name=downloaded.filename,
                    mime=downloaded.content_type,
                )

st.caption("made by Bronek xoxo")
st.caption("Questions? Message me on Slack.")
