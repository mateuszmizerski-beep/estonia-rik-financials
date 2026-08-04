"""Small, explicit client for RIK's e-Business Register SOAP/XML service."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
import mimetypes
from pathlib import Path
import threading
from typing import Iterable
from urllib.parse import urlparse
import xml.etree.ElementTree as ET

import requests


RIK_ENDPOINT = "https://ariregxmlv6.rik.ee/"
SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
PRODUCER_NS = "http://arireg.x-road.eu/producer/"
MAX_DOCUMENT_BYTES = 25 * 1024 * 1024
_QUERY_LOCK = threading.Lock()


class RikError(RuntimeError):
    """Base error raised for readable RIK API failures."""


class RikAuthenticationError(RikError):
    """Raised when RIK rejects the configured credentials."""


class RikResponseError(RikError):
    """Raised when RIK returns an invalid or unexpected response."""


@dataclass(frozen=True)
class CompanySummary:
    registry_code: str
    name: str
    legal_form: str | None = None
    status: str | None = None
    address: str | None = None


@dataclass(frozen=True)
class AnnualReportAvailability:
    report_type: str
    report_name: str
    fiscal_year: int
    period_start: date | None
    period_end: date | None


@dataclass(frozen=True)
class StatementLine:
    registry_code: str
    fiscal_year: int
    report_type: str
    report_name: str
    period_start: date | None
    period_end: date | None
    row_number: str
    row_name: str
    column_code: str
    column_name: str
    value: Decimal | None
    retrieved_at: datetime


@dataclass(frozen=True)
class CompanyDocument:
    document_id: str
    registry_code: str
    document_type: str
    document_name: str
    size_bytes: int | None
    status_date: date | None
    validity: str | None
    report_kind: str | None
    fiscal_year: int | None
    url: str
    retrieved_at: datetime

    @property
    def extension(self) -> str:
        return {"A": ".pdf", "D": ".bdoc", "X": ".xml", "P": ".pdf"}.get(
            self.document_type.upper(), ""
        )


@dataclass(frozen=True)
class DownloadedDocument:
    content: bytes
    filename: str
    content_type: str


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(element: ET.Element, name: str) -> list[ET.Element]:
    return [child for child in list(element) if _local_name(child.tag) == name]


def _first(element: ET.Element, name: str) -> ET.Element | None:
    for candidate in element.iter():
        if _local_name(candidate.tag) == name:
            return candidate
    return None


def _text(element: ET.Element, name: str, default: str = "") -> str:
    found = _first(element, name)
    return (found.text or "").strip() if found is not None else default


def _parse_date(value: str) -> date | None:
    cleaned = value.strip().rstrip("Z")
    if not cleaned:
        return None
    for pattern in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(cleaned[:10], pattern).date()
        except ValueError:
            continue
    return None


def _parse_int(value: str) -> int | None:
    try:
        return int(value.strip())
    except (AttributeError, TypeError, ValueError):
        return None


def _parse_decimal(value: str) -> Decimal | None:
    cleaned = value.strip().replace(" ", "").replace("\u00a0", "").replace(",", ".")
    if not cleaned:
        return None
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


class RikXmlClient:
    """RIK SOAP client with no credential or raw-response logging."""

    def __init__(
        self,
        username: str,
        password: str,
        *,
        endpoint: str = RIK_ENDPOINT,
        timeout: tuple[float, float] = (10.0, 45.0),
        session: requests.Session | None = None,
    ) -> None:
        if not username or not password:
            raise ValueError("RIK username and password are required.")
        self._username = username
        self._password = password
        self.endpoint = endpoint
        self.timeout = timeout
        self.session = session or requests.Session()
        self.session.headers.update(
            {
                "Accept": "text/xml, application/xml",
                "Content-Type": "text/xml; charset=utf-8",
                "User-Agent": "Gain-Estonia-Financials/1.0",
            }
        )

    def _envelope(self, operation: str, parameters: Iterable[tuple[str, object]]) -> bytes:
        envelope = ET.Element(f"{{{SOAP_NS}}}Envelope")
        body = ET.SubElement(envelope, f"{{{SOAP_NS}}}Body")
        request = ET.SubElement(body, f"{{{PRODUCER_NS}}}{operation}")
        request_body = ET.SubElement(request, f"{{{PRODUCER_NS}}}keha")
        auth_parameters: list[tuple[str, object]] = [
            ("ariregister_kasutajanimi", self._username),
            ("ariregister_parool", self._password),
        ]
        for name, value in [*auth_parameters, *parameters]:
            if value is None or value == "":
                continue
            node = ET.SubElement(request_body, f"{{{PRODUCER_NS}}}{name}")
            node.text = str(value)
        return ET.tostring(envelope, encoding="utf-8", xml_declaration=True)

    def _post(self, operation: str, parameters: Iterable[tuple[str, object]]) -> ET.Element:
        payload = self._envelope(operation, parameters)
        try:
            with _QUERY_LOCK:
                response = self.session.post(
                    self.endpoint,
                    data=payload,
                    headers={"SOAPAction": ""},
                    timeout=self.timeout,
                )
        except requests.RequestException as exc:
            raise RikError("RIK could not be reached. Please try again.") from exc

        if response.status_code in {401, 403}:
            raise RikAuthenticationError("RIK rejected the configured credentials.")
        if response.status_code >= 400:
            raise RikError(f"RIK returned HTTP {response.status_code}.")
        if len(response.content) > 15 * 1024 * 1024:
            raise RikResponseError("RIK returned an unexpectedly large XML response.")

        try:
            root = ET.fromstring(response.content)
        except ET.ParseError as exc:
            raise RikResponseError("RIK returned malformed XML.") from exc

        fault = _first(root, "Fault")
        if fault is not None:
            fault_text = _text(fault, "faultstring") or _text(fault, "Text") or "RIK rejected the query."
            lowered = fault_text.casefold()
            if any(token in lowered for token in ("parool", "password", "kasutaja", "auth")):
                raise RikAuthenticationError("RIK rejected the configured credentials.")
            raise RikError(fault_text[:300])

        response_body = _first(root, "Body")
        keha = _first(response_body, "keha") if response_body is not None else None
        if keha is None:
            raise RikResponseError("RIK response did not contain the expected response body.")
        return keha

    def get_company_summary(self, registry_code: str) -> CompanySummary:
        keha = self._post(
            "lihtandmed_v2",
            (("ariregistri_kood", registry_code), ("keel", "eng")),
        )
        companies = _first(keha, "ettevotjad")
        item = _first(companies, "item") if companies is not None else None
        if item is None:
            raise RikResponseError(f"No company was found for registry code {registry_code}.")
        return CompanySummary(
            registry_code=_text(item, "ariregistri_kood") or registry_code,
            name=_text(item, "evnimi") or registry_code,
            legal_form=_text(item, "oiguslik_vorm_tekstina") or None,
            status=_text(item, "staatus_tekstina") or None,
            address=_text(item, "aadress_ads__ads_normaliseeritud_taisaadress") or None,
        )

    def list_annual_reports(self, registry_code: str) -> list[AnnualReportAvailability]:
        keha = self._post(
            "majandusaastaAruanneteLoetelu_v1",
            (("ariregistri_kood", registry_code),),
        )
        reports: list[AnnualReportAvailability] = []
        for node in keha.iter():
            if _local_name(node.tag) != "majandusaasta_aruanded":
                continue
            year = _parse_int(_text(node, "aruande_aasta"))
            if year is None:
                continue
            reports.append(
                AnnualReportAvailability(
                    report_type=_text(node, "aruande_kood"),
                    report_name=_text(node, "aruande_nimetus"),
                    fiscal_year=year,
                    period_start=_parse_date(_text(node, "majandusaasta_algus")),
                    period_end=_parse_date(_text(node, "majandusaasta_lopp")),
                )
            )
        return sorted(reports, key=lambda item: (item.fiscal_year, item.report_type), reverse=True)

    def get_annual_report_lines(
        self,
        registry_code: str,
        report_type: str,
        year: int,
        language: str = "eng",
    ) -> list[StatementLine]:
        if language not in {"eng", "est"}:
            raise ValueError("language must be 'eng' or 'est'.")
        keha = self._post(
            "majandusaastaAruanneteKirjed_v1",
            (
                ("ariregistri_kood", registry_code),
                ("aruande_liik", report_type),
                ("aruandeaasta", year),
                ("keel", language),
            ),
        )
        response_type = _text(keha, "aruande_liik") or str(report_type)
        report_name = _text(keha, "aruande_nimetus")
        period_start = _parse_date(_text(keha, "majandusaasta_algus"))
        period_end = _parse_date(_text(keha, "majandusaasta_lopp"))
        lines: list[StatementLine] = []
        retrieved_at = datetime.now(timezone.utc)
        for row in keha.iter():
            if _local_name(row.tag) != "majandusaasta_aruanded_read":
                continue
            row_number = _text(row, "rea_nr")
            row_name = _text(row, "rea_nimetus")
            columns = _children(row, "majandusaasta_aruanded_veerud")
            for column in columns:
                lines.append(
                    StatementLine(
                        registry_code=registry_code,
                        fiscal_year=year,
                        report_type=response_type,
                        report_name=report_name,
                        period_start=period_start,
                        period_end=period_end,
                        row_number=row_number,
                        row_name=row_name,
                        column_code=_text(column, "veeru_kood"),
                        column_name=_text(column, "veeru_nimetus"),
                        value=_parse_decimal(_text(column, "vaartus")),
                        retrieved_at=retrieved_at,
                    )
                )
        return lines

    def list_company_documents(
        self, registry_code: str, year: int | None = None
    ) -> list[CompanyDocument]:
        keha = self._post(
            "ettevotjaDokumentideLoetelu_v1",
            (
                ("ariregistri_kood", registry_code),
                ("aruandeaasta", year),
                ("keel", "eng"),
            ),
        )
        documents: list[CompanyDocument] = []
        retrieved_at = datetime.now(timezone.utc)
        for node in keha.iter():
            if _local_name(node.tag) != "ettevotja_dokumendid":
                continue
            url = _text(node, "dokumendi_url")
            if not url:
                continue
            documents.append(
                CompanyDocument(
                    document_id=_text(node, "dokumendi_id"),
                    registry_code=_text(node, "ariregistri_kood") or registry_code,
                    document_type=_text(node, "dokumendi_liik"),
                    document_name=_text(node, "dokumendi_nimetus"),
                    size_bytes=_parse_int(_text(node, "dokumendi_suurus")),
                    status_date=_parse_date(_text(node, "dokumendi_seisu_kuupaev")),
                    validity=_text(node, "dokumendi_kehtivus") or None,
                    report_kind=_text(node, "aruande_liik") or None,
                    fiscal_year=_parse_int(_text(node, "aruandeaasta")),
                    url=url,
                    retrieved_at=retrieved_at,
                )
            )
        return sorted(
            documents,
            key=lambda item: (item.fiscal_year or 0, item.document_type, item.document_id),
            reverse=True,
        )

    def download_document(self, document: CompanyDocument) -> DownloadedDocument:
        if document.size_bytes and document.size_bytes > MAX_DOCUMENT_BYTES:
            raise RikError("The selected source document is larger than 25 MB.")
        try:
            with _QUERY_LOCK:
                response = self.session.get(document.url, timeout=self.timeout, stream=True)
                response.raise_for_status()
                chunks: list[bytes] = []
                total = 0
                for chunk in response.iter_content(chunk_size=128 * 1024):
                    if not chunk:
                        continue
                    total += len(chunk)
                    if total > MAX_DOCUMENT_BYTES:
                        raise RikError("The selected source document is larger than 25 MB.")
                    chunks.append(chunk)
        except requests.RequestException as exc:
            raise RikError("The source document could not be downloaded from RIK.") from exc

        content = b"".join(chunks)
        content_type = response.headers.get("Content-Type", "application/octet-stream").split(";", 1)[0]
        suffix = document.extension or mimetypes.guess_extension(content_type) or ""
        path_name = Path(urlparse(document.url).path).name
        if path_name and "." in path_name:
            filename = path_name
        else:
            year = document.fiscal_year or "source"
            filename = f"{document.registry_code}_{year}_{document.document_type}{suffix}"
        return DownloadedDocument(content=content, filename=filename, content_type=content_type)
