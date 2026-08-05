"""Parse RIK annual-report XBRL packages into the Estonia extractor model."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

from estonia_extractor import (
    ACTIVITY_SEGMENT,
    GEOGRAPHY_SEGMENT,
    EstonianReport,
    geography_label,
    load_terms,
    normalize,
)


class XbrlParseError(ValueError):
    """Raised when a RIK XBRL download is missing or malformed."""


@dataclass(frozen=True)
class XbrlContext:
    context_id: str
    start: date | None
    end: date | None
    instant: date | None
    dimensions: tuple[tuple[str, str], ...]

    @property
    def year(self) -> int | None:
        reporting_date = self.instant or self.end
        return reporting_date.year if reporting_date else None

    @property
    def is_duration(self) -> bool:
        return self.start is not None and self.end is not None

    @property
    def is_instant(self) -> bool:
        return self.instant is not None


@dataclass(frozen=True)
class XbrlFact:
    concept: str
    context: XbrlContext
    value: Decimal
    unit: str | None
    order: int


CONSOLIDATED_CONCEPTS: dict[str, tuple[str, ...]] = {
    "Revenue": ("RevenueConsolidated",),
    "Other income": ("OtherIncomeConsolidated", "OtherOperatingIncomeConsolidated"),
    "COGS": ("CostOfSalesConsolidated", "RawMaterialsAndConsumablesUsedConsolidated", "GoodsRawMaterialsAndServicesConsolidated"),
    "Reported EBIT": ("TotalProfitLossConsolidated",),
    "D&A": ("DepreciationAndImpairmentLossReversalConsolidated",),
    "Fixed assets": ("NonCurrentAssetsConsolidated",),
    "Current assets": ("CurrentAssetsConsolidated",),
    "Stocks / inventories": ("InventoriesConsolidated", "InventoriesTotalConsolidated"),
    "Trade debtors / receivables": ("AccountsReceivableConsolidated", "AccountsReceivablesConsolidated"),
    "Trade creditors / payables": ("TradePayablesTotalConsolidated",),
    "Cash and cash equivalents": ("CashAndCashEquivalentsConsolidated", "CashAndCashEquivalentsTotalConsolidated"),
    "Debt LT": ("LongTermLoanLiabilitiesConsolidated",),
    "Debt ST": ("ShortTermLoanLiabilitiesConsolidated",),
    "CAPEX": ("InvestingActivitiesPurchaseOfPropertyPlantAndEquipmentAndIntangibleAssetsConsolidated",),
    "Investments in tangible assets": ("PropertyPlantAndEquipmentAcquisitionsAndAdditionsConsolidated",),
    "Investments in intangible assets": ("IntangibleAssetsAcquisitionsAndAdditionsConsolidated",),
    "FTEs": ("AverageNumberOfEmployeesInFullTimeEquivalentUnitsConsolidated",),
}

STANDALONE_CONCEPTS: dict[str, tuple[str, ...]] = {
    "Revenue": ("Revenue",),
    "Other income": ("OtherIncome", "OtherOperatingIncome"),
    "COGS": ("CostOfSales", "RawMaterialsAndConsumablesUsed", "GoodsRawMaterialsAndServices"),
    "Reported EBIT": ("TotalProfitLoss",),
    "D&A": ("DepreciationAndImpairmentLossReversal",),
    "Fixed assets": ("NonCurrentAssets",),
    "Current assets": ("CurrentAssets",),
    "Stocks / inventories": ("Inventories", "InventoriesTotal"),
    "Trade debtors / receivables": ("AccountsReceivable", "AccountsReceivables"),
    "Trade creditors / payables": ("TradePayablesTotal",),
    "Cash and cash equivalents": ("CashAndCashEquivalents", "CashAndCashEquivalentsTotal"),
    "Debt LT": ("LongTermLoanLiabilities",),
    "Debt ST": ("ShortTermLoanLiabilities",),
    "CAPEX": ("InvestingActivitiesPurchaseOfPropertyPlantAndEquipmentAndIntangibleAssets",),
    "Investments in tangible assets": ("PropertyPlantAndEquipmentAcquisitionsAndAdditions",),
    "Investments in intangible assets": ("IntangibleAssetsAcquisitionsAndAdditions",),
    "FTEs": ("AverageNumberOfEmployeesInFullTimeEquivalentUnits",),
}

INSTANT_ITEMS = {
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
}

DIMENSIONAL_TOTAL_ITEMS = {"Trade debtors / receivables", "Trade creditors / payables"}
ANCHOR_ITEMS = ("Revenue", "Reported EBIT", "Fixed assets", "Current assets")
ACTIVITY_TRANSLATIONS = {
    "enda kinnisvara müük": "Sale of own real estate",
    "inseneeritööd": "Engineering services",
    "jaekaubandus": "Retail trade",
    "kauba müük": "Sales of goods",
    "keemiatoodete müük": "Sales of chemical products",
    "laevaehitus": "Shipbuilding",
    "laevandus": "Shipping",
    "laevaremont": "Ship repair",
    "metalli müük": "Metal sales",
    "metallikonstruktsioonid": "Metal structures",
    "muu ostukauba müük": "Other purchased goods sales",
    "muud teenused": "Other services",
    "programmeerimine": "Programming",
    "päikesepaneelide müük paigaldamine ja lahenduste pakkumine": "Solar panel sales, installation and solutions",
    "päikesepaneelide paigaldamine ja lahenduste pakkumine": "Solar panel installation and solutions",
    "renditulu": "Rental income",
    "rent": "Rental",
    "sadamateenused": "Port services",
    "stivideerimise teenused": "Stevedoring services",
    "hulgikaubandus": "Wholesale trade",
    "toitlustus": "Catering",
    "transpordi ja lõikamisteenus": "Transport and cutting services",
    "transporditeenused": "Transport services",
    "tsinkimine": "Galvanizing",
    "vanaraua müük": "Scrap metal sales",
    "üüritulu": "Rental income",
}
D_AND_A_COMPONENTS = {
    False: (
        "CostOfGoodsSoldDepreciation",
        "DistributionExpenseDepreciation",
        "AdministrativeExpenseDepreciation",
    ),
    True: (
        "CostOfGoodsSoldDepreciationConsolidated",
        "DistributionExpenseDepreciationConsolidated",
        "AdministrativeExpenseDepreciationConsolidated",
    ),
}


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _parse_date(value: str | None) -> date | None:
    try:
        return date.fromisoformat((value or "").strip()[:10])
    except ValueError:
        return None


def _decimal(value: str | None) -> Decimal | None:
    cleaned = (value or "").strip().replace("\u00a0", "").replace(" ", "")
    if not cleaned:
        return None
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def _instance_bytes(content: bytes) -> bytes:
    if not content:
        raise XbrlParseError("RIK returned an empty XBRL document.")
    if content.startswith(b"PK"):
        try:
            with ZipFile(BytesIO(content)) as archive:
                candidates = [
                    name for name in archive.namelist() if name.casefold().endswith(".xbrl")
                ]
                if not candidates:
                    raise XbrlParseError("The RIK XBRL package did not contain an instance file.")
                return archive.read(sorted(candidates)[0])
        except BadZipFile as exc:
            raise XbrlParseError("RIK returned a damaged XBRL package.") from exc
    stripped = content.lstrip()
    if stripped.startswith(b"<") and b"xbrl" in stripped[:4000].lower():
        return content
    if stripped.startswith(b"<!DOCTYPE html") or b"<html" in stripped[:1000].lower():
        raise XbrlParseError("RIK temporarily returned a web page instead of the XBRL package.")
    raise XbrlParseError("RIK did not return a recognizable XBRL package.")


def _contexts(root: ET.Element) -> dict[str, XbrlContext]:
    contexts: dict[str, XbrlContext] = {}
    for element in root.iter():
        if _local_name(element.tag) != "context":
            continue
        context_id = element.attrib.get("id", "")
        start = end = instant = None
        dimensions: list[tuple[str, str]] = []
        for child in element.iter():
            child_name = _local_name(child.tag)
            if child_name == "startDate":
                start = _parse_date(child.text)
            elif child_name == "endDate":
                end = _parse_date(child.text)
            elif child_name == "instant":
                instant = _parse_date(child.text)
            elif child_name in {"explicitMember", "typedMember"}:
                dimensions.append(
                    (
                        child.attrib.get("dimension", ""),
                        " ".join("".join(child.itertext()).split()),
                    )
                )
        if context_id:
            contexts[context_id] = XbrlContext(
                context_id,
                start,
                end,
                instant,
                tuple(dimensions),
            )
    return contexts


def _facts(root: ET.Element, contexts: dict[str, XbrlContext]) -> list[XbrlFact]:
    facts: list[XbrlFact] = []
    for order, element in enumerate(root.iter()):
        context_id = element.attrib.get("contextRef")
        if context_id not in contexts:
            continue
        value = _decimal(element.text)
        if value is None:
            continue
        facts.append(
            XbrlFact(
                _local_name(element.tag),
                contexts[context_id],
                value,
                element.attrib.get("unitRef"),
                order,
            )
        )
    return facts


def _has_consolidated_scope(facts: list[XbrlFact], report_year: int) -> bool:
    available = {fact.concept for fact in facts if fact.context.year == report_year}
    anchors = sum(
        bool(set(CONSOLIDATED_CONCEPTS[item]) & available) for item in ANCHOR_ITEMS
    )
    return anchors >= 3


def _dimension_score(fact: XbrlFact, item: str) -> tuple[int, Decimal, int]:
    dimension_text = " ".join(value for _dimension, value in fact.context.dimensions).casefold()
    if item in DIMENSIONAL_TOTAL_ITEMS:
        if "totalabstract" in dimension_text:
            rank = 0
        elif not fact.context.dimensions:
            rank = 1
        else:
            rank = 2
    else:
        rank = 0 if not fact.context.dimensions else 2
    return rank, -abs(fact.value), fact.order


def _select_fact(
    facts: list[XbrlFact],
    concepts: tuple[str, ...],
    item: str,
    year: int,
) -> XbrlFact | None:
    for concept in concepts:
        candidates = [
            fact
            for fact in facts
            if fact.concept == concept
            and fact.context.year == year
            and (fact.context.is_instant if item in INSTANT_ITEMS else fact.context.is_duration)
        ]
        if candidates:
            return sorted(candidates, key=lambda fact: _dimension_score(fact, item))[0]
    return None


def _source(fact: XbrlFact, report_year: int) -> str:
    period = fact.context.instant or fact.context.end
    dimension = ""
    if fact.context.dimensions:
        members = ", ".join(value for _name, value in fact.context.dimensions)
        dimension = f" | dimension {members}"
    return (
        f"RIK XBRL | AR{report_year} | concept {fact.concept} | "
        f"context {fact.context.context_id} | period {period.isoformat() if period else 'unknown'}"
        f"{dimension}"
    )


def _period_for_report(
    facts: list[XbrlFact], concepts: dict[str, tuple[str, ...]], report_year: int
) -> tuple[date | None, date | None]:
    revenue = _select_fact(facts, concepts["Revenue"], "Revenue", report_year)
    if revenue is not None:
        return revenue.context.start, revenue.context.end
    candidates = [
        fact.context
        for fact in facts
        if fact.context.year == report_year
        and fact.context.is_duration
        and not fact.context.dimensions
    ]
    if not candidates:
        return None, date(report_year, 12, 31)
    candidates.sort(
        key=lambda context: (
            -((context.end - context.start).days if context.start and context.end else 0),
            context.context_id,
        )
    )
    return candidates[0].start, candidates[0].end


def _goodwill_amortisation(
    facts: list[XbrlFact], consolidated: bool, year: int
) -> XbrlFact | None:
    concept = "IntangibleAssetsDepreciationConsolidated" if consolidated else "IntangibleAssetsDepreciation"
    candidates = []
    for fact in facts:
        dimension_text = " ".join(value for _name, value in fact.context.dimensions).casefold()
        if (
            fact.concept == concept
            and fact.context.year == year
            and "goodwill" in dimension_text
        ):
            candidates.append(fact)
    return max(candidates, key=lambda fact: (abs(fact.value), -fact.order), default=None)


def _derived_da(
    facts: list[XbrlFact], consolidated: bool, year: int, report_year: int
) -> tuple[Decimal, str] | None:
    components: list[XbrlFact] = []
    for concept in D_AND_A_COMPONENTS[consolidated]:
        fact = _select_fact(facts, (concept,), "D&A", year)
        if fact is not None:
            components.append(fact)
    if not components:
        return None
    value = sum((fact.value for fact in components), Decimal("0"))
    source = (
        f"RIK XBRL | AR{report_year} | derived D&A from "
        + ", ".join(f"{fact.concept} ({fact.context.context_id})" for fact in components)
    )
    return value, source


def _container_matches_scope(name: str, consolidated: bool) -> bool:
    has_suffix = "Consolidated" in name
    return has_suffix if consolidated else not has_suffix


def _segment_label(container: ET.Element, contexts: dict[str, XbrlContext], report_year: int) -> str | None:
    labels: list[tuple[int, str]] = []
    for child in list(container):
        context = contexts.get(child.attrib.get("contextRef", ""))
        text = " ".join("".join(child.itertext()).split())
        if not text or _decimal(text) is not None:
            continue
        rank = 0 if context and context.year == report_year else 1
        labels.append((rank, text))
    return sorted(labels)[0][1] if labels else None


def _add_segments(
    report: EstonianReport,
    root: ET.Element,
    contexts: dict[str, XbrlContext],
    report_year: int,
    consolidated: bool,
) -> None:
    activity_order = geography_order = 0
    for container in root.iter():
        children = list(container)
        if not children:
            continue
        name = _local_name(container.tag)
        is_activity = "NetSalesByOperatingActivities" in name
        is_geography = "NetSalesByGeographicalLocation" in name
        if not (is_activity or is_geography) or not _container_matches_scope(name, consolidated):
            continue
        label = _segment_label(container, contexts, report_year)
        if not label:
            continue
        if is_activity:
            activity_order += 1
            segment_by = ACTIVITY_SEGMENT
            group = "Activity"
            translated = ACTIVITY_TRANSLATIONS.get(normalize(label), label)
            segment_label = translated
            is_rest = False
            order = activity_order
        else:
            geography_order += 1
            segment_by = GEOGRAPHY_SEGMENT
            group = "EU" if "InEuropeanUnion" in name else "WORLD"
            normalized_label = normalize(label)
            if normalized_label in {"muud", "other"}:
                segment_label = "Rest of EU" if group == "EU" else "Rest of the world"
                is_rest = True
            else:
                parsed_label, parsed_group, parsed_rest = geography_label(label, group)
                if parsed_label is None:
                    continue
                segment_label = parsed_label
                group = parsed_group or group
                is_rest = bool(parsed_rest)
            order = geography_order

        for child in children:
            context = contexts.get(child.attrib.get("contextRef", ""))
            value = _decimal(child.text)
            if context is None or context.year is None or value is None:
                continue
            concept = _local_name(child.tag)
            report.set_segment_value(
                context.year,
                segment_by,
                segment_label,
                value,
                (
                    f"RIK XBRL | AR{report_year} | concept {concept} | "
                    f"context {context.context_id} | segment {label}"
                ),
                group,
                order,
                is_rest=is_rest,
            )


def parse_xbrl_report(
    content: bytes,
    *,
    registry_code: str,
    company_name: str,
    report_year: int,
    source_name: str = "RIK annual report XBRL",
) -> EstonianReport:
    """Return current and comparative values from a RIK XBRL package."""
    instance = _instance_bytes(content)
    try:
        root = ET.fromstring(instance)
    except ET.ParseError as exc:
        raise XbrlParseError("The RIK XBRL instance was malformed.") from exc

    contexts = _contexts(root)
    facts = _facts(root, contexts)
    if not contexts or not facts:
        raise XbrlParseError("The RIK XBRL instance did not contain usable facts.")
    consolidated = _has_consolidated_scope(facts, report_year)
    concepts = CONSOLIDATED_CONCEPTS if consolidated else STANDALONE_CONCEPTS
    period_start, period_end = _period_for_report(facts, concepts, report_year)
    report = EstonianReport(
        source_path=Path(source_name),
        payload_name=f"RIK {'consolidated ' if consolidated else ''}XBRL annual report FY{report_year}",
        lines=[],
        terms=load_terms(None),
        period_start=period_start,
        period_end=period_end,
        company=company_name,
        accounting_basis="consolidated" if consolidated else "reported",
    )
    years = sorted(
        {
            context.year
            for context in contexts.values()
            if context.year is not None and context.year in {report_year, report_year - 1}
        },
        reverse=True,
    )
    for year in years:
        for item, item_concepts in concepts.items():
            fact = _select_fact(facts, item_concepts, item, year)
            if fact is not None:
                report.set_value(year, item, fact.value, _source(fact, report_year))
        if report.get_value(year, "D&A") is None:
            derived_da = _derived_da(facts, consolidated, year, report_year)
            if derived_da is not None:
                report.set_value(year, "D&A", derived_da[0], derived_da[1])
        goodwill = _goodwill_amortisation(facts, consolidated, year)
        if goodwill is not None:
            report.set_value(
                year,
                "Goodwill amortisation",
                goodwill.value,
                _source(goodwill, report_year),
            )
    _add_segments(report, root, contexts, report_year, consolidated)
    if sum(report.get_value(report_year, item) is not None for item in ANCHOR_ITEMS) < 3:
        raise XbrlParseError("The RIK XBRL instance lacked a complete primary financial statement block.")
    return report
