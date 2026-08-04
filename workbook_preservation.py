"""Preserve Excel extension features that openpyxl cannot round-trip."""

from __future__ import annotations

import os
from pathlib import Path
import re
import tempfile
import zipfile


WORKSHEET_PART = "xl/worksheets/sheet1.xml"
XR_NAMESPACE = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"


def _extract_ext_list(xml: bytes) -> bytes | None:
    match = re.search(rb"<extLst>.*?</extLst>", xml, flags=re.DOTALL)
    return match.group(0) if match else None


def _inject_extended_validations(source_xml: bytes, output_xml: bytes) -> bytes:
    source_ext = _extract_ext_list(source_xml)
    if source_ext is None or b"dataValidations" not in source_ext:
        return output_xml

    output_ext = _extract_ext_list(output_xml)
    if output_ext is not None:
        source_extensions = re.findall(rb"<ext\b.*?</ext>", source_ext, flags=re.DOTALL)
        merged = output_ext.replace(b"</extLst>", b"".join(source_extensions) + b"</extLst>")
        output_xml = output_xml.replace(output_ext, merged, 1)
    else:
        output_xml = output_xml.replace(b"</worksheet>", source_ext + b"</worksheet>", 1)

    root_start = output_xml.find(b"<worksheet")
    if root_start < 0:
        raise ValueError("Generated workbook worksheet XML is invalid.")
    root_end = output_xml.find(b">", root_start)
    root_tag = output_xml[root_start : root_end + 1]
    if b"xmlns:xr=" not in root_tag:
        replacement = root_tag[:-1] + f' xmlns:xr="{XR_NAMESPACE}">'.encode("utf-8")
        output_xml = output_xml[:root_start] + replacement + output_xml[root_end + 1 :]
    return output_xml


def restore_extended_validations(template_path: Path, output_path: Path) -> None:
    """Copy the template's x14 data validations back into a generated workbook."""
    template_path = Path(template_path)
    output_path = Path(output_path)
    with zipfile.ZipFile(template_path, "r") as source_zip:
        source_xml = source_zip.read(WORKSHEET_PART)
    with zipfile.ZipFile(output_path, "r") as output_zip:
        output_xml = output_zip.read(WORKSHEET_PART)
        replacement_xml = _inject_extended_validations(source_xml, output_xml)
        if replacement_xml == output_xml:
            return
        entries = [(item, output_zip.read(item.filename)) for item in output_zip.infolist()]

    handle, temporary_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}-", suffix=".xlsx", dir=output_path.parent
    )
    os.close(handle)
    temporary_path = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary_path, "w") as target_zip:
            for item, data in entries:
                target_zip.writestr(
                    item,
                    replacement_xml if item.filename == WORKSHEET_PART else data,
                )
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def count_extended_validations(workbook_path: Path) -> int:
    with zipfile.ZipFile(workbook_path, "r") as workbook_zip:
        xml = workbook_zip.read(WORKSHEET_PART)
    return len(re.findall(rb"<x14:dataValidation\b", xml))
