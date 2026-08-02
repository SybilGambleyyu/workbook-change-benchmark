"""Validation for WCAB's observable fixture truth contract."""

from __future__ import annotations

import base64
import binascii
import io
import json
import posixpath
import re
import struct
from collections import deque
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from .build import CASE_IDS, FIXTURE_SCHEMA_VERSION

_CELL_REFERENCE = re.compile(
    r"(?:(?:'(?P<quoted>(?:[^']|'')+)'|(?P<plain>[A-Za-z_][A-Za-z0-9_. ]*))!)?"
    r"(?P<cell>\$?[A-Z]{1,3}\$?[1-9][0-9]*)"
)
_DYNAMIC_REFERENCE_FUNCTION = re.compile(r"(?i)\b(?P<function>INDIRECT|OFFSET)\s*\(")
_SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOCUMENT_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_OFFICE_2010_SPREADSHEET_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_SLICER_CACHE_RELATIONSHIP = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
_DATA_MASHUP_NS = "http://schemas.microsoft.com/DataMashup"
_POWER_QUERY_CUSTOM_XML_RELATIONSHIP = f"{_DOCUMENT_RELATIONSHIPS_NS}/customXml"
_POWER_QUERY_STREAM_MAX_ENCODED_BYTES = 128 * 1024
_POWER_QUERY_PACKAGE_MAX_UNCOMPRESSED_BYTES = 128 * 1024
_DYNAMIC_ARRAY_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
_DRAWINGML_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_DRAWINGML_SPREADSHEET_DRAWING_NS = (
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
)
_POWER_QUERY_M_FILTER = re.compile(
    r"\Asection (?P<section>[A-Za-z_][A-Za-z0-9_]*);\n\n"
    r"shared (?P<query>[A-Za-z_][A-Za-z0-9_]*) = let\n"
    r'    Source = Excel\.CurrentWorkbook\(\)\{\[Name="(?P<source_table>[A-Za-z_][A-Za-z0-9_]*)"\]\}\[Content\],\n'
    r'    FilteredRows = Table\.SelectRows\(Source, each \[(?P<filter_column>[A-Za-z_][A-Za-z0-9_]*)\] = "(?P<filter_value>[A-Za-z_][A-Za-z0-9_]*)"\)\n'
    r"in\n"
    r"    FilteredRows;\n\Z"
)


class FixtureValidationError(ValueError):
    """One or more fixture assertions did not match the generated workbooks."""


def _load_truth(path: Path) -> dict[str, Any]:
    try:
        result = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise FixtureValidationError(f"{path}: cannot load truth manifest: {error}") from error
    if not isinstance(result, dict):
        raise FixtureValidationError(f"{path}: truth manifest must be a JSON object")
    return result


def _load_workbook(path: Path) -> Workbook:
    try:
        return load_workbook(path, data_only=False, keep_links=True)
    except Exception as error:  # Openpyxl has several parse-specific exception types.
        raise FixtureValidationError(f"{path}: cannot load workbook: {error}") from error


def _cell_kind(cell: Any) -> str:
    if cell.value is None:
        return "blank"
    if cell.data_type == "f":
        return "formula"
    return "value"


def _dynamic_reference_functions(formula: str) -> tuple[str, ...]:
    """Return dynamic-reference functions in first-appearance order.

    This is deliberately a narrow fixture validator, not a general formula
    parser.  It only establishes the direct, generated-workbook invariant that
    a declared dynamic-reference boundary is present in the candidate.
    """

    return tuple(
        dict.fromkeys(
            match.group("function").upper()
            for match in _DYNAMIC_REFERENCE_FUNCTION.finditer(formula)
        )
    )


def _defined_name_text(workbook: Workbook, name: str) -> str | None:
    definition = workbook.defined_names.get(name)
    return None if definition is None else definition.attr_text


def _external_data_connection_refresh_on_load(path: Path, connection_id: int) -> bool | None:
    """Read a relationship-backed connection's explicit refresh-on-open flag.

    This intentionally validates only the generated fixture's small raw-OOXML
    contract. It neither resolves a target nor loads a workbook through a
    library that might discard the connection part.
    """

    try:
        with ZipFile(path) as archive:
            relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
            relationship_exists = any(
                relationship.get("Type") == f"{_DOCUMENT_RELATIONSHIPS_NS}/connections"
                and relationship.get("Target") == "connections.xml"
                for relationship in relationships.findall(relationship_tag)
            )
            if not relationship_exists:
                return None
            connections = ElementTree.fromstring(archive.read("xl/connections.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    connection_tag = f"{{{_SPREADSHEETML_NS}}}connection"
    matches = [
        connection
        for connection in connections.findall(connection_tag)
        if connection.get("id") == str(connection_id)
    ]
    if len(matches) != 1:
        return None
    flag = matches[0].get("refreshOnLoad")
    return {"0": False, "1": True}.get(flag)


def _workbook_properties(path: Path) -> dict[str, str] | None:
    """Read the stored ``workbookPr`` attributes without loading link targets."""

    try:
        with ZipFile(path) as archive:
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    properties = workbook.find(f"{{{_SPREADSHEETML_NS}}}workbookPr")
    return None if properties is None else dict(properties.attrib)


def _calculation_properties(path: Path) -> dict[str, str] | None:
    """Read stored ``calcPr`` attributes without evaluating workbook formulas."""

    try:
        with ZipFile(path) as archive:
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    properties = workbook.find(f"{{{_SPREADSHEETML_NS}}}calcPr")
    return None if properties is None else dict(properties.attrib)


def _ooxml_boolean(value: str | None) -> bool | None:
    """Parse the strict boolean values used by WCAB's raw OOXML contracts."""

    return {"0": False, "1": True, "false": False, "true": True}.get(value)


def _calculation_iteration_state(path: Path) -> tuple[bool, int, Decimal] | None:
    """Read WCAB's explicit stored iteration switch and shared bounds."""

    properties = _calculation_properties(path)
    if properties is None:
        return None
    iterate = _ooxml_boolean(properties.get("iterate"))
    try:
        count = int(properties["iterateCount"])
        delta = Decimal(properties["iterateDelta"])
    except (KeyError, InvalidOperation, ValueError):
        return None
    return None if iterate is None else (iterate, count, delta)


def _calculation_full_precision(path: Path) -> bool | None:
    """Read the explicit stored full-precision calculation control."""

    properties = _calculation_properties(path)
    return None if properties is None else _ooxml_boolean(properties.get("fullPrecision"))


def _relationship_target(
    relationships: ElementTree.Element,
    relationship_id: str,
    *,
    relationship_type: str | None = None,
) -> str | None:
    """Return one generated-package relationship target without following it."""

    relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
    matches = [
        relationship
        for relationship in relationships.findall(relationship_tag)
        if relationship.get("Id") == relationship_id
        and (relationship_type is None or relationship.get("Type") == relationship_type)
    ]
    if len(matches) != 1:
        return None
    target = matches[0].get("Target")
    if not isinstance(target, str) or not target or target.startswith("../"):
        return None
    return target.lstrip("/")


def _relationship_part_member(
    relationships: ElementTree.Element,
    relationship_id: str,
    source_member: str,
    *,
    relationship_type: str,
) -> str | None:
    """Resolve one local relationship to a safe package-member name.

    PivotTable relationships conventionally use ``../`` targets, unlike the
    workbook-to-worksheet links handled by :func:`_relationship_target`. This
    helper follows only an in-package relative target and rejects external or
    escaping relationships before opening a part.
    """

    if not source_member.startswith("xl/"):
        return None
    relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
    matches = [
        relationship
        for relationship in relationships.findall(relationship_tag)
        if relationship.get("Id") == relationship_id
        and relationship.get("Type") == relationship_type
        and relationship.get("TargetMode") != "External"
    ]
    if len(matches) != 1:
        return None
    target = matches[0].get("Target")
    if not isinstance(target, str) or not target or target.startswith("//"):
        return None
    if target.startswith("/"):
        member = target.lstrip("/")
    else:
        member = posixpath.normpath(f"{posixpath.dirname(source_member)}/{target}")
    if (
        not member.startswith("xl/")
        or member.startswith("../")
        or any(segment in {"", ".", ".."} for segment in member.split("/"))
    ):
        return None
    return member


def _worksheet_member_for_sheet(archive: ZipFile, sheet_name: str) -> str | None:
    """Resolve one generated worksheet name to its local OOXML part.

    This follows only the workbook-local relationship needed by WCAB's raw
    fixture contracts. It rejects traversal targets and never opens a linked
    workbook or evaluates a formula.
    """

    try:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.find(f"{{{_SPREADSHEETML_NS}}}sheets")
        if sheets is None:
            return None
        sheet_tag = f"{{{_SPREADSHEETML_NS}}}sheet"
        matches = [sheet for sheet in sheets.findall(sheet_tag) if sheet.get("name") == sheet_name]
        if len(matches) != 1:
            return None
        relationship_id = matches[0].get(f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id")
        if not isinstance(relationship_id, str):
            return None
        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = _relationship_target(
            relationships,
            relationship_id,
            relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/worksheet",
        )
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    if target is None:
        return None
    return target if target.startswith("xl/") else f"xl/{target}"


def _raw_pivot_cache_refresh_state(path: Path, pivot_sheet: str) -> dict[str, Any] | None:
    """Read WCAB's compact relationship-backed PivotCache/PivotTable graph.

    This is intentionally a narrow reader for the generated fixture. It
    establishes the stored open-time control and the local cache-to-PivotTable
    binding without opening a client, refreshing data, or interpreting a
    PivotTable display value.
    """

    try:
        with ZipFile(path) as archive:
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            workbook_relationships = ElementTree.fromstring(
                archive.read("xl/_rels/workbook.xml.rels")
            )
            pivot_caches_tag = f"{{{_SPREADSHEETML_NS}}}pivotCaches"
            pivot_cache_tag = f"{{{_SPREADSHEETML_NS}}}pivotCache"
            pivot_cache_sets = workbook.findall(pivot_caches_tag)
            if len(pivot_cache_sets) != 1:
                return None
            pivot_caches = pivot_cache_sets[0]
            if len(pivot_caches) != 1 or pivot_caches[0].tag != pivot_cache_tag:
                return None
            pivot_cache = pivot_caches[0]
            try:
                cache_id = int(pivot_cache.get("cacheId", ""))
            except ValueError:
                return None
            cache_relationship_id = pivot_cache.get(f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id")
            if not isinstance(cache_relationship_id, str):
                return None
            definition_member = _relationship_part_member(
                workbook_relationships,
                cache_relationship_id,
                "xl/workbook.xml",
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotCacheDefinition",
            )
            if definition_member is None:
                return None
            cache_definition = ElementTree.fromstring(archive.read(definition_member))
            cache_records_relationship_id = cache_definition.get(
                f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id"
            )
            if not isinstance(cache_records_relationship_id, str):
                return None
            definition_directory, definition_filename = definition_member.rsplit("/", maxsplit=1)
            cache_definition_relationships = ElementTree.fromstring(
                archive.read(f"{definition_directory}/_rels/{definition_filename}.rels")
            )
            cache_records_member = _relationship_part_member(
                cache_definition_relationships,
                cache_records_relationship_id,
                definition_member,
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotCacheRecords",
            )
            if cache_records_member is None:
                return None
            cache_records = ElementTree.fromstring(archive.read(cache_records_member))

            report_member = _worksheet_member_for_sheet(archive, pivot_sheet)
            if report_member is None:
                return None
            report = ElementTree.fromstring(archive.read(report_member))
            report_directory, report_filename = report_member.rsplit("/", maxsplit=1)
            report_relationships = ElementTree.fromstring(
                archive.read(f"{report_directory}/_rels/{report_filename}.rels")
            )
            pivot_table_parts_tag = f"{{{_SPREADSHEETML_NS}}}pivotTableParts"
            pivot_table_part_tag = f"{{{_SPREADSHEETML_NS}}}pivotTablePart"
            pivot_table_parts = report.findall(pivot_table_parts_tag)
            if len(pivot_table_parts) != 1 or len(pivot_table_parts[0]) != 1:
                return None
            pivot_table_part = pivot_table_parts[0][0]
            if pivot_table_part.tag != pivot_table_part_tag:
                return None
            pivot_table_relationship_id = pivot_table_part.get(
                f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id"
            )
            if not isinstance(pivot_table_relationship_id, str):
                return None
            pivot_table_member = _relationship_part_member(
                report_relationships,
                pivot_table_relationship_id,
                report_member,
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotTable",
            )
            if pivot_table_member is None:
                return None
            pivot_table = ElementTree.fromstring(archive.read(pivot_table_member))
            pivot_directory, pivot_filename = pivot_table_member.rsplit("/", maxsplit=1)
            pivot_relationships = ElementTree.fromstring(
                archive.read(f"{pivot_directory}/_rels/{pivot_filename}.rels")
            )
            bound_definition_member = _relationship_part_member(
                pivot_relationships,
                "rIdWCABPivotCache",
                pivot_table_member,
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotCacheDefinition",
            )
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None

    cache_definition_tag = f"{{{_SPREADSHEETML_NS}}}pivotCacheDefinition"
    cache_source_tag = f"{{{_SPREADSHEETML_NS}}}cacheSource"
    worksheet_source_tag = f"{{{_SPREADSHEETML_NS}}}worksheetSource"
    cache_records_tag = f"{{{_SPREADSHEETML_NS}}}pivotCacheRecords"
    record_tag = f"{{{_SPREADSHEETML_NS}}}r"
    shared_item_index_tag = f"{{{_SPREADSHEETML_NS}}}x"
    pivot_table_tag = f"{{{_SPREADSHEETML_NS}}}pivotTableDefinition"
    location_tag = f"{{{_SPREADSHEETML_NS}}}location"
    if (
        cache_definition.tag != cache_definition_tag
        or cache_records.tag != cache_records_tag
        or pivot_table.tag != pivot_table_tag
        or bound_definition_member != definition_member
    ):
        return None
    cache_sources = cache_definition.findall(cache_source_tag)
    if len(cache_sources) != 1 or len(cache_sources[0]) != 1:
        return None
    cache_source = cache_sources[0]
    worksheet_source = cache_source[0]
    if worksheet_source.tag != worksheet_source_tag:
        return None
    locations = pivot_table.findall(location_tag)
    if len(locations) != 1:
        return None
    records = cache_records.findall(record_tag)
    try:
        declared_record_count = int(cache_records.get("count", ""))
        record_indexes = tuple(
            tuple(int(value.get("v", "")) for value in record) for record in records
        )
    except (TypeError, ValueError):
        return None
    if declared_record_count != len(records) or any(
        len(record) != 2 or any(value.tag != shared_item_index_tag for value in record)
        for record in records
    ):
        return None
    return {
        "cache_id": cache_id,
        "definition_member": definition_member,
        "definition_attributes": tuple(sorted(cache_definition.attrib.items())),
        "cache_source_attributes": tuple(sorted(cache_source.attrib.items())),
        "worksheet_source_attributes": tuple(sorted(worksheet_source.attrib.items())),
        "cache_records_member": cache_records_member,
        "cache_records_attributes": tuple(sorted(cache_records.attrib.items())),
        "cache_record_indexes": record_indexes,
        "report_member": report_member,
        "pivot_table_member": pivot_table_member,
        "pivot_table_attributes": tuple(sorted(pivot_table.attrib.items())),
        "location_attributes": tuple(sorted(locations[0].attrib.items())),
    }


def _pivot_cache_definition_without_refresh_on_load(path: Path, pivot_sheet: str) -> bytes | None:
    """Return WCAB's cache definition with only its open-time flag removed."""

    state = _raw_pivot_cache_refresh_state(path, pivot_sheet)
    if state is None:
        return None
    try:
        with ZipFile(path) as archive:
            definition = ElementTree.fromstring(archive.read(state["definition_member"]))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    definition.attrib.pop("refreshOnLoad", None)
    return ElementTree.tostring(definition, encoding="utf-8", xml_declaration=True)


def _raw_pivot_data_field_aggregation_state(path: Path, pivot_sheet: str) -> dict[str, Any] | None:
    """Read the sole generated PivotTable value-field aggregation declaration.

    The shared PivotCache reader establishes the relationship-backed graph.
    This narrower layer then reads exactly one ``dataFields/dataField`` entry;
    it never recalculates, refreshes, or renders the PivotTable.
    """

    state = _raw_pivot_cache_refresh_state(path, pivot_sheet)
    if state is None:
        return None
    try:
        with ZipFile(path) as archive:
            pivot_table = ElementTree.fromstring(archive.read(state["pivot_table_member"]))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    data_fields_tag = f"{{{_SPREADSHEETML_NS}}}dataFields"
    data_field_tag = f"{{{_SPREADSHEETML_NS}}}dataField"
    data_fields = pivot_table.findall(data_fields_tag)
    if (
        len(data_fields) != 1
        or data_fields[0].get("count") != "1"
        or len(data_fields[0]) != 1
        or data_fields[0][0].tag != data_field_tag
    ):
        return None
    data_field = data_fields[0][0]
    try:
        source_index = int(data_field.get("fld", ""))
    except ValueError:
        return None
    subtotal = data_field.get("subtotal")
    if not isinstance(subtotal, str):
        return None
    return {
        **state,
        "data_field_source_index": source_index,
        "data_field_attributes": tuple(sorted(data_field.attrib.items())),
    }


def _pivot_table_without_data_field_subtotal(path: Path, pivot_sheet: str) -> bytes | None:
    """Return the generated PivotTable definition without its one subtotal attr."""

    state = _raw_pivot_data_field_aggregation_state(path, pivot_sheet)
    if state is None:
        return None
    try:
        with ZipFile(path) as archive:
            pivot_table = ElementTree.fromstring(archive.read(state["pivot_table_member"]))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    data_fields = pivot_table.findall(f"{{{_SPREADSHEETML_NS}}}dataFields")
    if len(data_fields) != 1 or len(data_fields[0]) != 1:
        return None
    data_fields[0][0].attrib.pop("subtotal", None)
    return ElementTree.tostring(pivot_table, encoding="utf-8", xml_declaration=True)


def _raw_pivot_slicer_selection_state(path: Path, pivot_sheet: str) -> dict[str, Any] | None:
    """Read WCAB's compact local PivotTable Slicer-cache selection graph.

    The reader establishes only a stored relationship-backed declaration. It
    does not create a Slicer view, apply a filter, refresh a cache, calculate a
    PivotTable, or infer any displayed report result.
    """

    state = _raw_pivot_cache_refresh_state(path, pivot_sheet)
    if state is None:
        return None
    try:
        with ZipFile(path) as archive:
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            workbook_relationships = ElementTree.fromstring(
                archive.read("xl/_rels/workbook.xml.rels")
            )
            sheets = workbook.find(f"{{{_SPREADSHEETML_NS}}}sheets")
            report_sheets = (
                [
                    sheet
                    for sheet in sheets.findall(f"{{{_SPREADSHEETML_NS}}}sheet")
                    if sheet.get("name") == pivot_sheet
                ]
                if sheets is not None
                else []
            )
            if len(report_sheets) != 1:
                return None
            report_sheet_id = report_sheets[0].get("sheetId")
            try:
                if not isinstance(report_sheet_id, str) or int(report_sheet_id) <= 0:
                    return None
            except ValueError:
                return None

            workbook_extension_lists = workbook.findall(f"{{{_SPREADSHEETML_NS}}}extLst")
            if len(workbook_extension_lists) != 1:
                return None
            workbook_extensions = workbook_extension_lists[0]
            slicer_extensions = [
                extension
                for extension in workbook_extensions.findall(f"{{{_SPREADSHEETML_NS}}}ext")
                if extension.get("uri") == "{BBE1A952-AA13-448E-AADC-164F8A28A991}"
            ]
            if len(slicer_extensions) != 1:
                return None
            slicer_cache_sets = slicer_extensions[0].findall(
                f"{{{_OFFICE_2010_SPREADSHEET_NS}}}slicerCaches"
            )
            if len(slicer_cache_sets) != 1 or len(slicer_cache_sets[0]) != 1:
                return None
            slicer_declaration = slicer_cache_sets[0][0]
            if slicer_declaration.tag != f"{{{_OFFICE_2010_SPREADSHEET_NS}}}slicerCache":
                return None
            slicer_relationship_id = slicer_declaration.get(f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id")
            if not isinstance(slicer_relationship_id, str):
                return None
            slicer_member = _relationship_part_member(
                workbook_relationships,
                slicer_relationship_id,
                "xl/workbook.xml",
                relationship_type=_SLICER_CACHE_RELATIONSHIP,
            )
            if slicer_member is None:
                return None
            slicer_cache = ElementTree.fromstring(archive.read(slicer_member))

            content_types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
            slicer_content_types = [
                override
                for override in content_types.findall(f"{{{_CONTENT_TYPES_NS}}}Override")
                if override.get("PartName") == f"/{slicer_member}"
            ]
            if len(slicer_content_types) != 1:
                return None

            cache_definition = ElementTree.fromstring(archive.read(state["definition_member"]))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None

    if slicer_cache.tag != f"{{{_OFFICE_2010_SPREADSHEET_NS}}}slicerCacheDefinition":
        return None
    pivot_table_sets = slicer_cache.findall(f"{{{_OFFICE_2010_SPREADSHEET_NS}}}pivotTables")
    data_sets = slicer_cache.findall(f"{{{_OFFICE_2010_SPREADSHEET_NS}}}data")
    if (
        len(pivot_table_sets) != 1
        or pivot_table_sets[0].get("count") != "1"
        or len(pivot_table_sets[0]) != 1
        or pivot_table_sets[0][0].tag != f"{{{_OFFICE_2010_SPREADSHEET_NS}}}pivotTable"
        or len(data_sets) != 1
        or len(data_sets[0]) != 1
        or data_sets[0][0].tag != f"{{{_OFFICE_2010_SPREADSHEET_NS}}}tabular"
    ):
        return None
    slicer_pivot_table = pivot_table_sets[0][0]
    tabular = data_sets[0][0]
    item_sets = tabular.findall(f"{{{_OFFICE_2010_SPREADSHEET_NS}}}items")
    if len(item_sets) != 1:
        return None
    items = item_sets[0]
    slicer_items = list(items)
    if any(item.tag != f"{{{_OFFICE_2010_SPREADSHEET_NS}}}i" or len(item) for item in slicer_items):
        return None
    try:
        item_count = int(items.get("count", ""))
        item_indexes = tuple(int(item.get("x", "")) for item in slicer_items)
    except ValueError:
        return None
    selected_indexes = tuple(
        item_index
        for item_index, item in zip(item_indexes, slicer_items, strict=True)
        if _ooxml_boolean(item.get("s")) is True
    )
    if (
        item_count != len(slicer_items)
        or item_count <= 0
        or item_indexes != tuple(range(item_count))
        or len(selected_indexes) != 1
    ):
        return None

    cache_extensions = cache_definition.findall(f"{{{_SPREADSHEETML_NS}}}extLst")
    if len(cache_extensions) != 1:
        return None
    pivot_cache_extensions = [
        extension
        for extension in cache_extensions[0].findall(f"{{{_SPREADSHEETML_NS}}}ext")
        if extension.get("uri") == "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}"
    ]
    if len(pivot_cache_extensions) != 1 or len(pivot_cache_extensions[0]) != 1:
        return None
    pivot_cache_extension = pivot_cache_extensions[0][0]
    if pivot_cache_extension.tag != f"{{{_OFFICE_2010_SPREADSHEET_NS}}}pivotCacheDefinition":
        return None

    cache_field_sets = cache_definition.findall(f"{{{_SPREADSHEETML_NS}}}cacheFields")
    if len(cache_field_sets) != 1:
        return None
    cache_fields = list(cache_field_sets[0])
    source_name = slicer_cache.get("sourceName")
    source_fields = [
        (index, field)
        for index, field in enumerate(cache_fields)
        if field.tag == f"{{{_SPREADSHEETML_NS}}}cacheField" and field.get("name") == source_name
    ]
    if len(source_fields) != 1:
        return None
    source_field_index, source_field = source_fields[0]
    shared_item_sets = source_field.findall(f"{{{_SPREADSHEETML_NS}}}sharedItems")
    if len(shared_item_sets) != 1:
        return None
    shared_items = list(shared_item_sets[0])
    shared_item_values = tuple(item.get("v") for item in shared_items)
    if (
        shared_item_sets[0].get("count") != str(len(shared_items))
        or any(item.tag != f"{{{_SPREADSHEETML_NS}}}s" for item in shared_items)
        or any(not isinstance(value, str) for value in shared_item_values)
        or item_count != len(shared_item_values)
    ):
        return None
    selected_item_index = selected_indexes[0]
    return {
        **state,
        "pivot_sheet_id": int(report_sheet_id),
        "slicer_member": slicer_member,
        "slicer_declaration_attributes": tuple(sorted(slicer_declaration.attrib.items())),
        "slicer_extension_attributes": tuple(sorted(slicer_extensions[0].attrib.items())),
        "slicer_content_type_attributes": tuple(sorted(slicer_content_types[0].attrib.items())),
        "slicer_attributes": tuple(sorted(slicer_cache.attrib.items())),
        "slicer_pivot_table_attributes": tuple(sorted(slicer_pivot_table.attrib.items())),
        "slicer_tabular_attributes": tuple(sorted(tabular.attrib.items())),
        "slicer_item_attributes": tuple(
            tuple(sorted(item.attrib.items())) for item in slicer_items
        ),
        "slicer_source_field_index": source_field_index,
        "slicer_source_shared_item_values": shared_item_values,
        "slicer_selected_item_index": selected_item_index,
        "slicer_selected_value": shared_item_values[selected_item_index],
        "pivot_cache_slicer_extension_attributes": tuple(
            sorted(pivot_cache_extension.attrib.items())
        ),
    }


def _slicer_cache_without_selection_state(path: Path, pivot_sheet: str) -> bytes | None:
    """Return WCAB's one Slicer cache with only selected-state attrs erased."""

    state = _raw_pivot_slicer_selection_state(path, pivot_sheet)
    if state is None:
        return None
    try:
        with ZipFile(path) as archive:
            slicer_cache = ElementTree.fromstring(archive.read(state["slicer_member"]))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    item_sets = slicer_cache.findall(
        f"{{{_OFFICE_2010_SPREADSHEET_NS}}}data/"
        f"{{{_OFFICE_2010_SPREADSHEET_NS}}}tabular/"
        f"{{{_OFFICE_2010_SPREADSHEET_NS}}}items"
    )
    if len(item_sets) != 1 or len(item_sets[0]) != 2:
        return None
    for item in item_sets[0]:
        item.attrib.pop("s", None)
    return ElementTree.tostring(slicer_cache, encoding="utf-8", xml_declaration=True)


def _power_query_length_prefixed_fields(payload: bytes) -> tuple[int, tuple[bytes, ...]] | None:
    """Read the four compact fields from WCAB's generated Data Mashup stream."""

    if len(payload) < 4:
        return None
    version = struct.unpack_from("<I", payload)[0]
    cursor = 4
    fields: list[bytes] = []
    for _ in range(4):
        if cursor + 4 > len(payload):
            return None
        field_size = struct.unpack_from("<I", payload, cursor)[0]
        cursor += 4
        if field_size > len(payload) - cursor:
            return None
        fields.append(payload[cursor : cursor + field_size])
        cursor += field_size
    return (version, tuple(fields)) if cursor == len(payload) else None


def _raw_power_query_m_filter_state(path: Path) -> dict[str, Any] | None:
    """Read WCAB's compact, connection-only local-table Data Mashup contract.

    This reader accepts only the small generated envelope, with bounded nested
    ZIP members and a deliberately narrow M formula pattern. It never executes
    M, refreshes a query, materializes output, or evaluates a workbook formula.
    """

    try:
        with ZipFile(path) as archive:
            relationships = ElementTree.fromstring(archive.read("_rels/.rels"))
            relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
            custom_xml_relationships = [
                relationship
                for relationship in relationships.findall(relationship_tag)
                if relationship.get("Id") == "rIdWCABPowerQuery"
                and relationship.get("Type") == _POWER_QUERY_CUSTOM_XML_RELATIONSHIP
                and relationship.get("Target") == "customXml/item1.xml"
                and relationship.get("TargetMode") != "External"
            ]
            if len(custom_xml_relationships) != 1:
                return None

            content_types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
            xml_defaults = [
                item
                for item in content_types.findall(f"{{{_CONTENT_TYPES_NS}}}Default")
                if item.get("Extension") == "xml" and item.get("ContentType") == "application/xml"
            ]
            if len(xml_defaults) != 1:
                return None

            mashup_root = ElementTree.fromstring(archive.read("customXml/item1.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    if mashup_root.tag != f"{{{_DATA_MASHUP_NS}}}DataMashup" or mashup_root.attrib:
        return None
    encoded_stream = mashup_root.text
    if (
        not isinstance(encoded_stream, str)
        or not encoded_stream
        or len(encoded_stream.encode("ascii", "ignore")) > _POWER_QUERY_STREAM_MAX_ENCODED_BYTES
    ):
        return None
    try:
        decoded_stream = base64.b64decode(encoded_stream, validate=True)
    except (binascii.Error, ValueError):
        return None
    stream_fields = _power_query_length_prefixed_fields(decoded_stream)
    if stream_fields is None:
        return None
    stream_version, (package_payload, permissions_payload, metadata_payload, permission_binding) = (
        stream_fields
    )
    try:
        with ZipFile(io.BytesIO(package_payload)) as package:
            package_entries = package.infolist()
            expected_package_members = (
                "Config/Package.xml",
                "Formulas/Section1.m",
                "[Content_Types].xml",
            )
            if (
                tuple(sorted(entry.filename for entry in package_entries))
                != expected_package_members
                or any(
                    entry.file_size > _POWER_QUERY_PACKAGE_MAX_UNCOMPRESSED_BYTES
                    for entry in package_entries
                )
                or sum(entry.file_size for entry in package_entries)
                > _POWER_QUERY_PACKAGE_MAX_UNCOMPRESSED_BYTES
                or package.testzip() is not None
            ):
                return None
            package_content_types = package.read("[Content_Types].xml")
            package_configuration = package.read("Config/Package.xml")
            formula = package.read("Formulas/Section1.m").decode("utf-8")
    except (BadZipFile, KeyError, OSError, RuntimeError, UnicodeDecodeError):
        return None
    formula_match = _POWER_QUERY_M_FILTER.fullmatch(formula)
    if formula_match is None:
        return None
    if len(metadata_payload) < 12:
        return None
    metadata_version, metadata_xml_size = struct.unpack_from("<II", metadata_payload)
    metadata_xml_end = 8 + metadata_xml_size
    if metadata_xml_end + 4 > len(metadata_payload):
        return None
    metadata_content_size = struct.unpack_from("<I", metadata_payload, metadata_xml_end)[0]
    metadata_content_start = metadata_xml_end + 4
    if metadata_content_size != 0 or metadata_content_start != len(metadata_payload):
        return None
    try:
        metadata_root = ElementTree.fromstring(metadata_payload[8:metadata_xml_end])
        permissions_root = ElementTree.fromstring(permissions_payload)
    except ElementTree.ParseError:
        return None
    if (
        metadata_root.tag != f"{{{_DATA_MASHUP_NS}}}LocalPackageMetadataFile"
        or metadata_root.attrib
    ):
        return None
    item_sets = metadata_root.findall(f"{{{_DATA_MASHUP_NS}}}Items")
    if len(item_sets) != 1 or len(item_sets[0]) != 1:
        return None
    metadata_item = item_sets[0][0]
    if metadata_item.tag != f"{{{_DATA_MASHUP_NS}}}Item" or metadata_item.attrib:
        return None
    metadata_item_parts = list(metadata_item)
    if [part.tag for part in metadata_item_parts] != [
        f"{{{_DATA_MASHUP_NS}}}ItemLocation",
        f"{{{_DATA_MASHUP_NS}}}StableEntries",
    ]:
        return None
    item_location, stable_entries = metadata_item_parts
    if len(item_location) != 2 or len(stable_entries) != 1:
        return None
    item_type, item_path = list(item_location)
    stable_entry = stable_entries[0]
    if (
        item_type.tag != f"{{{_DATA_MASHUP_NS}}}ItemType"
        or item_type.text != "Formula"
        or item_path.tag != f"{{{_DATA_MASHUP_NS}}}ItemPath"
        or not isinstance(item_path.text, str)
        or stable_entry.tag != f"{{{_DATA_MASHUP_NS}}}Entry"
        or tuple(sorted(stable_entry.attrib.items())) != (("Type", "FillEnabled"), ("Value", "l0"))
        or len(stable_entry)
    ):
        return None
    if permissions_root.tag != "PermissionList" or permissions_root.attrib:
        return None
    permission_children = list(permissions_root)
    if any(child.attrib or len(child) for child in permission_children):
        return None
    permission_values = {child.tag: child.text for child in permission_children}
    if permission_values != {
        "CanEvaluateFuturePackages": "false",
        "FirewallEnabled": "true",
    }:
        return None
    return {
        "data_mashup_member": "customXml/item1.xml",
        "root_relationship_attributes": tuple(sorted(custom_xml_relationships[0].attrib.items())),
        "stream_version": stream_version,
        "package_members": tuple(sorted(entry.filename for entry in package_entries)),
        "package_content_types": package_content_types,
        "package_configuration": package_configuration,
        "formula": formula,
        "query_section": formula_match.group("section"),
        "query_name": formula_match.group("query"),
        "source_table": formula_match.group("source_table"),
        "filter_column": formula_match.group("filter_column"),
        "filter_value": formula_match.group("filter_value"),
        "metadata_version": metadata_version,
        "metadata_item_path": item_path.text,
        "fill_enabled": False,
        "firewall_enabled": True,
        "future_packages_allowed": False,
        "permission_binding": permission_binding,
    }


def _chart_anchor_cell(anchor: ElementTree.Element) -> str | None:
    """Return one zero-offset worksheet-drawing anchor as an A1 coordinate."""

    from_tag = f"{{{_DRAWINGML_SPREADSHEET_DRAWING_NS}}}from"
    anchor_from = anchor.find(from_tag)
    if anchor_from is None:
        return None
    values: dict[str, int] = {}
    for name in ("col", "colOff", "row", "rowOff"):
        element = anchor_from.find(f"{{{_DRAWINGML_SPREADSHEET_DRAWING_NS}}}{name}")
        try:
            values[name] = int(element.text) if element is not None and element.text else -1
        except ValueError:
            return None
    if values["col"] < 0 or values["row"] < 0 or values["colOff"] != 0 or values["rowOff"] != 0:
        return None
    column = values["col"] + 1
    label = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        label = chr(ord("A") + remainder) + label
    return f"{label}{values['row'] + 1}"


def _raw_chart_series_value_reference_state(path: Path, chart_sheet: str) -> dict[str, Any] | None:
    """Read WCAB's one worksheet-drawing chart and its stored references.

    This follows the generated worksheet -> drawing -> chart relationship chain
    only. It does not evaluate a chart formula, calculate source cells, or
    render a chart.
    """

    try:
        with ZipFile(path) as archive:
            worksheet_member = _worksheet_member_for_sheet(archive, chart_sheet)
            if worksheet_member is None:
                return None
            worksheet = ElementTree.fromstring(archive.read(worksheet_member))
            drawing_references = worksheet.findall(f"{{{_SPREADSHEETML_NS}}}drawing")
            if len(drawing_references) != 1:
                return None
            drawing_relationship_id = drawing_references[0].get(
                f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id"
            )
            if not isinstance(drawing_relationship_id, str):
                return None
            worksheet_directory, worksheet_filename = worksheet_member.rsplit("/", maxsplit=1)
            worksheet_relationships = ElementTree.fromstring(
                archive.read(f"{worksheet_directory}/_rels/{worksheet_filename}.rels")
            )
            drawing_member = _relationship_part_member(
                worksheet_relationships,
                drawing_relationship_id,
                worksheet_member,
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/drawing",
            )
            if drawing_member is None:
                return None
            drawing = ElementTree.fromstring(archive.read(drawing_member))
            anchors = drawing.findall(f"{{{_DRAWINGML_SPREADSHEET_DRAWING_NS}}}oneCellAnchor")
            if len(anchors) != 1:
                return None
            anchor = _chart_anchor_cell(anchors[0])
            if anchor is None:
                return None
            chart_references = list(drawing.iter(f"{{{_DRAWINGML_CHART_NS}}}chart"))
            if len(chart_references) != 1:
                return None
            chart_relationship_id = chart_references[0].get(f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id")
            if not isinstance(chart_relationship_id, str):
                return None
            drawing_directory, drawing_filename = drawing_member.rsplit("/", maxsplit=1)
            drawing_relationships = ElementTree.fromstring(
                archive.read(f"{drawing_directory}/_rels/{drawing_filename}.rels")
            )
            chart_member = _relationship_part_member(
                drawing_relationships,
                chart_relationship_id,
                drawing_member,
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/chart",
            )
            if chart_member is None:
                return None
            chart = ElementTree.fromstring(archive.read(chart_member))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None

    if chart.tag != f"{{{_DRAWINGML_CHART_NS}}}chartSpace":
        return None
    series = chart.findall(f".//{{{_DRAWINGML_CHART_NS}}}ser")
    if len(series) != 1:
        return None
    title_references = series[0].findall(
        f"{{{_DRAWINGML_CHART_NS}}}tx/{{{_DRAWINGML_CHART_NS}}}strRef/{{{_DRAWINGML_CHART_NS}}}f"
    )
    category_references = series[0].findall(
        f"{{{_DRAWINGML_CHART_NS}}}cat/{{{_DRAWINGML_CHART_NS}}}numRef/{{{_DRAWINGML_CHART_NS}}}f"
    )
    value_references = series[0].findall(
        f"{{{_DRAWINGML_CHART_NS}}}val/{{{_DRAWINGML_CHART_NS}}}numRef/{{{_DRAWINGML_CHART_NS}}}f"
    )
    if (
        len(title_references) != 1
        or len(category_references) != 1
        or len(value_references) != 1
        or not isinstance(title_references[0].text, str)
        or not isinstance(category_references[0].text, str)
        or not isinstance(value_references[0].text, str)
    ):
        return None
    return {
        "worksheet_member": worksheet_member,
        "drawing_member": drawing_member,
        "chart_member": chart_member,
        "anchor": anchor,
        "series_title_reference": title_references[0].text,
        "category_reference": category_references[0].text,
        "value_reference": value_references[0].text,
    }


def _chart_without_value_reference(path: Path, chart_member: str) -> bytes | None:
    """Return one generated chart XML tree with its sole value reference erased."""

    try:
        with ZipFile(path) as archive:
            chart = ElementTree.fromstring(archive.read(chart_member))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    series = chart.findall(f".//{{{_DRAWINGML_CHART_NS}}}ser")
    if len(series) != 1:
        return None
    value_references = series[0].findall(
        f"{{{_DRAWINGML_CHART_NS}}}val/{{{_DRAWINGML_CHART_NS}}}numRef/{{{_DRAWINGML_CHART_NS}}}f"
    )
    if len(value_references) != 1:
        return None
    value_references[0].text = None
    return ElementTree.tostring(chart, encoding="utf-8", xml_declaration=True)


def _formula_cached_result_state(
    path: Path, sheet_name: str, coordinate: str
) -> tuple[str, str, str, str, bytes] | None:
    """Read one formula expression and its raw saved ``<v>`` result.

    The tuple is ``(worksheet_member, formula, result_type, result_text,
    worksheet_without_result)``. It is intentionally a narrow raw-OOXML
    reader: the final item lets the validator prove that the fixture changed
    only the selected result text, without treating that text as an evaluated
    or business-correct outcome.
    """

    try:
        with ZipFile(path) as archive:
            worksheet_member = _worksheet_member_for_sheet(archive, sheet_name)
            if worksheet_member is None:
                return None
            worksheet = ElementTree.fromstring(archive.read(worksheet_member))
            cell_tag = f"{{{_SPREADSHEETML_NS}}}c"
            cells = [cell for cell in worksheet.iter(cell_tag) if cell.get("r") == coordinate]
            if len(cells) != 1:
                return None
            cell = cells[0]
            formula = cell.find(f"{{{_SPREADSHEETML_NS}}}f")
            cached_result = cell.find(f"{{{_SPREADSHEETML_NS}}}v")
            if formula is None or cached_result is None or cached_result.text is None:
                return None
            formula_text = f"={formula.text or ''}"
            cell_type = cell.get("t")
            result_type = "numeric" if cell_type in {None, "n"} else cell_type
            result_text = cached_result.text
            cached_result.text = None
            worksheet_without_result = ElementTree.tostring(
                worksheet, encoding="utf-8", xml_declaration=True
            )
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    return (
        worksheet_member,
        formula_text,
        result_type,
        result_text,
        worksheet_without_result,
    )


def _raw_cell_state(
    path: Path, sheet_name: str, coordinate: str
) -> tuple[str, str | None, str | None, str | None, str | None] | None:
    """Read one generated cell's raw type, style, formula, and value text.

    This is intentionally narrower than a worksheet reader. It lets date-system
    fixtures prove that a stored serial and formula text did not change before a
    client applies any epoch-based conversion.
    """

    try:
        with ZipFile(path) as archive:
            worksheet_member = _worksheet_member_for_sheet(archive, sheet_name)
            if worksheet_member is None:
                return None
            worksheet = ElementTree.fromstring(archive.read(worksheet_member))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    cell_tag = f"{{{_SPREADSHEETML_NS}}}c"
    matches = [cell for cell in worksheet.iter(cell_tag) if cell.get("r") == coordinate]
    if len(matches) != 1:
        return None
    cell = matches[0]
    formula = cell.find(f"{{{_SPREADSHEETML_NS}}}f")
    value = cell.find(f"{{{_SPREADSHEETML_NS}}}v")
    return (
        worksheet_member,
        cell.get("t"),
        cell.get("s"),
        f"={formula.text or ''}" if formula is not None else None,
        value.text if value is not None else None,
    )


def _raw_auto_filter_state(
    path: Path, sheet_name: str
) -> (
    tuple[
        str,
        tuple[tuple[str, str], ...],
        int,
        tuple[tuple[str, str], ...],
        tuple[tuple[str, str], ...],
        tuple[str, ...],
    ]
    | None
):
    """Read WCAB's one explicit worksheet AutoFilter criterion.

    The narrow reader proves a raw stored control transition without applying a
    filter, changing row visibility, or calculating a ``SUBTOTAL`` result.
    """

    try:
        with ZipFile(path) as archive:
            worksheet_member = _worksheet_member_for_sheet(archive, sheet_name)
            if worksheet_member is None:
                return None
            worksheet = ElementTree.fromstring(archive.read(worksheet_member))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    auto_filter_tag = f"{{{_SPREADSHEETML_NS}}}autoFilter"
    filter_column_tag = f"{{{_SPREADSHEETML_NS}}}filterColumn"
    filters_tag = f"{{{_SPREADSHEETML_NS}}}filters"
    filter_tag = f"{{{_SPREADSHEETML_NS}}}filter"
    auto_filters = worksheet.findall(auto_filter_tag)
    if len(auto_filters) != 1:
        return None
    auto_filter = auto_filters[0]
    filter_columns = auto_filter.findall(filter_column_tag)
    if len(filter_columns) != 1:
        return None
    filter_column = filter_columns[0]
    if len(auto_filter) != 1 or auto_filter[0] is not filter_column:
        return None
    try:
        column_id = int(filter_column.get("colId", ""))
    except ValueError:
        return None
    filters = filter_column.findall(filters_tag)
    if len(filters) != 1:
        return None
    if len(filter_column) != 1 or filter_column[0] is not filters[0]:
        return None
    filter_elements = filters[0].findall(filter_tag)
    if len(filter_elements) != 1 or len(filters[0]) != 1 or filters[0][0] is not filter_elements[0]:
        return None
    values: list[str] = []
    for filter_element in filter_elements:
        value = filter_element.get("val")
        if not isinstance(value, str):
            return None
        values.append(value)
    if len(values) != 1:
        return None
    return (
        worksheet_member,
        tuple(sorted(auto_filter.attrib.items())),
        column_id,
        tuple(sorted(filter_column.attrib.items())),
        tuple(sorted(filters[0].attrib.items())),
        tuple(values),
    )


def _worksheet_without_auto_filter_criteria(path: Path, sheet_name: str) -> bytes | None:
    """Return one raw worksheet with its sole AutoFilter criterion removed."""

    if _raw_auto_filter_state(path, sheet_name) is None:
        return None
    try:
        with ZipFile(path) as archive:
            worksheet_member = _worksheet_member_for_sheet(archive, sheet_name)
            if worksheet_member is None:
                return None
            worksheet = ElementTree.fromstring(archive.read(worksheet_member))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    auto_filter_tag = f"{{{_SPREADSHEETML_NS}}}autoFilter"
    filter_column_tag = f"{{{_SPREADSHEETML_NS}}}filterColumn"
    auto_filters = worksheet.findall(auto_filter_tag)
    if len(auto_filters) != 1:
        return None
    auto_filter = auto_filters[0]
    filter_columns = auto_filter.findall(filter_column_tag)
    if len(filter_columns) != 1:
        return None
    auto_filter.remove(filter_columns[0])
    return ElementTree.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def _raw_cell_number_format(path: Path, sheet_name: str, coordinate: str) -> str | None:
    """Resolve the generated cell's custom number format from raw styles XML."""

    state = _raw_cell_state(path, sheet_name, coordinate)
    if state is None:
        return None
    try:
        style_index = int(state[2] or "0")
        with ZipFile(path) as archive:
            styles = ElementTree.fromstring(archive.read("xl/styles.xml"))
    except (BadZipFile, KeyError, OSError, ValueError, ElementTree.ParseError):
        return None
    cell_xfs = styles.findall(f"./{{{_SPREADSHEETML_NS}}}cellXfs/{{{_SPREADSHEETML_NS}}}xf")
    if not 0 <= style_index < len(cell_xfs):
        return None
    try:
        number_format_id = int(cell_xfs[style_index].get("numFmtId", ""))
    except ValueError:
        return None
    custom_formats: dict[int, str] = {}
    for number_format in styles.findall(
        f"./{{{_SPREADSHEETML_NS}}}numFmts/{{{_SPREADSHEETML_NS}}}numFmt"
    ):
        try:
            custom_format_id = int(number_format.get("numFmtId", ""))
        except ValueError:
            continue
        format_code = number_format.get("formatCode")
        if format_code is not None:
            custom_formats[custom_format_id] = format_code
    return custom_formats.get(number_format_id)


def _workbook_date_system_state(path: Path) -> tuple[bool, bool] | None:
    """Read WCAB's explicit raw workbook serial-date controls."""

    properties = _workbook_properties(path)
    if properties is None:
        return None
    date_1904 = _ooxml_boolean(properties.get("date1904"))
    date_compatibility = _ooxml_boolean(properties.get("dateCompatibility"))
    return (
        (date_1904, date_compatibility)
        if date_1904 is not None and date_compatibility is not None
        else None
    )


def _workbook_without_date_system_controls(path: Path) -> bytes | None:
    """Return raw workbook XML with only WCAB's two date controls removed."""

    try:
        with ZipFile(path) as archive:
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    properties = workbook.find(f"{{{_SPREADSHEETML_NS}}}workbookPr")
    if properties is None:
        return None
    properties.attrib.pop("date1904", None)
    properties.attrib.pop("dateCompatibility", None)
    return ElementTree.tostring(workbook, encoding="utf-8", xml_declaration=True)


def _xlsx_member_differences(baseline_path: Path, candidate_path: Path) -> set[str] | None:
    """Return changed member names when two small fixture archives align."""

    try:
        with ZipFile(baseline_path) as baseline, ZipFile(candidate_path) as candidate:
            baseline_members = {
                entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
            }
            candidate_members = {
                entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
            }
    except (BadZipFile, OSError):
        return None
    if set(baseline_members) != set(candidate_members):
        return None
    return {
        member
        for member in baseline_members
        if baseline_members[member] != candidate_members[member]
    }


def _dynamic_array_cell_metadata_indexes(archive: ZipFile) -> set[int] | None:
    """Read the generated package's cell-metadata bindings for dynamic arrays.

    This mirrors only the compact OOXML relationship used by WCAB's fixture:
    a workbook ``sheetMetadata`` relationship, an ``XLDAPR`` future-metadata
    record with ``fDynamic=true``, and one-based ``cm`` indices into
    ``cellMetadata``. It intentionally does not calculate formulas or infer a
    dynamic result extent.
    """

    try:
        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
    targets = [
        relationship.get("Target")
        for relationship in relationships.findall(relationship_tag)
        if relationship.get("Type") == f"{_DOCUMENT_RELATIONSHIPS_NS}/sheetMetadata"
    ]
    if not targets:
        return set()
    if len(targets) != 1 or targets[0] != "metadata.xml":
        return None
    try:
        metadata = ElementTree.fromstring(archive.read("xl/metadata.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None

    def tag(name: str) -> str:
        return f"{{{_SPREADSHEETML_NS}}}{name}"

    metadata_types = [
        item.get("name", "")
        for item in metadata.findall(f"./{tag('metadataTypes')}/{tag('metadataType')}")
    ]
    dynamic_future_indexes: dict[str, set[int]] = {}
    dynamic_properties_tag = f"{{{_DYNAMIC_ARRAY_NS}}}dynamicArrayProperties"
    for future_metadata in metadata.findall(tag("futureMetadata")):
        name = future_metadata.get("name")
        if not name:
            continue
        indexes = {
            index
            for index, block in enumerate(future_metadata.findall(tag("bk")))
            if any(
                item.tag == dynamic_properties_tag
                and item.get("fDynamic", "").casefold() in {"1", "true"}
                for item in block.iter()
            )
        }
        if indexes:
            dynamic_future_indexes[name] = indexes

    cell_metadata = metadata.find(tag("cellMetadata"))
    if cell_metadata is None:
        return set()
    dynamic_cells: set[int] = set()
    for cell_index, block in enumerate(cell_metadata.findall(tag("bk")), start=1):
        for record in block.findall(tag("rc")):
            try:
                type_index = int(record.get("t", ""))
                value_index = int(record.get("v", ""))
            except ValueError:
                continue
            if not 1 <= type_index <= len(metadata_types):
                continue
            if value_index in dynamic_future_indexes.get(metadata_types[type_index - 1], set()):
                dynamic_cells.add(cell_index)
                break
    return dynamic_cells


def _array_formula_state(
    path: Path, sheet_name: str, coordinate: str
) -> tuple[str, str | None, str | None] | None:
    """Classify a generated raw-OOXML array anchor as CSE or dynamic.

    The returned tuple is ``(mode, formula_text, output_range)``. This helper
    is deliberately limited to the generated fixture's anchor representation;
    any incomplete metadata becomes ``unclassified`` rather than a guess.
    """

    try:
        with ZipFile(path) as archive:
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            sheets = workbook.find(f"{{{_SPREADSHEETML_NS}}}sheets")
            if sheets is None:
                return None
            sheet_tag = f"{{{_SPREADSHEETML_NS}}}sheet"
            matching_sheets = [
                sheet for sheet in sheets.findall(sheet_tag) if sheet.get("name") == sheet_name
            ]
            if len(matching_sheets) != 1:
                return None
            relationship_id = matching_sheets[0].get(f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id")
            if not isinstance(relationship_id, str):
                return None
            relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            target = _relationship_target(
                relationships,
                relationship_id,
                relationship_type=f"{_DOCUMENT_RELATIONSHIPS_NS}/worksheet",
            )
            if target is None:
                return None
            worksheet_member = target if target.startswith("xl/") else f"xl/{target}"
            worksheet = ElementTree.fromstring(archive.read(worksheet_member))
            cell_tag = f"{{{_SPREADSHEETML_NS}}}c"
            cells = [cell for cell in worksheet.iter(cell_tag) if cell.get("r") == coordinate]
            if len(cells) != 1:
                return None
            formula = cells[0].find(f"{{{_SPREADSHEETML_NS}}}f")
            if formula is None:
                return None
            formula_text = f"={formula.text or ''}"
            output_range = formula.get("ref")
            if formula.get("t") != "array":
                return "ordinary", formula_text, output_range
            metadata_index = cells[0].get("cm")
            if metadata_index is None:
                return "legacy_cse", formula_text, output_range
            try:
                cell_metadata_index = int(metadata_index)
            except ValueError:
                return "unclassified", formula_text, output_range
            dynamic_cells = _dynamic_array_cell_metadata_indexes(archive)
            if dynamic_cells is None:
                return "unclassified", formula_text, output_range
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    mode = "dynamic" if cell_metadata_index in dynamic_cells else "unclassified"
    return mode, formula_text, output_range


def _direct_graph(workbook: Workbook) -> dict[tuple[str, str], set[tuple[str, str]]]:
    """Return a small static local A1 dependency graph for the generated fixtures.

    This deliberately recognizes only ordinary direct A1 references.  The
    benchmark labels this as a lower bound rather than using it as an Excel
    calculation engine.
    """

    graph: dict[tuple[str, str], set[tuple[str, str]]] = {}
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                # External references are separate benchmark facts.  Avoid
                # treating an external sheet label as a local dependency.
                if "[" in cell.value:
                    continue
                dependent = (worksheet.title, cell.coordinate)
                for match in _CELL_REFERENCE.finditer(cell.value):
                    sheet = match.group("quoted") or match.group("plain") or worksheet.title
                    sheet = sheet.replace("''", "'")
                    if sheet not in workbook.sheetnames:
                        continue
                    source = (sheet, match.group("cell").replace("$", ""))
                    graph.setdefault(source, set()).add(dependent)
    return graph


def _reachable(
    graph: dict[tuple[str, str], set[tuple[str, str]]], source: tuple[str, str]
) -> set[tuple[str, str]]:
    discovered: set[tuple[str, str]] = set()
    pending = deque(graph.get(source, ()))
    while pending:
        current = pending.popleft()
        if current in discovered:
            continue
        discovered.add(current)
        pending.extend(graph.get(current, ()))
    return discovered


def _workbook_pair(
    case_dir: Path, truth: dict[str, Any], workbook: str | None = None
) -> tuple[Path, Path]:
    topology = truth.get("topology")
    if topology == "pair":
        if workbook is not None:
            raise FixtureValidationError(
                f"{case_dir}: pair fixture unexpectedly names workbook {workbook!r}"
            )
        return case_dir / "baseline.xlsx", case_dir / "candidate.xlsx"
    if topology == "portfolio":
        if not workbook:
            raise FixtureValidationError(f"{case_dir}: portfolio fact must name a workbook")
        return case_dir / "baseline" / workbook, case_dir / "candidate" / workbook
    raise FixtureValidationError(f"{case_dir}: unsupported topology {topology!r}")


def _assert(condition: bool, message: str, errors: list[str]) -> None:
    if not condition:
        errors.append(message)


def _validate_fact(
    case_dir: Path, truth: dict[str, Any], fact: dict[str, Any], errors: list[str]
) -> None:
    kind = fact.get("kind")
    baseline_path, candidate_path = _workbook_pair(case_dir, truth, fact.get("workbook"))
    baseline = _load_workbook(baseline_path)
    candidate = _load_workbook(candidate_path)

    if kind in {
        "formula_to_value",
        "formula_changed",
        "value_changed",
        "external_formula_added",
        "formula_cell_unlocked",
        "dynamic_formula_reference_added",
    }:
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        if sheet_name not in baseline.sheetnames or sheet_name not in candidate.sheetnames:
            errors.append(f"{truth['id']}: sheet {sheet_name!r} is absent")
            return
        before = baseline[sheet_name][coordinate]
        after = candidate[sheet_name][coordinate]
        if kind == "formula_to_value":
            _assert(
                _cell_kind(before) == "formula" and _cell_kind(after) == "value",
                f"{truth['id']}: expected {sheet_name}!{coordinate} formula-to-value change",
                errors,
            )
        elif kind == "formula_changed":
            _assert(
                _cell_kind(before) == "formula"
                and _cell_kind(after) == "formula"
                and before.value != after.value,
                f"{truth['id']}: expected {sheet_name}!{coordinate} changed formula",
                errors,
            )
        elif kind == "value_changed":
            _assert(
                _cell_kind(before) == "value"
                and _cell_kind(after) == "value"
                and before.value != after.value,
                f"{truth['id']}: expected {sheet_name}!{coordinate} changed literal value",
                errors,
            )
        elif kind == "external_formula_added":
            _assert(
                _cell_kind(after) == "formula" and "[" in str(after.value),
                f"{truth['id']}: expected external formula at {sheet_name}!{coordinate}",
                errors,
            )
        elif kind == "dynamic_formula_reference_added":
            expected_functions = fact.get("functions")
            before_functions = _dynamic_reference_functions(str(before.value))
            after_functions = _dynamic_reference_functions(str(after.value))
            _assert(
                _cell_kind(before) == "formula"
                and _cell_kind(after) == "formula"
                and before.value != after.value
                and not before_functions
                and isinstance(expected_functions, list)
                and all(isinstance(function, str) for function in expected_functions)
                and after_functions == tuple(expected_functions),
                f"{truth['id']}: expected dynamic reference functions {expected_functions!r} at {sheet_name}!{coordinate}",
                errors,
            )
        else:
            _assert(
                baseline[sheet_name].protection.sheet
                and candidate[sheet_name].protection.sheet
                and _cell_kind(before) == "formula"
                and bool(before.protection.locked)
                and not bool(after.protection.locked),
                f"{truth['id']}: expected protected formula {sheet_name}!{coordinate} to be explicitly unlocked",
                errors,
            )
        return

    if kind == "defined_name_changed":
        name = fact.get("name")
        _assert(
            _defined_name_text(baseline, name) != _defined_name_text(candidate, name),
            f"{truth['id']}: expected defined name {name!r} to change",
            errors,
        )
        return

    if kind in {"data_validation_count_changed", "conditional_formatting_count_changed"}:
        sheet_name = fact.get("sheet")
        if sheet_name not in baseline.sheetnames or sheet_name not in candidate.sheetnames:
            errors.append(f"{truth['id']}: sheet {sheet_name!r} is absent")
            return
        if kind == "data_validation_count_changed":
            before_count = len(baseline[sheet_name].data_validations.dataValidation)
            after_count = len(candidate[sheet_name].data_validations.dataValidation)
        else:
            before_count = len(baseline[sheet_name].conditional_formatting)
            after_count = len(candidate[sheet_name].conditional_formatting)
        _assert(
            before_count == fact.get("baseline_count")
            and after_count == fact.get("candidate_count"),
            f"{truth['id']}: expected {kind} {fact.get('baseline_count')} -> {fact.get('candidate_count')}, got {before_count} -> {after_count}",
            errors,
        )
        return

    if kind == "sheet_visibility_changed":
        sheet_name = fact.get("sheet")
        _assert(
            sheet_name in baseline.sheetnames
            and sheet_name in candidate.sheetnames
            and baseline[sheet_name].sheet_state == fact.get("baseline_state")
            and candidate[sheet_name].sheet_state == fact.get("candidate_state"),
            f"{truth['id']}: expected {sheet_name!r} visibility change",
            errors,
        )
        return

    if kind == "manual_calculation_incomplete":
        _assert(
            candidate.calculation.calcMode == "manual"
            and candidate.calculation.calcCompleted is False,
            f"{truth['id']}: expected manual, incomplete calculation metadata",
            errors,
        )
        return

    if kind == "iterative_calculation_enabled":
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        formula = fact.get("formula")
        before_state = _calculation_iteration_state(baseline_path)
        after_state = _calculation_iteration_state(candidate_path)
        before_properties = _calculation_properties(baseline_path)
        after_properties = _calculation_properties(candidate_path)
        before_formula = (
            baseline[sheet_name][coordinate]
            if isinstance(sheet_name, str)
            and isinstance(coordinate, str)
            and sheet_name in baseline.sheetnames
            else None
        )
        after_formula = (
            candidate[sheet_name][coordinate]
            if isinstance(sheet_name, str)
            and isinstance(coordinate, str)
            and sheet_name in candidate.sheetnames
            else None
        )
        expected_count = fact.get("iteration_count")
        expected_delta = fact.get("iteration_delta")
        try:
            declared_delta = Decimal(str(expected_delta))
        except InvalidOperation:
            declared_delta = None
        ignored_properties = {"iterate", "iterateCount", "iterateDelta"}
        before_non_iteration_properties = (
            {
                key: value
                for key, value in before_properties.items()
                if key not in ignored_properties
            }
            if before_properties is not None
            else None
        )
        after_non_iteration_properties = (
            {key: value for key, value in after_properties.items() if key not in ignored_properties}
            if after_properties is not None
            else None
        )
        source = (sheet_name, coordinate)
        graph = _direct_graph(candidate)
        _assert(
            isinstance(sheet_name, str)
            and isinstance(coordinate, str)
            and isinstance(formula, str)
            and isinstance(expected_count, int)
            and not isinstance(expected_count, bool)
            and declared_delta is not None
            and fact.get("baseline_iterate") is False
            and fact.get("candidate_iterate") is True
            and before_formula is not None
            and after_formula is not None
            and _cell_kind(before_formula) == "formula"
            and _cell_kind(after_formula) == "formula"
            and before_formula.value == formula
            and after_formula.value == formula
            and source in graph.get(source, set())
            and before_state == (False, expected_count, declared_delta)
            and after_state == (True, expected_count, declared_delta)
            and before_non_iteration_properties == after_non_iteration_properties,
            f"{truth['id']}: expected unchanged direct circular formula and calcPr iterate false -> true with declared bounds only",
            errors,
        )
        return

    if kind == "precision_as_displayed_enabled":
        input_sheet = fact.get("input_sheet")
        input_cell = fact.get("input_cell")
        formula_sheet = fact.get("formula_sheet")
        formula_cell = fact.get("formula_cell")
        formula = fact.get("formula")
        number_format = fact.get("number_format")
        expected_value = fact.get("input_value")
        try:
            declared_value = Decimal(str(expected_value))
        except InvalidOperation:
            declared_value = None
        before_properties = _calculation_properties(baseline_path)
        after_properties = _calculation_properties(candidate_path)
        before_non_precision_properties = (
            {key: value for key, value in before_properties.items() if key != "fullPrecision"}
            if before_properties is not None
            else None
        )
        after_non_precision_properties = (
            {key: value for key, value in after_properties.items() if key != "fullPrecision"}
            if after_properties is not None
            else None
        )
        if (
            not isinstance(input_sheet, str)
            or not isinstance(input_cell, str)
            or not isinstance(formula_sheet, str)
            or not isinstance(formula_cell, str)
            or input_sheet not in baseline.sheetnames
            or input_sheet not in candidate.sheetnames
            or formula_sheet not in baseline.sheetnames
            or formula_sheet not in candidate.sheetnames
        ):
            errors.append(f"{truth['id']}: precision-as-displayed locations are absent")
            return
        before_input = baseline[input_sheet][input_cell]
        after_input = candidate[input_sheet][input_cell]
        before_formula = baseline[formula_sheet][formula_cell]
        after_formula = candidate[formula_sheet][formula_cell]
        try:
            before_value = Decimal(str(before_input.value))
            after_value = Decimal(str(after_input.value))
        except InvalidOperation:
            before_value = after_value = None
        graph = _direct_graph(candidate)
        _assert(
            isinstance(formula, str)
            and isinstance(number_format, str)
            and declared_value is not None
            and declared_value.is_finite()
            and fact.get("baseline_full_precision") is True
            and fact.get("candidate_full_precision") is False
            and _calculation_full_precision(baseline_path) is True
            and _calculation_full_precision(candidate_path) is False
            and before_non_precision_properties == after_non_precision_properties
            and _cell_kind(before_input) == "value"
            and _cell_kind(after_input) == "value"
            and before_value == declared_value
            and after_value == declared_value
            and before_input.number_format == number_format
            and after_input.number_format == number_format
            and _cell_kind(before_formula) == "formula"
            and _cell_kind(after_formula) == "formula"
            and before_formula.value == formula
            and after_formula.value == formula
            and (formula_sheet, formula_cell) in graph.get((input_sheet, input_cell), set()),
            f"{truth['id']}: expected unchanged stored input/formula and calcPr fullPrecision true -> false only",
            errors,
        )
        return

    if kind == "auto_filter_criteria_changed":
        sheet_name = fact.get("sheet")
        filter_reference = fact.get("filter_ref")
        filter_column_id = fact.get("filter_column_id")
        baseline_filter_value = fact.get("baseline_filter_value")
        candidate_filter_value = fact.get("candidate_filter_value")
        subtotal_cell = fact.get("subtotal_cell")
        subtotal_formula = fact.get("subtotal_formula")
        dashboard_sheet = fact.get("dashboard_sheet")
        dashboard_cell = fact.get("dashboard_cell")
        dashboard_formula = fact.get("dashboard_formula")
        before_filter = (
            _raw_auto_filter_state(baseline_path, sheet_name)
            if isinstance(sheet_name, str)
            else None
        )
        after_filter = (
            _raw_auto_filter_state(candidate_path, sheet_name)
            if isinstance(sheet_name, str)
            else None
        )
        before_subtotal = (
            _raw_cell_state(baseline_path, sheet_name, subtotal_cell)
            if isinstance(sheet_name, str) and isinstance(subtotal_cell, str)
            else None
        )
        after_subtotal = (
            _raw_cell_state(candidate_path, sheet_name, subtotal_cell)
            if isinstance(sheet_name, str) and isinstance(subtotal_cell, str)
            else None
        )
        before_dashboard = (
            _raw_cell_state(baseline_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        after_dashboard = (
            _raw_cell_state(candidate_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        graph = _direct_graph(candidate)
        _assert(
            isinstance(sheet_name, str)
            and isinstance(filter_reference, str)
            and type(filter_column_id) is int
            and isinstance(baseline_filter_value, str)
            and isinstance(candidate_filter_value, str)
            and baseline_filter_value != candidate_filter_value
            and isinstance(subtotal_cell, str)
            and isinstance(subtotal_formula, str)
            and isinstance(dashboard_sheet, str)
            and isinstance(dashboard_cell, str)
            and isinstance(dashboard_formula, str)
            and before_filter is not None
            and after_filter is not None
            and before_filter[0] == after_filter[0]
            and before_filter[1] == (("ref", filter_reference),)
            and after_filter[1] == (("ref", filter_reference),)
            and before_filter[2] == filter_column_id
            and after_filter[2] == filter_column_id
            and before_filter[3] == after_filter[3]
            and before_filter[4] == after_filter[4]
            and before_filter[5] == (baseline_filter_value,)
            and after_filter[5] == (candidate_filter_value,)
            and before_subtotal is not None
            and after_subtotal is not None
            and before_subtotal == after_subtotal
            and before_subtotal[3] == subtotal_formula
            and before_dashboard is not None
            and after_dashboard is not None
            and before_dashboard == after_dashboard
            and before_dashboard[3] == dashboard_formula
            and _worksheet_without_auto_filter_criteria(baseline_path, sheet_name)
            == _worksheet_without_auto_filter_criteria(candidate_path, sheet_name)
            and _xlsx_member_differences(baseline_path, candidate_path) == {before_filter[0]}
            and (dashboard_sheet, dashboard_cell) in graph.get((sheet_name, subtotal_cell), set()),
            f"{truth['id']}: expected one raw AutoFilter criterion change only with stable formulas and dependency edge",
            errors,
        )
        return

    if kind == "workbook_date_system_changed":
        serial_sheet = fact.get("serial_sheet")
        serial_cell = fact.get("serial_cell")
        formula_sheet = fact.get("formula_sheet")
        formula_cell = fact.get("formula_cell")
        formula = fact.get("formula")
        dashboard_sheet = fact.get("dashboard_sheet")
        dashboard_cell = fact.get("dashboard_cell")
        dashboard_formula = fact.get("dashboard_formula")
        number_format = fact.get("number_format")
        try:
            declared_serial = Decimal(str(fact.get("serial_value")))
        except InvalidOperation:
            declared_serial = None
        before_date_system = _workbook_date_system_state(baseline_path)
        after_date_system = _workbook_date_system_state(candidate_path)
        before_input = (
            _raw_cell_state(baseline_path, serial_sheet, serial_cell)
            if isinstance(serial_sheet, str) and isinstance(serial_cell, str)
            else None
        )
        after_input = (
            _raw_cell_state(candidate_path, serial_sheet, serial_cell)
            if isinstance(serial_sheet, str) and isinstance(serial_cell, str)
            else None
        )
        before_formula = (
            _raw_cell_state(baseline_path, formula_sheet, formula_cell)
            if isinstance(formula_sheet, str) and isinstance(formula_cell, str)
            else None
        )
        after_formula = (
            _raw_cell_state(candidate_path, formula_sheet, formula_cell)
            if isinstance(formula_sheet, str) and isinstance(formula_cell, str)
            else None
        )
        before_dashboard = (
            _raw_cell_state(baseline_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        after_dashboard = (
            _raw_cell_state(candidate_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        try:
            before_serial = Decimal(before_input[4]) if before_input and before_input[4] else None
            after_serial = Decimal(after_input[4]) if after_input and after_input[4] else None
        except InvalidOperation:
            before_serial = after_serial = None
        graph = _direct_graph(candidate)
        _assert(
            fact.get("baseline_date_1904") is False
            and fact.get("candidate_date_1904") is True
            and fact.get("date_compatibility") is True
            and isinstance(serial_sheet, str)
            and isinstance(serial_cell, str)
            and isinstance(formula_sheet, str)
            and isinstance(formula_cell, str)
            and isinstance(formula, str)
            and isinstance(dashboard_sheet, str)
            and isinstance(dashboard_cell, str)
            and isinstance(dashboard_formula, str)
            and isinstance(number_format, str)
            and declared_serial is not None
            and declared_serial.is_finite()
            and before_date_system == (False, True)
            and after_date_system == (True, True)
            and before_input is not None
            and after_input is not None
            and before_input == after_input
            and before_input[1] in {None, "n"}
            and before_input[3] is None
            and before_serial == declared_serial
            and after_serial == declared_serial
            and before_formula is not None
            and after_formula is not None
            and before_formula == after_formula
            and before_formula[3] == formula
            and before_dashboard is not None
            and after_dashboard is not None
            and before_dashboard == after_dashboard
            and before_dashboard[3] == dashboard_formula
            and _raw_cell_number_format(baseline_path, serial_sheet, serial_cell) == number_format
            and _raw_cell_number_format(candidate_path, serial_sheet, serial_cell) == number_format
            and _workbook_without_date_system_controls(baseline_path)
            == _workbook_without_date_system_controls(candidate_path)
            and _xlsx_member_differences(baseline_path, candidate_path) == {"xl/workbook.xml"}
            and (formula_sheet, formula_cell) in graph.get((serial_sheet, serial_cell), set())
            and (dashboard_sheet, dashboard_cell)
            in graph.get((formula_sheet, formula_cell), set()),
            f"{truth['id']}: expected raw workbookPr date1904 false -> true only with stable serial, style, formulas, and dependency edges",
            errors,
        )
        return

    if kind == "formula_cached_result_changed":
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        formula = fact.get("formula")
        input_sheet = fact.get("input_sheet")
        input_cell = fact.get("input_cell")
        result_type = fact.get("result_type")
        try:
            declared_input = Decimal(str(fact.get("input_value")))
            declared_before_result = Decimal(str(fact.get("baseline_cached_result")))
            declared_after_result = Decimal(str(fact.get("candidate_cached_result")))
        except InvalidOperation:
            declared_input = declared_before_result = declared_after_result = None
        before_state = (
            _formula_cached_result_state(baseline_path, sheet_name, coordinate)
            if isinstance(sheet_name, str) and isinstance(coordinate, str)
            else None
        )
        after_state = (
            _formula_cached_result_state(candidate_path, sheet_name, coordinate)
            if isinstance(sheet_name, str) and isinstance(coordinate, str)
            else None
        )
        if before_state is None or after_state is None:
            errors.append(f"{truth['id']}: formula cached-result location is absent or malformed")
            return
        (
            before_member,
            before_formula_text,
            before_result_type,
            before_result_text,
            before_without_result,
        ) = before_state
        (
            after_member,
            after_formula_text,
            after_result_type,
            after_result_text,
            after_without_result,
        ) = after_state
        try:
            before_result = Decimal(before_result_text)
            after_result = Decimal(after_result_text)
        except InvalidOperation:
            before_result = after_result = None
        if (
            not isinstance(input_sheet, str)
            or not isinstance(input_cell, str)
            or input_sheet not in baseline.sheetnames
            or input_sheet not in candidate.sheetnames
        ):
            errors.append(f"{truth['id']}: formula cached-result input location is absent")
            return
        before_input = baseline[input_sheet][input_cell]
        after_input = candidate[input_sheet][input_cell]
        try:
            before_input_value = Decimal(str(before_input.value))
            after_input_value = Decimal(str(after_input.value))
        except InvalidOperation:
            before_input_value = after_input_value = None
        graph = _direct_graph(candidate)
        _assert(
            isinstance(formula, str)
            and result_type == "numeric"
            and declared_input is not None
            and declared_input.is_finite()
            and declared_before_result is not None
            and declared_before_result.is_finite()
            and declared_after_result is not None
            and declared_after_result.is_finite()
            and declared_before_result != declared_after_result
            and before_formula_text == formula
            and after_formula_text == formula
            and before_result_type == result_type
            and after_result_type == result_type
            and before_result == declared_before_result
            and after_result == declared_after_result
            and before_result is not None
            and after_result is not None
            and before_result.is_finite()
            and after_result.is_finite()
            and _cell_kind(before_input) == "value"
            and _cell_kind(after_input) == "value"
            and before_input_value == declared_input
            and after_input_value == declared_input
            and _cell_kind(baseline[sheet_name][coordinate]) == "formula"
            and _cell_kind(candidate[sheet_name][coordinate]) == "formula"
            and before_member == after_member
            and before_without_result == after_without_result
            and _xlsx_member_differences(baseline_path, candidate_path) == {before_member}
            and _calculation_properties(baseline_path) == _calculation_properties(candidate_path)
            and (sheet_name, coordinate) in graph.get((input_sheet, input_cell), set()),
            f"{truth['id']}: expected only raw numeric formula cache {sheet_name}!{coordinate} to change with formula, input, and controls unchanged",
            errors,
        )
        return

    if kind == "chart_series_value_reference_changed":
        chart_sheet = fact.get("chart_sheet")
        chart_anchor = fact.get("chart_anchor")
        source_sheet = fact.get("source_sheet")
        series_title_reference = fact.get("series_title_ref")
        category_reference = fact.get("category_ref")
        baseline_value_reference = fact.get("baseline_value_ref")
        candidate_value_reference = fact.get("candidate_value_ref")
        before_state = (
            _raw_chart_series_value_reference_state(baseline_path, chart_sheet)
            if isinstance(chart_sheet, str)
            else None
        )
        after_state = (
            _raw_chart_series_value_reference_state(candidate_path, chart_sheet)
            if isinstance(chart_sheet, str)
            else None
        )
        _assert(
            isinstance(chart_sheet, str)
            and isinstance(chart_anchor, str)
            and isinstance(source_sheet, str)
            and isinstance(series_title_reference, str)
            and isinstance(category_reference, str)
            and isinstance(baseline_value_reference, str)
            and isinstance(candidate_value_reference, str)
            and baseline_value_reference != candidate_value_reference
            and chart_sheet in baseline.sheetnames
            and chart_sheet in candidate.sheetnames
            and source_sheet in baseline.sheetnames
            and source_sheet in candidate.sheetnames
            and before_state is not None
            and after_state is not None
            and before_state["worksheet_member"] == after_state["worksheet_member"]
            and before_state["drawing_member"] == after_state["drawing_member"]
            and before_state["chart_member"] == after_state["chart_member"]
            and before_state["anchor"] == chart_anchor
            and after_state["anchor"] == chart_anchor
            and before_state["series_title_reference"] == series_title_reference
            and after_state["series_title_reference"] == series_title_reference
            and before_state["category_reference"] == category_reference
            and after_state["category_reference"] == category_reference
            and before_state["value_reference"] == baseline_value_reference
            and after_state["value_reference"] == candidate_value_reference
            and all(
                reference.startswith(f"'{source_sheet}'!")
                for reference in (
                    series_title_reference,
                    category_reference,
                    baseline_value_reference,
                    candidate_value_reference,
                )
            )
            and _chart_without_value_reference(baseline_path, before_state["chart_member"])
            == _chart_without_value_reference(candidate_path, after_state["chart_member"])
            and _xlsx_member_differences(baseline_path, candidate_path)
            == {before_state["chart_member"]}
            and _calculation_properties(baseline_path) == _calculation_properties(candidate_path),
            f"{truth['id']}: expected one relationship-bound chart value-series reference to change with stable worksheets and chart bindings",
            errors,
        )
        return

    if kind == "power_query_m_filter_changed":
        data_mashup_part = fact.get("data_mashup_part")
        source_sheet = fact.get("source_sheet")
        source_table = fact.get("source_table")
        source_reference = fact.get("source_ref")
        query_section = fact.get("query_section")
        query_name = fact.get("query_name")
        filter_column = fact.get("filter_column")
        baseline_filter_value = fact.get("baseline_filter_value")
        candidate_filter_value = fact.get("candidate_filter_value")

        def source_table_state(
            workbook: Workbook,
        ) -> tuple[str, tuple[tuple[Any, ...], ...]] | None:
            if (
                not isinstance(source_sheet, str)
                or not isinstance(source_table, str)
                or not isinstance(source_reference, str)
                or source_sheet not in workbook.sheetnames
            ):
                return None
            worksheet = workbook[source_sheet]
            try:
                table = worksheet.tables[source_table]
                cells = worksheet[source_reference]
            except (KeyError, ValueError):
                return None
            if not isinstance(cells, tuple) or any(not isinstance(row, tuple) for row in cells):
                return None
            return table.ref, tuple(tuple(cell.value for cell in row) for row in cells)

        before_state = _raw_power_query_m_filter_state(baseline_path)
        after_state = _raw_power_query_m_filter_state(candidate_path)
        before_table_state = source_table_state(baseline)
        after_table_state = source_table_state(candidate)
        expected_root_relationship_attributes = tuple(
            sorted(
                {
                    "Id": "rIdWCABPowerQuery",
                    "Type": _POWER_QUERY_CUSTOM_XML_RELATIONSHIP,
                    "Target": "customXml/item1.xml",
                }.items()
            )
        )
        expected_package_members = (
            "Config/Package.xml",
            "Formulas/Section1.m",
            "[Content_Types].xml",
        )
        expected_source_rows = (
            ("Region", "Amount"),
            ("North", 10),
            ("South", 20),
            ("North", 15),
            ("South", 25),
        )

        def expected_formula(filter_value: str) -> str:
            return (
                f"section {query_section};\n\n"
                f"shared {query_name} = let\n"
                f'    Source = Excel.CurrentWorkbook(){{[Name="{source_table}"]}}[Content],\n'
                f'    FilteredRows = Table.SelectRows(Source, each [{filter_column}] = "{filter_value}")\n'
                "in\n"
                "    FilteredRows;\n"
            )

        mutable_fields = {"formula", "filter_value"}
        _assert(
            data_mashup_part == "customXml/item1.xml"
            and source_sheet == "Source"
            and source_table == "SourceData"
            and source_reference == "A1:B5"
            and query_section == "Section1"
            and query_name == "RegionQuery"
            and filter_column == "Region"
            and baseline_filter_value == "North"
            and candidate_filter_value == "South"
            and fact.get("fill_enabled") is False
            and fact.get("firewall_enabled") is True
            and fact.get("future_packages_allowed") is False
            and before_state is not None
            and after_state is not None
            and before_table_state == after_table_state
            and before_table_state == (source_reference, expected_source_rows)
            and {key: value for key, value in before_state.items() if key not in mutable_fields}
            == {key: value for key, value in after_state.items() if key not in mutable_fields}
            and before_state["data_mashup_member"] == data_mashup_part
            and after_state["data_mashup_member"] == data_mashup_part
            and before_state["root_relationship_attributes"]
            == expected_root_relationship_attributes
            and after_state["root_relationship_attributes"] == expected_root_relationship_attributes
            and before_state["stream_version"] == 0
            and after_state["stream_version"] == 0
            and before_state["package_members"] == expected_package_members
            and after_state["package_members"] == expected_package_members
            and before_state["package_content_types"] == b"<Types/>"
            and after_state["package_content_types"] == b"<Types/>"
            and before_state["package_configuration"]
            == b"<Package>WCAB connection-only local table query</Package>"
            and after_state["package_configuration"]
            == b"<Package>WCAB connection-only local table query</Package>"
            and before_state["formula"] == expected_formula(baseline_filter_value)
            and after_state["formula"] == expected_formula(candidate_filter_value)
            and before_state["query_section"] == query_section
            and after_state["query_section"] == query_section
            and before_state["query_name"] == query_name
            and after_state["query_name"] == query_name
            and before_state["source_table"] == source_table
            and after_state["source_table"] == source_table
            and before_state["filter_column"] == filter_column
            and after_state["filter_column"] == filter_column
            and before_state["filter_value"] == baseline_filter_value
            and after_state["filter_value"] == candidate_filter_value
            and before_state["metadata_version"] == 0
            and after_state["metadata_version"] == 0
            and before_state["metadata_item_path"] == f"{query_section}/{query_name}"
            and after_state["metadata_item_path"] == f"{query_section}/{query_name}"
            and before_state["fill_enabled"] is False
            and after_state["fill_enabled"] is False
            and before_state["firewall_enabled"] is True
            and after_state["firewall_enabled"] is True
            and before_state["future_packages_allowed"] is False
            and after_state["future_packages_allowed"] is False
            and before_state["permission_binding"] == b""
            and after_state["permission_binding"] == b""
            and _xlsx_member_differences(baseline_path, candidate_path) == {data_mashup_part}
            and _calculation_properties(baseline_path) == _calculation_properties(candidate_path),
            f"{truth['id']}: expected only one compact connection-only local-table Power Query M filter literal to change with stable source, metadata, permissions, and calculation controls",
            errors,
        )
        return

    if kind == "pivot_slicer_selection_changed":
        cache_id = fact.get("cache_id")
        source_type = fact.get("source_type")
        source_sheet = fact.get("source_sheet")
        source_reference = fact.get("source_ref")
        pivot_sheet = fact.get("pivot_sheet")
        pivot_reference = fact.get("pivot_ref")
        pivot_output_cell = fact.get("pivot_output_cell")
        dashboard_sheet = fact.get("dashboard_sheet")
        dashboard_cell = fact.get("dashboard_cell")
        dashboard_formula = fact.get("dashboard_formula")
        slicer_name = fact.get("slicer_name")
        slicer_source_name = fact.get("slicer_source_name")
        slicer_pivot_table_name = fact.get("slicer_pivot_table_name")
        slicer_pivot_tab_id = fact.get("slicer_pivot_tab_id")
        item_count = fact.get("item_count")
        baseline_selected_item_index = fact.get("baseline_selected_item_index")
        candidate_selected_item_index = fact.get("candidate_selected_item_index")
        baseline_selected_value = fact.get("baseline_selected_value")
        candidate_selected_value = fact.get("candidate_selected_value")
        before_state = (
            _raw_pivot_slicer_selection_state(baseline_path, pivot_sheet)
            if isinstance(pivot_sheet, str)
            else None
        )
        after_state = (
            _raw_pivot_slicer_selection_state(candidate_path, pivot_sheet)
            if isinstance(pivot_sheet, str)
            else None
        )
        before_output = (
            _raw_cell_state(baseline_path, pivot_sheet, pivot_output_cell)
            if isinstance(pivot_sheet, str) and isinstance(pivot_output_cell, str)
            else None
        )
        after_output = (
            _raw_cell_state(candidate_path, pivot_sheet, pivot_output_cell)
            if isinstance(pivot_sheet, str) and isinstance(pivot_output_cell, str)
            else None
        )
        before_dashboard = (
            _raw_cell_state(baseline_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        after_dashboard = (
            _raw_cell_state(candidate_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        expected_definition_attributes = tuple(
            sorted(
                {
                    "recordCount": "4",
                    "refreshOnLoad": "0",
                    "saveData": "1",
                    f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id": "rIdWCABPivotRecords",
                }.items()
            )
        )
        expected_pivot_attributes = (
            tuple(
                sorted(
                    {
                        "cacheId": str(cache_id),
                        "dataCaption": "Amount",
                        "name": "WCAB Pivot Report",
                    }.items()
                )
            )
            if type(cache_id) is int
            else None
        )
        expected_location_attributes = (
            tuple(
                sorted(
                    {
                        "ref": pivot_reference,
                        "firstHeaderRow": "1",
                        "firstDataRow": "2",
                        "firstDataCol": "1",
                    }.items()
                )
            )
            if isinstance(pivot_reference, str)
            else None
        )
        expected_cache_source_attributes = (
            (("type", source_type),) if isinstance(source_type, str) else None
        )
        expected_worksheet_source_attributes = (
            tuple(sorted({"ref": source_reference, "sheet": source_sheet}.items()))
            if isinstance(source_reference, str) and isinstance(source_sheet, str)
            else None
        )
        expected_slicer_declaration_attributes = (
            (f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id", "rIdWCABSlicerCache"),
        )
        expected_slicer_extension_attributes = (("uri", "{BBE1A952-AA13-448E-AADC-164F8A28A991}"),)
        expected_slicer_content_type_attributes = (
            ("ContentType", "application/vnd.ms-excel.slicerCache+xml"),
            ("PartName", "/xl/slicerCaches/slicerCache1.xml"),
        )
        expected_slicer_attributes = (
            tuple(sorted({"name": slicer_name, "sourceName": slicer_source_name}.items()))
            if isinstance(slicer_name, str) and isinstance(slicer_source_name, str)
            else None
        )
        expected_slicer_pivot_table_attributes = (
            tuple(
                sorted({"name": slicer_pivot_table_name, "tabId": str(slicer_pivot_tab_id)}.items())
            )
            if isinstance(slicer_pivot_table_name, str) and type(slicer_pivot_tab_id) is int
            else None
        )
        expected_slicer_tabular_attributes = (
            (("pivotCacheId", str(cache_id)),) if type(cache_id) is int else None
        )
        expected_baseline_slicer_items = (
            (
                tuple(sorted({"x": "0", "s": "1"}.items())),
                tuple(sorted({"x": "1", "s": "0"}.items())),
            )
            if baseline_selected_item_index == 0 and candidate_selected_item_index == 1
            else None
        )
        expected_candidate_slicer_items = (
            (
                tuple(sorted({"x": "0", "s": "0"}.items())),
                tuple(sorted({"x": "1", "s": "1"}.items())),
            )
            if baseline_selected_item_index == 0 and candidate_selected_item_index == 1
            else None
        )
        graph = _direct_graph(candidate)
        selection_fields = {
            "slicer_item_attributes",
            "slicer_selected_item_index",
            "slicer_selected_value",
        }
        _assert(
            type(cache_id) is int
            and cache_id > 0
            and isinstance(source_type, str)
            and isinstance(source_sheet, str)
            and isinstance(source_reference, str)
            and isinstance(pivot_sheet, str)
            and isinstance(pivot_reference, str)
            and isinstance(pivot_output_cell, str)
            and isinstance(dashboard_sheet, str)
            and isinstance(dashboard_cell, str)
            and isinstance(dashboard_formula, str)
            and isinstance(slicer_name, str)
            and isinstance(slicer_source_name, str)
            and isinstance(slicer_pivot_table_name, str)
            and type(slicer_pivot_tab_id) is int
            and type(item_count) is int
            and item_count == 2
            and type(baseline_selected_item_index) is int
            and type(candidate_selected_item_index) is int
            and baseline_selected_item_index == 0
            and candidate_selected_item_index == 1
            and isinstance(baseline_selected_value, str)
            and isinstance(candidate_selected_value, str)
            and baseline_selected_value != candidate_selected_value
            and source_sheet in baseline.sheetnames
            and source_sheet in candidate.sheetnames
            and pivot_sheet in baseline.sheetnames
            and pivot_sheet in candidate.sheetnames
            and dashboard_sheet in baseline.sheetnames
            and dashboard_sheet in candidate.sheetnames
            and before_state is not None
            and after_state is not None
            and {key: value for key, value in before_state.items() if key not in selection_fields}
            == {key: value for key, value in after_state.items() if key not in selection_fields}
            and before_state["cache_id"] == cache_id
            and after_state["cache_id"] == cache_id
            and before_state["definition_attributes"] == expected_definition_attributes
            and after_state["definition_attributes"] == expected_definition_attributes
            and before_state["cache_source_attributes"] == expected_cache_source_attributes
            and after_state["cache_source_attributes"] == expected_cache_source_attributes
            and before_state["worksheet_source_attributes"] == expected_worksheet_source_attributes
            and after_state["worksheet_source_attributes"] == expected_worksheet_source_attributes
            and before_state["cache_records_attributes"] == (("count", "4"),)
            and after_state["cache_records_attributes"] == (("count", "4"),)
            and before_state["cache_record_indexes"] == ((0, 0), (0, 1), (1, 2), (1, 3))
            and after_state["cache_record_indexes"] == ((0, 0), (0, 1), (1, 2), (1, 3))
            and before_state["pivot_table_attributes"] == expected_pivot_attributes
            and after_state["pivot_table_attributes"] == expected_pivot_attributes
            and before_state["location_attributes"] == expected_location_attributes
            and after_state["location_attributes"] == expected_location_attributes
            and before_state["pivot_sheet_id"] == slicer_pivot_tab_id
            and after_state["pivot_sheet_id"] == slicer_pivot_tab_id
            and before_state["slicer_member"] == "xl/slicerCaches/slicerCache1.xml"
            and after_state["slicer_member"] == "xl/slicerCaches/slicerCache1.xml"
            and before_state["slicer_declaration_attributes"]
            == expected_slicer_declaration_attributes
            and after_state["slicer_declaration_attributes"]
            == expected_slicer_declaration_attributes
            and before_state["slicer_extension_attributes"] == expected_slicer_extension_attributes
            and after_state["slicer_extension_attributes"] == expected_slicer_extension_attributes
            and before_state["slicer_content_type_attributes"]
            == expected_slicer_content_type_attributes
            and after_state["slicer_content_type_attributes"]
            == expected_slicer_content_type_attributes
            and before_state["slicer_attributes"] == expected_slicer_attributes
            and after_state["slicer_attributes"] == expected_slicer_attributes
            and before_state["slicer_pivot_table_attributes"]
            == expected_slicer_pivot_table_attributes
            and after_state["slicer_pivot_table_attributes"]
            == expected_slicer_pivot_table_attributes
            and before_state["slicer_tabular_attributes"] == expected_slicer_tabular_attributes
            and after_state["slicer_tabular_attributes"] == expected_slicer_tabular_attributes
            and before_state["slicer_source_field_index"] == 0
            and after_state["slicer_source_field_index"] == 0
            and before_state["slicer_source_shared_item_values"]
            == (baseline_selected_value, candidate_selected_value)
            and after_state["slicer_source_shared_item_values"]
            == (baseline_selected_value, candidate_selected_value)
            and before_state["slicer_item_attributes"] == expected_baseline_slicer_items
            and after_state["slicer_item_attributes"] == expected_candidate_slicer_items
            and before_state["slicer_selected_item_index"] == baseline_selected_item_index
            and after_state["slicer_selected_item_index"] == candidate_selected_item_index
            and before_state["slicer_selected_value"] == baseline_selected_value
            and after_state["slicer_selected_value"] == candidate_selected_value
            and before_state["pivot_cache_slicer_extension_attributes"]
            == (("pivotCacheId", str(cache_id)),)
            and after_state["pivot_cache_slicer_extension_attributes"]
            == (("pivotCacheId", str(cache_id)),)
            and before_output is not None
            and after_output is not None
            and before_output == after_output
            and before_output[3] is None
            and before_output[4] is not None
            and before_dashboard is not None
            and after_dashboard is not None
            and before_dashboard == after_dashboard
            and before_dashboard[3] == dashboard_formula
            and _slicer_cache_without_selection_state(baseline_path, pivot_sheet)
            == _slicer_cache_without_selection_state(candidate_path, pivot_sheet)
            and _xlsx_member_differences(baseline_path, candidate_path)
            == {before_state["slicer_member"]}
            and _calculation_properties(baseline_path) == _calculation_properties(candidate_path)
            and (dashboard_sheet, dashboard_cell)
            in graph.get((pivot_sheet, pivot_output_cell), set()),
            f"{truth['id']}: expected one relationship-bound PivotTable Slicer selection to move with stable source, cache, stored output, and dashboard formula",
            errors,
        )
        return

    if kind == "pivot_data_field_aggregation_changed":
        cache_id = fact.get("cache_id")
        source_type = fact.get("source_type")
        source_sheet = fact.get("source_sheet")
        source_reference = fact.get("source_ref")
        pivot_sheet = fact.get("pivot_sheet")
        pivot_reference = fact.get("pivot_ref")
        pivot_output_cell = fact.get("pivot_output_cell")
        dashboard_sheet = fact.get("dashboard_sheet")
        dashboard_cell = fact.get("dashboard_cell")
        dashboard_formula = fact.get("dashboard_formula")
        data_field_source_index = fact.get("data_field_source_index")
        baseline_subtotal = fact.get("baseline_subtotal")
        candidate_subtotal = fact.get("candidate_subtotal")
        before_state = (
            _raw_pivot_data_field_aggregation_state(baseline_path, pivot_sheet)
            if isinstance(pivot_sheet, str)
            else None
        )
        after_state = (
            _raw_pivot_data_field_aggregation_state(candidate_path, pivot_sheet)
            if isinstance(pivot_sheet, str)
            else None
        )
        before_output = (
            _raw_cell_state(baseline_path, pivot_sheet, pivot_output_cell)
            if isinstance(pivot_sheet, str) and isinstance(pivot_output_cell, str)
            else None
        )
        after_output = (
            _raw_cell_state(candidate_path, pivot_sheet, pivot_output_cell)
            if isinstance(pivot_sheet, str) and isinstance(pivot_output_cell, str)
            else None
        )
        before_dashboard = (
            _raw_cell_state(baseline_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        after_dashboard = (
            _raw_cell_state(candidate_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        expected_definition_attributes = tuple(
            sorted(
                {
                    "recordCount": "4",
                    "refreshOnLoad": "0",
                    "saveData": "1",
                    f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id": "rIdWCABPivotRecords",
                }.items()
            )
        )
        expected_pivot_attributes = (
            tuple(
                sorted(
                    {
                        "cacheId": str(cache_id),
                        "dataCaption": "Amount",
                        "name": "WCAB Pivot Report",
                    }.items()
                )
            )
            if type(cache_id) is int
            else None
        )
        expected_location_attributes = (
            tuple(
                sorted(
                    {
                        "ref": pivot_reference,
                        "firstHeaderRow": "1",
                        "firstDataRow": "2",
                        "firstDataCol": "1",
                    }.items()
                )
            )
            if isinstance(pivot_reference, str)
            else None
        )
        expected_cache_source_attributes = (
            (("type", source_type),) if isinstance(source_type, str) else None
        )
        expected_worksheet_source_attributes = (
            tuple(sorted({"ref": source_reference, "sheet": source_sheet}.items()))
            if isinstance(source_reference, str) and isinstance(source_sheet, str)
            else None
        )
        expected_baseline_data_field_attributes = (
            tuple(
                sorted(
                    {
                        "fld": str(data_field_source_index),
                        "subtotal": baseline_subtotal,
                    }.items()
                )
            )
            if type(data_field_source_index) is int and isinstance(baseline_subtotal, str)
            else None
        )
        expected_candidate_data_field_attributes = (
            tuple(
                sorted(
                    {
                        "fld": str(data_field_source_index),
                        "subtotal": candidate_subtotal,
                    }.items()
                )
            )
            if type(data_field_source_index) is int and isinstance(candidate_subtotal, str)
            else None
        )
        graph = _direct_graph(candidate)
        _assert(
            type(cache_id) is int
            and cache_id > 0
            and isinstance(source_type, str)
            and isinstance(source_sheet, str)
            and isinstance(source_reference, str)
            and isinstance(pivot_sheet, str)
            and isinstance(pivot_reference, str)
            and isinstance(pivot_output_cell, str)
            and isinstance(dashboard_sheet, str)
            and isinstance(dashboard_cell, str)
            and isinstance(dashboard_formula, str)
            and type(data_field_source_index) is int
            and data_field_source_index >= 0
            and isinstance(baseline_subtotal, str)
            and isinstance(candidate_subtotal, str)
            and baseline_subtotal != candidate_subtotal
            and source_sheet in baseline.sheetnames
            and source_sheet in candidate.sheetnames
            and pivot_sheet in baseline.sheetnames
            and pivot_sheet in candidate.sheetnames
            and dashboard_sheet in baseline.sheetnames
            and dashboard_sheet in candidate.sheetnames
            and before_state is not None
            and after_state is not None
            and {
                key: value for key, value in before_state.items() if key != "data_field_attributes"
            }
            == {key: value for key, value in after_state.items() if key != "data_field_attributes"}
            and before_state["cache_id"] == cache_id
            and after_state["cache_id"] == cache_id
            and before_state["definition_attributes"] == expected_definition_attributes
            and after_state["definition_attributes"] == expected_definition_attributes
            and before_state["cache_source_attributes"] == expected_cache_source_attributes
            and after_state["cache_source_attributes"] == expected_cache_source_attributes
            and before_state["worksheet_source_attributes"] == expected_worksheet_source_attributes
            and after_state["worksheet_source_attributes"] == expected_worksheet_source_attributes
            and before_state["cache_records_attributes"] == (("count", "4"),)
            and after_state["cache_records_attributes"] == (("count", "4"),)
            and before_state["cache_record_indexes"] == ((0, 0), (0, 1), (1, 2), (1, 3))
            and after_state["cache_record_indexes"] == ((0, 0), (0, 1), (1, 2), (1, 3))
            and before_state["pivot_table_attributes"] == expected_pivot_attributes
            and after_state["pivot_table_attributes"] == expected_pivot_attributes
            and before_state["location_attributes"] == expected_location_attributes
            and after_state["location_attributes"] == expected_location_attributes
            and before_state["data_field_source_index"] == data_field_source_index
            and after_state["data_field_source_index"] == data_field_source_index
            and before_state["data_field_attributes"] == expected_baseline_data_field_attributes
            and after_state["data_field_attributes"] == expected_candidate_data_field_attributes
            and before_output is not None
            and after_output is not None
            and before_output == after_output
            and before_output[3] is None
            and before_output[4] is not None
            and before_dashboard is not None
            and after_dashboard is not None
            and before_dashboard == after_dashboard
            and before_dashboard[3] == dashboard_formula
            and _pivot_table_without_data_field_subtotal(baseline_path, pivot_sheet)
            == _pivot_table_without_data_field_subtotal(candidate_path, pivot_sheet)
            and _xlsx_member_differences(baseline_path, candidate_path)
            == {before_state["pivot_table_member"]}
            and _calculation_properties(baseline_path) == _calculation_properties(candidate_path)
            and (dashboard_sheet, dashboard_cell)
            in graph.get((pivot_sheet, pivot_output_cell), set()),
            f"{truth['id']}: expected one relationship-bound PivotTable value aggregation to change with stable source, cache, stored output, and dashboard formula",
            errors,
        )
        return

    if kind == "pivot_cache_refresh_on_load_changed":
        cache_id = fact.get("cache_id")
        source_type = fact.get("source_type")
        source_sheet = fact.get("source_sheet")
        source_reference = fact.get("source_ref")
        pivot_sheet = fact.get("pivot_sheet")
        pivot_reference = fact.get("pivot_ref")
        pivot_output_cell = fact.get("pivot_output_cell")
        dashboard_sheet = fact.get("dashboard_sheet")
        dashboard_cell = fact.get("dashboard_cell")
        dashboard_formula = fact.get("dashboard_formula")
        before_state = (
            _raw_pivot_cache_refresh_state(baseline_path, pivot_sheet)
            if isinstance(pivot_sheet, str)
            else None
        )
        after_state = (
            _raw_pivot_cache_refresh_state(candidate_path, pivot_sheet)
            if isinstance(pivot_sheet, str)
            else None
        )
        before_output = (
            _raw_cell_state(baseline_path, pivot_sheet, pivot_output_cell)
            if isinstance(pivot_sheet, str) and isinstance(pivot_output_cell, str)
            else None
        )
        after_output = (
            _raw_cell_state(candidate_path, pivot_sheet, pivot_output_cell)
            if isinstance(pivot_sheet, str) and isinstance(pivot_output_cell, str)
            else None
        )
        before_dashboard = (
            _raw_cell_state(baseline_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        after_dashboard = (
            _raw_cell_state(candidate_path, dashboard_sheet, dashboard_cell)
            if isinstance(dashboard_sheet, str) and isinstance(dashboard_cell, str)
            else None
        )
        expected_baseline_definition = tuple(
            sorted(
                {
                    "recordCount": "4",
                    "refreshOnLoad": "0",
                    "saveData": "1",
                    f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id": "rIdWCABPivotRecords",
                }.items()
            )
        )
        expected_candidate_definition = tuple(
            sorted(
                {
                    "recordCount": "4",
                    "refreshOnLoad": "1",
                    "saveData": "1",
                    f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id": "rIdWCABPivotRecords",
                }.items()
            )
        )
        expected_pivot_attributes = (
            tuple(
                sorted(
                    {
                        "cacheId": str(cache_id),
                        "dataCaption": "Amount",
                        "name": "WCAB Pivot Report",
                    }.items()
                )
            )
            if type(cache_id) is int
            else None
        )
        expected_location_attributes = (
            tuple(
                sorted(
                    {
                        "ref": pivot_reference,
                        "firstHeaderRow": "1",
                        "firstDataRow": "2",
                        "firstDataCol": "1",
                    }.items()
                )
            )
            if isinstance(pivot_reference, str)
            else None
        )
        expected_cache_source_attributes = (
            (("type", source_type),) if isinstance(source_type, str) else None
        )
        expected_worksheet_source_attributes = (
            tuple(sorted({"ref": source_reference, "sheet": source_sheet}.items()))
            if isinstance(source_reference, str) and isinstance(source_sheet, str)
            else None
        )
        graph = _direct_graph(candidate)
        _assert(
            type(cache_id) is int
            and cache_id > 0
            and isinstance(source_type, str)
            and isinstance(source_sheet, str)
            and isinstance(source_reference, str)
            and isinstance(pivot_sheet, str)
            and isinstance(pivot_reference, str)
            and isinstance(pivot_output_cell, str)
            and isinstance(dashboard_sheet, str)
            and isinstance(dashboard_cell, str)
            and isinstance(dashboard_formula, str)
            and fact.get("baseline_refresh_on_load") is False
            and fact.get("candidate_refresh_on_load") is True
            and before_state is not None
            and after_state is not None
            and before_state["cache_id"] == cache_id
            and after_state["cache_id"] == cache_id
            and before_state["definition_member"] == after_state["definition_member"]
            and before_state["definition_attributes"] == expected_baseline_definition
            and after_state["definition_attributes"] == expected_candidate_definition
            and before_state["cache_source_attributes"] == expected_cache_source_attributes
            and after_state["cache_source_attributes"] == expected_cache_source_attributes
            and before_state["worksheet_source_attributes"] == expected_worksheet_source_attributes
            and after_state["worksheet_source_attributes"] == expected_worksheet_source_attributes
            and before_state["cache_records_member"] == after_state["cache_records_member"]
            and before_state["cache_records_attributes"] == (("count", "4"),)
            and after_state["cache_records_attributes"] == (("count", "4"),)
            and before_state["cache_record_indexes"] == ((0, 0), (0, 1), (1, 2), (1, 3))
            and after_state["cache_record_indexes"] == ((0, 0), (0, 1), (1, 2), (1, 3))
            and before_state["report_member"] == after_state["report_member"]
            and before_state["pivot_table_member"] == after_state["pivot_table_member"]
            and before_state["pivot_table_attributes"] == expected_pivot_attributes
            and after_state["pivot_table_attributes"] == expected_pivot_attributes
            and before_state["location_attributes"] == expected_location_attributes
            and after_state["location_attributes"] == expected_location_attributes
            and before_output is not None
            and after_output is not None
            and before_output == after_output
            and before_output[3] is None
            and before_output[4] is not None
            and before_dashboard is not None
            and after_dashboard is not None
            and before_dashboard == after_dashboard
            and before_dashboard[3] == dashboard_formula
            and _pivot_cache_definition_without_refresh_on_load(baseline_path, pivot_sheet)
            == _pivot_cache_definition_without_refresh_on_load(candidate_path, pivot_sheet)
            and _xlsx_member_differences(baseline_path, candidate_path)
            == {before_state["definition_member"]}
            and _calculation_properties(baseline_path) == _calculation_properties(candidate_path)
            and (dashboard_sheet, dashboard_cell)
            in graph.get((pivot_sheet, pivot_output_cell), set()),
            f"{truth['id']}: expected one relationship-bound PivotCache refresh-on-open false -> true only with stable stored report/dashboard cells",
            errors,
        )
        return

    if kind == "external_data_connection_refresh_on_load_changed":
        connection_id = fact.get("connection_id")
        before_refresh_on_load = (
            _external_data_connection_refresh_on_load(baseline_path, connection_id)
            if isinstance(connection_id, int)
            else None
        )
        after_refresh_on_load = (
            _external_data_connection_refresh_on_load(candidate_path, connection_id)
            if isinstance(connection_id, int)
            else None
        )
        _assert(
            isinstance(connection_id, int)
            and fact.get("baseline_refresh_on_load") is False
            and fact.get("candidate_refresh_on_load") is True
            and before_refresh_on_load is fact.get("baseline_refresh_on_load")
            and after_refresh_on_load is fact.get("candidate_refresh_on_load"),
            f"{truth['id']}: expected connection {connection_id!r} refresh-on-open false -> true",
            errors,
        )
        return

    if kind == "external_workbook_link_update_policy_changed":
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        formula = fact.get("formula")
        before_properties = _workbook_properties(baseline_path)
        after_properties = _workbook_properties(candidate_path)
        before_formula = (
            baseline[sheet_name][coordinate]
            if isinstance(sheet_name, str)
            and isinstance(coordinate, str)
            and sheet_name in baseline.sheetnames
            else None
        )
        after_formula = (
            candidate[sheet_name][coordinate]
            if isinstance(sheet_name, str)
            and isinstance(coordinate, str)
            and sheet_name in candidate.sheetnames
            else None
        )
        before_without_policy = (
            {key: value for key, value in before_properties.items() if key != "updateLinks"}
            if before_properties is not None
            else None
        )
        after_without_policy = (
            {key: value for key, value in after_properties.items() if key != "updateLinks"}
            if after_properties is not None
            else None
        )
        _assert(
            isinstance(sheet_name, str)
            and isinstance(coordinate, str)
            and isinstance(formula, str)
            and "[" in formula
            and before_formula is not None
            and after_formula is not None
            and _cell_kind(before_formula) == "formula"
            and _cell_kind(after_formula) == "formula"
            and before_formula.value == formula
            and after_formula.value == formula
            and fact.get("baseline_update_links") == "never"
            and fact.get("candidate_update_links") == "always"
            and before_properties is not None
            and after_properties is not None
            and before_properties.get("updateLinks") == fact.get("baseline_update_links")
            and after_properties.get("updateLinks") == fact.get("candidate_update_links")
            and before_without_policy == after_without_policy,
            f"{truth['id']}: expected unchanged {sheet_name}!{coordinate} external formula and workbook updateLinks never -> always only",
            errors,
        )
        return

    if kind == "array_formula_mode_changed":
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        before = (
            _array_formula_state(baseline_path, sheet_name, coordinate)
            if isinstance(sheet_name, str) and isinstance(coordinate, str)
            else None
        )
        after = (
            _array_formula_state(candidate_path, sheet_name, coordinate)
            if isinstance(sheet_name, str) and isinstance(coordinate, str)
            else None
        )
        expected_before = (
            fact.get("baseline_mode"),
            fact.get("formula"),
            fact.get("baseline_output_range"),
        )
        expected_after = (
            fact.get("candidate_mode"),
            fact.get("formula"),
            fact.get("candidate_output_range"),
        )
        _assert(
            before == expected_before
            and after == expected_after
            and fact.get("baseline_mode") == "legacy_cse"
            and fact.get("candidate_mode") == "dynamic",
            f"{truth['id']}: expected unchanged {sheet_name}!{coordinate} formula to switch legacy CSE -> dynamic array mode",
            errors,
        )
        return

    if kind == "static_cycle_introduced":
        graph = _direct_graph(candidate)
        cells = [(cell["sheet"], cell["cell"]) for cell in fact.get("cells", [])]
        for cell in cells:
            _assert(
                cell in _reachable(graph, cell),
                f"{truth['id']}: expected a static cycle through {cell[0]}!{cell[1]}",
                errors,
            )
        return

    if kind == "three_d_scope_changed":
        formula_sheet = fact.get("formula_sheet")
        formula_cell = fact.get("formula_cell")
        _assert(
            baseline[formula_sheet][formula_cell].value
            == candidate[formula_sheet][formula_cell].value,
            f"{truth['id']}: 3-D formula text should remain unchanged",
            errors,
        )
        order = candidate.sheetnames
        try:
            after_index = order.index(fact["after_sheet"])
            inserted_index = order.index(fact["inserted_sheet"])
            before_index = order.index(fact["before_sheet"])
        except ValueError:
            errors.append(f"{truth['id']}: missing declared 3-D scope sheet")
            return
        _assert(
            after_index < inserted_index < before_index,
            f"{truth['id']}: inserted sheet is not inside the declared 3-D span",
            errors,
        )
        return

    if kind == "structural_formula_rewrite":
        before = fact.get("baseline", {})
        after = fact.get("candidate", {})
        _assert(
            baseline[before["sheet"]][before["cell"]].value == before["formula"]
            and candidate[after["sheet"]][after["cell"]].value == after["formula"],
            f"{truth['id']}: declared structural formula rewrite is not present",
            errors,
        )
        return

    if kind == "structured_table_scope_changed":
        table_sheet = fact.get("table_sheet")
        table_name = fact.get("table")
        formula_sheet = fact.get("formula_sheet")
        formula_cell = fact.get("formula_cell")
        if (
            table_sheet not in baseline.sheetnames
            or table_sheet not in candidate.sheetnames
            or formula_sheet not in baseline.sheetnames
            or formula_sheet not in candidate.sheetnames
        ):
            errors.append(f"{truth['id']}: declared table or formula sheet is absent")
            return
        try:
            before_table = baseline[table_sheet].tables[table_name]
            after_table = candidate[table_sheet].tables[table_name]
        except KeyError:
            errors.append(f"{truth['id']}: declared Excel Table {table_name!r} is absent")
            return
        before_formula = baseline[formula_sheet][formula_cell].value
        after_formula = candidate[formula_sheet][formula_cell].value
        _assert(
            before_table.ref == fact.get("baseline_ref")
            and after_table.ref == fact.get("candidate_ref")
            and before_formula == after_formula
            and f"{table_name}[" in str(after_formula),
            f"{truth['id']}: expected structured-reference formula to retain text while table scope changes",
            errors,
        )
        return

    if kind == "portfolio_value_changed":
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        before = baseline[sheet_name][coordinate]
        after = candidate[sheet_name][coordinate]
        _assert(
            _cell_kind(before) == "value"
            and _cell_kind(after) == "value"
            and before.value != after.value,
            f"{truth['id']}: expected portfolio value change at {sheet_name}!{coordinate}",
            errors,
        )
        return

    if kind == "portfolio_external_reference":
        sheet_name = fact.get("sheet")
        coordinate = fact.get("cell")
        value = candidate[sheet_name][coordinate].value
        target = fact.get("target_workbook")
        _assert(
            _cell_kind(candidate[sheet_name][coordinate]) == "formula"
            and f"[{target}]" in str(value),
            f"{truth['id']}: expected local external reference to {target!r}",
            errors,
        )
        return

    errors.append(f"{truth['id']}: unsupported fact kind {kind!r}")


def _validate_coverage_expectations(
    case_dir: Path, truth: dict[str, Any], errors: list[str]
) -> None:
    expectations = truth.get("coverage_expectations")
    if not isinstance(expectations, list):
        errors.append(f"{truth['id']}: coverage_expectations must be an array")
        return
    for expectation in expectations:
        if not isinstance(expectation, dict):
            errors.append(f"{truth['id']}: coverage expectation must be an object")
            continue
        kind = expectation.get("kind")
        if kind not in {
            "dynamic_reference_static_coverage",
            "dynamic_reference_driver_changed",
        }:
            errors.append(f"{truth['id']}: unsupported coverage expectation kind {kind!r}")
            continue
        try:
            baseline_path, candidate_path = _workbook_pair(case_dir, truth)
        except FixtureValidationError as error:
            errors.append(str(error))
            continue
        baseline = _load_workbook(baseline_path)
        candidate = _load_workbook(candidate_path)
        expected_functions = expectation.get("functions")
        if kind == "dynamic_reference_static_coverage":
            sheet_name = expectation.get("sheet")
            coordinate = expectation.get("cell")
            if (
                not isinstance(sheet_name, str)
                or not isinstance(coordinate, str)
                or sheet_name not in baseline.sheetnames
                or sheet_name not in candidate.sheetnames
            ):
                errors.append(f"{truth['id']}: dynamic-reference coverage location is absent")
                continue
            before = baseline[sheet_name][coordinate]
            after = candidate[sheet_name][coordinate]
            _assert(
                _cell_kind(before) == "formula"
                and _cell_kind(after) == "formula"
                and not _dynamic_reference_functions(str(before.value))
                and isinstance(expected_functions, list)
                and all(isinstance(function, str) for function in expected_functions)
                and _dynamic_reference_functions(str(after.value)) == tuple(expected_functions),
                f"{truth['id']}: expected static-dependency coverage boundary at {sheet_name}!{coordinate}",
                errors,
            )
            continue

        driver = expectation.get("driver")
        formula = expectation.get("formula")
        if not isinstance(driver, dict) or not isinstance(formula, dict):
            errors.append(
                f"{truth['id']}: dynamic-reference driver coverage needs driver and formula"
            )
            continue
        driver_sheet = driver.get("sheet")
        driver_cell = driver.get("cell")
        formula_sheet = formula.get("sheet")
        formula_cell = formula.get("cell")
        locations = (driver_sheet, driver_cell, formula_sheet, formula_cell)
        if (
            not all(isinstance(location, str) for location in locations)
            or driver_sheet not in baseline.sheetnames
            or driver_sheet not in candidate.sheetnames
            or formula_sheet not in baseline.sheetnames
            or formula_sheet not in candidate.sheetnames
        ):
            errors.append(f"{truth['id']}: dynamic-reference driver coverage location is absent")
            continue
        before_driver = baseline[driver_sheet][driver_cell]
        after_driver = candidate[driver_sheet][driver_cell]
        before_formula = baseline[formula_sheet][formula_cell]
        after_formula = candidate[formula_sheet][formula_cell]
        graph = _direct_graph(candidate)
        _assert(
            _cell_kind(before_driver) == "value"
            and _cell_kind(after_driver) == "value"
            and before_driver.value != after_driver.value
            and _cell_kind(before_formula) == "formula"
            and _cell_kind(after_formula) == "formula"
            and before_formula.value == after_formula.value
            and isinstance(expected_functions, list)
            and all(isinstance(function, str) for function in expected_functions)
            and _dynamic_reference_functions(str(after_formula.value)) == tuple(expected_functions)
            and (formula_sheet, formula_cell) in graph.get((driver_sheet, driver_cell), set()),
            f"{truth['id']}: expected changed dynamic-reference driver {driver_sheet}!{driver_cell} to feed unchanged formula {formula_sheet}!{formula_cell}",
            errors,
        )


def _validate_impacts(case_dir: Path, truth: dict[str, Any], errors: list[str]) -> None:
    if truth.get("topology") != "pair":
        return
    candidate = _load_workbook(case_dir / "candidate.xlsx")
    graph = _direct_graph(candidate)
    for impact in truth.get("must_reach", []):
        source = impact.get("source", {})
        source_key = (source.get("sheet"), source.get("cell"))
        reachable = _reachable(graph, source_key)
        for target in impact.get("targets", []):
            target_key = (target.get("sheet"), target.get("cell"))
            _assert(
                target_key in reachable,
                f"{truth['id']}: static impact {source_key[0]}!{source_key[1]} -> {target_key[0]}!{target_key[1]} is absent",
                errors,
            )


def validate_case(case_dir: str | Path) -> list[str]:
    """Validate one case directory and return its errors without raising."""

    directory = Path(case_dir)
    truth_path = directory / "truth.json"
    truth = _load_truth(truth_path)
    errors: list[str] = []
    _assert(
        truth.get("schema_version") == FIXTURE_SCHEMA_VERSION,
        f"{directory}: unsupported schema version {truth.get('schema_version')!r}",
        errors,
    )
    _assert(
        truth.get("id") in CASE_IDS,
        f"{directory}: unknown case id {truth.get('id')!r}",
        errors,
    )
    _assert(
        truth.get("review_expectation") in {"allow", "review", "block"},
        f"{truth.get('id', directory)}: invalid review expectation",
        errors,
    )
    _assert(
        isinstance(truth.get("coverage_expectations"), list),
        f"{truth.get('id', directory)}: coverage_expectations must be an array",
        errors,
    )
    if truth.get("topology") == "pair":
        for filename in ("baseline.xlsx", "candidate.xlsx"):
            _assert(
                (directory / filename).is_file(), f"{truth.get('id')}: missing {filename}", errors
            )
    elif truth.get("topology") == "portfolio":
        _assert(
            (directory / "baseline").is_dir(),
            f"{truth.get('id')}: missing baseline portfolio",
            errors,
        )
        _assert(
            (directory / "candidate").is_dir(),
            f"{truth.get('id')}: missing candidate portfolio",
            errors,
        )
    else:
        errors.append(f"{truth.get('id', directory)}: invalid topology {truth.get('topology')!r}")
        return errors

    for fact in truth.get("facts", []):
        if not isinstance(fact, dict):
            errors.append(f"{truth.get('id')}: fact must be an object")
            continue
        _validate_fact(directory, truth, fact, errors)
    _validate_coverage_expectations(directory, truth, errors)
    _validate_impacts(directory, truth, errors)
    return errors


def validate_all(root: str | Path) -> dict[str, list[str]]:
    """Validate all fixture manifests below *root*.

    Returns a mapping of case IDs to validation errors.  An empty mapping means
    the entire tree satisfies its documented truth contract.
    """

    fixture_root = Path(root)
    manifests = sorted(fixture_root.rglob("truth.json"))
    if len(manifests) != len(CASE_IDS):
        return {"_catalog": [f"expected {len(CASE_IDS)} truth manifests, found {len(manifests)}"]}
    failures: dict[str, list[str]] = {}
    seen: set[str] = set()
    for manifest in manifests:
        truth = _load_truth(manifest)
        case_id = str(truth.get("id", manifest.parent))
        seen.add(case_id)
        errors = validate_case(manifest.parent)
        if errors:
            failures[case_id] = errors
    missing = sorted(set(CASE_IDS) - seen)
    if missing:
        failures["_catalog"] = [f"missing case IDs: {', '.join(missing)}"]
    return failures
