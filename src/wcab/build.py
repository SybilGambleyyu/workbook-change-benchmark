"""Deterministic generation of the WCAB fixture tree.

The fixtures are intentionally generated rather than copied from real business
workbooks.  That makes every assertion inspectable and the redistribution
rights unambiguous.
"""

from __future__ import annotations

import base64
import io
import json
import re
import struct
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.worksheet.table import Table, TableStyleInfo

from .manifest import write_manifest

FIXTURE_SCHEMA_VERSION = 3
_FIXED_TIMESTAMP = datetime(2024, 1, 1, tzinfo=timezone.utc)
_CORE_MODIFIED = re.compile(rb"(<dcterms:modified\b[^>]*>)[^<]*(</dcterms:modified>)")
WorkbookFactory = Callable[[], Workbook]
WorkbookMutator = Callable[[Workbook], None]
ArchiveMutator = Callable[[dict[str, bytes]], None]
_SPREADSHEETML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOCUMENT_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_DYNAMIC_ARRAY_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
_EXTERNAL_WORKBOOK_LINK_FORMULA = "='[WCABSource.xlsx]Inputs'!$B$2"
_ITERATIVE_CALCULATION_FORMULA = "=(B2+Inputs!$B$2)/2"
_ITERATION_COUNT = 100
_ITERATION_DELTA = 0.001
_PRECISION_AS_DISPLAYED_INPUT = 10.005
_PRECISION_AS_DISPLAYED_NUMBER_FORMAT = "0.00"
_PRECISION_AS_DISPLAYED_FORMULA = "=Inputs!$B$2*2"
_FORMULA_CACHED_RESULT_INPUT = 10
_FORMULA_CACHED_RESULT_FORMULA = "=Inputs!$B$2*2"
_FORMULA_CACHED_RESULT_BASELINE = 20
_FORMULA_CACHED_RESULT_CANDIDATE = 25
_WORKBOOK_DATE_SYSTEM_SERIAL = 45292
_WORKBOOK_DATE_SYSTEM_NUMBER_FORMAT = "yyyy-mm-dd"
_WORKBOOK_DATE_SYSTEM_FORMULA = "=Inputs!$B$2+30"
_WORKBOOK_DATE_SYSTEM_DASHBOARD_FORMULA = "=Model!$B$2"
_AUTO_FILTER_REF = "A1:B5"
_AUTO_FILTER_COLUMN_ID = 0
_AUTO_FILTER_BASELINE_VALUE = "North"
_AUTO_FILTER_CANDIDATE_VALUE = "South"
_AUTO_FILTER_SUBTOTAL_FORMULA = "=SUBTOTAL(109,B2:B5)"
_AUTO_FILTER_DASHBOARD_FORMULA = "=Report!$D$2"
_PIVOT_CACHE_ID = 1
_PIVOT_CACHE_SOURCE_SHEET = "Source"
_PIVOT_CACHE_SOURCE_REF = "A1:B5"
_PIVOT_REPORT_SHEET = "Report"
_PIVOT_REPORT_REF = "A1:B2"
_PIVOT_CACHE_DASHBOARD_FORMULA = "=Report!$B$2"
_PIVOT_DATA_FIELD_SOURCE_INDEX = 1
_PIVOT_DATA_FIELD_BASELINE_SUBTOTAL = "sum"
_PIVOT_DATA_FIELD_CANDIDATE_SUBTOTAL = "average"
_OFFICE_2010_SPREADSHEET_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
_PIVOT_SLICER_CACHE_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
)
_PIVOT_SLICER_CACHE_CONTENT_TYPE = "application/vnd.ms-excel.slicerCache+xml"
_PIVOT_SLICER_CACHE_EXTENSION_URI = "{BBE1A952-AA13-448E-AADC-164F8A28A991}"
_PIVOT_SLICER_PIVOT_CACHE_EXTENSION_URI = "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}"
_PIVOT_SLICER_NAME = "WCAB Region slicer"
_PIVOT_SLICER_SOURCE_NAME = "Region"
_PIVOT_SLICER_PIVOT_TABLE_NAME = "WCAB Pivot Report"
_PIVOT_SLICER_PIVOT_TAB_ID = 2
_PIVOT_SLICER_ITEM_COUNT = 2
_PIVOT_SLICER_BASELINE_SELECTED_INDEX = 0
_PIVOT_SLICER_CANDIDATE_SELECTED_INDEX = 1
_DATA_MASHUP_NS = "http://schemas.microsoft.com/DataMashup"
_POWER_QUERY_CUSTOM_XML_MEMBER = "customXml/item1.xml"
_POWER_QUERY_ROOT_RELATIONSHIP_ID = "rIdWCABPowerQuery"
_POWER_QUERY_SOURCE_SHEET = "Source"
_POWER_QUERY_SOURCE_TABLE = "SourceData"
_POWER_QUERY_SOURCE_REF = "A1:B5"
_POWER_QUERY_SECTION = "Section1"
_POWER_QUERY_NAME = "RegionQuery"
_POWER_QUERY_FILTER_COLUMN = "Region"
_POWER_QUERY_BASELINE_FILTER_VALUE = "North"
_POWER_QUERY_CANDIDATE_FILTER_VALUE = "South"
_SCENARIO_MANAGER_SHEET = "Inputs"
_SCENARIO_MANAGER_DASHBOARD_SHEET = "Dashboard"
_SCENARIO_MANAGER_CHANGING_CELL = "B2"
_SCENARIO_MANAGER_STABLE_INPUT_CELL = "B3"
_SCENARIO_MANAGER_RESULT_CELL = "D2"
_SCENARIO_MANAGER_DASHBOARD_CELL = "B4"
_SCENARIO_MANAGER_WORKSHEET_INPUT_VALUE = 0.1
_SCENARIO_MANAGER_WORKSHEET_STABLE_INPUT_VALUE = 125
_SCENARIO_MANAGER_FORMULA = "=B2*B3"
_SCENARIO_MANAGER_DASHBOARD_FORMULA = "=Inputs!$D$2"
_SCENARIO_MANAGER_NAME = "WCAB downside"
_SCENARIO_MANAGER_COMMENT = "Synthetic downside assumption set"
_SCENARIO_MANAGER_USER = "WCAB"
_SCENARIO_MANAGER_BASELINE_STORED_VALUE = "0.08"
_SCENARIO_MANAGER_CANDIDATE_STORED_VALUE = "0.16"
_SCENARIO_MANAGER_STABLE_STORED_VALUE = "125"
_SCENARIO_MANAGER_INPUT_NUMBER_FORMAT_ID = 10
_WHAT_IF_DATA_TABLE_SHEET = "Sensitivity"
_WHAT_IF_DATA_TABLE_MODEL_SHEET = "Model"
_WHAT_IF_DATA_TABLE_DASHBOARD_SHEET = "Dashboard"
_WHAT_IF_DATA_TABLE_MASTER_CELL = "D3"
_WHAT_IF_DATA_TABLE_OUTPUT_RANGE = "D3:D5"
_WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL = "B2"
_WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_CELL = "B3"
_WHAT_IF_DATA_TABLE_SCALE_CELL = "B4"
_WHAT_IF_DATA_TABLE_INPUT_VALUE_RANGE = "C3:C5"
_WHAT_IF_DATA_TABLE_PRIMARY_INPUT_VALUE = 0.08
_WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_VALUE = 0.12
_WHAT_IF_DATA_TABLE_SCALE_VALUE = 100
_WHAT_IF_DATA_TABLE_GRID_VALUES = (0.04, 0.08, 0.12)
_WHAT_IF_DATA_TABLE_OUTPUT_FORMULA_CELL = "D2"
_WHAT_IF_DATA_TABLE_OUTPUT_FORMULA = "=Model!$B$2"
_WHAT_IF_DATA_TABLE_MODEL_CELL = "B2"
_WHAT_IF_DATA_TABLE_MODEL_FORMULA = "=Sensitivity!$B$2*Sensitivity!$B$3*Sensitivity!$B$4"
_WHAT_IF_DATA_TABLE_DASHBOARD_CELL = "B4"
_WHAT_IF_DATA_TABLE_DASHBOARD_FORMULA = "=Model!$B$2"
_DATA_VALIDATION_LIST_SHEET = "Inputs"
_DATA_VALIDATION_LIST_SOURCE_SHEET = "Lists"
_DATA_VALIDATION_LIST_MODEL_SHEET = "Model"
_DATA_VALIDATION_LIST_DASHBOARD_SHEET = "Dashboard"
_DATA_VALIDATION_LIST_TARGET_RANGE = "B2"
_DATA_VALIDATION_LIST_BASELINE_SOURCE_FORMULA = "=Lists!$A$2:$A$4"
_DATA_VALIDATION_LIST_CANDIDATE_SOURCE_FORMULA = "=Lists!$B$2:$B$4"
_DATA_VALIDATION_LIST_BASELINE_SOURCE_RANGE = "A2:A4"
_DATA_VALIDATION_LIST_CANDIDATE_SOURCE_RANGE = "B2:B4"
_DATA_VALIDATION_LIST_BASELINE_SOURCE_VALUES = ("Draft", "Review", "Approved")
_DATA_VALIDATION_LIST_CANDIDATE_SOURCE_VALUES = ("Draft", "Suspended", "Rejected")
_DATA_VALIDATION_LIST_INPUT_VALUE = "Draft"
_DATA_VALIDATION_LIST_MODEL_CELL = "B2"
_DATA_VALIDATION_LIST_MODEL_FORMULA = "=Inputs!$B$2"
_DATA_VALIDATION_LIST_DASHBOARD_CELL = "B4"
_DATA_VALIDATION_LIST_DASHBOARD_FORMULA = "=Model!$B$2"
_DATA_VALIDATION_LIST_ERROR_STYLE = "stop"
_DATA_VALIDATION_LIST_ERROR_TITLE = "Invalid status"
_DATA_VALIDATION_LIST_ERROR = "Choose an approved status."
_DATA_VALIDATION_LIST_PROMPT_TITLE = "Approved status"
_DATA_VALIDATION_LIST_PROMPT = "Choose a documented status."
_CONDITIONAL_FORMATTING_THRESHOLD_SHEET = "Operations"
_CONDITIONAL_FORMATTING_THRESHOLD_RANGE = "B2:B4"
_CONDITIONAL_FORMATTING_THRESHOLD_BASELINE_FORMULA = "100"
_CONDITIONAL_FORMATTING_THRESHOLD_CANDIDATE_FORMULA = "50"
_CONDITIONAL_FORMATTING_THRESHOLD_VALUES = (10, 75, 120)
_CONDITIONAL_FORMATTING_THRESHOLD_FILL_RGB = "FFFFC7CE"
_NUMBER_FORMAT_VISIBILITY_SHEET = "Operations"
_NUMBER_FORMAT_VISIBILITY_CELL = "B2"
_NUMBER_FORMAT_VISIBILITY_VALUE = 0.125
_NUMBER_FORMAT_VISIBILITY_BASELINE_FORMAT = "0.0%;[Red](0.0%);-"
_NUMBER_FORMAT_VISIBILITY_CANDIDATE_FORMAT = ";;;"
_NUMBER_FORMAT_VISIBILITY_CUSTOM_ID = 164
_NUMBER_FORMAT_VISIBILITY_FORMULA_CELL = "B3"
_NUMBER_FORMAT_VISIBILITY_FORMULA = "=B2"
_CHART_SERIES_SOURCE_SHEET = "Source"
_CHART_SERIES_DASHBOARD_SHEET = "Dashboard"
_CHART_SERIES_ANCHOR = "D2"
_CHART_SERIES_TITLE_REFERENCE = "'Source'!B1"
_CHART_SERIES_CATEGORY_REFERENCE = "'Source'!$A$2:$A$4"
_CHART_SERIES_BASELINE_VALUE_REFERENCE = "'Source'!$B$2:$B$4"
_CHART_SERIES_CANDIDATE_VALUE_REFERENCE = "'Source'!$C$2:$C$4"
_DRAWINGML_CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _configure_workbook(workbook: Workbook, *, title: str) -> None:
    workbook.properties.creator = "WCAB"
    workbook.properties.lastModifiedBy = "WCAB"
    workbook.properties.title = title
    workbook.properties.subject = "Synthetic workbook change-assurance fixture"
    workbook.properties.created = _FIXED_TIMESTAMP
    workbook.properties.modified = _FIXED_TIMESTAMP
    workbook.calculation.fullCalcOnLoad = False
    workbook.calculation.forceFullCalc = False


def _write_canonical_xlsx_members(path: Path, members: dict[str, bytes]) -> None:
    """Write OOXML package members with stable ordering and metadata."""

    core_properties = members.get("docProps/core.xml")
    if core_properties is not None:
        members["docProps/core.xml"] = _CORE_MODIFIED.sub(
            rb"\g<1>2024-01-01T00:00:00Z\g<2>", core_properties
        )
    with NamedTemporaryFile(dir=path.parent, prefix=f".{path.name}.", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED, compresslevel=9) as target:
            for name in sorted(members):
                info = ZipInfo(filename=name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                info.create_system = 3
                info.external_attr = 0o100644 << 16
                target.writestr(info, members[name])
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def _canonicalize_xlsx(path: Path) -> None:
    """Rewrite a generated XLSX with stable ZIP order and timestamps.

    Openpyxl's XML generation is deterministic for the supported dependency
    versions, but ZIP member timestamps otherwise make byte-for-byte fixture
    reproduction impossible.
    """

    with ZipFile(path, "r") as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
    _write_canonical_xlsx_members(path, members)


def _rewrite_xlsx_parts(path: Path, mutate: ArchiveMutator) -> None:
    """Apply a deterministic, raw-OOXML fixture mutation to an XLSX package."""

    with ZipFile(path, "r") as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
    mutate(members)
    _write_canonical_xlsx_members(path, members)


def _save_workbook(workbook: Workbook, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    _canonicalize_xlsx(path)


def _write_json(path: Path, value: dict[str, Any]) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _truth(
    *,
    case_id: str,
    title: str,
    family: str,
    review_expectation: str,
    facts: list[dict[str, Any]],
    must_reach: list[dict[str, Any]] | None = None,
    coverage: list[str] | None = None,
    coverage_expectations: list[dict[str, Any]] | None = None,
    topology: str = "pair",
) -> dict[str, Any]:
    return {
        "schema_version": FIXTURE_SCHEMA_VERSION,
        "id": case_id,
        "title": title,
        "family": family,
        "topology": topology,
        "review_expectation": review_expectation,
        "facts": facts,
        "must_reach": must_reach or [],
        "coverage": coverage or [],
        "coverage_expectations": coverage_expectations or [],
    }


def _write_pair(
    directory: Path,
    factory: WorkbookFactory,
    mutate_candidate: WorkbookMutator,
    truth: dict[str, Any],
) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    _save_workbook(factory(), directory / "baseline.xlsx")
    candidate = factory()
    mutate_candidate(candidate)
    _save_workbook(candidate, directory / "candidate.xlsx")
    _write_json(directory / "truth.json", truth)


def _finance_workbook() -> Workbook:
    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB finance forecast")
    assumptions = workbook.active
    assumptions.title = "Assumptions"
    assumptions.append(["Driver", "2024", "2025", "2026"])
    assumptions.append(["Units", 100, 115, 130])
    assumptions.append(["Price", 8.5, 8.7, 9.0])
    assumptions.append(["Cost per unit", 3.2, 3.4, 3.5])
    assumptions["B10"] = "Discount rate"
    assumptions["E10"] = 0.10
    assumptions["B11"] = "Stress discount rate"
    assumptions["E11"] = 0.12

    revenue = workbook.create_sheet("Revenue")
    revenue.append(["Metric", "2024", "2025", "2026"])
    for column, assumption_column in zip(("B", "C", "D"), ("B", "C", "D"), strict=True):
        revenue[f"{column}5"] = f"=Assumptions!{assumption_column}2"
        revenue[f"{column}6"] = f"=Assumptions!{assumption_column}3"
        revenue[f"{column}8"] = f"={column}5*{column}6"
        revenue[f"{column}9"] = f"={column}5*Assumptions!{assumption_column}4"
        revenue[f"{column}10"] = f"={column}8-{column}9"
    revenue["A5"] = "Units"
    revenue["A6"] = "Price"
    revenue["A8"] = "Revenue"
    revenue["A9"] = "Cost of goods"
    revenue["A10"] = "Gross profit"

    summary = workbook.create_sheet("Summary")
    summary.append(["Metric", "2024", "2025", "2026"])
    for column in ("B", "C", "D"):
        summary[f"{column}5"] = f"=Revenue!{column}8"
        summary[f"{column}6"] = f"=Revenue!{column}10"
        summary[f"{column}8"] = f"=Assumptions!{column}2"
        summary[f"{column}10"] = f"={column}6*(1-DiscountRate)"
    summary["A5"] = "Revenue"
    summary["A6"] = "Gross profit"
    summary["A8"] = "Units source"
    summary["A10"] = "Discounted gross profit"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Board dashboard"
    dashboard["A4"] = "2025 gross profit"
    dashboard["B4"] = "=Summary!C6"
    dashboard["A5"] = "2025 discounted gross profit"
    dashboard["B5"] = "=Summary!C10"

    workbook.defined_names.add(
        DefinedName("DiscountRate", attr_text="Assumptions!$E$10", comment="WCAB synthetic driver")
    )
    return workbook


def _operations_workbook() -> Workbook:
    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB operations tracker")
    sheet = workbook.active
    sheet.title = "Operations"
    sheet.append(
        ["Region", "SKU", "Units", "Unit cost", "Unit price", "Margin", "Status", "Regional sales"]
    )
    rows = [
        ("North", "A-100", 12, 4, 9),
        ("North", "B-200", 10, 5, 11),
        ("South", "A-100", 9, 4, 9),
        ("South", "B-200", 16, 5, 11),
    ]
    for row_number, (region, sku, units, cost, price) in enumerate(rows, start=6):
        sheet.cell(row_number, 1, region)
        sheet.cell(row_number, 2, sku)
        sheet.cell(row_number, 3, units)
        sheet.cell(row_number, 4, cost)
        sheet.cell(row_number, 5, price)
        sheet.cell(row_number, 6, f"=C{row_number}*(E{row_number}-D{row_number})")
        sheet.cell(row_number, 8, f"=SUMIFS($F$6:$F$9,$A$6:$A$9,A{row_number})")

    validation = DataValidation(type="whole", operator="between", formula1="0", formula2="1000")
    validation.promptTitle = "Units"
    validation.prompt = "Enter a whole number between 0 and 1000."
    sheet.add_data_validation(validation)
    validation.add("C6:C9")
    high_margin = PatternFill(fill_type="solid", fgColor="C6EFCE")
    sheet.conditional_formatting.add(
        "F6:F9", CellIsRule(operator="greaterThan", formula=["50"], fill=high_margin)
    )
    return workbook


def _data_validation_list_source_workbook() -> Workbook:
    """Build a list-controlled input with two stable local source lists.

    The candidate is later changed at the raw SpreadsheetML ``formula1``
    declaration only. The source lists, current input, and ordinary formula
    path remain unchanged so this fixture records a future entry-control
    boundary rather than a calculated model result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB list validation source fixture")
    inputs = workbook.active
    inputs.title = _DATA_VALIDATION_LIST_SHEET
    inputs["A1"] = "Status input"
    inputs["A2"] = "Current status"
    inputs[_DATA_VALIDATION_LIST_TARGET_RANGE] = _DATA_VALIDATION_LIST_INPUT_VALUE
    validation = DataValidation(
        type="list",
        formula1=_DATA_VALIDATION_LIST_BASELINE_SOURCE_FORMULA,
        allow_blank=False,
        showErrorMessage=True,
        errorStyle=_DATA_VALIDATION_LIST_ERROR_STYLE,
        errorTitle=_DATA_VALIDATION_LIST_ERROR_TITLE,
        error=_DATA_VALIDATION_LIST_ERROR,
        showInputMessage=False,
        promptTitle=_DATA_VALIDATION_LIST_PROMPT_TITLE,
        prompt=_DATA_VALIDATION_LIST_PROMPT,
    )
    validation.showDropDown = False
    validation.add(_DATA_VALIDATION_LIST_TARGET_RANGE)
    inputs.add_data_validation(validation)

    lists = workbook.create_sheet(_DATA_VALIDATION_LIST_SOURCE_SHEET)
    lists["A1"] = "Approved statuses"
    lists["B1"] = "Restricted statuses"
    for row, (baseline_value, candidate_value) in enumerate(
        zip(
            _DATA_VALIDATION_LIST_BASELINE_SOURCE_VALUES,
            _DATA_VALIDATION_LIST_CANDIDATE_SOURCE_VALUES,
            strict=True,
        ),
        start=2,
    ):
        lists.cell(row, 1, baseline_value)
        lists.cell(row, 2, candidate_value)

    model = workbook.create_sheet(_DATA_VALIDATION_LIST_MODEL_SHEET)
    model["A1"] = "Selected status"
    model[_DATA_VALIDATION_LIST_MODEL_CELL] = _DATA_VALIDATION_LIST_MODEL_FORMULA

    dashboard = workbook.create_sheet(_DATA_VALIDATION_LIST_DASHBOARD_SHEET)
    dashboard["A1"] = "Board status"
    dashboard["A4"] = "Selected status"
    dashboard[_DATA_VALIDATION_LIST_DASHBOARD_CELL] = _DATA_VALIDATION_LIST_DASHBOARD_FORMULA
    return workbook


def _conditional_formatting_threshold_workbook() -> Workbook:
    """Build one stable exception-highlight rule with stored metric values.

    Conditional-formatting formulas live alongside worksheet controls rather
    than as ordinary cell formulas. The candidate is later changed only at the
    rule's raw threshold formula, which lets the fixture measure a visual
    control transition without asking a spreadsheet client to render it.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB conditional-formatting threshold fixture")
    operations = workbook.active
    operations.title = _CONDITIONAL_FORMATTING_THRESHOLD_SHEET
    operations["A1"] = "Metric"
    operations["B1"] = "Exception review metric"
    for row, value in enumerate(_CONDITIONAL_FORMATTING_THRESHOLD_VALUES, start=2):
        operations.cell(row, 1, f"Period {row - 1}")
        operations.cell(row, 2, value)
    exception_fill = PatternFill(
        fill_type="solid", fgColor=_CONDITIONAL_FORMATTING_THRESHOLD_FILL_RGB
    )
    operations.conditional_formatting.add(
        _CONDITIONAL_FORMATTING_THRESHOLD_RANGE,
        CellIsRule(
            operator="greaterThan",
            formula=[_CONDITIONAL_FORMATTING_THRESHOLD_BASELINE_FORMULA],
            fill=exception_fill,
        ),
    )
    return workbook


def _number_format_visibility_workbook() -> Workbook:
    """Build one reported metric with a custom display format.

    The candidate is later changed only in the raw custom numFmt declaration in
    styles.xml. Its target cell, style index, numeric value, and neighboring
    formula therefore remain stable; the fixture records stored display
    metadata without asking a spreadsheet client to render it.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB custom number-format fixture")
    operations = workbook.active
    operations.title = _NUMBER_FORMAT_VISIBILITY_SHEET
    operations["A1"] = "Metric"
    operations["B1"] = "Reported margin"
    operations["A2"] = "Base case"
    operations[_NUMBER_FORMAT_VISIBILITY_CELL] = _NUMBER_FORMAT_VISIBILITY_VALUE
    operations[
        _NUMBER_FORMAT_VISIBILITY_CELL
    ].number_format = _NUMBER_FORMAT_VISIBILITY_BASELINE_FORMAT
    operations["A3"] = "Raw model value"
    operations[_NUMBER_FORMAT_VISIBILITY_FORMULA_CELL] = _NUMBER_FORMAT_VISIBILITY_FORMULA
    return workbook


def _auto_filter_criteria_workbook(filter_value: str) -> Workbook:
    """Build a report with an active stored AutoFilter criterion.

    The generated pair differs only in the criterion text. Its ``SUBTOTAL``
    formula and downstream consumer deliberately remain fixed: WCAB records
    the stored filter boundary and never calculates which rows or result an
    Excel client would display.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB AutoFilter criterion fixture")
    report = workbook.active
    report.title = "Report"
    report.append(["Region", "Amount"])
    for region, amount in (("North", 100), ("North", 200), ("South", 300), ("South", 400)):
        report.append([region, amount])
    report["D1"] = "Visible amount total"
    report["D2"] = _AUTO_FILTER_SUBTOTAL_FORMULA
    report.auto_filter.ref = _AUTO_FILTER_REF
    report.auto_filter.add_filter_column(_AUTO_FILTER_COLUMN_ID, [filter_value])

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A4"] = "Visible regional amount total"
    dashboard["B4"] = _AUTO_FILTER_DASHBOARD_FORMULA
    return workbook


def _governance_workbook() -> Workbook:
    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB governed calculation")
    sheet = workbook.active
    sheet.title = "Controls"
    sheet.append(["Control", "Units", "Rate", "Calculated amount"])
    sheet["A10"] = "Approved estimate"
    sheet["B10"] = 12
    sheet["C10"] = 5
    sheet["D10"] = "=B10*C10"
    sheet["D11"] = "=D10+1"
    sheet.protection.sheet = True
    sheet.protection.enable()
    hidden = workbook.create_sheet("ReviewControls")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "Internal review notes"
    return workbook


def _three_d_workbook(*, include_adjustment: bool = False) -> Workbook:
    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB 3-D scope fixture")
    for name, amount in (("Jan", 10), ("Feb", 20), ("Mar", 30)):
        sheet = workbook.create_sheet(name)
        sheet["A1"] = name
        sheet["B5"] = amount
    if include_adjustment:
        adjustment = workbook.create_sheet("FebAdjustment", 2)
        adjustment["A1"] = "Inserted adjustment period"
        adjustment["B5"] = 5
    summary = workbook.create_sheet("Summary")
    summary["A1"] = "3-D total"
    summary["B5"] = "=SUM(Jan:Mar!B5)"
    workbook.remove(workbook["Sheet"])
    return workbook


def _refactor_workbook(*, rewritten: bool = False) -> Workbook:
    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB structural formula rewrite")
    model = workbook.active
    model.title = "Model"
    model["A1"] = "Input model"
    model["B5"] = 2
    if rewritten:
        model["C5"] = "Inserted descriptive column"
        model["D5"] = 3
        model["E5"] = "=B5*D5"
        output_cell = "E5"
    else:
        model["C5"] = 3
        model["D5"] = "=B5*C5"
        output_cell = "D5"
    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Output"
    summary["B2"] = f"=Model!{output_cell}"
    return workbook


def _structured_table_workbook() -> Workbook:
    """Build a table-backed model with a stable structured-reference summary."""

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB structured table scope fixture")
    ledger = workbook.active
    ledger.title = "Ledger"
    ledger.append(["Region", "Units", "Price", "Amount"])
    for row_number, (region, units, price) in enumerate(
        (("North", 10, 8), ("South", 12, 9), ("West", 7, 11)), start=2
    ):
        ledger.append([region, units, price, f"=B{row_number}*C{row_number}"])
    table = Table(displayName="SalesLedger", ref="A1:D4")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ledger.add_table(table)

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Structured table total"
    summary["B2"] = "=SUM(SalesLedger[Amount])"
    return workbook


def _power_query_local_table_workbook() -> Workbook:
    """Build a local Excel Table consumed by a connection-only M query.

    The Data Mashup payload is attached after openpyxl writes the ordinary
    worksheet/table package. Its formula is a stored query definition, not a
    request to execute Power Query or materialize a query result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB Power Query local-table filter fixture")
    source = workbook.active
    source.title = _POWER_QUERY_SOURCE_SHEET
    source.append([_POWER_QUERY_FILTER_COLUMN, "Amount"])
    for region, amount in (("North", 10), ("South", 20), ("North", 15), ("South", 25)):
        source.append([region, amount])
    table = Table(displayName=_POWER_QUERY_SOURCE_TABLE, ref=_POWER_QUERY_SOURCE_REF)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    source.add_table(table)

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Connection-only query output is not materialized"
    return workbook


def _scenario_manager_workbook() -> Workbook:
    """Build a small model whose alternate inputs live in Scenario Manager.

    The Scenario Manager declaration is attached after openpyxl saves the
    ordinary workbook. The visible input cells and formula path stay fixed;
    the benchmark concerns a stored alternate input value, not an applied
    scenario or a calculated result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB Scenario Manager stored-input fixture")
    inputs = workbook.active
    inputs.title = _SCENARIO_MANAGER_SHEET
    inputs["A1"] = "Scenario-controlled inputs"
    inputs["A2"] = "Margin"
    inputs[_SCENARIO_MANAGER_CHANGING_CELL] = _SCENARIO_MANAGER_WORKSHEET_INPUT_VALUE
    inputs["A3"] = "Volume"
    inputs[_SCENARIO_MANAGER_STABLE_INPUT_CELL] = _SCENARIO_MANAGER_WORKSHEET_STABLE_INPUT_VALUE
    inputs["C2"] = "Scenario summary result"
    inputs[_SCENARIO_MANAGER_RESULT_CELL] = _SCENARIO_MANAGER_FORMULA

    dashboard = workbook.create_sheet(_SCENARIO_MANAGER_DASHBOARD_SHEET)
    dashboard["A1"] = "Board scenario output"
    dashboard["A4"] = "Scenario result"
    dashboard[_SCENARIO_MANAGER_DASHBOARD_CELL] = _SCENARIO_MANAGER_DASHBOARD_FORMULA
    return workbook


def _what_if_data_table_workbook() -> Workbook:
    """Build one column-oriented What-If Data Table without calculating it.

    The data-table master is a raw SpreadsheetML formula control rather than
    an ordinary Excel Table or a calculated value. Both possible input cells
    remain part of a stable direct model path; the fixture changes only which
    local input reference the sensitivity declaration uses.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB What-If Data Table input fixture")
    sensitivity = workbook.active
    sensitivity.title = _WHAT_IF_DATA_TABLE_SHEET
    sensitivity["A1"] = "Sensitivity controls"
    sensitivity["A2"] = "Primary rate"
    sensitivity[_WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL] = _WHAT_IF_DATA_TABLE_PRIMARY_INPUT_VALUE
    sensitivity["A3"] = "Alternative rate"
    sensitivity[_WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_CELL] = (
        _WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_VALUE
    )
    sensitivity["A4"] = "Scale"
    sensitivity[_WHAT_IF_DATA_TABLE_SCALE_CELL] = _WHAT_IF_DATA_TABLE_SCALE_VALUE
    sensitivity["C2"] = "Scenario rate"
    sensitivity[_WHAT_IF_DATA_TABLE_OUTPUT_FORMULA_CELL] = _WHAT_IF_DATA_TABLE_OUTPUT_FORMULA
    for coordinate, value in zip(("C3", "C4", "C5"), _WHAT_IF_DATA_TABLE_GRID_VALUES, strict=True):
        sensitivity[coordinate] = value
    sensitivity[_WHAT_IF_DATA_TABLE_MASTER_CELL] = DataTableFormula(
        ref=_WHAT_IF_DATA_TABLE_OUTPUT_RANGE,
        ca=True,
        dt2D=False,
        dtr=False,
        r1=_WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL,
    )

    model = workbook.create_sheet(_WHAT_IF_DATA_TABLE_MODEL_SHEET)
    model["A1"] = "Sensitivity result"
    model[_WHAT_IF_DATA_TABLE_MODEL_CELL] = _WHAT_IF_DATA_TABLE_MODEL_FORMULA

    dashboard = workbook.create_sheet(_WHAT_IF_DATA_TABLE_DASHBOARD_SHEET)
    dashboard["A1"] = "Board output"
    dashboard[_WHAT_IF_DATA_TABLE_DASHBOARD_CELL] = _WHAT_IF_DATA_TABLE_DASHBOARD_FORMULA
    return workbook


def _dynamic_reference_workbook() -> Workbook:
    """Build a small model whose candidate can introduce ``INDIRECT``.

    The text driver is deliberately editable workbook data.  The fixture does
    not calculate it; it tests whether a tool makes the resulting static
    dependency-coverage boundary visible.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB dynamic reference fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Address driver"
    inputs["E12"] = "Revenue!C8"

    revenue = workbook.create_sheet("Revenue")
    revenue["A1"] = "Revenue sources"
    revenue["B8"] = 100
    revenue["C8"] = 120

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Selected revenue"
    summary["B2"] = "=Revenue!$C$8"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Board output"
    dashboard["B4"] = "=Summary!$B$2"
    return workbook


def _dynamic_reference_driver_workbook(*, function: str, driver_value: str | int) -> Workbook:
    """Build a model with a stable dynamic formula and an editable driver.

    ``INDIRECT`` changes a reference by changing text; ``OFFSET`` changes one
    by changing displacement. In both cases the formula is deliberately stable
    across the fixture pair, so reviewers must not equate unchanged formula
    text with unchanged effective dependency selection.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title=f"WCAB {function} driver fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Dynamic reference driver"
    inputs["E12"] = driver_value

    revenue = workbook.create_sheet("Revenue")
    revenue["A1"] = "Revenue sources"
    revenue["B8"] = 100
    revenue["C8"] = 120

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Selected revenue"
    if function == "INDIRECT":
        summary["B2"] = "=INDIRECT(Inputs!$E$12)"
    elif function == "OFFSET":
        summary["B2"] = "=OFFSET(Revenue!$B$8,0,Inputs!$E$12)"
    else:  # Internal callers use the two explicitly supported fixture forms.
        raise ValueError(f"unsupported dynamic-reference fixture function {function!r}")

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Board output"
    dashboard["B4"] = "=Summary!$B$2"
    return workbook


def _external_data_refresh_workbook() -> Workbook:
    """Build a workbook whose raw connection control is added after saving.

    The cells intentionally remain ordinary static data and formulas. The case
    isolates the stored external-data refresh setting rather than pretending to
    fetch a source or calculate an imported result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB external data refresh fixture")
    data = workbook.active
    data.title = "ImportedData"
    data["A1"] = "Saved imported revenue"
    data["B2"] = 100

    summary = workbook.create_sheet("Summary")
    summary["A1"] = "Saved imported revenue"
    summary["B2"] = "=ImportedData!$B$2"

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Board output"
    dashboard["B4"] = "=Summary!$B$2"
    return workbook


def _pivot_cache_refresh_workbook() -> Workbook:
    """Build an ordinary workbook around one raw PivotCache control.

    The raw PivotTable package is added after saving because openpyxl preserves
    PivotTable parts but does not create them. The source data, stored report
    cells, and direct dashboard consumer stay fixed across the pair: WCAB
    records only the cache's stored refresh-on-open request, never a rendered
    PivotTable result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB PivotCache refresh fixture")
    source = workbook.active
    source.title = _PIVOT_CACHE_SOURCE_SHEET
    source.append(["Region", "Amount"])
    for region, amount in (("North", 100), ("North", 200), ("South", 300), ("South", 400)):
        source.append([region, amount])

    report = workbook.create_sheet(_PIVOT_REPORT_SHEET)
    report["A1"] = "Region"
    report["B1"] = "Pivot amount"
    report["A2"] = "North"
    report["B2"] = 300

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A4"] = "Reported pivot amount"
    dashboard["B4"] = _PIVOT_CACHE_DASHBOARD_FORMULA
    return workbook


def _chart_series_reference_workbook() -> Workbook:
    """Build a dashboard chart with two stable local value columns.

    The candidate's raw chart-series value reference is switched after saving.
    The source cells, chart anchor, title reference, category reference, and
    all other package members remain stable. WCAB records the stored chart
    binding, not a rendered chart or a calculated visual result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB chart series reference fixture")
    source = workbook.active
    source.title = _CHART_SERIES_SOURCE_SHEET
    source.append(["Quarter", "Quarterly revenue", "Quarterly revenue"])
    source.append(["Q1", 100, 140])
    source.append(["Q2", 120, 160])
    source.append(["Q3", 140, 180])

    dashboard = workbook.create_sheet(_CHART_SERIES_DASHBOARD_SHEET)
    dashboard["A1"] = "Quarterly revenue chart"
    chart = BarChart()
    chart.title = "Quarterly revenue"
    chart.add_data(Reference(source, min_col=2, min_row=1, max_row=4), titles_from_data=True)
    chart.set_categories(Reference(source, min_col=1, min_row=2, max_row=4))
    dashboard.add_chart(chart, _CHART_SERIES_ANCHOR)
    return workbook


def _external_workbook_link_policy_workbook() -> Workbook:
    """Build a workbook with one unchanged, synthetic external-link formula.

    The source workbook intentionally does not exist. The paired case changes
    only the raw workbook-open policy, so neither generation nor validation
    can resolve the link, execute its formula, or claim a saved result.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB external workbook link policy fixture")
    model = workbook.active
    model.title = "LinkedModel"
    model["A1"] = "Synthetic external workbook driver"
    model["B2"] = _EXTERNAL_WORKBOOK_LINK_FORMULA

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Board output"
    dashboard["B4"] = "=LinkedModel!$B$2"
    return workbook


def _iterative_calculation_workbook() -> Workbook:
    """Build an unchanged, intentionally circular model with explicit bounds.

    The baseline records that iteration is disabled; the candidate changes only
    that stored calculation control. The direct self-reference is deliberate
    evidence of why the control matters, not a request for WCAB to calculate
    the model or assert a converged value.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB iterative calculation fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Convergence target"
    inputs["B2"] = 10

    model = workbook.create_sheet("Model")
    model["A1"] = "Iterative circular model"
    model["B2"] = _ITERATIVE_CALCULATION_FORMULA

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Board output"
    dashboard["B4"] = "=Model!$B$2"

    workbook.calculation.iterate = False
    workbook.calculation.iterateCount = _ITERATION_COUNT
    workbook.calculation.iterateDelta = _ITERATION_DELTA
    return workbook


def _precision_as_displayed_workbook() -> Workbook:
    """Build a model whose stored precision control can change in isolation.

    The generated pair deliberately retains the input's stored 10.005 value,
    its two-decimal display format, and every formula. A desktop Excel client
    can apply its own precision-as-displayed behavior later, but fixture
    generation never opens Excel, calculates a formula, or rewrites that value.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB precision-as-displayed fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Stored input with a two-decimal display"
    inputs["B2"] = _PRECISION_AS_DISPLAYED_INPUT
    inputs["B2"].number_format = _PRECISION_AS_DISPLAYED_NUMBER_FORMAT

    model = workbook.create_sheet("Model")
    model["A1"] = "Doubled input"
    model["B2"] = _PRECISION_AS_DISPLAYED_FORMULA
    model["B2"].number_format = _PRECISION_AS_DISPLAYED_NUMBER_FORMAT

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Published output"
    dashboard["B4"] = "=Model!$B$2"
    dashboard["B4"].number_format = _PRECISION_AS_DISPLAYED_NUMBER_FORMAT

    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullPrecision = True
    workbook.calculation.calcCompleted = True
    workbook.calculation.calcOnSave = True
    return workbook


def _workbook_date_system_workbook() -> Workbook:
    """Build a model whose raw date base can change without a cell edit.

    The input deliberately remains a numeric serial with a date display format.
    The pair is later mutated only in ``workbookPr``: WCAB does not calculate a
    formula, convert the serial to a date, or assert what an Excel client will
    display after it opens the package.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB workbook date-system fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Stored date serial"
    inputs["B2"] = _WORKBOOK_DATE_SYSTEM_SERIAL
    inputs["B2"].number_format = _WORKBOOK_DATE_SYSTEM_NUMBER_FORMAT

    model = workbook.create_sheet("Model")
    model["A1"] = "Date serial plus 30"
    model["B2"] = _WORKBOOK_DATE_SYSTEM_FORMULA
    model["B2"].number_format = _WORKBOOK_DATE_SYSTEM_NUMBER_FORMAT

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Published date"
    dashboard["B4"] = _WORKBOOK_DATE_SYSTEM_DASHBOARD_FORMULA
    dashboard["B4"].number_format = _WORKBOOK_DATE_SYSTEM_NUMBER_FORMAT
    return workbook


def _set_workbook_date_system(path: Path, *, date_1904: bool, date_compatibility: bool) -> None:
    """Set only the two raw ``workbookPr`` date-system controls."""

    def mutate(members: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(members["xl/workbook.xml"])
        properties = workbook.find(f"{{{_SPREADSHEETML_NS}}}workbookPr")
        if properties is None:
            raise ValueError("date-system fixture has no workbookPr element")
        properties.set("date1904", "1" if date_1904 else "0")
        properties.set("dateCompatibility", "1" if date_compatibility else "0")
        members["xl/workbook.xml"] = ElementTree.tostring(
            workbook, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _formula_cached_result_workbook() -> Workbook:
    """Build a model whose saved formula result can change in isolation.

    Formula text and the direct input remain ordinary workbook content. The
    generated pair receives a raw ``<v>`` value for ``Model!B2`` only after
    saving, since openpyxl intentionally writes formula text rather than a
    cached result. The stable manual calculation metadata is descriptive save
    state, not a request for WCAB to evaluate either formula.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB stored formula result fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "Unchanged direct input"
    inputs["B2"] = _FORMULA_CACHED_RESULT_INPUT

    model = workbook.create_sheet("Model")
    model["A1"] = "Saved formula result"
    model["B2"] = _FORMULA_CACHED_RESULT_FORMULA

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Local formula consumer"
    dashboard["B4"] = "=Model!$B$2"

    workbook.calculation.calcMode = "manual"
    workbook.calculation.calcCompleted = True
    workbook.calculation.calcOnSave = False
    return workbook


def _set_formula_cached_result(path: Path, *, result: int) -> None:
    """Set the raw numeric ``<v>`` cache for the generated ``Model!B2``.

    This narrowly mutates the serialized worksheet rather than calculating or
    opening the workbook in an Excel client. The structural guards make a
    changed generator dependency fail loudly instead of silently broadening
    the fixture's claim.
    """

    def mutate(members: dict[str, bytes]) -> None:
        worksheet_member = "xl/worksheets/sheet2.xml"
        worksheet = ElementTree.fromstring(members[worksheet_member])
        cell_tag = f"{{{_SPREADSHEETML_NS}}}c"
        formula_tag = f"{{{_SPREADSHEETML_NS}}}f"
        value_tag = f"{{{_SPREADSHEETML_NS}}}v"
        cells = [cell for cell in worksheet.iter(cell_tag) if cell.get("r") == "B2"]
        if len(cells) != 1:
            raise ValueError("formula-cache fixture has an unexpected Model!B2 cell")
        cell = cells[0]
        formula = cell.find(formula_tag)
        if formula is None or f"={formula.text or ''}" != _FORMULA_CACHED_RESULT_FORMULA:
            raise ValueError("formula-cache fixture has an unexpected Model!B2 formula")
        if cell.get("t") not in {None, "n"}:
            raise ValueError("formula-cache fixture has a nonnumeric Model!B2 cell")
        cached_value = cell.find(value_tag)
        if cached_value is None:
            cached_value = ElementTree.SubElement(cell, value_tag)
        cached_value.text = str(result)
        members[worksheet_member] = ElementTree.tostring(
            worksheet, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _array_formula_semantics_workbook() -> Workbook:
    """Build a fixed legacy CSE array with a direct downstream consumer.

    The candidate receives raw dynamic-array metadata after saving. Its anchor
    formula and currently stored output range intentionally remain unchanged:
    the review-relevant change is the formula mode, not a claimed calculation.
    """

    workbook = Workbook()
    _configure_workbook(workbook, title="WCAB array formula semantics fixture")
    inputs = workbook.active
    inputs.title = "Inputs"
    inputs["A1"] = "a"
    inputs["A2"] = "bb"
    inputs["A3"] = "ccc"

    model = workbook.create_sheet("Model")
    model["A1"] = "Array formula anchor"
    model["B1"].value = ArrayFormula(ref="B1:B3", text="=LEN(Inputs!A1:A3)")

    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Anchor consumer"
    dashboard["B2"] = "=Model!B1"
    return workbook


def _add_dynamic_array_metadata(path: Path) -> None:
    """Mark the generated ``Model!B1`` array formula as a dynamic array.

    Openpyxl serializes the legacy CSE anchor but does not write the Office
    dynamic-array cell-metadata binding. This compact, public OOXML shape
    follows the serialization documented by XlsxWriter: the formula cell gets
    ``cm=1`` and that cell-metadata record resolves to an ``XLDAPR`` record
    whose ``fDynamic`` property is true.
    """

    def mutate(members: dict[str, bytes]) -> None:
        def insert_before_closing(member: str, closing: bytes, addition: bytes) -> None:
            contents = members[member]
            if contents.count(closing) != 1:
                raise ValueError(f"array formula fixture has unexpected {member} markup")
            members[member] = contents.replace(closing, addition + closing, 1)

        anchor = b'<c r="B1"><f t="array" ref="B1:B3">'
        marked_anchor = b'<c r="B1" cm="1"><f t="array" ref="B1:B3">'
        worksheet_member = "xl/worksheets/sheet2.xml"
        if members[worksheet_member].count(anchor) != 1:
            raise ValueError("array formula fixture has an unexpected Model!B1 contract")
        members[worksheet_member] = members[worksheet_member].replace(anchor, marked_anchor, 1)

        insert_before_closing(
            "[Content_Types].xml",
            b"</Types>",
            (
                b'<Override PartName="/xl/metadata.xml" '
                b'ContentType="application/vnd.openxmlformats-officedocument.'
                b'spreadsheetml.sheetMetadata+xml"/>'
            ),
        )
        insert_before_closing(
            "xl/_rels/workbook.xml.rels",
            b"</Relationships>",
            (
                b'<Relationship Id="rIdWCABDynamicArrayMetadata" '
                b'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                b'relationships/sheetMetadata" Target="metadata.xml"/>'
            ),
        )
        members["xl/metadata.xml"] = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            b'xmlns:xda="' + _DYNAMIC_ARRAY_NS.encode("ascii") + b'">'
            b'<metadataTypes count="1"><metadataType name="XLDAPR" '
            b'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" '
            b'merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
            b'clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
            b'</metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst>'
            b'<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
            b'<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
            b'</ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk>'
            b'<rc t="1" v="0"/></bk></cellMetadata></metadata>'
        )

    _rewrite_xlsx_parts(path, mutate)


def _add_external_data_connection(path: Path, *, refresh_on_load: bool) -> None:
    """Add one relationship-backed, non-routable external-data connection.

    A reserved ``.invalid`` URL makes the generated package inspectable while
    ensuring the benchmark fixture never names a real provider. The only
    baseline/candidate difference is the connection's ``refreshOnLoad`` flag.
    """

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(members: dict[str, bytes]) -> None:
        content_types = ElementTree.fromstring(members["[Content_Types].xml"])
        override_tag = f"{{{_CONTENT_TYPES_NS}}}Override"
        if not any(item.get("PartName") == "/xl/connections.xml" for item in content_types):
            ElementTree.SubElement(
                content_types,
                override_tag,
                {
                    "PartName": "/xl/connections.xml",
                    "ContentType": (
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.connections+xml"
                    ),
                },
            )
        members["[Content_Types].xml"] = serialize(content_types)

        relationships = ElementTree.fromstring(members["xl/_rels/workbook.xml.rels"])
        relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
        ElementTree.SubElement(
            relationships,
            relationship_tag,
            {
                "Id": "rIdWCABExternalDataConnection",
                "Type": f"{_DOCUMENT_RELATIONSHIPS_NS}/connections",
                "Target": "connections.xml",
            },
        )
        members["xl/_rels/workbook.xml.rels"] = serialize(relationships)

        connections = ElementTree.Element(f"{{{_SPREADSHEETML_NS}}}connections")
        connection = ElementTree.SubElement(
            connections,
            f"{{{_SPREADSHEETML_NS}}}connection",
            {
                "id": "1",
                "name": "WCAB synthetic external-data connection",
                "type": "4",
                "refreshedVersion": "1",
                "refreshOnLoad": "1" if refresh_on_load else "0",
            },
        )
        ElementTree.SubElement(
            connection,
            f"{{{_SPREADSHEETML_NS}}}webPr",
            {"url": "https://example.invalid/wcab-external-data-refresh"},
        )
        members["xl/connections.xml"] = serialize(connections)

    _rewrite_xlsx_parts(path, mutate)


def _add_pivot_cache_refresh_control(path: Path, *, refresh_on_load: bool) -> None:
    """Attach one small PivotTable/PivotCache graph with a raw open-time flag.

    The fixture has a local worksheet source and no external provider. It is a
    compact, relationship-backed package shape that openpyxl can read, but this
    generator never asks Excel or openpyxl to refresh, render, or calculate the
    PivotTable. The sole baseline/candidate distinction is
    ``pivotCacheDefinition/@refreshOnLoad``.
    """

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(members: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(members["xl/workbook.xml"])
        sheets = workbook.find(f"{{{_SPREADSHEETML_NS}}}sheets")
        sheet_tag = f"{{{_SPREADSHEETML_NS}}}sheet"
        relationship_id_attribute = f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id"
        report_sheets = (
            [
                sheet
                for sheet in sheets.findall(sheet_tag)
                if sheet.get("name") == _PIVOT_REPORT_SHEET
            ]
            if sheets is not None
            else []
        )
        if len(report_sheets) != 1:
            raise ValueError("PivotCache fixture has an unexpected Report worksheet")
        report_relationship_id = report_sheets[0].get(relationship_id_attribute)
        if not isinstance(report_relationship_id, str):
            raise ValueError("PivotCache fixture Report worksheet has no relationship")

        workbook_relationships = ElementTree.fromstring(members["xl/_rels/workbook.xml.rels"])
        relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
        report_relationships = [
            relationship
            for relationship in workbook_relationships.findall(relationship_tag)
            if relationship.get("Id") == report_relationship_id
            and relationship.get("Type") == f"{_DOCUMENT_RELATIONSHIPS_NS}/worksheet"
        ]
        if len(report_relationships) != 1:
            raise ValueError("PivotCache fixture cannot resolve Report worksheet relationship")
        report_target = report_relationships[0].get("Target")
        if (
            not isinstance(report_target, str)
            or not report_target
            or report_target.startswith("../")
        ):
            raise ValueError("PivotCache fixture has an unsafe Report worksheet target")
        if report_target.startswith("/"):
            if not report_target.startswith("/xl/"):
                raise ValueError("PivotCache fixture has a non-workbook Report worksheet target")
            report_member = report_target.lstrip("/")
        else:
            report_member = f"xl/{report_target.lstrip('./')}"
        if not report_member.startswith("xl/worksheets/") or ".." in report_member.split("/"):
            raise ValueError("PivotCache fixture has an unsafe Report worksheet member")
        if report_member not in members:
            raise ValueError("PivotCache fixture Report worksheet member is absent")

        pivot_caches_tag = f"{{{_SPREADSHEETML_NS}}}pivotCaches"
        if workbook.find(pivot_caches_tag) is not None:
            raise ValueError("PivotCache fixture unexpectedly already has pivot caches")
        pivot_caches = ElementTree.SubElement(workbook, pivot_caches_tag)
        ElementTree.SubElement(
            pivot_caches,
            f"{{{_SPREADSHEETML_NS}}}pivotCache",
            {
                "cacheId": str(_PIVOT_CACHE_ID),
                relationship_id_attribute: "rIdWCABPivotCache",
            },
        )
        members["xl/workbook.xml"] = serialize(workbook)

        if any(
            relationship.get("Id") == "rIdWCABPivotCache"
            for relationship in workbook_relationships.findall(relationship_tag)
        ):
            raise ValueError("PivotCache fixture relationship ID is already in use")
        ElementTree.SubElement(
            workbook_relationships,
            relationship_tag,
            {
                "Id": "rIdWCABPivotCache",
                "Type": f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotCacheDefinition",
                "Target": "pivotCache/pivotCacheDefinition1.xml",
            },
        )
        members["xl/_rels/workbook.xml.rels"] = serialize(workbook_relationships)

        report = ElementTree.fromstring(members[report_member])
        pivot_table_parts_tag = f"{{{_SPREADSHEETML_NS}}}pivotTableParts"
        if report.find(pivot_table_parts_tag) is not None:
            raise ValueError("PivotCache fixture Report worksheet already has pivot tables")
        pivot_table_parts = ElementTree.SubElement(report, pivot_table_parts_tag, {"count": "1"})
        ElementTree.SubElement(
            pivot_table_parts,
            f"{{{_SPREADSHEETML_NS}}}pivotTablePart",
            {relationship_id_attribute: "rIdWCABPivotTable"},
        )
        members[report_member] = serialize(report)

        report_directory, report_filename = report_member.rsplit("/", maxsplit=1)
        report_relationship_member = f"{report_directory}/_rels/{report_filename}.rels"
        if report_relationship_member in members:
            report_relationships_root = ElementTree.fromstring(members[report_relationship_member])
        else:
            report_relationships_root = ElementTree.Element(
                f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationships"
            )
        if any(
            relationship.get("Id") == "rIdWCABPivotTable"
            for relationship in report_relationships_root.findall(relationship_tag)
        ):
            raise ValueError("PivotCache fixture Report relationship ID is already in use")
        ElementTree.SubElement(
            report_relationships_root,
            relationship_tag,
            {
                "Id": "rIdWCABPivotTable",
                "Type": f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotTable",
                "Target": "../pivotTables/pivotTable1.xml",
            },
        )
        members[report_relationship_member] = serialize(report_relationships_root)

        pivot_table_member = "xl/pivotTables/pivotTable1.xml"
        pivot_table = ElementTree.Element(
            f"{{{_SPREADSHEETML_NS}}}pivotTableDefinition",
            {
                "name": "WCAB Pivot Report",
                "cacheId": str(_PIVOT_CACHE_ID),
                "dataCaption": "Amount",
            },
        )
        ElementTree.SubElement(
            pivot_table,
            f"{{{_SPREADSHEETML_NS}}}location",
            {
                "ref": _PIVOT_REPORT_REF,
                "firstHeaderRow": "1",
                "firstDataRow": "2",
                "firstDataCol": "1",
            },
        )
        pivot_fields = ElementTree.SubElement(
            pivot_table, f"{{{_SPREADSHEETML_NS}}}pivotFields", {"count": "2"}
        )
        ElementTree.SubElement(
            pivot_fields,
            f"{{{_SPREADSHEETML_NS}}}pivotField",
            {"axis": "axisRow", "showAll": "0"},
        )
        ElementTree.SubElement(
            pivot_fields,
            f"{{{_SPREADSHEETML_NS}}}pivotField",
            {"dataField": "1", "showAll": "0"},
        )
        row_fields = ElementTree.SubElement(
            pivot_table, f"{{{_SPREADSHEETML_NS}}}rowFields", {"count": "1"}
        )
        ElementTree.SubElement(row_fields, f"{{{_SPREADSHEETML_NS}}}field", {"x": "0"})
        data_fields = ElementTree.SubElement(
            pivot_table, f"{{{_SPREADSHEETML_NS}}}dataFields", {"count": "1"}
        )
        ElementTree.SubElement(
            data_fields,
            f"{{{_SPREADSHEETML_NS}}}dataField",
            {"fld": "1", "subtotal": "sum"},
        )
        members[pivot_table_member] = serialize(pivot_table)

        pivot_relationships = ElementTree.Element(f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationships")
        ElementTree.SubElement(
            pivot_relationships,
            relationship_tag,
            {
                "Id": "rIdWCABPivotCache",
                "Type": f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotCacheDefinition",
                "Target": "../pivotCache/pivotCacheDefinition1.xml",
            },
        )
        members["xl/pivotTables/_rels/pivotTable1.xml.rels"] = serialize(pivot_relationships)

        cache_definition = ElementTree.Element(
            f"{{{_SPREADSHEETML_NS}}}pivotCacheDefinition",
            {
                "recordCount": "4",
                "refreshOnLoad": "1" if refresh_on_load else "0",
                "saveData": "1",
                relationship_id_attribute: "rIdWCABPivotRecords",
            },
        )
        cache_source = ElementTree.SubElement(
            cache_definition, f"{{{_SPREADSHEETML_NS}}}cacheSource", {"type": "worksheet"}
        )
        ElementTree.SubElement(
            cache_source,
            f"{{{_SPREADSHEETML_NS}}}worksheetSource",
            {"ref": _PIVOT_CACHE_SOURCE_REF, "sheet": _PIVOT_CACHE_SOURCE_SHEET},
        )
        cache_fields = ElementTree.SubElement(
            cache_definition, f"{{{_SPREADSHEETML_NS}}}cacheFields", {"count": "2"}
        )
        for name, values in (
            ("Region", (("s", "North"), ("s", "South"))),
            (
                "Amount",
                (("n", "100"), ("n", "200"), ("n", "300"), ("n", "400")),
            ),
        ):
            cache_field = ElementTree.SubElement(
                cache_fields,
                f"{{{_SPREADSHEETML_NS}}}cacheField",
                {"name": name, "numFmtId": "0"},
            )
            shared_items = ElementTree.SubElement(
                cache_field,
                f"{{{_SPREADSHEETML_NS}}}sharedItems",
                {"count": str(len(values))},
            )
            for element_name, value in values:
                ElementTree.SubElement(
                    shared_items, f"{{{_SPREADSHEETML_NS}}}{element_name}", {"v": value}
                )
        members["xl/pivotCache/pivotCacheDefinition1.xml"] = serialize(cache_definition)

        cache_definition_relationships = ElementTree.Element(
            f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationships"
        )
        ElementTree.SubElement(
            cache_definition_relationships,
            relationship_tag,
            {
                "Id": "rIdWCABPivotRecords",
                "Type": f"{_DOCUMENT_RELATIONSHIPS_NS}/pivotCacheRecords",
                "Target": "pivotCacheRecords1.xml",
            },
        )
        members["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = serialize(
            cache_definition_relationships
        )

        cache_records = ElementTree.Element(
            f"{{{_SPREADSHEETML_NS}}}pivotCacheRecords", {"count": "4"}
        )
        for region_index, amount_index in ((0, 0), (0, 1), (1, 2), (1, 3)):
            record = ElementTree.SubElement(cache_records, f"{{{_SPREADSHEETML_NS}}}r")
            ElementTree.SubElement(record, f"{{{_SPREADSHEETML_NS}}}x", {"v": str(region_index)})
            ElementTree.SubElement(record, f"{{{_SPREADSHEETML_NS}}}x", {"v": str(amount_index)})
        members["xl/pivotCache/pivotCacheRecords1.xml"] = serialize(cache_records)

        content_types = ElementTree.fromstring(members["[Content_Types].xml"])
        override_tag = f"{{{_CONTENT_TYPES_NS}}}Override"
        for part_name, content_type in (
            (
                "/xl/pivotTables/pivotTable1.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
            ),
            (
                "/xl/pivotCache/pivotCacheDefinition1.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml."
                "pivotCacheDefinition+xml",
            ),
            (
                "/xl/pivotCache/pivotCacheRecords1.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
            ),
        ):
            if any(
                item.get("PartName") == part_name for item in content_types.findall(override_tag)
            ):
                raise ValueError(f"PivotCache fixture already has content type {part_name!r}")
            ElementTree.SubElement(
                content_types,
                override_tag,
                {"PartName": part_name, "ContentType": content_type},
            )
        members["[Content_Types].xml"] = serialize(content_types)

    _rewrite_xlsx_parts(path, mutate)


def _power_query_nested_zip(parts: dict[str, bytes]) -> bytes:
    """Build a deterministic compact package embedded in a Data Mashup."""

    payload = io.BytesIO()
    with ZipFile(payload, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for name in sorted(parts):
            info = ZipInfo(filename=name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            archive.writestr(info, parts[name])
    return payload.getvalue()


def _power_query_m_filter_formula(filter_value: str) -> str:
    """Return WCAB's compact connection-only local-table M program."""

    if filter_value not in {
        _POWER_QUERY_BASELINE_FILTER_VALUE,
        _POWER_QUERY_CANDIDATE_FILTER_VALUE,
    }:
        raise ValueError(f"unsupported Power Query filter value {filter_value!r}")
    return (
        f"section {_POWER_QUERY_SECTION};\n\n"
        f"shared {_POWER_QUERY_NAME} = let\n"
        f'    Source = Excel.CurrentWorkbook(){{[Name="{_POWER_QUERY_SOURCE_TABLE}"]}}[Content],\n'
        f'    FilteredRows = Table.SelectRows(Source, each [{_POWER_QUERY_FILTER_COLUMN}] = "{filter_value}")\n'
        "in\n"
        "    FilteredRows;\n"
    )


def _power_query_local_table_data_mashup(filter_value: str) -> bytes:
    """Return a deterministic compact Data Mashup with one stored M program.

    The package has no external source, load target, embedded content, or
    user-bound permission payload. It describes a connection-only query over
    the local generated Excel Table and is not evaluated while building.
    """

    package_parts = _power_query_nested_zip(
        {
            "[Content_Types].xml": b"<Types/>",
            "Config/Package.xml": b"<Package>WCAB connection-only local table query</Package>",
            f"Formulas/{_POWER_QUERY_SECTION}.m": _power_query_m_filter_formula(
                filter_value
            ).encode("utf-8"),
        }
    )
    metadata_xml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        f'<LocalPackageMetadataFile xmlns="{_DATA_MASHUP_NS}">\n'
        "  <Items>\n"
        "    <Item>\n"
        "      <ItemLocation>\n"
        "        <ItemType>Formula</ItemType>\n"
        f"        <ItemPath>{_POWER_QUERY_SECTION}/{_POWER_QUERY_NAME}</ItemPath>\n"
        "      </ItemLocation>\n"
        "      <StableEntries>\n"
        '        <Entry Type="FillEnabled" Value="l0" />\n'
        "      </StableEntries>\n"
        "    </Item>\n"
        "  </Items>\n"
        "</LocalPackageMetadataFile>\n"
    ).encode()
    metadata = struct.pack("<II", 0, len(metadata_xml)) + metadata_xml + struct.pack("<I", 0)
    permissions = (
        b'<?xml version="1.0" encoding="utf-8"?>\n'
        b"<PermissionList>\n"
        b"  <CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\n"
        b"  <FirewallEnabled>true</FirewallEnabled>\n"
        b"</PermissionList>\n"
    )
    stream = struct.pack("<I", 0)
    for field in (package_parts, permissions, metadata, b""):
        stream += struct.pack("<I", len(field)) + field
    root = ElementTree.Element(f"{{{_DATA_MASHUP_NS}}}DataMashup")
    root.text = base64.b64encode(stream).decode("ascii")
    return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)


def _add_power_query_local_table_filter(path: Path, *, filter_value: str) -> None:
    """Attach one generated local-table Data Mashup query to an XLSX package."""

    # Validate before editing the package so callers cannot create an ambiguous
    # fixture with a value outside its explicit two-value truth contract.
    _power_query_m_filter_formula(filter_value)

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(members: dict[str, bytes]) -> None:
        relationships = ElementTree.fromstring(members["_rels/.rels"])
        relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
        custom_xml_relationship_type = f"{_DOCUMENT_RELATIONSHIPS_NS}/customXml"
        if any(
            relationship.get("Id") == _POWER_QUERY_ROOT_RELATIONSHIP_ID
            or relationship.get("Type") == custom_xml_relationship_type
            for relationship in relationships.findall(relationship_tag)
        ):
            raise ValueError("Power Query fixture unexpectedly already has custom XML")
        ElementTree.SubElement(
            relationships,
            relationship_tag,
            {
                "Id": _POWER_QUERY_ROOT_RELATIONSHIP_ID,
                "Type": custom_xml_relationship_type,
                "Target": _POWER_QUERY_CUSTOM_XML_MEMBER,
            },
        )
        members["_rels/.rels"] = serialize(relationships)

        content_types = ElementTree.fromstring(members["[Content_Types].xml"])
        default_tag = f"{{{_CONTENT_TYPES_NS}}}Default"
        xml_defaults = [
            item
            for item in content_types.findall(default_tag)
            if item.get("Extension") == "xml" and item.get("ContentType") == "application/xml"
        ]
        if len(xml_defaults) != 1:
            raise ValueError("Power Query fixture lacks the default XML content type")
        members[_POWER_QUERY_CUSTOM_XML_MEMBER] = _power_query_local_table_data_mashup(filter_value)

    _rewrite_xlsx_parts(path, mutate)


def _add_scenario_manager_stored_input(path: Path, *, stored_value: str) -> None:
    """Attach WCAB's one raw Scenario Manager declaration to a worksheet.

    Scenario Manager is represented directly inside the worksheet part rather
    than as ordinary cells. This helper creates one selected, locked scenario
    with two stored input records; callers vary only the first record's value.
    It does not ask a spreadsheet client to show a scenario or calculate the
    resulting workbook.
    """

    if stored_value not in {
        _SCENARIO_MANAGER_BASELINE_STORED_VALUE,
        _SCENARIO_MANAGER_CANDIDATE_STORED_VALUE,
    }:
        raise ValueError(f"unsupported Scenario Manager stored value {stored_value!r}")

    def mutate(members: dict[str, bytes]) -> None:
        worksheet_member = "xl/worksheets/sheet1.xml"
        worksheet = ElementTree.fromstring(members[worksheet_member])
        scenarios_tag = f"{{{_SPREADSHEETML_NS}}}scenarios"
        if worksheet.findall(scenarios_tag):
            raise ValueError("Scenario Manager fixture unexpectedly already has scenarios")
        sheet_data_tag = f"{{{_SPREADSHEETML_NS}}}sheetData"
        sheet_data_indexes = [
            index for index, child in enumerate(worksheet) if child.tag == sheet_data_tag
        ]
        if len(sheet_data_indexes) != 1:
            raise ValueError("Scenario Manager fixture lacks one sheetData element")

        scenarios = ElementTree.Element(
            scenarios_tag,
            {"current": "0", "show": "0", "sqref": _SCENARIO_MANAGER_RESULT_CELL},
        )
        scenario = ElementTree.SubElement(
            scenarios,
            f"{{{_SPREADSHEETML_NS}}}scenario",
            {
                "name": _SCENARIO_MANAGER_NAME,
                "locked": "1",
                "count": "2",
                "user": _SCENARIO_MANAGER_USER,
                "comment": _SCENARIO_MANAGER_COMMENT,
            },
        )
        ElementTree.SubElement(
            scenario,
            f"{{{_SPREADSHEETML_NS}}}inputCells",
            {
                "r": _SCENARIO_MANAGER_CHANGING_CELL,
                "val": stored_value,
                "numFmtId": str(_SCENARIO_MANAGER_INPUT_NUMBER_FORMAT_ID),
            },
        )
        ElementTree.SubElement(
            scenario,
            f"{{{_SPREADSHEETML_NS}}}inputCells",
            {
                "r": _SCENARIO_MANAGER_STABLE_INPUT_CELL,
                "val": _SCENARIO_MANAGER_STABLE_STORED_VALUE,
            },
        )
        worksheet.insert(sheet_data_indexes[0] + 1, scenarios)
        members[worksheet_member] = ElementTree.tostring(
            worksheet, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _set_what_if_data_table_input_reference(path: Path, *, input_cell: str) -> None:
    """Set the one generated Data Table master's local input reference.

    ``f/@r1`` is raw SpreadsheetML metadata on the Data Table master. The
    helper edits that one attribute after openpyxl serializes the ordinary
    workbook, so the pair does not manufacture or recalculate a table result.
    """

    if input_cell not in {
        _WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL,
        _WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_CELL,
    }:
        raise ValueError(f"unsupported What-If Data Table input cell {input_cell!r}")

    def mutate(members: dict[str, bytes]) -> None:
        worksheet_member = "xl/worksheets/sheet1.xml"
        worksheet = ElementTree.fromstring(members[worksheet_member])
        cell_tag = f"{{{_SPREADSHEETML_NS}}}c"
        formula_tag = f"{{{_SPREADSHEETML_NS}}}f"
        masters = [
            cell
            for cell in worksheet.iter(cell_tag)
            if cell.get("r") == _WHAT_IF_DATA_TABLE_MASTER_CELL
        ]
        if len(masters) != 1:
            raise ValueError("What-If Data Table fixture lacks one master cell")
        formulas = masters[0].findall(formula_tag)
        if (
            len(formulas) != 1
            or formulas[0].get("t") != "dataTable"
            or formulas[0].get("ref") != _WHAT_IF_DATA_TABLE_OUTPUT_RANGE
            or formulas[0].get("r1") != _WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL
        ):
            raise ValueError("What-If Data Table fixture has an unexpected master declaration")
        formulas[0].set("r1", input_cell)
        members[worksheet_member] = ElementTree.tostring(
            worksheet, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _set_data_validation_list_source(path: Path, *, source_formula: str) -> None:
    """Set WCAB's one list-validation source declaration without saving cells.

    Excel stores a list validation's source as the ``formula1`` child of the
    target's ``dataValidation`` record. Editing that small raw declaration
    after openpyxl saves the baseline shape makes the candidate's package
    boundary explicit: only the Inputs worksheet part is allowed to differ.
    """

    if source_formula not in {
        _DATA_VALIDATION_LIST_BASELINE_SOURCE_FORMULA,
        _DATA_VALIDATION_LIST_CANDIDATE_SOURCE_FORMULA,
    }:
        raise ValueError(f"unsupported list-validation source {source_formula!r}")

    def mutate(members: dict[str, bytes]) -> None:
        worksheet_member = "xl/worksheets/sheet1.xml"
        worksheet = ElementTree.fromstring(members[worksheet_member])
        data_validations_tag = f"{{{_SPREADSHEETML_NS}}}dataValidations"
        data_validation_tag = f"{{{_SPREADSHEETML_NS}}}dataValidation"
        formula1_tag = f"{{{_SPREADSHEETML_NS}}}formula1"
        containers = worksheet.findall(data_validations_tag)
        if len(containers) != 1:
            raise ValueError("list-validation fixture lacks one dataValidations container")
        validations = list(containers[0])
        if len(validations) != 1 or validations[0].tag != data_validation_tag:
            raise ValueError("list-validation fixture lacks one target validation")
        validation = validations[0]
        formulas = validation.findall(formula1_tag)
        if (
            validation.get("sqref") != _DATA_VALIDATION_LIST_TARGET_RANGE
            or validation.get("type") != "list"
            or len(validation) != 1
            or len(formulas) != 1
            or formulas[0].attrib
            or len(formulas[0])
            or formulas[0].text != _DATA_VALIDATION_LIST_BASELINE_SOURCE_FORMULA
        ):
            raise ValueError("list-validation fixture has an unexpected source declaration")
        formulas[0].text = source_formula
        members[worksheet_member] = ElementTree.tostring(
            worksheet, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _set_conditional_formatting_threshold(path: Path, *, formula: str) -> None:
    """Set WCAB's one conditional-formatting threshold without saving cells.

    The stored ``formula`` child of a ``cellIs`` conditional-formatting rule
    defines the exception threshold. Editing it after openpyxl writes the
    baseline shape keeps the candidate's boundary narrow: only its operations
    worksheet part may differ, and no conditional formatting is evaluated.
    """

    if formula not in {
        _CONDITIONAL_FORMATTING_THRESHOLD_BASELINE_FORMULA,
        _CONDITIONAL_FORMATTING_THRESHOLD_CANDIDATE_FORMULA,
    }:
        raise ValueError(f"unsupported conditional-formatting threshold {formula!r}")

    def mutate(members: dict[str, bytes]) -> None:
        worksheet_member = "xl/worksheets/sheet1.xml"
        worksheet = ElementTree.fromstring(members[worksheet_member])
        conditional_formatting_tag = f"{{{_SPREADSHEETML_NS}}}conditionalFormatting"
        rule_tag = f"{{{_SPREADSHEETML_NS}}}cfRule"
        formula_tag = f"{{{_SPREADSHEETML_NS}}}formula"
        controls = worksheet.findall(conditional_formatting_tag)
        if (
            len(controls) != 1
            or controls[0].get("sqref") != _CONDITIONAL_FORMATTING_THRESHOLD_RANGE
        ):
            raise ValueError("conditional-formatting fixture lacks one target control")
        rules = list(controls[0])
        if len(rules) != 1 or rules[0].tag != rule_tag:
            raise ValueError("conditional-formatting fixture lacks one rule")
        rule = rules[0]
        formulas = rule.findall(formula_tag)
        if (
            tuple(sorted(rule.attrib.items()))
            != (
                ("dxfId", "0"),
                ("operator", "greaterThan"),
                ("priority", "1"),
                ("type", "cellIs"),
            )
            or len(rule) != 1
            or len(formulas) != 1
            or formulas[0].attrib
            or len(formulas[0])
            or formulas[0].text != _CONDITIONAL_FORMATTING_THRESHOLD_BASELINE_FORMULA
        ):
            raise ValueError("conditional-formatting fixture has an unexpected threshold rule")
        formulas[0].text = formula
        members[worksheet_member] = ElementTree.tostring(
            worksheet, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _set_number_format_visibility_format(path: Path, *, format_code: str) -> None:
    """Set WCAB's one custom format code without saving cells or styles.

    Custom number formats are declared in styles.xml. Replacing only the
    selected formatCode after openpyxl saves the baseline shape keeps the
    candidate boundary narrow: the target cell retains the same style index and
    only the referenced custom display declaration changes.
    """

    if format_code != _NUMBER_FORMAT_VISIBILITY_CANDIDATE_FORMAT:
        raise ValueError(f"unsupported number format {format_code!r}")

    def mutate(members: dict[str, bytes]) -> None:
        styles_member = "xl/styles.xml"
        styles = ElementTree.fromstring(members[styles_member])
        num_fmts_tag = f"{{{_SPREADSHEETML_NS}}}numFmts"
        num_fmt_tag = f"{{{_SPREADSHEETML_NS}}}numFmt"
        containers = styles.findall(num_fmts_tag)
        if len(containers) != 1 or tuple(sorted(containers[0].attrib.items())) != (("count", "1"),):
            raise ValueError("number-format fixture lacks one custom-number-format container")
        number_formats = containers[0].findall(num_fmt_tag)
        if len(number_formats) != 1 or tuple(sorted(number_formats[0].attrib.items())) != (
            ("formatCode", _NUMBER_FORMAT_VISIBILITY_BASELINE_FORMAT),
            ("numFmtId", str(_NUMBER_FORMAT_VISIBILITY_CUSTOM_ID)),
        ):
            raise ValueError("number-format fixture has an unexpected custom format")
        number_formats[0].set("formatCode", format_code)
        members[styles_member] = ElementTree.tostring(
            styles, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _add_pivot_slicer_selection(path: Path, *, selected_item_index: int) -> None:
    """Attach one local PivotTable slicer with one explicitly selected cache item.

    ``slicerCacheDefinition/data/tabular/items/i/@s`` is a stored Slicer-cache
    state.  This package-only fixture deliberately records that declaration
    after openpyxl writes ordinary worksheet cells; it neither creates a Slicer
    drawing nor asks a spreadsheet client to apply, refresh, calculate, or
    render the filter.
    """

    if type(selected_item_index) is not int or selected_item_index not in range(
        _PIVOT_SLICER_ITEM_COUNT
    ):
        raise ValueError(f"unsupported PivotTable slicer selection {selected_item_index!r}")

    def serialize(root: ElementTree.Element) -> bytes:
        return ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)

    def mutate(members: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(members["xl/workbook.xml"])
        relationship_id_attribute = f"{{{_DOCUMENT_RELATIONSHIPS_NS}}}id"
        relationship_tag = f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship"
        sheets = workbook.find(f"{{{_SPREADSHEETML_NS}}}sheets")
        report_sheets = (
            [
                sheet
                for sheet in sheets.findall(f"{{{_SPREADSHEETML_NS}}}sheet")
                if sheet.get("name") == _PIVOT_REPORT_SHEET
            ]
            if sheets is not None
            else []
        )
        if len(report_sheets) != 1:
            raise ValueError("PivotTable slicer fixture has an unexpected Report worksheet")
        report_sheet_id = report_sheets[0].get("sheetId")
        try:
            if (
                not isinstance(report_sheet_id, str)
                or int(report_sheet_id) != _PIVOT_SLICER_PIVOT_TAB_ID
            ):
                raise ValueError
        except ValueError as error:
            raise ValueError(
                "PivotTable slicer fixture Report worksheet has no sheet id"
            ) from error

        pivot_caches = workbook.findall(f"{{{_SPREADSHEETML_NS}}}pivotCaches")
        if len(pivot_caches) != 1 or len(pivot_caches[0]) != 1:
            raise ValueError("PivotTable slicer fixture must contain exactly one PivotCache")
        pivot_cache = pivot_caches[0][0]
        if pivot_cache.tag != f"{{{_SPREADSHEETML_NS}}}pivotCache":
            raise ValueError("PivotTable slicer fixture has an unexpected PivotCache declaration")
        if workbook.find(f"{{{_SPREADSHEETML_NS}}}extLst") is not None:
            raise ValueError(
                "PivotTable slicer fixture unexpectedly already has workbook extensions"
            )

        workbook_extensions = ElementTree.SubElement(workbook, f"{{{_SPREADSHEETML_NS}}}extLst")
        workbook_extension = ElementTree.SubElement(
            workbook_extensions,
            f"{{{_SPREADSHEETML_NS}}}ext",
            {"uri": _PIVOT_SLICER_CACHE_EXTENSION_URI},
        )
        slicer_caches = ElementTree.SubElement(
            workbook_extension, f"{{{_OFFICE_2010_SPREADSHEET_NS}}}slicerCaches"
        )
        ElementTree.SubElement(
            slicer_caches,
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}slicerCache",
            {relationship_id_attribute: "rIdWCABSlicerCache"},
        )
        members["xl/workbook.xml"] = serialize(workbook)

        workbook_relationships = ElementTree.fromstring(members["xl/_rels/workbook.xml.rels"])
        if any(
            relationship.get("Id") == "rIdWCABSlicerCache"
            for relationship in workbook_relationships.findall(relationship_tag)
        ):
            raise ValueError("PivotTable slicer fixture relationship ID is already in use")
        ElementTree.SubElement(
            workbook_relationships,
            relationship_tag,
            {
                "Id": "rIdWCABSlicerCache",
                "Type": _PIVOT_SLICER_CACHE_RELATIONSHIP,
                "Target": "slicerCaches/slicerCache1.xml",
            },
        )
        members["xl/_rels/workbook.xml.rels"] = serialize(workbook_relationships)

        cache_definition_member = "xl/pivotCache/pivotCacheDefinition1.xml"
        cache_definition = ElementTree.fromstring(members[cache_definition_member])
        if cache_definition.tag != f"{{{_SPREADSHEETML_NS}}}pivotCacheDefinition":
            raise ValueError("PivotTable slicer fixture has an unexpected cache definition")
        if cache_definition.find(f"{{{_SPREADSHEETML_NS}}}extLst") is not None:
            raise ValueError("PivotTable slicer fixture unexpectedly already has cache extensions")
        cache_extensions = ElementTree.SubElement(
            cache_definition, f"{{{_SPREADSHEETML_NS}}}extLst"
        )
        cache_extension = ElementTree.SubElement(
            cache_extensions,
            f"{{{_SPREADSHEETML_NS}}}ext",
            {"uri": _PIVOT_SLICER_PIVOT_CACHE_EXTENSION_URI},
        )
        ElementTree.SubElement(
            cache_extension,
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}pivotCacheDefinition",
            {"pivotCacheId": str(_PIVOT_CACHE_ID)},
        )
        members[cache_definition_member] = serialize(cache_definition)

        slicer_member = "xl/slicerCaches/slicerCache1.xml"
        slicer_cache = ElementTree.Element(
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}slicerCacheDefinition",
            {"name": _PIVOT_SLICER_NAME, "sourceName": _PIVOT_SLICER_SOURCE_NAME},
        )
        pivot_tables = ElementTree.SubElement(
            slicer_cache,
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}pivotTables",
            {"count": "1"},
        )
        ElementTree.SubElement(
            pivot_tables,
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}pivotTable",
            {"tabId": report_sheet_id, "name": _PIVOT_SLICER_PIVOT_TABLE_NAME},
        )
        data = ElementTree.SubElement(slicer_cache, f"{{{_OFFICE_2010_SPREADSHEET_NS}}}data")
        tabular = ElementTree.SubElement(
            data,
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}tabular",
            {"pivotCacheId": str(_PIVOT_CACHE_ID)},
        )
        items = ElementTree.SubElement(
            tabular,
            f"{{{_OFFICE_2010_SPREADSHEET_NS}}}items",
            {"count": str(_PIVOT_SLICER_ITEM_COUNT)},
        )
        for item_index in range(_PIVOT_SLICER_ITEM_COUNT):
            ElementTree.SubElement(
                items,
                f"{{{_OFFICE_2010_SPREADSHEET_NS}}}i",
                {"x": str(item_index), "s": "1" if item_index == selected_item_index else "0"},
            )
        members[slicer_member] = serialize(slicer_cache)

        content_types = ElementTree.fromstring(members["[Content_Types].xml"])
        override_tag = f"{{{_CONTENT_TYPES_NS}}}Override"
        if any(
            item.get("PartName") == f"/{slicer_member}"
            for item in content_types.findall(override_tag)
        ):
            raise ValueError("PivotTable slicer fixture already has a slicer content type")
        ElementTree.SubElement(
            content_types,
            override_tag,
            {"PartName": f"/{slicer_member}", "ContentType": _PIVOT_SLICER_CACHE_CONTENT_TYPE},
        )
        members["[Content_Types].xml"] = serialize(content_types)

    _rewrite_xlsx_parts(path, mutate)


def _set_pivot_data_field_subtotal(path: Path, *, subtotal: str) -> None:
    """Set the sole generated PivotTable value field's aggregation control.

    ``dataField/@subtotal`` stores the aggregate function for a PivotTable
    value field.  This is deliberately a raw OOXML change after openpyxl has
    written the ordinary workbook, so the fixture records a PivotTable
    declaration rather than a recalculated or rendered output.
    """

    if subtotal not in {
        _PIVOT_DATA_FIELD_BASELINE_SUBTOTAL,
        _PIVOT_DATA_FIELD_CANDIDATE_SUBTOTAL,
    }:
        raise ValueError(f"unsupported PivotTable data-field subtotal {subtotal!r}")

    def mutate(members: dict[str, bytes]) -> None:
        pivot_members = sorted(
            name
            for name in members
            if name.startswith("xl/pivotTables/")
            and name.endswith(".xml")
            and "/_rels/" not in name
        )
        if pivot_members != ["xl/pivotTables/pivotTable1.xml"]:
            raise ValueError("pivot aggregation fixture must contain exactly one PivotTable part")
        pivot_table = ElementTree.fromstring(members[pivot_members[0]])
        data_fields_tag = f"{{{_SPREADSHEETML_NS}}}dataFields"
        data_field_tag = f"{{{_SPREADSHEETML_NS}}}dataField"
        data_fields = pivot_table.findall(data_fields_tag)
        if (
            len(data_fields) != 1
            or data_fields[0].get("count") != "1"
            or len(data_fields[0]) != 1
            or data_fields[0][0].tag != data_field_tag
            or data_fields[0][0].get("fld") != str(_PIVOT_DATA_FIELD_SOURCE_INDEX)
        ):
            raise ValueError("pivot aggregation fixture has an unexpected data-field declaration")
        data_fields[0][0].set("subtotal", subtotal)
        members[pivot_members[0]] = ElementTree.tostring(
            pivot_table, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _set_chart_series_value_reference(path: Path, *, value_reference: str) -> None:
    """Set the sole generated chart's stored numeric-series reference.

    The mutation is deliberately limited to the one ``c:ser/c:val/c:numRef/c:f``
    text node. Rewriting it after openpyxl saves the package keeps this fixture
    about a DrawingML chart binding rather than a worksheet edit or a rendered
    chart cache.
    """

    if value_reference not in {
        _CHART_SERIES_BASELINE_VALUE_REFERENCE,
        _CHART_SERIES_CANDIDATE_VALUE_REFERENCE,
    }:
        raise ValueError(f"unsupported chart-series value reference {value_reference!r}")

    def mutate(members: dict[str, bytes]) -> None:
        chart_members = sorted(
            name
            for name in members
            if name.startswith("xl/charts/") and name.endswith(".xml") and "/_rels/" not in name
        )
        if len(chart_members) != 1:
            raise ValueError("chart-series fixture must contain exactly one chart part")
        chart_member = chart_members[0]
        chart = ElementTree.fromstring(members[chart_member])
        series = chart.findall(f".//{{{_DRAWINGML_CHART_NS}}}ser")
        if len(series) != 1:
            raise ValueError("chart-series fixture must contain exactly one series")
        formulas = series[0].findall(
            f"{{{_DRAWINGML_CHART_NS}}}val/"
            f"{{{_DRAWINGML_CHART_NS}}}numRef/"
            f"{{{_DRAWINGML_CHART_NS}}}f"
        )
        if len(formulas) != 1:
            raise ValueError("chart-series fixture must contain exactly one numeric value formula")
        formulas[0].text = value_reference
        members[chart_member] = ElementTree.tostring(chart, encoding="utf-8", xml_declaration=True)

    _rewrite_xlsx_parts(path, mutate)


def _set_external_workbook_link_update_policy(path: Path, *, update_links: str) -> None:
    """Set the stored workbook-open policy for external-workbook links.

    ``workbookPr/@updateLinks`` is distinct from the relationship-backed
    connection control used by the earlier external-data case. The allowed
    values here deliberately make the generated transition unambiguous:
    never update linked workbooks at open versus always update them at open.
    """

    if update_links not in {"never", "always"}:
        raise ValueError(f"unsupported external-workbook update policy {update_links!r}")

    def mutate(members: dict[str, bytes]) -> None:
        workbook = ElementTree.fromstring(members["xl/workbook.xml"])
        properties_tag = f"{{{_SPREADSHEETML_NS}}}workbookPr"
        properties = workbook.find(properties_tag)
        if properties is None:
            properties = ElementTree.Element(properties_tag)
            workbook.insert(0, properties)
        properties.set("updateLinks", update_links)
        members["xl/workbook.xml"] = ElementTree.tostring(
            workbook, encoding="utf-8", xml_declaration=True
        )

    _rewrite_xlsx_parts(path, mutate)


def _portfolio_workbook(*, driver_value: int) -> tuple[Workbook, Workbook]:
    drivers = Workbook()
    _configure_workbook(drivers, title="WCAB portfolio drivers")
    inputs = drivers.active
    inputs.title = "Inputs"
    inputs["A1"] = "Approved driver"
    inputs["B2"] = driver_value

    model = Workbook()
    _configure_workbook(model, title="WCAB portfolio model")
    summary = model.active
    summary.title = "Summary"
    summary["A1"] = "Externally supplied driver"
    summary["B2"] = "='[drivers.xlsx]Inputs'!$B$2"
    summary["A3"] = "Derived output"
    summary["B3"] = "=B2*10"
    return drivers, model


def _build_finance_formula_to_value(root: Path) -> None:
    truth = _truth(
        case_id="finance.formula_to_value",
        title="Revenue formula is replaced with a hard-coded value",
        family="finance",
        review_expectation="block",
        facts=[{"kind": "formula_to_value", "sheet": "Revenue", "cell": "C8"}],
        must_reach=[
            {
                "source": {"sheet": "Revenue", "cell": "C8"},
                "targets": [
                    {"sheet": "Revenue", "cell": "C10"},
                    {"sheet": "Summary", "cell": "C5"},
                    {"sheet": "Summary", "cell": "C6"},
                    {"sheet": "Dashboard", "cell": "B4"},
                ],
            }
        ],
        coverage=["Static dependency lower bound; numerical recalculation is out of scope."],
    )
    _write_pair(
        root / "finance" / "formula_to_value",
        _finance_workbook,
        lambda wb: setattr(wb["Revenue"]["C8"], "value", 999),
        truth,
    )


def _build_finance_wrong_period_reference(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Revenue"]["C8"] = "=B5*C6"

    truth = _truth(
        case_id="finance.wrong_period_reference",
        title="Revenue formula uses a prior-period units reference",
        family="finance",
        review_expectation="block",
        facts=[{"kind": "formula_changed", "sheet": "Revenue", "cell": "C8"}],
        must_reach=[
            {
                "source": {"sheet": "Revenue", "cell": "C8"},
                "targets": [
                    {"sheet": "Revenue", "cell": "C10"},
                    {"sheet": "Summary", "cell": "C5"},
                    {"sheet": "Dashboard", "cell": "B4"},
                ],
            }
        ],
        coverage=[
            "The benchmark asserts the reference drift, not the business correctness of every formula."
        ],
    )
    _write_pair(root / "finance" / "wrong_period_reference", _finance_workbook, mutate, truth)


def _build_finance_input_change(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Assumptions"]["C2"] = 140

    truth = _truth(
        case_id="finance.input_value_change",
        title="Approved input changes and propagates into outputs",
        family="finance",
        review_expectation="review",
        facts=[{"kind": "value_changed", "sheet": "Assumptions", "cell": "C2"}],
        must_reach=[
            {
                "source": {"sheet": "Assumptions", "cell": "C2"},
                "targets": [
                    {"sheet": "Revenue", "cell": "C5"},
                    {"sheet": "Revenue", "cell": "C8"},
                    {"sheet": "Revenue", "cell": "C10"},
                    {"sheet": "Summary", "cell": "C6"},
                    {"sheet": "Dashboard", "cell": "B4"},
                ],
            }
        ],
        coverage=[
            "An input change is not inherently an error; it remains material review evidence."
        ],
    )
    _write_pair(root / "finance" / "input_value_change", _finance_workbook, mutate, truth)


def _build_finance_external_formula(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Summary"]["C8"] = "='[MarketData.xlsx]Rates'!$B$2"

    truth = _truth(
        case_id="finance.external_formula_reference",
        title="An internal driver is replaced by an external workbook reference",
        family="finance",
        review_expectation="block",
        facts=[
            {"kind": "formula_changed", "sheet": "Summary", "cell": "C8"},
            {"kind": "external_formula_added", "sheet": "Summary", "cell": "C8"},
        ],
        coverage=[
            "The external workbook is intentionally absent and must not be opened by a benchmark tool."
        ],
    )
    _write_pair(root / "finance" / "external_formula_reference", _finance_workbook, mutate, truth)


def _build_finance_defined_name(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook.defined_names["DiscountRate"] = DefinedName(
            "DiscountRate", attr_text="Assumptions!$E$11", comment="WCAB synthetic driver"
        )

    truth = _truth(
        case_id="finance.defined_name_redirect",
        title="A named rate redirects to a different assumption",
        family="finance",
        review_expectation="block",
        facts=[{"kind": "defined_name_changed", "name": "DiscountRate"}],
        coverage=[
            "The assertion compares the defined-name destination; name resolution beyond this fixture is out of scope."
        ],
    )
    _write_pair(root / "finance" / "defined_name_redirect", _finance_workbook, mutate, truth)


def _build_operations_formula_interruption(root: Path) -> None:
    truth = _truth(
        case_id="operations.copied_formula_interruption",
        title="One copied margin formula is replaced by a manual number",
        family="operations",
        review_expectation="block",
        facts=[{"kind": "formula_to_value", "sheet": "Operations", "cell": "F7"}],
        coverage=[
            "Immediate formula peers provide the copied-block context; no formula execution is required."
        ],
    )
    _write_pair(
        root / "operations" / "copied_formula_interruption",
        _operations_workbook,
        lambda wb: setattr(wb["Operations"]["F7"], "value", 42),
        truth,
    )


def _build_operations_sumifs_shape(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Operations"]["H6"] = "=SUMIFS($F$6:$F$9,$A$6:$A$8,A6)"

    truth = _truth(
        case_id="operations.sumifs_range_shape",
        title="A conditional aggregate has mismatched static ranges",
        family="operations",
        review_expectation="block",
        facts=[{"kind": "formula_changed", "sheet": "Operations", "cell": "H6"}],
        coverage=[
            "The range-shape defect is deliberately static and does not require evaluating SUMIFS."
        ],
    )
    _write_pair(root / "operations" / "sumifs_range_shape", _operations_workbook, mutate, truth)


def _build_operations_data_validation(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Operations"].data_validations.dataValidation.clear()

    truth = _truth(
        case_id="operations.data_validation_removed",
        title="Input validation is removed from an operational tracker",
        family="operations",
        review_expectation="block",
        facts=[
            {
                "kind": "data_validation_count_changed",
                "sheet": "Operations",
                "baseline_count": 1,
                "candidate_count": 0,
            }
        ],
        coverage=[
            "The contract observes validation presence, not whether a future user enters an invalid value."
        ],
    )
    _write_pair(
        root / "operations" / "data_validation_removed", _operations_workbook, mutate, truth
    )


def _build_operations_data_validation_list_source(root: Path) -> None:
    """Build a stored list-source retarget without altering input/result cells."""

    directory = root / "operations" / "data_validation_list_source_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_data_validation_list_source_workbook(), baseline)
    _save_workbook(_data_validation_list_source_workbook(), candidate)
    _set_data_validation_list_source(
        candidate, source_formula=_DATA_VALIDATION_LIST_CANDIDATE_SOURCE_FORMULA
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="operations.data_validation_list_source_changed",
            title="A list validation switches its permitted-status source",
            family="operations",
            review_expectation="block",
            facts=[
                {
                    "kind": "data_validation_list_source_changed",
                    "validation_sheet": _DATA_VALIDATION_LIST_SHEET,
                    "target_range": _DATA_VALIDATION_LIST_TARGET_RANGE,
                    "validation_type": "list",
                    "baseline_source_formula": _DATA_VALIDATION_LIST_BASELINE_SOURCE_FORMULA,
                    "candidate_source_formula": _DATA_VALIDATION_LIST_CANDIDATE_SOURCE_FORMULA,
                    "allow_blank": False,
                    "dropdown_hidden": False,
                    "show_error_message": True,
                    "error_style": _DATA_VALIDATION_LIST_ERROR_STYLE,
                    "error_title": _DATA_VALIDATION_LIST_ERROR_TITLE,
                    "error": _DATA_VALIDATION_LIST_ERROR,
                    "show_input_message": False,
                    "prompt_title": _DATA_VALIDATION_LIST_PROMPT_TITLE,
                    "prompt": _DATA_VALIDATION_LIST_PROMPT,
                    "source_sheet": _DATA_VALIDATION_LIST_SOURCE_SHEET,
                    "baseline_source_range": _DATA_VALIDATION_LIST_BASELINE_SOURCE_RANGE,
                    "candidate_source_range": _DATA_VALIDATION_LIST_CANDIDATE_SOURCE_RANGE,
                    "baseline_source_values": list(_DATA_VALIDATION_LIST_BASELINE_SOURCE_VALUES),
                    "candidate_source_values": list(_DATA_VALIDATION_LIST_CANDIDATE_SOURCE_VALUES),
                    "input_cell": _DATA_VALIDATION_LIST_TARGET_RANGE,
                    "input_value": _DATA_VALIDATION_LIST_INPUT_VALUE,
                    "model_sheet": _DATA_VALIDATION_LIST_MODEL_SHEET,
                    "model_cell": _DATA_VALIDATION_LIST_MODEL_CELL,
                    "model_formula": _DATA_VALIDATION_LIST_MODEL_FORMULA,
                    "dashboard_sheet": _DATA_VALIDATION_LIST_DASHBOARD_SHEET,
                    "dashboard_cell": _DATA_VALIDATION_LIST_DASHBOARD_CELL,
                    "dashboard_formula": _DATA_VALIDATION_LIST_DASHBOARD_FORMULA,
                }
            ],
            must_reach=[
                {
                    "source": {
                        "sheet": _DATA_VALIDATION_LIST_SHEET,
                        "cell": _DATA_VALIDATION_LIST_TARGET_RANGE,
                    },
                    "targets": [
                        {
                            "sheet": _DATA_VALIDATION_LIST_MODEL_SHEET,
                            "cell": _DATA_VALIDATION_LIST_MODEL_CELL,
                        },
                        {
                            "sheet": _DATA_VALIDATION_LIST_DASHBOARD_SHEET,
                            "cell": _DATA_VALIDATION_LIST_DASHBOARD_CELL,
                        },
                    ],
                }
            ],
            coverage=[
                "The pair changes only the raw formula1 source declaration in the Inputs worksheet's one list data-validation record. It does not edit the target range, rule metadata, source-list values, current input, ordinary formulas, or calculation properties.",
                "The observed contract is a stored validation source. WCAB does not evaluate the list formula, determine whether any future entry is valid, accept or reject an entry, calculate the workbook, or claim Excel-client behavior.",
                "The direct static Inputs-to-model-to-dashboard path is a lower bound only if a user later enters a value. It is not proof that a validation rule permits that value or of a resulting calculation.",
            ],
        ),
    )


def _build_operations_conditional_formatting(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        worksheet = workbook["Operations"]
        for rule_range in list(worksheet.conditional_formatting):
            del worksheet.conditional_formatting[str(rule_range.sqref)]

    truth = _truth(
        case_id="operations.conditional_formatting_removed",
        title="A material exception highlight is removed",
        family="operations",
        review_expectation="review",
        facts=[
            {
                "kind": "conditional_formatting_count_changed",
                "sheet": "Operations",
                "baseline_count": 1,
                "candidate_count": 0,
            }
        ],
        coverage=[
            "The benchmark observes the control removal but does not judge the intended visual design."
        ],
    )
    _write_pair(
        root / "operations" / "conditional_formatting_removed", _operations_workbook, mutate, truth
    )


def _build_operations_conditional_formatting_threshold(root: Path) -> None:
    """Build a stored exception-threshold transition without changing cells."""

    directory = root / "operations" / "conditional_formatting_threshold_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_conditional_formatting_threshold_workbook(), baseline)
    _save_workbook(_conditional_formatting_threshold_workbook(), candidate)
    _set_conditional_formatting_threshold(
        candidate, formula=_CONDITIONAL_FORMATTING_THRESHOLD_CANDIDATE_FORMULA
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="operations.conditional_formatting_threshold_changed",
            title="A conditional-formatting exception threshold is lowered",
            family="operations",
            review_expectation="review",
            facts=[
                {
                    "kind": "conditional_formatting_threshold_changed",
                    "sheet": _CONDITIONAL_FORMATTING_THRESHOLD_SHEET,
                    "target_range": _CONDITIONAL_FORMATTING_THRESHOLD_RANGE,
                    "priority": 1,
                    "rule_type": "cellIs",
                    "operator": "greaterThan",
                    "baseline_formula": _CONDITIONAL_FORMATTING_THRESHOLD_BASELINE_FORMULA,
                    "candidate_formula": _CONDITIONAL_FORMATTING_THRESHOLD_CANDIDATE_FORMULA,
                    "metric_values": list(_CONDITIONAL_FORMATTING_THRESHOLD_VALUES),
                    "fill_rgb": _CONDITIONAL_FORMATTING_THRESHOLD_FILL_RGB,
                }
            ],
            coverage=[
                "The pair changes only the raw formula child in one cellIs conditional-formatting rule. It does not edit the target range, priority, operator, differential fill, worksheet values, calculation properties, or any other package member.",
                "The observed contract is a stored visual exception threshold. WCAB does not evaluate the rule, determine which cells a client formats, calculate the workbook, or claim Excel-client behavior.",
            ],
        ),
    )


def _build_operations_number_format_visibility(root: Path) -> None:
    """Build a custom-format transition without a cell or formula edit."""

    directory = root / "operations" / "number_format_value_hidden"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_number_format_visibility_workbook(), baseline)
    _save_workbook(_number_format_visibility_workbook(), candidate)
    _set_number_format_visibility_format(
        candidate, format_code=_NUMBER_FORMAT_VISIBILITY_CANDIDATE_FORMAT
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="operations.number_format_value_hidden",
            title="A reported margin receives a hide-value number format",
            family="operations",
            review_expectation="review",
            facts=[
                {
                    "kind": "cell_number_format_changed",
                    "sheet": _NUMBER_FORMAT_VISIBILITY_SHEET,
                    "cell": _NUMBER_FORMAT_VISIBILITY_CELL,
                    "value": _NUMBER_FORMAT_VISIBILITY_VALUE,
                    "custom_number_format_id": _NUMBER_FORMAT_VISIBILITY_CUSTOM_ID,
                    "baseline_format": _NUMBER_FORMAT_VISIBILITY_BASELINE_FORMAT,
                    "candidate_format": _NUMBER_FORMAT_VISIBILITY_CANDIDATE_FORMAT,
                    "formula_cell": _NUMBER_FORMAT_VISIBILITY_FORMULA_CELL,
                    "formula": _NUMBER_FORMAT_VISIBILITY_FORMULA,
                }
            ],
            coverage=[
                "The pair changes only one custom numFmt formatCode declaration in styles.xml. It does not edit the target cell's style index or stored numeric text, the neighboring formula, calculation properties, or any other package member.",
                "The observed contract is stored display metadata. WCAB does not render a number format, resolve locale or column-width behavior, decide what a client displays, calculate the workbook, or claim Excel-client behavior.",
            ],
        ),
    )


def _build_operations_auto_filter_criteria(root: Path) -> None:
    """Build an active-filter transition with stable cells and formulas."""

    directory = root / "operations" / "auto_filter_criteria_changed"
    directory.mkdir(parents=True, exist_ok=True)
    _save_workbook(
        _auto_filter_criteria_workbook(_AUTO_FILTER_BASELINE_VALUE), directory / "baseline.xlsx"
    )
    _save_workbook(
        _auto_filter_criteria_workbook(_AUTO_FILTER_CANDIDATE_VALUE), directory / "candidate.xlsx"
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="operations.auto_filter_criteria_changed",
            title="An active AutoFilter criterion changes a subtotal control",
            family="operations",
            review_expectation="block",
            facts=[
                {
                    "kind": "auto_filter_criteria_changed",
                    "sheet": "Report",
                    "filter_ref": _AUTO_FILTER_REF,
                    "filter_column_id": _AUTO_FILTER_COLUMN_ID,
                    "baseline_filter_value": _AUTO_FILTER_BASELINE_VALUE,
                    "candidate_filter_value": _AUTO_FILTER_CANDIDATE_VALUE,
                    "subtotal_cell": "D2",
                    "subtotal_formula": _AUTO_FILTER_SUBTOTAL_FORMULA,
                    "dashboard_sheet": "Dashboard",
                    "dashboard_cell": "B4",
                    "dashboard_formula": _AUTO_FILTER_DASHBOARD_FORMULA,
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": "Report", "cell": "D2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                }
            ],
            coverage=[
                "The pair changes only one raw worksheet AutoFilter criterion. It does not edit a cell value, formula text, style, calculation property, row state, or dependency edge.",
                "Excel filters control which rows are shown, and SUBTOTAL excludes rows omitted by a filter. WCAB does not apply the filter, calculate a subtotal, infer a visible row set, or claim what an Excel client will display or print.",
                "The validator reads raw OOXML filter declarations, formula text, and package-member differences. It treats the criterion transition as stored-control review evidence, not proof of a calculated outcome.",
            ],
        ),
    )


def _build_governance_visibility(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["ReviewControls"].sheet_state = "visible"

    truth = _truth(
        case_id="governance.hidden_sheet_revealed",
        title="A hidden review-control sheet is made visible",
        family="governance",
        review_expectation="review",
        facts=[
            {
                "kind": "sheet_visibility_changed",
                "sheet": "ReviewControls",
                "baseline_state": "hidden",
                "candidate_state": "visible",
            }
        ],
        coverage=[
            "Visibility is an observable workbook control; cell content is not part of this assertion."
        ],
    )
    _write_pair(root / "governance" / "hidden_sheet_revealed", _governance_workbook, mutate, truth)


def _build_governance_protection(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Controls"]["D10"].protection = Protection(locked=False)

    truth = _truth(
        case_id="governance.formula_cell_unlocked",
        title="A protected calculated cell is explicitly unlocked",
        family="governance",
        review_expectation="block",
        facts=[{"kind": "formula_cell_unlocked", "sheet": "Controls", "cell": "D10"}],
        coverage=[
            "This is a direct cell-protection assertion; style inheritance and allowed edit ranges are out of scope."
        ],
    )
    _write_pair(root / "governance" / "formula_cell_unlocked", _governance_workbook, mutate, truth)


def _build_governance_manual_calculation(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook.calculation.calcMode = "manual"
        workbook.calculation.calcCompleted = False

    truth = _truth(
        case_id="governance.manual_calculation_incomplete",
        title="A formula workbook records incomplete manual calculation",
        family="governance",
        review_expectation="block",
        facts=[{"kind": "manual_calculation_incomplete"}],
        coverage=[
            "Calculation metadata is evidence of save state, not proof that any cached result is wrong."
        ],
    )
    _write_pair(
        root / "governance" / "manual_calculation_incomplete", _governance_workbook, mutate, truth
    )


def _build_governance_iterative_calculation(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook.calculation.iterate = True

    truth = _truth(
        case_id="governance.iterative_calculation_enabled",
        title="An unchanged direct circular formula enables iterative calculation",
        family="governance",
        review_expectation="block",
        facts=[
            {
                "kind": "iterative_calculation_enabled",
                "sheet": "Model",
                "cell": "B2",
                "formula": _ITERATIVE_CALCULATION_FORMULA,
                "baseline_iterate": False,
                "candidate_iterate": True,
                "iteration_count": _ITERATION_COUNT,
                "iteration_delta": _ITERATION_DELTA,
            }
        ],
        must_reach=[
            {
                "source": {"sheet": "Model", "cell": "B2"},
                "targets": [{"sheet": "Dashboard", "cell": "B4"}],
            }
        ],
        coverage=[
            "The direct self-reference is intentional and remains unchanged. The fixture records calculation controls; it does not calculate the circular formula or claim a converged, cached, or business-correct result.",
            "The validator checks the explicit iterate flag plus the shared iteration count and delta. It does not predict the number of recalculations, convergence, numerical precision, or Excel-client behavior.",
            "The pair differs only in workbook calculation metadata; no worksheet value, formula text, or dependency edge is edited.",
        ],
    )
    _write_pair(
        root / "governance" / "iterative_calculation_enabled",
        _iterative_calculation_workbook,
        mutate,
        truth,
    )


def _build_governance_precision_as_displayed(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook.calculation.fullPrecision = False

    truth = _truth(
        case_id="governance.precision_as_displayed_enabled",
        title="Precision-as-displayed calculation is enabled without a cell edit",
        family="governance",
        review_expectation="block",
        facts=[
            {
                "kind": "precision_as_displayed_enabled",
                "input_sheet": "Inputs",
                "input_cell": "B2",
                "input_value": _PRECISION_AS_DISPLAYED_INPUT,
                "number_format": _PRECISION_AS_DISPLAYED_NUMBER_FORMAT,
                "formula_sheet": "Model",
                "formula_cell": "B2",
                "formula": _PRECISION_AS_DISPLAYED_FORMULA,
                "baseline_full_precision": True,
                "candidate_full_precision": False,
            }
        ],
        must_reach=[
            {
                "source": {"sheet": "Inputs", "cell": "B2"},
                "targets": [
                    {"sheet": "Model", "cell": "B2"},
                    {"sheet": "Dashboard", "cell": "B4"},
                ],
            }
        ],
        coverage=[
            "The pair records the raw calcPr fullPrecision switch that corresponds to Excel's precision-as-displayed setting. WCAB does not emulate Excel's calculation engine or claim that a client applied the control.",
            "Both generated packages retain the exact stored 10.005 input and its two-decimal number format. The validator does not claim that opening, calculating, or saving either workbook rounds a value or produces a particular formula result.",
            "The pair differs only in workbook calculation metadata; no worksheet value, formula text, number format, or dependency edge is edited.",
        ],
    )
    _write_pair(
        root / "governance" / "precision_as_displayed_enabled",
        _precision_as_displayed_workbook,
        mutate,
        truth,
    )


def _build_governance_workbook_date_system(root: Path) -> None:
    """Build a raw date-system-control transition with stable cell content."""

    directory = root / "governance" / "workbook_date_system_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_workbook_date_system_workbook(), baseline)
    _save_workbook(_workbook_date_system_workbook(), candidate)
    _set_workbook_date_system(baseline, date_1904=False, date_compatibility=True)
    _set_workbook_date_system(candidate, date_1904=True, date_compatibility=True)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="governance.workbook_date_system_changed",
            title="Workbook serial-date controls change without a cell edit",
            family="governance",
            review_expectation="block",
            facts=[
                {
                    "kind": "workbook_date_system_changed",
                    "baseline_date_1904": False,
                    "candidate_date_1904": True,
                    "date_compatibility": True,
                    "serial_sheet": "Inputs",
                    "serial_cell": "B2",
                    "serial_value": _WORKBOOK_DATE_SYSTEM_SERIAL,
                    "number_format": _WORKBOOK_DATE_SYSTEM_NUMBER_FORMAT,
                    "formula_sheet": "Model",
                    "formula_cell": "B2",
                    "formula": _WORKBOOK_DATE_SYSTEM_FORMULA,
                    "dashboard_sheet": "Dashboard",
                    "dashboard_cell": "B4",
                    "dashboard_formula": _WORKBOOK_DATE_SYSTEM_DASHBOARD_FORMULA,
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": "Inputs", "cell": "B2"},
                    "targets": [
                        {"sheet": "Model", "cell": "B2"},
                        {"sheet": "Dashboard", "cell": "B4"},
                    ],
                },
                {
                    "source": {"sheet": "Model", "cell": "B2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                },
            ],
            coverage=[
                "The pair changes only raw workbookPr date1904 from false to true while dateCompatibility remains explicitly true. It does not edit a worksheet cell, formula, style, calculation property, or dependency edge.",
                "The stored numeric serial and date number format remain unchanged. WCAB does not calculate a formula, convert the serial into a date, infer a visible value, or claim what an Excel client will do on opening or saving either workbook.",
                "The validator reads raw OOXML controls, numeric cell text, formula text, number format, and package-member differences. It treats the date-system transition as a stored-control review boundary, not proof of a displayed outcome.",
            ],
        ),
    )


def _build_governance_formula_cached_result(root: Path) -> None:
    """Build a pair with exactly one unexplained saved formula-result change."""

    directory = root / "governance" / "formula_cached_result_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_formula_cached_result_workbook(), baseline)
    _save_workbook(_formula_cached_result_workbook(), candidate)
    _set_formula_cached_result(baseline, result=_FORMULA_CACHED_RESULT_BASELINE)
    _set_formula_cached_result(candidate, result=_FORMULA_CACHED_RESULT_CANDIDATE)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="governance.formula_cached_result_changed",
            title="A saved formula result changes without a formula or input edit",
            family="governance",
            review_expectation="block",
            facts=[
                {
                    "kind": "formula_cached_result_changed",
                    "sheet": "Model",
                    "cell": "B2",
                    "formula": _FORMULA_CACHED_RESULT_FORMULA,
                    "input_sheet": "Inputs",
                    "input_cell": "B2",
                    "input_value": _FORMULA_CACHED_RESULT_INPUT,
                    "result_type": "numeric",
                    "baseline_cached_result": _FORMULA_CACHED_RESULT_BASELINE,
                    "candidate_cached_result": _FORMULA_CACHED_RESULT_CANDIDATE,
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": "Inputs", "cell": "B2"},
                    "targets": [
                        {"sheet": "Model", "cell": "B2"},
                        {"sheet": "Dashboard", "cell": "B4"},
                    ],
                },
                {
                    "source": {"sheet": "Model", "cell": "B2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                },
            ],
            coverage=[
                "The pair changes exactly one raw OOXML CellValue (<v>) saved beside an unchanged formula; it does not calculate either formula or change a visible input, formula, dependency edge, or calculation control.",
                "A saved formula result is evidence of the last stored calculation, not proof that either result is current, stale, tampered, mathematically correct, or what an Excel client will display after opening or recalculating.",
                "The validator compares the raw formula cache, formula text, input, and package members. It does not execute a formula, infer volatile or external inputs, or claim a downstream displayed result changed.",
            ],
        ),
    )


def _build_governance_static_cycle(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Controls"]["D10"] = "=D11+1"
        workbook["Controls"]["D11"] = "=D10+1"

    truth = _truth(
        case_id="governance.static_cycle_introduced",
        title="Two direct formulas form a static cycle",
        family="governance",
        review_expectation="block",
        facts=[
            {
                "kind": "static_cycle_introduced",
                "cells": [
                    {"sheet": "Controls", "cell": "D10"},
                    {"sheet": "Controls", "cell": "D11"},
                ],
            }
        ],
        coverage=[
            "The cycle is direct and static; iterative calculation settings and dynamic references are out of scope."
        ],
    )
    _write_pair(
        root / "governance" / "static_cycle_introduced", _governance_workbook, mutate, truth
    )


def _build_governance_external_data_refresh(root: Path) -> None:
    """Build a pair that differs only in a raw connection refresh control."""

    directory = root / "governance" / "external_data_refresh_on_open"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_external_data_refresh_workbook(), baseline)
    _save_workbook(_external_data_refresh_workbook(), candidate)
    _add_external_data_connection(baseline, refresh_on_load=False)
    _add_external_data_connection(candidate, refresh_on_load=True)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="governance.external_data_refresh_on_open",
            title="An external-data connection starts refreshing when the workbook opens",
            family="governance",
            review_expectation="block",
            facts=[
                {
                    "kind": "external_data_connection_refresh_on_load_changed",
                    "connection_id": 1,
                    "baseline_refresh_on_load": False,
                    "candidate_refresh_on_load": True,
                }
            ],
            coverage=[
                "The connection endpoint is synthetic and non-routable; the benchmark never opens it or tests credentials, trust, or returned data.",
                "The observable contract is the stored refresh-on-open control, not a claim that downstream saved values will recalculate.",
            ],
        ),
    )


def _build_governance_pivot_cache_refresh(root: Path) -> None:
    """Build a local PivotCache whose open-time refresh request changes."""

    directory = root / "governance" / "pivot_cache_refresh_on_open"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_pivot_cache_refresh_workbook(), baseline)
    _save_workbook(_pivot_cache_refresh_workbook(), candidate)
    _add_pivot_cache_refresh_control(baseline, refresh_on_load=False)
    _add_pivot_cache_refresh_control(candidate, refresh_on_load=True)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="governance.pivot_cache_refresh_on_open",
            title="A PivotTable cache starts refreshing when the workbook opens",
            family="governance",
            review_expectation="block",
            facts=[
                {
                    "kind": "pivot_cache_refresh_on_load_changed",
                    "cache_id": _PIVOT_CACHE_ID,
                    "source_type": "worksheet",
                    "source_sheet": _PIVOT_CACHE_SOURCE_SHEET,
                    "source_ref": _PIVOT_CACHE_SOURCE_REF,
                    "pivot_sheet": _PIVOT_REPORT_SHEET,
                    "pivot_ref": _PIVOT_REPORT_REF,
                    "pivot_output_cell": "B2",
                    "dashboard_sheet": "Dashboard",
                    "dashboard_cell": "B4",
                    "dashboard_formula": _PIVOT_CACHE_DASHBOARD_FORMULA,
                    "baseline_refresh_on_load": False,
                    "candidate_refresh_on_load": True,
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": _PIVOT_REPORT_SHEET, "cell": "B2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                }
            ],
            coverage=[
                "The pair changes only raw pivotCacheDefinition/@refreshOnLoad. It does not edit source cells, stored PivotTable display cells, formula text, or calculation properties.",
                "PivotTable cache refresh-on-open is stored control evidence. WCAB does not open the workbook in Excel, refresh a cache, calculate, or render the PivotTable, or claim a changed report result.",
                "The validator reads relationship-backed raw OOXML, checks the cache/PivotTable binding and direct local dashboard edge, and treats that edge as a lower bound only if a client refresh changes the PivotTable display.",
            ],
        ),
    )


def _build_structural_pivot_data_field_aggregation(root: Path) -> None:
    """Build a PivotTable whose stored value aggregation switches in isolation."""

    directory = root / "structural" / "pivot_data_field_aggregation_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_pivot_cache_refresh_workbook(), baseline)
    _save_workbook(_pivot_cache_refresh_workbook(), candidate)
    _add_pivot_cache_refresh_control(baseline, refresh_on_load=False)
    _add_pivot_cache_refresh_control(candidate, refresh_on_load=False)
    _set_pivot_data_field_subtotal(baseline, subtotal=_PIVOT_DATA_FIELD_BASELINE_SUBTOTAL)
    _set_pivot_data_field_subtotal(candidate, subtotal=_PIVOT_DATA_FIELD_CANDIDATE_SUBTOTAL)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.pivot_data_field_aggregation_changed",
            title="A PivotTable value field switches from Sum to Average",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "pivot_data_field_aggregation_changed",
                    "cache_id": _PIVOT_CACHE_ID,
                    "source_type": "worksheet",
                    "source_sheet": _PIVOT_CACHE_SOURCE_SHEET,
                    "source_ref": _PIVOT_CACHE_SOURCE_REF,
                    "pivot_sheet": _PIVOT_REPORT_SHEET,
                    "pivot_ref": _PIVOT_REPORT_REF,
                    "pivot_output_cell": "B2",
                    "dashboard_sheet": "Dashboard",
                    "dashboard_cell": "B4",
                    "dashboard_formula": _PIVOT_CACHE_DASHBOARD_FORMULA,
                    "data_field_source_index": _PIVOT_DATA_FIELD_SOURCE_INDEX,
                    "baseline_subtotal": _PIVOT_DATA_FIELD_BASELINE_SUBTOTAL,
                    "candidate_subtotal": _PIVOT_DATA_FIELD_CANDIDATE_SUBTOTAL,
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": _PIVOT_REPORT_SHEET, "cell": "B2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                }
            ],
            coverage=[
                "The pair changes only raw pivotTableDefinition/dataFields/dataField/@subtotal from sum to average. It does not edit source cells, cache records, stored PivotTable display cells, formula text, refresh controls, or calculation properties.",
                "The observable contract is the stored aggregation declaration. WCAB does not calculate, refresh, or render the PivotTable, infer a changed displayed value, or claim client behavior.",
                "The validator follows the local cache/PivotTable relationship graph, verifies the stable source and direct dashboard edge, and treats that edge as a lower bound only if a client refresh changes the PivotTable display.",
            ],
        ),
    )


def _build_structural_pivot_slicer_selection(root: Path) -> None:
    """Build a PivotTable Slicer whose selected cache item switches in isolation."""

    directory = root / "structural" / "pivot_slicer_selection_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_pivot_cache_refresh_workbook(), baseline)
    _save_workbook(_pivot_cache_refresh_workbook(), candidate)
    _add_pivot_cache_refresh_control(baseline, refresh_on_load=False)
    _add_pivot_cache_refresh_control(candidate, refresh_on_load=False)
    _add_pivot_slicer_selection(baseline, selected_item_index=_PIVOT_SLICER_BASELINE_SELECTED_INDEX)
    _add_pivot_slicer_selection(
        candidate, selected_item_index=_PIVOT_SLICER_CANDIDATE_SELECTED_INDEX
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.pivot_slicer_selection_changed",
            title="A PivotTable slicer switches from North to South",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "pivot_slicer_selection_changed",
                    "cache_id": _PIVOT_CACHE_ID,
                    "source_type": "worksheet",
                    "source_sheet": _PIVOT_CACHE_SOURCE_SHEET,
                    "source_ref": _PIVOT_CACHE_SOURCE_REF,
                    "pivot_sheet": _PIVOT_REPORT_SHEET,
                    "pivot_ref": _PIVOT_REPORT_REF,
                    "pivot_output_cell": "B2",
                    "dashboard_sheet": "Dashboard",
                    "dashboard_cell": "B4",
                    "dashboard_formula": _PIVOT_CACHE_DASHBOARD_FORMULA,
                    "slicer_name": _PIVOT_SLICER_NAME,
                    "slicer_source_name": _PIVOT_SLICER_SOURCE_NAME,
                    "slicer_pivot_table_name": _PIVOT_SLICER_PIVOT_TABLE_NAME,
                    "slicer_pivot_tab_id": _PIVOT_SLICER_PIVOT_TAB_ID,
                    "item_count": _PIVOT_SLICER_ITEM_COUNT,
                    "baseline_selected_item_index": _PIVOT_SLICER_BASELINE_SELECTED_INDEX,
                    "candidate_selected_item_index": _PIVOT_SLICER_CANDIDATE_SELECTED_INDEX,
                    "baseline_selected_value": "North",
                    "candidate_selected_value": "South",
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": _PIVOT_REPORT_SHEET, "cell": "B2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                }
            ],
            coverage=[
                "The pair changes only raw slicerCacheDefinition/data/tabular/items/i/@s selection state. It does not edit source cells, cache records, stored PivotTable display cells, formula text, refresh controls, or calculation properties.",
                "The observable contract is the stored Slicer-cache selection declaration. This fixture has no Slicer drawing or view geometry; WCAB does not apply the filter, refresh, calculate, render, infer a changed display value, or claim client behavior.",
                "The validator follows the local workbook-to-Slicer-cache-to-PivotCache/PivotTable relationship graph, verifies the stable source and direct dashboard edge, and treats that edge as a lower bound only if a client applies the Slicer filter.",
            ],
        ),
    )


def _build_structural_power_query_m_filter(root: Path) -> None:
    """Build a connection-only local-table Power Query M filter transition."""

    directory = root / "structural" / "power_query_m_filter_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_power_query_local_table_workbook(), baseline)
    _save_workbook(_power_query_local_table_workbook(), candidate)
    _add_power_query_local_table_filter(baseline, filter_value=_POWER_QUERY_BASELINE_FILTER_VALUE)
    _add_power_query_local_table_filter(candidate, filter_value=_POWER_QUERY_CANDIDATE_FILTER_VALUE)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.power_query_m_filter_changed",
            title="A Power Query local-table filter switches from North to South",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "power_query_m_filter_changed",
                    "data_mashup_part": _POWER_QUERY_CUSTOM_XML_MEMBER,
                    "source_sheet": _POWER_QUERY_SOURCE_SHEET,
                    "source_table": _POWER_QUERY_SOURCE_TABLE,
                    "source_ref": _POWER_QUERY_SOURCE_REF,
                    "query_section": _POWER_QUERY_SECTION,
                    "query_name": _POWER_QUERY_NAME,
                    "filter_column": _POWER_QUERY_FILTER_COLUMN,
                    "baseline_filter_value": _POWER_QUERY_BASELINE_FILTER_VALUE,
                    "candidate_filter_value": _POWER_QUERY_CANDIDATE_FILTER_VALUE,
                    "fill_enabled": False,
                    "firewall_enabled": True,
                    "future_packages_allowed": False,
                }
            ],
            coverage=[
                "The pair changes only the stored M filter literal in one Data Mashup custom-XML part. It does not edit the local source table, worksheet cells, table definition, workbook calculation properties, connection settings, or a stored query result.",
                "The query is connection-only (FillEnabled=false) and reads the generated workbook's local SourceData table. WCAB does not execute M, apply the filter, refresh a query, materialize output, calculate formulas, or infer returned rows or client behavior.",
                "The validator follows the package-root customXml relationship, parses only this compact generated Data Mashup envelope, and verifies its one local Table source plus stable metadata and permission controls. It is not a general M parser or a proof of query execution semantics.",
            ],
        ),
    )


def _build_structural_scenario_manager_stored_input(root: Path) -> None:
    """Build a Scenario Manager alternate-input transition without cell churn."""

    directory = root / "structural" / "scenario_manager_stored_input_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_scenario_manager_workbook(), baseline)
    _save_workbook(_scenario_manager_workbook(), candidate)
    _add_scenario_manager_stored_input(
        baseline, stored_value=_SCENARIO_MANAGER_BASELINE_STORED_VALUE
    )
    _add_scenario_manager_stored_input(
        candidate, stored_value=_SCENARIO_MANAGER_CANDIDATE_STORED_VALUE
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.scenario_manager_stored_input_changed",
            title="A Scenario Manager downside input changes without a worksheet-cell edit",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "scenario_manager_stored_input_value_changed",
                    "scenario_sheet": _SCENARIO_MANAGER_SHEET,
                    "scenario_name": _SCENARIO_MANAGER_NAME,
                    "changing_cell": _SCENARIO_MANAGER_CHANGING_CELL,
                    "stable_input_cell": _SCENARIO_MANAGER_STABLE_INPUT_CELL,
                    "baseline_stored_value": _SCENARIO_MANAGER_BASELINE_STORED_VALUE,
                    "candidate_stored_value": _SCENARIO_MANAGER_CANDIDATE_STORED_VALUE,
                    "stable_stored_value": _SCENARIO_MANAGER_STABLE_STORED_VALUE,
                    "input_number_format_id": _SCENARIO_MANAGER_INPUT_NUMBER_FORMAT_ID,
                    "summary_ref": _SCENARIO_MANAGER_RESULT_CELL,
                    "worksheet_input_value": _SCENARIO_MANAGER_WORKSHEET_INPUT_VALUE,
                    "worksheet_stable_input_value": _SCENARIO_MANAGER_WORKSHEET_STABLE_INPUT_VALUE,
                    "result_cell": _SCENARIO_MANAGER_RESULT_CELL,
                    "result_formula": _SCENARIO_MANAGER_FORMULA,
                    "dashboard_sheet": _SCENARIO_MANAGER_DASHBOARD_SHEET,
                    "dashboard_cell": _SCENARIO_MANAGER_DASHBOARD_CELL,
                    "dashboard_formula": _SCENARIO_MANAGER_DASHBOARD_FORMULA,
                }
            ],
            must_reach=[
                {
                    "source": {
                        "sheet": _SCENARIO_MANAGER_SHEET,
                        "cell": _SCENARIO_MANAGER_CHANGING_CELL,
                    },
                    "targets": [
                        {
                            "sheet": _SCENARIO_MANAGER_SHEET,
                            "cell": _SCENARIO_MANAGER_RESULT_CELL,
                        },
                        {
                            "sheet": _SCENARIO_MANAGER_DASHBOARD_SHEET,
                            "cell": _SCENARIO_MANAGER_DASHBOARD_CELL,
                        },
                    ],
                }
            ],
            coverage=[
                "The pair changes only one raw scenarios/scenario/inputCells/@val record in the Inputs worksheet part. It does not edit visible worksheet cells, formula text, calculation properties, or the Scenario Manager selection, protection, comment, user, summary-reference, or number-format metadata.",
                "The observed contract is one stored alternate Scenario Manager input value. WCAB does not show or apply a scenario, calculate the model, predict a result, generate a Scenario Summary, or claim a client will use the stored input.",
                "The stable local formula path from the declared changing cell to the dashboard is a lower bound only if a client applies the scenario. The validator reads raw worksheet OOXML and formula text without executing either action.",
            ],
        ),
    )


def _build_structural_what_if_data_table_input(root: Path) -> None:
    """Build a Data Table input-reference switch without cell/result churn."""

    directory = root / "structural" / "what_if_data_table_input_reference_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_what_if_data_table_workbook(), baseline)
    _save_workbook(_what_if_data_table_workbook(), candidate)
    _set_what_if_data_table_input_reference(
        candidate, input_cell=_WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_CELL
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.what_if_data_table_input_reference_changed",
            title="A What-If Data Table switches its one-variable input cell",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "what_if_data_table_input_reference_changed",
                    "table_sheet": _WHAT_IF_DATA_TABLE_SHEET,
                    "master_cell": _WHAT_IF_DATA_TABLE_MASTER_CELL,
                    "output_range": _WHAT_IF_DATA_TABLE_OUTPUT_RANGE,
                    "baseline_input_cell": _WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL,
                    "candidate_input_cell": _WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_CELL,
                    "orientation": "column",
                    "recalculation_requested": True,
                    "input_value_range": _WHAT_IF_DATA_TABLE_INPUT_VALUE_RANGE,
                    "input_values": list(_WHAT_IF_DATA_TABLE_GRID_VALUES),
                    "primary_input_value": _WHAT_IF_DATA_TABLE_PRIMARY_INPUT_VALUE,
                    "alternate_input_value": _WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_VALUE,
                    "scale_cell": _WHAT_IF_DATA_TABLE_SCALE_CELL,
                    "scale_value": _WHAT_IF_DATA_TABLE_SCALE_VALUE,
                    "output_formula_cell": _WHAT_IF_DATA_TABLE_OUTPUT_FORMULA_CELL,
                    "output_formula": _WHAT_IF_DATA_TABLE_OUTPUT_FORMULA,
                    "model_sheet": _WHAT_IF_DATA_TABLE_MODEL_SHEET,
                    "model_cell": _WHAT_IF_DATA_TABLE_MODEL_CELL,
                    "model_formula": _WHAT_IF_DATA_TABLE_MODEL_FORMULA,
                    "dashboard_sheet": _WHAT_IF_DATA_TABLE_DASHBOARD_SHEET,
                    "dashboard_cell": _WHAT_IF_DATA_TABLE_DASHBOARD_CELL,
                    "dashboard_formula": _WHAT_IF_DATA_TABLE_DASHBOARD_FORMULA,
                }
            ],
            must_reach=[
                {
                    "source": {
                        "sheet": _WHAT_IF_DATA_TABLE_SHEET,
                        "cell": _WHAT_IF_DATA_TABLE_PRIMARY_INPUT_CELL,
                    },
                    "targets": [
                        {
                            "sheet": _WHAT_IF_DATA_TABLE_MODEL_SHEET,
                            "cell": _WHAT_IF_DATA_TABLE_MODEL_CELL,
                        },
                        {
                            "sheet": _WHAT_IF_DATA_TABLE_DASHBOARD_SHEET,
                            "cell": _WHAT_IF_DATA_TABLE_DASHBOARD_CELL,
                        },
                    ],
                },
                {
                    "source": {
                        "sheet": _WHAT_IF_DATA_TABLE_SHEET,
                        "cell": _WHAT_IF_DATA_TABLE_ALTERNATE_INPUT_CELL,
                    },
                    "targets": [
                        {
                            "sheet": _WHAT_IF_DATA_TABLE_MODEL_SHEET,
                            "cell": _WHAT_IF_DATA_TABLE_MODEL_CELL,
                        },
                        {
                            "sheet": _WHAT_IF_DATA_TABLE_DASHBOARD_SHEET,
                            "cell": _WHAT_IF_DATA_TABLE_DASHBOARD_CELL,
                        },
                    ],
                },
            ],
            coverage=[
                "The pair changes only the one-variable Data Table master's raw f/@r1 input-cell reference in the Sensitivity worksheet. It does not edit the table output range, orientation, recalculation request, visible cells, ordinary formula text, calculation properties, or saved table results.",
                "The table is a column-oriented one-variable What-If Data Table. WCAB records its stored declaration only; it does not substitute input values, recalculate the workbook or table, infer output values, resolve a circular dependency, or claim Excel-client behavior.",
                "Both possible input cells retain ordinary static local paths to the model and dashboard. Those are lower-bound formula edges, not a proof of which temporary value Excel substitutes or of any Data Table result.",
            ],
        ),
    )


def _build_governance_external_workbook_link_update_policy(root: Path) -> None:
    """Build a pair differing only in the global external-link open policy."""

    directory = root / "governance" / "external_workbook_link_update_on_open"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_external_workbook_link_policy_workbook(), baseline)
    _save_workbook(_external_workbook_link_policy_workbook(), candidate)
    _set_external_workbook_link_update_policy(baseline, update_links="never")
    _set_external_workbook_link_update_policy(candidate, update_links="always")
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="governance.external_workbook_link_update_on_open",
            title="An external-workbook link switches from never to always updating on open",
            family="governance",
            review_expectation="block",
            facts=[
                {
                    "kind": "external_workbook_link_update_policy_changed",
                    "sheet": "LinkedModel",
                    "cell": "B2",
                    "formula": _EXTERNAL_WORKBOOK_LINK_FORMULA,
                    "baseline_update_links": "never",
                    "candidate_update_links": "always",
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": "LinkedModel", "cell": "B2"},
                    "targets": [{"sheet": "Dashboard", "cell": "B4"}],
                }
            ],
            coverage=[
                "The external workbook name is synthetic and absent; the benchmark never opens, resolves, refreshes, or authenticates to it.",
                "The observed contract is only workbookPr/@updateLinks. It does not establish source reachability, trust, changed source data, a returned value, or recalculation success.",
                "The external-link formula and its local downstream reference remain unchanged; the validator compares stored XML and formula text without executing either formula.",
            ],
        ),
    )


def _build_structural_chart_series_reference(root: Path) -> None:
    """Build a dashboard chart whose local value-series binding changes."""

    directory = root / "structural" / "chart_series_reference_changed"
    directory.mkdir(parents=True, exist_ok=True)
    baseline = directory / "baseline.xlsx"
    candidate = directory / "candidate.xlsx"
    _save_workbook(_chart_series_reference_workbook(), baseline)
    _save_workbook(_chart_series_reference_workbook(), candidate)
    _set_chart_series_value_reference(
        baseline, value_reference=_CHART_SERIES_BASELINE_VALUE_REFERENCE
    )
    _set_chart_series_value_reference(
        candidate, value_reference=_CHART_SERIES_CANDIDATE_VALUE_REFERENCE
    )
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.chart_series_reference_changed",
            title="A dashboard chart switches to a different local value series",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "chart_series_value_reference_changed",
                    "chart_sheet": _CHART_SERIES_DASHBOARD_SHEET,
                    "chart_anchor": _CHART_SERIES_ANCHOR,
                    "source_sheet": _CHART_SERIES_SOURCE_SHEET,
                    "series_title_ref": _CHART_SERIES_TITLE_REFERENCE,
                    "category_ref": _CHART_SERIES_CATEGORY_REFERENCE,
                    "baseline_value_ref": _CHART_SERIES_BASELINE_VALUE_REFERENCE,
                    "candidate_value_ref": _CHART_SERIES_CANDIDATE_VALUE_REFERENCE,
                }
            ],
            coverage=[
                "The pair changes only the raw DrawingML c:ser/c:val/c:numRef/c:f value-reference text in one chart part. It does not edit worksheet cells, chart anchor, category/title references, or calculation properties.",
                "A chart-series source reference is stored report-binding evidence. WCAB does not open Excel, calculate values, refresh chart data, render a chart, infer a visual difference, or claim any client will use the stored reference.",
                "The validator follows the local worksheet-to-drawing-to-chart relationship chain and compares raw chart XML after removing the declared value reference. It does not resolve arbitrary chart formulas or add chart references to the formula dependency graph.",
            ],
        ),
    )


def _build_structural_array_formula_mode(root: Path) -> None:
    """Build a legacy-CSE to dynamic-array transition without formula text churn."""

    directory = root / "structural" / "array_formula_mode_changed"
    directory.mkdir(parents=True, exist_ok=True)
    _save_workbook(_array_formula_semantics_workbook(), directory / "baseline.xlsx")
    candidate = directory / "candidate.xlsx"
    _save_workbook(_array_formula_semantics_workbook(), candidate)
    _add_dynamic_array_metadata(candidate)
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.array_formula_mode_changed",
            title="An unchanged array formula switches from fixed CSE to dynamic spill semantics",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "array_formula_mode_changed",
                    "sheet": "Model",
                    "cell": "B1",
                    "formula": "=LEN(Inputs!A1:A3)",
                    "baseline_mode": "legacy_cse",
                    "candidate_mode": "dynamic",
                    "baseline_output_range": "B1:B3",
                    "candidate_output_range": "B1:B3",
                }
            ],
            must_reach=[
                {
                    "source": {"sheet": "Model", "cell": "B1"},
                    "targets": [{"sheet": "Dashboard", "cell": "B2"}],
                }
            ],
            coverage=[
                "The legacy CSE range and the candidate's currently stored output range coincide; a dynamic array can resize or be blocked when Excel recalculates.",
                "The validator checks stored formula mode and metadata only. It does not calculate a result, predict a future spill extent, identify blockers, or assert client-version compatibility.",
            ],
        ),
    )


def _build_structural_three_d_scope(root: Path) -> None:
    directory = root / "structural" / "three_d_scope_expansion"
    directory.mkdir(parents=True, exist_ok=True)
    _save_workbook(_three_d_workbook(), directory / "baseline.xlsx")
    _save_workbook(_three_d_workbook(include_adjustment=True), directory / "candidate.xlsx")
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.three_d_scope_expansion",
            title="A new period silently enters an unchanged 3-D formula span",
            family="structural",
            review_expectation="block",
            facts=[
                {
                    "kind": "three_d_scope_changed",
                    "formula_sheet": "Summary",
                    "formula_cell": "B5",
                    "inserted_sheet": "FebAdjustment",
                    "after_sheet": "Jan",
                    "before_sheet": "Feb",
                }
            ],
            coverage=[
                "The formula text is intentionally unchanged; the observable risk is tab-order scope."
            ],
        ),
    )


def _build_structural_formula_rewrite(root: Path) -> None:
    directory = root / "structural" / "formula_rewrite_after_column_insert"
    directory.mkdir(parents=True, exist_ok=True)
    _save_workbook(_refactor_workbook(), directory / "baseline.xlsx")
    _save_workbook(_refactor_workbook(rewritten=True), directory / "candidate.xlsx")
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="structural.formula_rewrite_after_column_insert",
            title="A column insertion rewrites formulas while retaining the declared inputs",
            family="structural",
            review_expectation="review",
            facts=[
                {
                    "kind": "structural_formula_rewrite",
                    "baseline": {"sheet": "Model", "cell": "D5", "formula": "=B5*C5"},
                    "candidate": {"sheet": "Model", "cell": "E5", "formula": "=B5*D5"},
                }
            ],
            coverage=[
                "The fixture declares equivalent logical input positions but does not prove general Excel semantic equivalence."
            ],
        ),
    )


def _build_structural_table_scope(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        ledger = workbook["Ledger"]
        ledger.append(["East", 15, 10, "=B5*C5"])
        ledger.tables["SalesLedger"].ref = "A1:D5"

    truth = _truth(
        case_id="structural.structured_table_scope_expansion",
        title="An Excel Table grows while a dependent structured reference stays unchanged",
        family="structural",
        review_expectation="block",
        facts=[
            {
                "kind": "structured_table_scope_changed",
                "table_sheet": "Ledger",
                "table": "SalesLedger",
                "baseline_ref": "A1:D4",
                "candidate_ref": "A1:D5",
                "formula_sheet": "Summary",
                "formula_cell": "B2",
            }
        ],
        coverage=[
            "The summary formula text intentionally remains unchanged; the observable risk is the stored Excel Table range.",
            "The contract does not calculate the structured reference, apply filters, or infer table-total semantics.",
        ],
    )
    _write_pair(
        root / "structural" / "structured_table_scope_expansion",
        _structured_table_workbook,
        mutate,
        truth,
    )


def _build_structural_dynamic_reference(root: Path) -> None:
    def mutate(workbook: Workbook) -> None:
        workbook["Summary"]["B2"] = "=INDIRECT(Inputs!$E$12)"

    truth = _truth(
        case_id="structural.dynamic_reference_introduced",
        title="A direct output reference becomes an INDIRECT-driven reference",
        family="structural",
        review_expectation="block",
        facts=[
            {
                "kind": "dynamic_formula_reference_added",
                "sheet": "Summary",
                "cell": "B2",
                "functions": ["INDIRECT"],
            }
        ],
        must_reach=[
            {
                "source": {"sheet": "Summary", "cell": "B2"},
                "targets": [{"sheet": "Dashboard", "cell": "B4"}],
            }
        ],
        coverage=[
            "INDIRECT returns a reference from workbook text; this fixture does not evaluate that text or claim complete target resolution.",
            "The scoreable coverage expectation requires an explicit static-dependency boundary disclosure, not a universal formula-evaluation claim.",
        ],
        coverage_expectations=[
            {
                "kind": "dynamic_reference_static_coverage",
                "sheet": "Summary",
                "cell": "B2",
                "functions": ["INDIRECT"],
            }
        ],
    )
    _write_pair(
        root / "structural" / "dynamic_reference_introduced",
        _dynamic_reference_workbook,
        mutate,
        truth,
    )


def _build_structural_dynamic_reference_drivers(root: Path) -> None:
    def write_case(
        *,
        case_name: str,
        case_id: str,
        function: str,
        baseline_driver: str | int,
        candidate_driver: str | int,
        title: str,
        coverage: list[str],
    ) -> None:
        truth = _truth(
            case_id=case_id,
            title=title,
            family="structural",
            review_expectation="block",
            facts=[{"kind": "value_changed", "sheet": "Inputs", "cell": "E12"}],
            must_reach=[
                {
                    "source": {"sheet": "Inputs", "cell": "E12"},
                    "targets": [
                        {"sheet": "Summary", "cell": "B2"},
                        {"sheet": "Dashboard", "cell": "B4"},
                    ],
                }
            ],
            coverage=coverage,
            coverage_expectations=[
                {
                    "kind": "dynamic_reference_driver_changed",
                    "driver": {"sheet": "Inputs", "cell": "E12"},
                    "formula": {"sheet": "Summary", "cell": "B2"},
                    "functions": [function],
                }
            ],
        )
        _write_pair(
            root / "structural" / case_name,
            lambda: _dynamic_reference_driver_workbook(
                function=function, driver_value=baseline_driver
            ),
            lambda workbook: setattr(workbook["Inputs"]["E12"], "value", candidate_driver),
            truth,
        )

    write_case(
        case_name="indirect_reference_driver_changed",
        case_id="structural.indirect_reference_driver_changed",
        function="INDIRECT",
        baseline_driver="Revenue!B8",
        candidate_driver="Revenue!C8",
        title="An unchanged INDIRECT formula receives a different address driver",
        coverage=[
            "INDIRECT returns a reference from workbook text; this fixture does not evaluate or calculate the selected address.",
            "The unchanged formula still directly reads the changed driver, but its effective target selection is outside a complete static dependency graph.",
        ],
    )
    write_case(
        case_name="offset_reference_driver_changed",
        case_id="structural.offset_reference_driver_changed",
        function="OFFSET",
        baseline_driver=0,
        candidate_driver=1,
        title="An unchanged OFFSET formula receives a different column displacement",
        coverage=[
            "OFFSET returns a reference displaced from another reference; this fixture does not calculate the selected cell.",
            "The unchanged formula still directly reads the changed driver, but its effective target selection is outside a complete static dependency graph.",
        ],
    )


def _build_portfolio_external_driver(root: Path) -> None:
    directory = root / "portfolio" / "external_driver_value_change"
    baseline = directory / "baseline"
    candidate = directory / "candidate"
    baseline.mkdir(parents=True, exist_ok=True)
    candidate.mkdir(parents=True, exist_ok=True)
    base_drivers, base_model = _portfolio_workbook(driver_value=10)
    candidate_drivers, candidate_model = _portfolio_workbook(driver_value=15)
    _save_workbook(base_drivers, baseline / "drivers.xlsx")
    _save_workbook(base_model, baseline / "model.xlsx")
    _save_workbook(candidate_drivers, candidate / "drivers.xlsx")
    _save_workbook(candidate_model, candidate / "model.xlsx")
    _write_json(
        directory / "truth.json",
        _truth(
            case_id="portfolio.external_driver_value_change",
            title="An upstream workbook driver changes while a downstream workbook consumes it",
            family="portfolio",
            review_expectation="review",
            topology="portfolio",
            facts=[
                {
                    "kind": "portfolio_value_changed",
                    "workbook": "drivers.xlsx",
                    "sheet": "Inputs",
                    "cell": "B2",
                },
                {
                    "kind": "portfolio_external_reference",
                    "workbook": "model.xlsx",
                    "sheet": "Summary",
                    "cell": "B2",
                    "target_workbook": "drivers.xlsx",
                },
            ],
            coverage=[
                "The target workbook is present only as a local sibling; no network or filesystem path resolution is permitted."
            ],
        ),
    )


_BUILDERS: tuple[Callable[[Path], None], ...] = (
    _build_finance_formula_to_value,
    _build_finance_wrong_period_reference,
    _build_finance_input_change,
    _build_finance_external_formula,
    _build_finance_defined_name,
    _build_operations_formula_interruption,
    _build_operations_sumifs_shape,
    _build_operations_data_validation,
    _build_operations_data_validation_list_source,
    _build_operations_conditional_formatting,
    _build_operations_conditional_formatting_threshold,
    _build_operations_number_format_visibility,
    _build_operations_auto_filter_criteria,
    _build_governance_visibility,
    _build_governance_protection,
    _build_governance_manual_calculation,
    _build_governance_iterative_calculation,
    _build_governance_precision_as_displayed,
    _build_governance_workbook_date_system,
    _build_governance_formula_cached_result,
    _build_governance_static_cycle,
    _build_governance_external_data_refresh,
    _build_governance_pivot_cache_refresh,
    _build_governance_external_workbook_link_update_policy,
    _build_structural_pivot_data_field_aggregation,
    _build_structural_pivot_slicer_selection,
    _build_structural_power_query_m_filter,
    _build_structural_scenario_manager_stored_input,
    _build_structural_what_if_data_table_input,
    _build_structural_chart_series_reference,
    _build_structural_array_formula_mode,
    _build_structural_three_d_scope,
    _build_structural_formula_rewrite,
    _build_structural_table_scope,
    _build_structural_dynamic_reference,
    _build_structural_dynamic_reference_drivers,
    _build_portfolio_external_driver,
)

CASE_IDS = (
    "finance.formula_to_value",
    "finance.wrong_period_reference",
    "finance.input_value_change",
    "finance.external_formula_reference",
    "finance.defined_name_redirect",
    "operations.copied_formula_interruption",
    "operations.sumifs_range_shape",
    "operations.data_validation_removed",
    "operations.data_validation_list_source_changed",
    "operations.conditional_formatting_removed",
    "operations.conditional_formatting_threshold_changed",
    "operations.number_format_value_hidden",
    "operations.auto_filter_criteria_changed",
    "governance.hidden_sheet_revealed",
    "governance.formula_cell_unlocked",
    "governance.manual_calculation_incomplete",
    "governance.iterative_calculation_enabled",
    "governance.precision_as_displayed_enabled",
    "governance.workbook_date_system_changed",
    "governance.formula_cached_result_changed",
    "governance.static_cycle_introduced",
    "governance.external_data_refresh_on_open",
    "governance.pivot_cache_refresh_on_open",
    "governance.external_workbook_link_update_on_open",
    "structural.pivot_data_field_aggregation_changed",
    "structural.pivot_slicer_selection_changed",
    "structural.power_query_m_filter_changed",
    "structural.scenario_manager_stored_input_changed",
    "structural.what_if_data_table_input_reference_changed",
    "structural.chart_series_reference_changed",
    "structural.array_formula_mode_changed",
    "structural.three_d_scope_expansion",
    "structural.formula_rewrite_after_column_insert",
    "structural.structured_table_scope_expansion",
    "structural.dynamic_reference_introduced",
    "structural.indirect_reference_driver_changed",
    "structural.offset_reference_driver_changed",
    "portfolio.external_driver_value_change",
)


def build_all(root: str | Path) -> list[Path]:
    """Build all WCAB fixtures below *root* and return their truth manifests."""

    fixture_root = Path(root).resolve()
    fixture_root.mkdir(parents=True, exist_ok=True)
    for builder in _BUILDERS:
        builder(fixture_root)
    manifests = sorted(fixture_root.rglob("truth.json"))
    if len(manifests) != len(CASE_IDS):
        raise RuntimeError(f"expected {len(CASE_IDS)} manifests, found {len(manifests)}")
    write_manifest(fixture_root, expected_ids=CASE_IDS)
    return manifests
