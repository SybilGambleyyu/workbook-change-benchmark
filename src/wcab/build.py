"""Deterministic generation of the WCAB fixture tree.

The fixtures are intentionally generated rather than copied from real business
workbooks.  That makes every assertion inspectable and the redistribution
rights unambiguous.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .manifest import write_manifest

FIXTURE_SCHEMA_VERSION = 3
_FIXED_TIMESTAMP = datetime(2024, 1, 1, tzinfo=timezone.utc)
_CORE_MODIFIED = re.compile(rb"(<dcterms:modified\b[^>]*>)[^<]*(</dcterms:modified>)")
WorkbookFactory = Callable[[], Workbook]
WorkbookMutator = Callable[[Workbook], None]


def _configure_workbook(workbook: Workbook, *, title: str) -> None:
    workbook.properties.creator = "WCAB"
    workbook.properties.lastModifiedBy = "WCAB"
    workbook.properties.title = title
    workbook.properties.subject = "Synthetic workbook change-assurance fixture"
    workbook.properties.created = _FIXED_TIMESTAMP
    workbook.properties.modified = _FIXED_TIMESTAMP
    workbook.calculation.fullCalcOnLoad = False
    workbook.calculation.forceFullCalc = False


def _canonicalize_xlsx(path: Path) -> None:
    """Rewrite a generated XLSX with stable ZIP order and timestamps.

    Openpyxl's XML generation is deterministic for the supported dependency
    versions, but ZIP member timestamps otherwise make byte-for-byte fixture
    reproduction impossible.
    """

    with ZipFile(path, "r") as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
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
    _build_operations_conditional_formatting,
    _build_governance_visibility,
    _build_governance_protection,
    _build_governance_manual_calculation,
    _build_governance_static_cycle,
    _build_structural_three_d_scope,
    _build_structural_formula_rewrite,
    _build_structural_table_scope,
    _build_structural_dynamic_reference,
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
    "operations.conditional_formatting_removed",
    "governance.hidden_sheet_revealed",
    "governance.formula_cell_unlocked",
    "governance.manual_calculation_incomplete",
    "governance.static_cycle_introduced",
    "structural.three_d_scope_expansion",
    "structural.formula_rewrite_after_column_insert",
    "structural.structured_table_scope_expansion",
    "structural.dynamic_reference_introduced",
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
