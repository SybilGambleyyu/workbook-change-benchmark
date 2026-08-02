"""Validation for WCAB's observable fixture truth contract."""

from __future__ import annotations

import json
import re
from collections import deque
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
_DYNAMIC_ARRAY_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"


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
