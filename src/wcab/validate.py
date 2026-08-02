"""Validation for WCAB's observable fixture truth contract."""

from __future__ import annotations

import json
import re
from collections import deque
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from .build import CASE_IDS, FIXTURE_SCHEMA_VERSION

_CELL_REFERENCE = re.compile(
    r"(?:(?:'(?P<quoted>(?:[^']|'')+)'|(?P<plain>[A-Za-z_][A-Za-z0-9_. ]*))!)?"
    r"(?P<cell>\$?[A-Z]{1,3}\$?[1-9][0-9]*)"
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


def _defined_name_text(workbook: Workbook, name: str) -> str | None:
    definition = workbook.defined_names.get(name)
    return None if definition is None else definition.attr_text


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
