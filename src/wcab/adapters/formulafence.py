"""Local-only FormulaFence adapter.

FormulaFence is not a dependency of WCAB. This adapter shells out to an
already installed FormulaFence executable and normalizes its JSON evidence so
the benchmark can demonstrate one concrete implementation without making its
report schema normative for other tools.
"""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any


class FormulaFenceAdapterError(RuntimeError):
    """FormulaFence could not produce a usable local JSON report."""


_FACT_TO_CHANGE: dict[str, tuple[str, str | None]] = {
    "formula_to_value": ("formula_to_value", "location"),
    "formula_changed": ("formula_changed", "location"),
    "value_changed": ("value_changed", "location"),
    "external_formula_added": ("external_workbook_link_surfaces_changed", None),
    "defined_name_changed": ("defined_name_changed", None),
    "data_validation_count_changed": ("data_validation_changed", None),
    "conditional_formatting_count_changed": ("conditional_formatting_changed", None),
    "sheet_visibility_changed": ("sheet_visibility_changed", None),
    "formula_cell_unlocked": ("cell_protection_assignments_changed", None),
    "manual_calculation_incomplete": ("calculation_settings_changed", None),
    "static_cycle_introduced": ("formula_changed", None),
    "three_d_scope_changed": ("three_d_reference_scope_changed", "formula_location"),
    "structured_table_scope_changed": ("table_definition_changed", None),
}

_LINT_EXPECTATIONS = {
    "operations.copied_formula_interruption": "FF082",
    "operations.sumifs_range_shape": "FF093",
    "governance.formula_cell_unlocked": "FF085",
    "governance.manual_calculation_incomplete": "FF086",
    "governance.static_cycle_introduced": "FF090",
}


def _resolve_executable(executable: str) -> str:
    resolved = shutil.which(executable)
    if resolved is None:
        candidate = Path(executable)
        if candidate.is_file():
            return str(candidate)
        raise FormulaFenceAdapterError(
            f"FormulaFence executable {executable!r} was not found; install it or pass --executable."
        )
    return resolved


def _run_json(arguments: list[str], *, executable: str, timeout: int = 120) -> dict[str, Any]:
    command = [_resolve_executable(executable), *arguments]
    with TemporaryDirectory(prefix="wcab-formulafence-") as temporary:
        report_path = Path(temporary) / "report.json"
        try:
            completed = subprocess.run(
                [*command, "--format", "json", "--output", str(report_path)],
                text=True,
                capture_output=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired as error:
            raise FormulaFenceAdapterError(
                f"FormulaFence exceeded the {timeout}-second adapter limit"
            ) from error
        if completed.returncode != 0:
            raise FormulaFenceAdapterError(
                f"FormulaFence exited {completed.returncode}: {completed.stderr.strip()}"
            )
        if not report_path.is_file():
            raise FormulaFenceAdapterError(
                f"FormulaFence did not write JSON (exit {completed.returncode}): {completed.stderr.strip()}"
            )
        try:
            report = json.loads(report_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise FormulaFenceAdapterError(f"FormulaFence emitted invalid JSON: {error}") from error
    if not isinstance(report, dict):
        raise FormulaFenceAdapterError("FormulaFence report root must be a JSON object")
    return report


def diff(
    baseline: str | Path, candidate: str | Path, *, executable: str = "formulafence"
) -> dict[str, Any]:
    """Run a local FormulaFence semantic diff and return its JSON report."""

    return _run_json(["diff", str(baseline), str(candidate)], executable=executable)


def lint(workbook: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Run local FormulaFence lint and return its JSON report."""

    return _run_json(["lint", str(workbook)], executable=executable)


def portfolio(
    baseline: str | Path, candidate: str | Path, *, executable: str = "formulafence"
) -> dict[str, Any]:
    """Run FormulaFence's local directory-portfolio comparison."""

    return _run_json(["portfolio", str(baseline), str(candidate)], executable=executable)


def _load_truth(case_dir: Path) -> dict[str, Any]:
    result = json.loads((case_dir / "truth.json").read_text(encoding="utf-8"))
    if not isinstance(result, dict):
        raise FormulaFenceAdapterError(f"{case_dir}: truth root must be an object")
    return result


def _fact_location(fact: dict[str, Any]) -> str | None:
    if fact["kind"] == "three_d_scope_changed":
        return f"{fact['formula_sheet']}!{fact['formula_cell']}"
    sheet = fact.get("sheet")
    cell = fact.get("cell")
    if isinstance(sheet, str) and isinstance(cell, str):
        return f"{sheet}!{cell}"
    return None


def _portfolio_entry(report: dict[str, Any], path: str) -> dict[str, Any] | None:
    workbooks = report.get("workbooks", [])
    if not isinstance(workbooks, list):
        raise FormulaFenceAdapterError("FormulaFence portfolio report has no workbooks list")
    for entry in workbooks:
        if isinstance(entry, dict) and entry.get("path") == path:
            return entry
    return None


def _evaluate_portfolio_case(
    directory: Path, truth: dict[str, Any], *, executable: str
) -> dict[str, Any]:
    report = portfolio(directory / "baseline", directory / "candidate", executable=executable)
    matched: list[str] = []
    missed: list[str] = []
    unmapped: list[str] = []
    for fact in truth.get("facts", []):
        kind = fact.get("kind")
        if kind == "portfolio_value_changed":
            entry = _portfolio_entry(report, str(fact.get("workbook")))
            changes = [] if entry is None else entry.get("changes", [])
            location = f"{fact.get('sheet')}!{fact.get('cell')}"
            observed = any(
                isinstance(change, dict)
                and change.get("kind") == "value_changed"
                and change.get("location") == location
                for change in changes
            )
        elif kind == "portfolio_external_reference":
            source_entry = _portfolio_entry(report, str(fact.get("target_workbook")))
            expected_downstream = fact.get("workbook")
            observed = False
            findings = [] if source_entry is None else source_entry.get("findings", [])
            for finding in findings:
                if not isinstance(finding, dict) or finding.get("rule_id") != "FF079":
                    continue
                samples = finding.get("details", {}).get("sample_impacts", [])
                if any(
                    isinstance(sample, dict) and sample.get("workbook") == expected_downstream
                    for sample in samples
                ):
                    observed = True
                    break
        else:
            unmapped.append(str(kind))
            continue
        (matched if observed else missed).append(str(kind))
    return {
        "case_id": truth.get("id"),
        "status": "matched" if not missed else "missed",
        "matched": matched,
        "missed": missed,
        "unmapped": unmapped,
        "change_count": report.get("summary", {}).get("change_count"),
    }


def evaluate_diff_case(case_dir: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Compare FormulaFence's diff evidence with mappable WCAB facts.

    Unmappable facts remain visible in the result instead of being silently
    counted as passes. Portfolio topology is intentionally left to a future
    adapter extension because FormulaFence uses a different portfolio command.
    """

    directory = Path(case_dir)
    truth = _load_truth(directory)
    if truth.get("topology") == "portfolio":
        return _evaluate_portfolio_case(directory, truth, executable=executable)
    if truth.get("topology") != "pair":
        raise FormulaFenceAdapterError(
            f"{directory}: unsupported topology {truth.get('topology')!r}"
        )

    report = diff(directory / "baseline.xlsx", directory / "candidate.xlsx", executable=executable)
    changes = report.get("changes", [])
    if not isinstance(changes, list):
        raise FormulaFenceAdapterError("FormulaFence report has no changes list")

    matched: list[str] = []
    missed: list[str] = []
    unmapped: list[str] = []
    for fact in truth.get("facts", []):
        kind = fact.get("kind")
        expectation = _FACT_TO_CHANGE.get(kind)
        if expectation is None:
            unmapped.append(str(kind))
            continue
        expected_kind, location_mode = expectation
        expected_location = _fact_location(fact) if location_mode is not None else None
        observed = any(
            change.get("kind") == expected_kind
            and (expected_location is None or change.get("location") == expected_location)
            for change in changes
            if isinstance(change, dict)
        )
        (matched if observed else missed).append(str(kind))
    return {
        "case_id": truth.get("id"),
        "status": "matched" if not missed else "missed",
        "matched": matched,
        "missed": missed,
        "unmapped": unmapped,
        "change_count": report.get("summary", {}).get("change_count"),
    }


def evaluate_reference_suite(
    root: str | Path, *, executable: str = "formulafence"
) -> dict[str, Any]:
    """Run the documented FormulaFence reference pass across WCAB fixtures."""

    fixture_root = Path(root)
    cases: list[dict[str, Any]] = []
    lint_results: list[dict[str, Any]] = []
    for truth_path in sorted(fixture_root.rglob("truth.json")):
        case_dir = truth_path.parent
        case_result = evaluate_diff_case(case_dir, executable=executable)
        cases.append(case_result)
        expected_rule = _LINT_EXPECTATIONS.get(str(case_result["case_id"]))
        if expected_rule is not None:
            report = lint(case_dir / "candidate.xlsx", executable=executable)
            findings = report.get("findings", [])
            observed_rules = {
                finding.get("rule_id") for finding in findings if isinstance(finding, dict)
            }
            lint_results.append(
                {
                    "case_id": case_result["case_id"],
                    "expected_rule": expected_rule,
                    "matched": expected_rule in observed_rules,
                    "observed_rules": sorted(
                        rule for rule in observed_rules if isinstance(rule, str)
                    ),
                }
            )
    missed = [case for case in cases if case["missed"]]
    lint_missed = [case for case in lint_results if not case["matched"]]
    mapped_fact_count = sum(len(case["matched"]) + len(case["missed"]) for case in cases)
    matched_fact_count = sum(len(case["matched"]) for case in cases)
    unmapped_fact_count = sum(len(case["unmapped"]) for case in cases)
    return {
        "tool": "FormulaFence",
        "case_count": len(cases),
        "matched_diff_case_count": sum(case["status"] == "matched" for case in cases),
        "mapped_diff_fact_count": mapped_fact_count,
        "matched_diff_fact_count": matched_fact_count,
        "unmapped_diff_fact_count": unmapped_fact_count,
        "diff_cases_with_misses": missed,
        "diff_cases": cases,
        "lint_rule_count": len(lint_results),
        "lint_rule_misses": lint_missed,
        "lint_rules": lint_results,
    }
