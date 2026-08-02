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

from ..build import CASE_IDS, FIXTURE_SCHEMA_VERSION
from ..score import OBSERVATION_SCHEMA_VERSION


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
    "dynamic_formula_reference_added": ("dynamic_formula_reference_added", "location"),
}

_COVERAGE_EXPECTATION_TO_RULE: dict[str, tuple[str, str | None]] = {
    "dynamic_reference_static_coverage": ("FF012", "location"),
}

_COVERAGE_EXPECTATION_TO_PROFILE_FEATURE = {
    "dynamic_reference_driver_changed": "dynamic_reference_cells",
}

_LINT_EXPECTATIONS = {
    "operations.copied_formula_interruption": "FF082",
    "operations.sumifs_range_shape": "FF093",
    "governance.formula_cell_unlocked": "FF085",
    "governance.manual_calculation_incomplete": "FF086",
    "governance.static_cycle_introduced": "FF090",
}

_PORTFOLIO_FACT_EVIDENCE = {
    "portfolio_value_changed": {"native_change_kind": "value_changed"},
    "portfolio_external_reference": {"native_rule_id": "FF079"},
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


def profile(workbook: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Run a local FormulaFence profile and return its JSON report."""

    return _run_json(["profile", str(workbook)], executable=executable)


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


def _nested_location(expectation: dict[str, Any], field: str) -> str | None:
    value = expectation.get(field)
    if not isinstance(value, dict):
        return None
    sheet = value.get("sheet")
    cell = value.get("cell")
    if isinstance(sheet, str) and isinstance(cell, str):
        return f"{sheet}!{cell}"
    return None


def _coverage_evidence(expectation: dict[str, Any]) -> dict[str, str] | None:
    kind = str(expectation.get("kind"))
    rule_mapping = _COVERAGE_EXPECTATION_TO_RULE.get(kind)
    if rule_mapping is not None:
        return {"native_rule_id": rule_mapping[0]}
    profile_feature = _COVERAGE_EXPECTATION_TO_PROFILE_FEATURE.get(kind)
    if profile_feature is not None:
        return {
            "native_change_kind": "value_changed",
            "native_profile_feature": profile_feature,
        }
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
        "matched_coverage_expectations": [],
        "missed_coverage_expectations": [],
        "unmapped_coverage_expectations": [],
        "change_count": report.get("summary", {}).get("change_count"),
    }


def evaluate_diff_case(case_dir: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Compare FormulaFence's diff evidence with mappable WCAB facts.

    Unmappable facts and coverage expectations remain visible in the result
    instead of being silently counted as passes. Portfolio facts are handled
    through FormulaFence's separate portfolio command.
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
        expected_functions = fact.get("functions")
        observed = any(
            change.get("kind") == expected_kind
            and (expected_location is None or change.get("location") == expected_location)
            and (
                expected_functions is None
                or (
                    isinstance(change.get("details"), dict)
                    and change["details"].get("functions") == expected_functions
                )
            )
            for change in changes
            if isinstance(change, dict)
        )
        (matched if observed else missed).append(str(kind))

    coverage_expectations = truth.get("coverage_expectations", [])
    if not isinstance(coverage_expectations, list):
        raise FormulaFenceAdapterError(f"{directory}: coverage_expectations must be an array")
    findings: list[Any] = []
    if coverage_expectations:
        raw_findings = report.get("findings", [])
        if not isinstance(raw_findings, list):
            raise FormulaFenceAdapterError("FormulaFence report has no findings list")
        findings = raw_findings
    candidate_profile: dict[str, Any] | None = None
    matched_coverage_expectations: list[dict[str, Any]] = []
    missed_coverage_expectations: list[dict[str, Any]] = []
    unmapped_coverage_expectations: list[dict[str, Any]] = []
    for expectation in coverage_expectations:
        if not isinstance(expectation, dict):
            raise FormulaFenceAdapterError(f"{directory}: coverage expectation must be an object")
        expectation_kind = str(expectation.get("kind"))
        rule_mapping = _COVERAGE_EXPECTATION_TO_RULE.get(expectation_kind)
        if rule_mapping is not None:
            expected_rule, location_mode = rule_mapping
            expected_location = _fact_location(expectation) if location_mode is not None else None
            expected_functions = expectation.get("functions")
            observed = any(
                isinstance(finding, dict)
                and finding.get("rule_id") == expected_rule
                and (expected_location is None or finding.get("location") == expected_location)
                and (
                    expected_functions is None
                    or (
                        isinstance(finding.get("details"), dict)
                        and finding["details"].get("functions") == expected_functions
                    )
                )
                for finding in findings
            )
        else:
            profile_feature = _COVERAGE_EXPECTATION_TO_PROFILE_FEATURE.get(expectation_kind)
            if profile_feature is None:
                unmapped_coverage_expectations.append(expectation)
                continue
            if candidate_profile is None:
                candidate_profile = profile(directory / "candidate.xlsx", executable=executable)
            features = candidate_profile.get("features")
            if not isinstance(features, dict):
                raise FormulaFenceAdapterError("FormulaFence profile has no features object")
            profile_entries = features.get(profile_feature)
            if not isinstance(profile_entries, list):
                raise FormulaFenceAdapterError(
                    f"FormulaFence profile feature {profile_feature!r} is not a list"
                )
            driver_location = _nested_location(expectation, "driver")
            formula_location = _nested_location(expectation, "formula")
            expected_functions = expectation.get("functions")
            driver_changed = any(
                isinstance(change, dict)
                and change.get("kind") == "value_changed"
                and change.get("location") == driver_location
                for change in changes
            )
            profile_observed = any(
                isinstance(entry, dict)
                and entry.get("location") == formula_location
                and entry.get("functions") == expected_functions
                for entry in profile_entries
            )
            observed = driver_changed and profile_observed
        (matched_coverage_expectations if observed else missed_coverage_expectations).append(
            expectation
        )
    return {
        "case_id": truth.get("id"),
        "status": "matched" if not missed else "missed",
        "matched": matched,
        "missed": missed,
        "unmapped": unmapped,
        "matched_coverage_expectations": matched_coverage_expectations,
        "missed_coverage_expectations": missed_coverage_expectations,
        "unmapped_coverage_expectations": unmapped_coverage_expectations,
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
    coverage_missed = [case for case in cases if case["missed_coverage_expectations"]]
    lint_missed = [case for case in lint_results if not case["matched"]]
    mapped_fact_count = sum(len(case["matched"]) + len(case["missed"]) for case in cases)
    matched_fact_count = sum(len(case["matched"]) for case in cases)
    unmapped_fact_count = sum(len(case["unmapped"]) for case in cases)
    mapped_coverage_expectation_count = sum(
        len(case["matched_coverage_expectations"]) + len(case["missed_coverage_expectations"])
        for case in cases
    )
    matched_coverage_expectation_count = sum(
        len(case["matched_coverage_expectations"]) for case in cases
    )
    unmapped_coverage_expectation_count = sum(
        len(case["unmapped_coverage_expectations"]) for case in cases
    )
    return {
        "tool": "FormulaFence",
        "case_count": len(cases),
        "matched_diff_case_count": sum(case["status"] == "matched" for case in cases),
        "mapped_diff_fact_count": mapped_fact_count,
        "matched_diff_fact_count": matched_fact_count,
        "unmapped_diff_fact_count": unmapped_fact_count,
        "diff_cases_with_misses": missed,
        "diff_cases": cases,
        "mapped_coverage_expectation_count": mapped_coverage_expectation_count,
        "matched_coverage_expectation_count": matched_coverage_expectation_count,
        "unmapped_coverage_expectation_count": unmapped_coverage_expectation_count,
        "coverage_expectations_with_misses": coverage_missed,
        "lint_rule_count": len(lint_results),
        "lint_rule_misses": lint_missed,
        "lint_rules": lint_results,
    }


def reference_observations(root: str | Path, *, executable: str = "formulafence") -> dict[str, Any]:
    """Emit scoreable normalized observations from the local FormulaFence adapter.

    This reuses the adapter's documented mapping from native FormulaFence
    evidence to WCAB facts. A matched fact is copied from WCAB truth only after
    the native change/rule mapping matched; the optional evidence object names
    that native mapping. FormulaFence is an analyzer rather than a policy
    engine, so this adapter intentionally leaves every review disposition null.
    """

    fixture_root = Path(root)
    directories: dict[str, tuple[Path, dict[str, Any]]] = {}
    for truth_path in fixture_root.rglob("truth.json"):
        truth = _load_truth(truth_path.parent)
        case_id = truth.get("id")
        if not isinstance(case_id, str):
            raise FormulaFenceAdapterError(f"{truth_path}: fixture case id must be a string")
        directories[case_id] = (truth_path.parent, truth)
    if set(directories) != set(CASE_IDS):
        raise FormulaFenceAdapterError("fixture tree does not match the WCAB case catalogue")

    cases: list[dict[str, Any]] = []
    for case_id in CASE_IDS:
        directory, truth = directories[case_id]
        result = evaluate_diff_case(directory, executable=executable)
        matched = set(result["matched"])
        facts: list[dict[str, Any]] = []
        for fact in truth.get("facts", []):
            kind = str(fact.get("kind"))
            if kind not in matched:
                continue
            if kind in _PORTFOLIO_FACT_EVIDENCE:
                evidence = _PORTFOLIO_FACT_EVIDENCE[kind]
            else:
                expectation = _FACT_TO_CHANGE.get(kind)
                if expectation is None:
                    continue
                evidence = {"native_change_kind": expectation[0]}
            facts.append({"evidence": evidence, "fact": fact})
        coverage_notes: list[str] = []
        if result["unmapped"]:
            coverage_notes.append(
                "FormulaFence adapter mapping intentionally unavailable for: "
                + ", ".join(sorted(result["unmapped"]))
            )
        unmapped_coverage_expectations = result.get("unmapped_coverage_expectations", [])
        if unmapped_coverage_expectations:
            coverage_notes.append(
                "FormulaFence adapter coverage mapping intentionally unavailable for: "
                + ", ".join(
                    sorted(
                        str(expectation.get("kind"))
                        for expectation in unmapped_coverage_expectations
                        if isinstance(expectation, dict)
                    )
                )
            )
        coverage_declarations: list[dict[str, Any]] = []
        for expectation in result.get("matched_coverage_expectations", []):
            if not isinstance(expectation, dict):
                continue
            evidence = _coverage_evidence(expectation)
            if evidence is None:
                continue
            coverage_declarations.append(
                {
                    "evidence": evidence,
                    "expectation": expectation,
                }
            )
        cases.append(
            {
                "coverage": {"declarations": coverage_declarations, "notes": coverage_notes},
                "facts": facts,
                "id": case_id,
                "review": None,
                "status": "analyzed",
            }
        )
    return {
        "benchmark": {"fixture_schema_versions": [FIXTURE_SCHEMA_VERSION]},
        "cases": cases,
        "schema_version": OBSERVATION_SCHEMA_VERSION,
        "tool": {"name": "FormulaFence"},
    }
