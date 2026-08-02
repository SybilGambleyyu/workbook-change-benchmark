"""Tool-neutral observation validation and scoring for WCAB."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any

from .build import CASE_IDS, FIXTURE_SCHEMA_VERSION

OBSERVATION_SCHEMA_VERSION = 2
_CASE_STATUSES = frozenset({"analyzed", "unsupported", "error"})
_REVIEW_DISPOSITIONS = frozenset({"allow", "review", "block"})


class ObservationError(ValueError):
    """An observation document cannot be scored against WCAB."""


def _load_json(path: Path, *, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ObservationError(f"{path}: cannot load {label}: {error}") from error
    if not isinstance(value, dict):
        raise ObservationError(f"{path}: {label} must be a JSON object")
    return value


def _case_truths(fixture_root: str | Path) -> dict[str, dict[str, Any]]:
    root = Path(fixture_root).resolve()
    manifests = sorted(root.rglob("truth.json"))
    truths: dict[str, dict[str, Any]] = {}
    for manifest in manifests:
        truth = _load_json(manifest, label="truth manifest")
        case_id = truth.get("id")
        if not isinstance(case_id, str) or not case_id:
            raise ObservationError(f"{manifest}: truth manifest needs a non-empty id")
        if truth.get("schema_version") != FIXTURE_SCHEMA_VERSION:
            raise ObservationError(
                f"{manifest}: expected fixture schema {FIXTURE_SCHEMA_VERSION}, "
                f"found {truth.get('schema_version')!r}"
            )
        if case_id in truths:
            raise ObservationError(f"{manifest}: duplicate case id {case_id!r}")
        facts = truth.get("facts")
        if not isinstance(facts, list) or not all(isinstance(fact, dict) for fact in facts):
            raise ObservationError(f"{manifest}: facts must be an array of objects")
        coverage_expectations = truth.get("coverage_expectations")
        if not isinstance(coverage_expectations, list) or not all(
            isinstance(expectation, dict) for expectation in coverage_expectations
        ):
            raise ObservationError(f"{manifest}: coverage_expectations must be an array of objects")
        if truth.get("review_expectation") not in _REVIEW_DISPOSITIONS:
            raise ObservationError(f"{manifest}: invalid review expectation")
        truths[case_id] = truth
    if set(truths) != set(CASE_IDS):
        missing = sorted(set(CASE_IDS) - set(truths))
        unexpected = sorted(set(truths) - set(CASE_IDS))
        details = []
        if missing:
            details.append(f"missing: {', '.join(missing)}")
        if unexpected:
            details.append(f"unexpected: {', '.join(unexpected)}")
        raise ObservationError(f"fixture tree does not match WCAB catalogue ({'; '.join(details)})")
    return truths


def _canonical_item(item: dict[str, Any]) -> str:
    return json.dumps(item, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def _match_contract_items(
    expected_items: list[dict[str, Any]], reported_items: list[dict[str, Any]]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """Return matched, missing, and unrecognized targeted contract items."""

    expected_counts = Counter(_canonical_item(item) for item in expected_items)
    reported_counts = Counter(_canonical_item(item) for item in reported_items)
    matched_keys = Counter(expected_counts & reported_counts)
    matched: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    for item in expected_items:
        key = _canonical_item(item)
        if matched_keys[key]:
            matched.append(item)
            matched_keys[key] -= 1
        else:
            missing.append(item)
    unrecognized: list[dict[str, Any]] = []
    remaining_expected = Counter(expected_counts)
    for item in reported_items:
        key = _canonical_item(item)
        if remaining_expected[key]:
            remaining_expected[key] -= 1
        else:
            unrecognized.append(item)
    return matched, missing, unrecognized


def _require_string(value: Any, *, path: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ObservationError(f"{path}: expected a non-empty string")
    return value


def _validate_observations(
    observations: dict[str, Any], *, expected_ids: set[str]
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    if observations.get("schema_version") != OBSERVATION_SCHEMA_VERSION:
        raise ObservationError(f"observation schema_version must be {OBSERVATION_SCHEMA_VERSION}")
    benchmark = observations.get("benchmark")
    if not isinstance(benchmark, dict):
        raise ObservationError("observations.benchmark must be an object")
    expected_versions = [FIXTURE_SCHEMA_VERSION]
    if benchmark.get("fixture_schema_versions") != expected_versions:
        raise ObservationError(
            f"observations.benchmark.fixture_schema_versions must equal {expected_versions}"
        )
    tool = observations.get("tool")
    if not isinstance(tool, dict):
        raise ObservationError("observations.tool must be an object")
    _require_string(tool.get("name"), path="observations.tool.name")
    if "version" in tool and tool["version"] is not None and not isinstance(tool["version"], str):
        raise ObservationError("observations.tool.version must be a string when provided")

    raw_cases = observations.get("cases")
    if not isinstance(raw_cases, list):
        raise ObservationError("observations.cases must be an array")
    cases: dict[str, dict[str, Any]] = {}
    for index, raw_case in enumerate(raw_cases):
        path = f"observations.cases[{index}]"
        if not isinstance(raw_case, dict):
            raise ObservationError(f"{path}: case must be an object")
        case_id = _require_string(raw_case.get("id"), path=f"{path}.id")
        if case_id not in expected_ids:
            raise ObservationError(f"{path}: unknown WCAB case id {case_id!r}")
        if case_id in cases:
            raise ObservationError(f"{path}: duplicate WCAB case id {case_id!r}")
        status = raw_case.get("status")
        if status not in _CASE_STATUSES:
            raise ObservationError(
                f"{path}.status must be one of {', '.join(sorted(_CASE_STATUSES))}"
            )
        facts = raw_case.get("facts", [])
        if not isinstance(facts, list):
            raise ObservationError(f"{path}.facts must be an array")
        normalized_facts: list[dict[str, Any]] = []
        for fact_index, item in enumerate(facts):
            fact_path = f"{path}.facts[{fact_index}]"
            if not isinstance(item, dict):
                raise ObservationError(f"{fact_path}: observation must be an object")
            fact = item.get("fact")
            if not isinstance(fact, dict):
                raise ObservationError(f"{fact_path}.fact must be an object")
            _require_string(fact.get("kind"), path=f"{fact_path}.fact.kind")
            evidence = item.get("evidence")
            if evidence is not None and not isinstance(evidence, dict):
                raise ObservationError(f"{fact_path}.evidence must be an object when provided")
            normalized_facts.append(item)
        review = raw_case.get("review")
        if review is not None and review not in _REVIEW_DISPOSITIONS:
            raise ObservationError(f"{path}.review must be null or a WCAB review disposition")
        coverage = raw_case.get("coverage")
        if not isinstance(coverage, dict):
            raise ObservationError(f"{path}.coverage must be an object")
        declarations = coverage.get("declarations")
        if not isinstance(declarations, list):
            raise ObservationError(f"{path}.coverage.declarations must be an array")
        normalized_declarations: list[dict[str, Any]] = []
        for declaration_index, declaration in enumerate(declarations):
            declaration_path = f"{path}.coverage.declarations[{declaration_index}]"
            if not isinstance(declaration, dict):
                raise ObservationError(f"{declaration_path}: declaration must be an object")
            expectation = declaration.get("expectation")
            if not isinstance(expectation, dict):
                raise ObservationError(f"{declaration_path}.expectation must be an object")
            _require_string(expectation.get("kind"), path=f"{declaration_path}.expectation.kind")
            evidence = declaration.get("evidence")
            if evidence is not None and not isinstance(evidence, dict):
                raise ObservationError(
                    f"{declaration_path}.evidence must be an object when provided"
                )
            normalized_declarations.append(declaration)
        notes = coverage.get("notes")
        if not isinstance(notes, list) or not all(isinstance(item, str) for item in notes):
            raise ObservationError(f"{path}.coverage.notes must be an array of strings")
        error = raw_case.get("error")
        if status == "analyzed" and error is not None:
            raise ObservationError(f"{path}: analyzed cases cannot contain an error")
        if status in {"unsupported", "error"}:
            if normalized_facts:
                raise ObservationError(f"{path}: unsupported or errored cases cannot report facts")
            if review is not None:
                raise ObservationError(
                    f"{path}: unsupported or errored cases cannot report a review"
                )
            if normalized_declarations:
                raise ObservationError(
                    f"{path}: unsupported or errored cases cannot report coverage declarations"
                )
        if status == "error":
            _require_string(error, path=f"{path}.error")
        elif error is not None:
            raise ObservationError(f"{path}.error is only valid when status is error")
        cases[case_id] = {
            "coverage": {"declarations": normalized_declarations, "notes": notes},
            "error": error,
            "facts": normalized_facts,
            "id": case_id,
            "review": review,
            "status": status,
        }
    return tool, cases


def observation_template(fixture_root: str | Path) -> dict[str, Any]:
    """Return a valid, intentionally unsupported observation-report skeleton."""

    _case_truths(fixture_root)
    return {
        "benchmark": {"fixture_schema_versions": [FIXTURE_SCHEMA_VERSION]},
        "cases": [
            {
                "coverage": {"declarations": [], "notes": []},
                "facts": [],
                "id": case_id,
                "review": None,
                "status": "unsupported",
            }
            for case_id in CASE_IDS
        ],
        "schema_version": OBSERVATION_SCHEMA_VERSION,
        "tool": {"name": "replace-with-tool-name"},
    }


def load_observations(path: str | Path) -> dict[str, Any]:
    """Load an observation report without scoring it."""

    return _load_json(Path(path), label="observation report")


def score_observations(fixture_root: str | Path, observations: dict[str, Any]) -> dict[str, Any]:
    """Score normalized tool observations against WCAB's declared contract.

    WCAB facts and coverage expectations are intentionally targeted rather
    than a complete inventory of every workbook difference. The resulting
    report therefore gives expected recall and lists unrecognized observations
    separately; it does not label unlisted observations as false positives or
    claim a precision score.
    """

    truths = _case_truths(fixture_root)
    tool, reported_cases = _validate_observations(observations, expected_ids=set(truths))
    case_results: list[dict[str, Any]] = []
    expected_fact_count = 0
    matched_fact_count = 0
    unrecognized_fact_count = 0
    expected_coverage_expectation_count = 0
    matched_coverage_declaration_count = 0
    unrecognized_coverage_declaration_count = 0
    analyzed_case_count = 0
    unsupported_case_count = 0
    errored_case_count = 0
    not_reported_case_count = 0
    review_reported_count = 0
    review_matched_count = 0
    review_mismatched_count = 0
    complete_case_count = 0

    for case_id in CASE_IDS:
        truth = truths[case_id]
        expected_facts = truth["facts"]
        expected_coverage_expectations = truth["coverage_expectations"]
        expected_fact_count += len(expected_facts)
        expected_coverage_expectation_count += len(expected_coverage_expectations)
        report = reported_cases.get(case_id)
        if report is None:
            not_reported_case_count += 1
            status = "not_reported"
            reported_facts: list[dict[str, Any]] = []
            review = None
            coverage: dict[str, Any] = {"declarations": [], "notes": []}
            error = None
        else:
            status = report["status"]
            reported_facts = [item["fact"] for item in report["facts"]]
            review = report["review"]
            coverage = report["coverage"]
            error = report["error"]
            if status == "analyzed":
                analyzed_case_count += 1
            elif status == "unsupported":
                unsupported_case_count += 1
            else:
                errored_case_count += 1

        matched_facts, missing_facts, unrecognized_facts = _match_contract_items(
            expected_facts, reported_facts
        )
        reported_coverage_expectations = [
            declaration["expectation"] for declaration in coverage["declarations"]
        ]
        (
            matched_coverage_declarations,
            missing_coverage_expectations,
            unrecognized_coverage_declarations,
        ) = _match_contract_items(expected_coverage_expectations, reported_coverage_expectations)

        matched_fact_count += len(matched_facts)
        unrecognized_fact_count += len(unrecognized_facts)
        matched_coverage_declaration_count += len(matched_coverage_declarations)
        unrecognized_coverage_declaration_count += len(unrecognized_coverage_declarations)
        review_expected = truth["review_expectation"]
        review_matched = status == "analyzed" and review == review_expected
        if review is not None:
            review_reported_count += 1
            if review_matched:
                review_matched_count += 1
            else:
                review_mismatched_count += 1
        complete = (
            status == "analyzed"
            and not missing_facts
            and not unrecognized_facts
            and not missing_coverage_expectations
            and not unrecognized_coverage_declarations
            and review_matched
        )
        if complete:
            complete_case_count += 1
        case_results.append(
            {
                "case_id": case_id,
                "coverage": coverage,
                "error": error,
                "expected_coverage_expectations": expected_coverage_expectations,
                "expected_review": review_expected,
                "matched_coverage_declarations": matched_coverage_declarations,
                "matched_facts": matched_facts,
                "missing_coverage_expectations": missing_coverage_expectations,
                "missing_facts": missing_facts,
                "reported_review": review,
                "review_matched": review_matched,
                "status": status,
                "unrecognized_coverage_declarations": unrecognized_coverage_declarations,
                "unrecognized_facts": unrecognized_facts,
            }
        )

    case_count = len(CASE_IDS)
    return {
        "benchmark": {
            "case_count": case_count,
            "fixture_schema_versions": [FIXTURE_SCHEMA_VERSION],
            "expected_coverage_expectation_count": expected_coverage_expectation_count,
            "expected_fact_count": expected_fact_count,
        },
        "cases": case_results,
        "observation_schema_version": OBSERVATION_SCHEMA_VERSION,
        "summary": {
            "analyzed_case_count": analyzed_case_count,
            "analysis_coverage": analyzed_case_count / case_count,
            "complete_case_count": complete_case_count,
            "coverage_disclosure_recall": (
                matched_coverage_declaration_count / expected_coverage_expectation_count
                if expected_coverage_expectation_count
                else 1.0
            ),
            "error_case_count": errored_case_count,
            "expected_coverage_expectation_count": expected_coverage_expectation_count,
            "expected_fact_count": expected_fact_count,
            "fact_recall": matched_fact_count / expected_fact_count,
            "matched_fact_count": matched_fact_count,
            "matched_coverage_declaration_count": matched_coverage_declaration_count,
            "missing_coverage_expectation_count": (
                expected_coverage_expectation_count - matched_coverage_declaration_count
            ),
            "missing_fact_count": expected_fact_count - matched_fact_count,
            "not_reported_case_count": not_reported_case_count,
            "review_agreement": review_matched_count / case_count,
            "review_matched_count": review_matched_count,
            "review_mismatched_count": review_mismatched_count,
            "review_not_reported_count": case_count - review_reported_count,
            "unrecognized_fact_count": unrecognized_fact_count,
            "unrecognized_coverage_declaration_count": unrecognized_coverage_declaration_count,
            "unsupported_case_count": unsupported_case_count,
        },
        "tool": tool,
    }


def write_json(path: str | Path, value: dict[str, Any]) -> None:
    """Write a stable, human-readable JSON document."""

    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
