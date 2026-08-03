from __future__ import annotations

import json
from pathlib import Path

import pytest

from wcab.adapters import formulafence
from wcab.build import CASE_IDS, build_all
from wcab.cli import main
from wcab.score import ObservationError, observation_template, score_observations, write_json


def _truths(root: Path) -> dict[str, dict]:
    return {
        truth["id"]: truth
        for path in root.rglob("truth.json")
        if (truth := json.loads(path.read_text(encoding="utf-8")))
    }


def _perfect_observations(root: Path) -> dict:
    truths = _truths(root)
    observations = observation_template(root)
    for case in observations["cases"]:
        truth = truths[case["id"]]
        case["status"] = "analyzed"
        case["facts"] = [{"fact": fact} for fact in truth["facts"]]
        case["coverage"]["declarations"] = [
            {"expectation": expectation} for expectation in truth["coverage_expectations"]
        ]
        case["review"] = truth["review_expectation"]
    observations["tool"] = {"name": "example-adapter", "version": "test"}
    return observations


def test_observation_template_covers_the_catalogue(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    template = observation_template(fixture_root)
    assert template["benchmark"]["fixture_schema_versions"] == [3]
    assert [case["id"] for case in template["cases"]] == list(CASE_IDS)
    assert {case["status"] for case in template["cases"]} == {"unsupported"}


def test_score_reports_full_expected_fact_recall_and_policy_agreement(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    result = score_observations(fixture_root, _perfect_observations(fixture_root))
    summary = result["summary"]
    assert summary["expected_fact_count"] == 51
    assert summary["matched_fact_count"] == 51
    assert summary["missing_fact_count"] == 0
    assert summary["fact_recall"] == 1.0
    assert summary["expected_coverage_expectation_count"] == 3
    assert summary["matched_coverage_declaration_count"] == 3
    assert summary["coverage_disclosure_recall"] == 1.0
    assert summary["review_agreement"] == 1.0
    assert summary["complete_case_count"] == len(CASE_IDS)
    assert summary["unrecognized_fact_count"] == 0


def test_score_keeps_unrecognized_facts_separate_from_expected_recall(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    observations = _perfect_observations(fixture_root)
    observations["cases"][0]["facts"].append(
        {"fact": {"kind": "formula_changed", "sheet": "Missing", "cell": "A1"}}
    )
    result = score_observations(fixture_root, observations)
    summary = result["summary"]
    assert summary["fact_recall"] == 1.0
    assert summary["unrecognized_fact_count"] == 1
    assert summary["complete_case_count"] == len(CASE_IDS) - 1


def test_score_records_unsupported_cases_without_treating_them_as_passes(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    result = score_observations(fixture_root, observation_template(fixture_root))
    summary = result["summary"]
    assert summary["analyzed_case_count"] == 0
    assert summary["unsupported_case_count"] == len(CASE_IDS)
    assert summary["matched_fact_count"] == 0
    assert summary["missing_fact_count"] == 51
    assert summary["missing_coverage_expectation_count"] == 3
    assert summary["review_not_reported_count"] == len(CASE_IDS)


def test_score_rejects_duplicate_case_observations(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    observations = observation_template(fixture_root)
    observations["cases"].append(dict(observations["cases"][0]))
    with pytest.raises(ObservationError, match="duplicate"):
        score_observations(fixture_root, observations)


def test_score_requires_the_dynamic_reference_coverage_disclosure(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    observations = _perfect_observations(fixture_root)
    dynamic_case = next(
        case
        for case in observations["cases"]
        if case["id"] == "structural.dynamic_reference_introduced"
    )
    dynamic_case["coverage"]["declarations"] = []
    result = score_observations(fixture_root, observations)
    summary = result["summary"]
    assert summary["coverage_disclosure_recall"] == pytest.approx(2 / 3)
    assert summary["missing_coverage_expectation_count"] == 1
    assert summary["complete_case_count"] == len(CASE_IDS) - 1


def test_score_rejects_coverage_declarations_for_unsupported_cases(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    observations = observation_template(fixture_root)
    observations["cases"][0]["coverage"]["declarations"] = [
        {"expectation": {"kind": "dynamic_reference_static_coverage"}}
    ]
    with pytest.raises(ObservationError, match="cannot report coverage declarations"):
        score_observations(fixture_root, observations)


def test_score_cli_strictly_rejects_the_unsupported_template(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    observations_path = tmp_path / "observations.json"
    write_json(observations_path, observation_template(fixture_root))
    assert (
        main(
            [
                "score",
                "--fixtures",
                str(fixture_root),
                "--observations",
                str(observations_path),
                "--strict",
            ]
        )
        == 1
    )


def test_score_cli_strictly_rejects_a_missing_coverage_disclosure(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    observations = _perfect_observations(fixture_root)
    dynamic_case = next(
        case
        for case in observations["cases"]
        if case["id"] == "structural.dynamic_reference_introduced"
    )
    dynamic_case["coverage"]["declarations"] = []
    observations_path = tmp_path / "observations.json"
    write_json(observations_path, observations)
    assert (
        main(
            [
                "score",
                "--fixtures",
                str(fixture_root),
                "--observations",
                str(observations_path),
                "--strict",
            ]
        )
        == 1
    )


def test_formulafence_reference_observations_are_scoreable(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_evaluate(case_dir, **_kwargs):
        truth = json.loads((Path(case_dir) / "truth.json").read_text(encoding="utf-8"))
        matched = [fact["kind"] for fact in truth["facts"]]
        unmapped = []
        matched_coverage_expectations = truth["coverage_expectations"]
        if truth["id"] == "structural.formula_rewrite_after_column_insert":
            matched = []
            unmapped = ["structural_formula_rewrite"]
        return {
            "case_id": truth["id"],
            "matched": matched,
            "missed": [],
            "status": "matched",
            "unmapped": unmapped,
            "matched_coverage_expectations": matched_coverage_expectations,
            "missed_coverage_expectations": [],
            "unmapped_coverage_expectations": [],
        }

    monkeypatch.setattr(formulafence, "evaluate_diff_case", fake_evaluate)
    observations = formulafence.reference_observations(fixture_root)
    result = score_observations(fixture_root, observations)
    summary = result["summary"]
    assert summary["analyzed_case_count"] == len(CASE_IDS)
    assert summary["matched_fact_count"] == 50
    assert summary["missing_fact_count"] == 1
    assert summary["matched_coverage_declaration_count"] == 3
    assert summary["coverage_disclosure_recall"] == 1.0
    assert summary["review_not_reported_count"] == len(CASE_IDS)
