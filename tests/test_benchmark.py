from __future__ import annotations

import json
from hashlib import sha256
from pathlib import Path

from wcab.adapters import formulafence
from wcab.build import CASE_IDS, build_all
from wcab.manifest import case_rows, manifest_text
from wcab.validate import validate_all, validate_case

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _tree_hashes(root: Path) -> dict[str, str]:
    return {
        path.relative_to(root).as_posix(): sha256(path.read_bytes()).hexdigest()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }


def test_committed_fixtures_validate() -> None:
    assert validate_all(PROJECT_ROOT / "fixtures") == {}


def test_build_is_byte_reproducible(tmp_path: Path) -> None:
    first = tmp_path / "first"
    second = tmp_path / "second"
    assert len(build_all(first)) == len(CASE_IDS)
    assert len(build_all(second)) == len(CASE_IDS)
    assert _tree_hashes(first) == _tree_hashes(second)


def test_committed_tree_matches_regeneration(tmp_path: Path) -> None:
    regenerated = tmp_path / "regenerated"
    build_all(regenerated)
    assert _tree_hashes(PROJECT_ROOT / "fixtures") == _tree_hashes(regenerated)


def test_truth_ids_are_complete_and_unique() -> None:
    manifests = sorted((PROJECT_ROOT / "fixtures").rglob("truth.json"))
    identifiers = [json.loads(path.read_text(encoding="utf-8"))["id"] for path in manifests]
    assert tuple(sorted(identifiers)) == tuple(sorted(CASE_IDS))
    assert len(set(identifiers)) == len(identifiers)


def test_committed_manifest_matches_fixture_tree() -> None:
    fixture_root = PROJECT_ROOT / "fixtures"
    rows = case_rows(fixture_root, expected_ids=CASE_IDS)
    assert [row["id"] for row in rows] == sorted(CASE_IDS)
    assert all(row["baseline_files"] and row["candidate_files"] for row in rows)
    assert all("coverage_expectations" in row for row in rows)
    dynamic_row = next(
        row for row in rows if row["id"] == "structural.dynamic_reference_introduced"
    )
    assert dynamic_row["coverage_expectations"] == [
        {
            "kind": "dynamic_reference_static_coverage",
            "sheet": "Summary",
            "cell": "B2",
            "functions": ["INDIRECT"],
        }
    ]
    driver_row = next(
        row for row in rows if row["id"] == "structural.indirect_reference_driver_changed"
    )
    assert driver_row["coverage_expectations"] == [
        {
            "kind": "dynamic_reference_driver_changed",
            "driver": {"sheet": "Inputs", "cell": "E12"},
            "formula": {"sheet": "Summary", "cell": "B2"},
            "functions": ["INDIRECT"],
        }
    ]
    external_data_row = next(
        row for row in rows if row["id"] == "governance.external_data_refresh_on_open"
    )
    assert external_data_row["facts"] == [
        {
            "kind": "external_data_connection_refresh_on_load_changed",
            "connection_id": 1,
            "baseline_refresh_on_load": False,
            "candidate_refresh_on_load": True,
        }
    ]
    assert (fixture_root / "manifest.jsonl").read_text(encoding="utf-8") == manifest_text(rows)


def test_manifest_is_reproducible_with_fixture_build(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    first = (fixture_root / "manifest.jsonl").read_text(encoding="utf-8")
    build_all(fixture_root)
    assert (fixture_root / "manifest.jsonl").read_text(encoding="utf-8") == first


def test_validator_rejects_a_false_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "finance" / "formula_to_value"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["cell"] = "A1"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_structured_table_scope(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "structured_table_scope_expansion"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_ref"] = "A1:D6"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_dynamic_reference_coverage_expectation(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "dynamic_reference_introduced"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["coverage_expectations"][0]["functions"] = ["OFFSET"]
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_dynamic_reference_driver_expectation(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "indirect_reference_driver_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["coverage_expectations"][0]["formula"]["cell"] = "B3"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_false_external_data_refresh_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_data_refresh_on_open"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_refresh_on_load"] = False
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_formulafence_adapter_maps_a_pair_fact(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [{"kind": "formula_to_value", "location": "Revenue!C8"}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(fixture_root / "finance" / "formula_to_value")
    assert result["status"] == "matched"
    assert result["matched"] == ["formula_to_value"]


def test_formulafence_adapter_maps_a_structured_table_scope_fact(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [{"kind": "table_definition_changed", "location": None}],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "structured_table_scope_expansion"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["structured_table_scope_changed"]


def test_formulafence_adapter_maps_dynamic_fact_and_coverage_expectation(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 2},
            "changes": [
                {
                    "kind": "dynamic_formula_reference_added",
                    "location": "Summary!B2",
                    "details": {"functions": ["INDIRECT"]},
                }
            ],
            "findings": [
                {
                    "rule_id": "FF012",
                    "location": "Summary!B2",
                    "details": {"functions": ["INDIRECT"]},
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "dynamic_reference_introduced"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["dynamic_formula_reference_added"]
    assert len(result["matched_coverage_expectations"]) == 1
    assert result["missed_coverage_expectations"] == []


def test_formulafence_adapter_maps_dynamic_driver_coverage_from_profile(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [{"kind": "value_changed", "location": "Inputs!E12"}],
        }

    def fake_profile(*_args, **_kwargs):
        return {
            "features": {
                "dynamic_reference_cells": [{"location": "Summary!B2", "functions": ["INDIRECT"]}]
            }
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    monkeypatch.setattr(formulafence, "profile", fake_profile)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "indirect_reference_driver_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["value_changed"]
    assert len(result["matched_coverage_expectations"]) == 1
    assert result["missed_coverage_expectations"] == []


def test_formulafence_adapter_maps_external_data_refresh_fact(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_connections_changed",
                    "location": None,
                    "details": {
                        "before": [{"id": 1, "refresh_on_load": False}],
                        "after": [{"id": 1, "refresh_on_load": True}],
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_data_refresh_on_open"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["external_data_connection_refresh_on_load_changed"]


def test_formulafence_adapter_requires_the_exact_external_data_refresh_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_connections_changed",
                    "location": None,
                    "details": {
                        "before": [{"id": 1, "refresh_on_load": False}],
                        "after": [{"id": 1, "refresh_on_load": False}],
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_data_refresh_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_data_connection_refresh_on_load_changed"]


def test_formulafence_adapter_maps_a_portfolio_impact(monkeypatch, tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_portfolio(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "workbooks": [
                {
                    "path": "drivers.xlsx",
                    "changes": [{"kind": "value_changed", "location": "Inputs!B2"}],
                    "findings": [
                        {
                            "rule_id": "FF079",
                            "details": {"sample_impacts": [{"workbook": "model.xlsx"}]},
                        }
                    ],
                }
            ],
        }

    monkeypatch.setattr(formulafence, "portfolio", fake_portfolio)
    result = formulafence.evaluate_diff_case(
        fixture_root / "portfolio" / "external_driver_value_change"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["portfolio_value_changed", "portfolio_external_reference"]
