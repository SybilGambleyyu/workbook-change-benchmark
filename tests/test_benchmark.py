from __future__ import annotations

import json
from hashlib import sha256
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

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
    external_link_policy_row = next(
        row for row in rows if row["id"] == "governance.external_workbook_link_update_on_open"
    )
    assert external_link_policy_row["facts"] == [
        {
            "kind": "external_workbook_link_update_policy_changed",
            "sheet": "LinkedModel",
            "cell": "B2",
            "formula": "='[WCABSource.xlsx]Inputs'!$B$2",
            "baseline_update_links": "never",
            "candidate_update_links": "always",
        }
    ]
    iterative_calculation_row = next(
        row for row in rows if row["id"] == "governance.iterative_calculation_enabled"
    )
    assert iterative_calculation_row["facts"] == [
        {
            "kind": "iterative_calculation_enabled",
            "sheet": "Model",
            "cell": "B2",
            "formula": "=(B2+Inputs!$B$2)/2",
            "baseline_iterate": False,
            "candidate_iterate": True,
            "iteration_count": 100,
            "iteration_delta": 0.001,
        }
    ]
    array_row = next(row for row in rows if row["id"] == "structural.array_formula_mode_changed")
    assert array_row["facts"] == [
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


def test_validator_rejects_a_false_external_workbook_link_policy_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_workbook_link_update_on_open"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_update_links"] = "never"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_external_workbook_link_policy(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = (
        fixture_root / "governance" / "external_workbook_link_update_on_open" / "candidate.xlsx"
    )
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'updateLinks="always"', b'updateLinks="never"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_external_workbook_link_policy_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "external_workbook_link_update_on_open"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_iterative_calculation_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "iterative_calculation_enabled"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["iteration_count"] = 99
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_corrupted_iterative_calculation_control(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "iterative_calculation_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'iterate="1"', b'iterate="0"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_validator_rejects_an_unrelated_iterative_calculation_control_change(
    tmp_path: Path,
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "governance" / "iterative_calculation_enabled" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/workbook.xml"] = members["xl/workbook.xml"].replace(
        b'forceFullCalc="0"', b'forceFullCalc="1"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


def test_iterative_calculation_pair_changes_only_workbook_xml(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "governance" / "iterative_calculation_enabled"
    with ZipFile(case / "baseline.xlsx") as baseline, ZipFile(case / "candidate.xlsx") as candidate:
        baseline_members = {
            entry.filename: baseline.read(entry.filename) for entry in baseline.infolist()
        }
        candidate_members = {
            entry.filename: candidate.read(entry.filename) for entry in candidate.infolist()
        }
    assert set(baseline_members) == set(candidate_members)
    assert [
        member
        for member in sorted(baseline_members)
        if baseline_members[member] != candidate_members[member]
    ] == ["xl/workbook.xml"]


def test_validator_rejects_a_false_array_formula_mode_fact(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    case = fixture_root / "structural" / "array_formula_mode_changed"
    truth_path = case / "truth.json"
    truth = json.loads(truth_path.read_text(encoding="utf-8"))
    truth["facts"][0]["candidate_mode"] = "legacy_cse"
    truth_path.write_text(json.dumps(truth), encoding="utf-8")
    assert validate_case(case)


def test_validator_rejects_a_missing_dynamic_array_metadata_binding(tmp_path: Path) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)
    candidate = fixture_root / "structural" / "array_formula_mode_changed" / "candidate.xlsx"
    with ZipFile(candidate) as archive:
        members = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    members["xl/metadata.xml"] = members["xl/metadata.xml"].replace(
        b'fDynamic="1"', b'fDynamic="0"', 1
    )
    staging = candidate.with_suffix(".corrupt.xlsx")
    with ZipFile(staging, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    staging.replace(candidate)
    assert validate_case(candidate.parent)


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


def test_formulafence_adapter_maps_the_exact_external_workbook_link_policy(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_refresh_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "update_links": "never",
                            "allow_refresh_query": False,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                        "after": {
                            "update_links": "always",
                            "allow_refresh_query": False,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_update_on_open"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["external_workbook_link_update_policy_changed"]


def test_formulafence_adapter_rejects_an_inexact_external_workbook_link_policy(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "external_data_refresh_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "update_links": "never",
                            "allow_refresh_query": False,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                        "after": {
                            "update_links": "always",
                            "allow_refresh_query": True,
                            "refresh_all_connections": False,
                            "save_external_link_values": True,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "external_workbook_link_update_on_open"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["external_workbook_link_update_policy_changed"]


def test_formulafence_adapter_maps_the_exact_iterative_calculation_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "calculation_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": False,
                            "iterateCount": 100,
                            "iterateDelta": 0.001,
                        },
                        "after": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": True,
                            "iterateCount": 100,
                            "iterateDelta": 0.001,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "iterative_calculation_enabled"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["iterative_calculation_enabled"]


def test_formulafence_adapter_rejects_an_inexact_iterative_calculation_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "calculation_settings_changed",
                    "location": None,
                    "details": {
                        "before": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": False,
                            "iterateCount": 100,
                            "iterateDelta": 0.001,
                        },
                        "after": {
                            "forceFullCalc": False,
                            "fullCalcOnLoad": False,
                            "iterate": True,
                            "iterateCount": 100,
                            "iterateDelta": 0.01,
                        },
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "governance" / "iterative_calculation_enabled"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["iterative_calculation_enabled"]


def test_formulafence_adapter_maps_the_exact_array_formula_mode_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "array_formula_mode_changed",
                    "location": "Model!B1",
                    "details": {
                        "before": {"mode": "legacy_cse", "output_range": "B1:B3"},
                        "after": {"mode": "dynamic", "output_range": "B1:B3"},
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "array_formula_mode_changed"
    )
    assert result["status"] == "matched"
    assert result["matched"] == ["array_formula_mode_changed"]


def test_formulafence_adapter_rejects_an_inexact_array_formula_mode_transition(
    monkeypatch, tmp_path: Path
) -> None:
    fixture_root = tmp_path / "fixtures"
    build_all(fixture_root)

    def fake_diff(*_args, **_kwargs):
        return {
            "summary": {"change_count": 1},
            "changes": [
                {
                    "kind": "array_formula_mode_changed",
                    "location": "Model!B1",
                    "details": {
                        "before": {"mode": "legacy_cse", "output_range": "B1:B3"},
                        "after": {"mode": "dynamic", "output_range": "B1:B4"},
                    },
                }
            ],
        }

    monkeypatch.setattr(formulafence, "diff", fake_diff)
    result = formulafence.evaluate_diff_case(
        fixture_root / "structural" / "array_formula_mode_changed"
    )
    assert result["status"] == "missed"
    assert result["matched"] == []
    assert result["missed"] == ["array_formula_mode_changed"]


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
