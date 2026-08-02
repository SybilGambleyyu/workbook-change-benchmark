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
