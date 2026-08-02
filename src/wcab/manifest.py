"""Machine-readable, integrity-addressed catalogues for WCAB fixtures."""

from __future__ import annotations

import json
from collections.abc import Iterable
from hashlib import sha256
from pathlib import Path
from typing import Any


class ManifestError(ValueError):
    """The fixture tree cannot be represented as a complete WCAB manifest."""


def _load_truth(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ManifestError(f"{path}: cannot load truth manifest: {error}") from error
    if not isinstance(value, dict):
        raise ManifestError(f"{path}: truth manifest must be a JSON object")
    return value


def _file_record(path: Path, fixture_root: Path) -> dict[str, Any]:
    if not path.is_file():
        raise ManifestError(f"expected fixture file is absent: {path}")
    payload = path.read_bytes()
    return {
        "bytes": len(payload),
        "path": path.relative_to(fixture_root).as_posix(),
        "sha256": sha256(payload).hexdigest(),
    }


def _role_files(
    case_dir: Path, fixture_root: Path, topology: str, role: str
) -> list[dict[str, Any]]:
    if topology == "pair":
        return [_file_record(case_dir / f"{role}.xlsx", fixture_root)]
    if topology == "portfolio":
        role_root = case_dir / role
        if not role_root.is_dir():
            raise ManifestError(f"expected portfolio directory is absent: {role_root}")
        return [
            _file_record(path, fixture_root)
            for path in sorted(role_root.rglob("*"))
            if path.is_file()
        ]
    raise ManifestError(f"{case_dir}: unsupported topology {topology!r}")


def case_rows(
    fixture_root: str | Path, *, expected_ids: Iterable[str] | None = None
) -> list[dict[str, Any]]:
    """Return stable, one-row-per-case records for a fixture tree.

    The rows preserve the full truth contract and add exact paths, byte counts,
    and SHA-256 digests for every baseline and candidate workbook. They do not
    evaluate formulas or turn the benchmark's static boundaries into claims of
    Excel semantic equivalence.
    """

    root = Path(fixture_root).resolve()
    manifests = sorted(root.rglob("truth.json"))
    rows: list[dict[str, Any]] = []
    identifiers: list[str] = []
    for truth_path in manifests:
        truth = _load_truth(truth_path)
        case_dir = truth_path.parent
        case_id = truth.get("id")
        if not isinstance(case_id, str) or not case_id:
            raise ManifestError(f"{truth_path}: id must be a non-empty string")
        topology = truth.get("topology")
        if not isinstance(topology, str):
            raise ManifestError(f"{truth_path}: topology must be a string")
        identifiers.append(case_id)
        rows.append(
            {
                "baseline_files": _role_files(case_dir, root, topology, "baseline"),
                "candidate_files": _role_files(case_dir, root, topology, "candidate"),
                "case_path": case_dir.relative_to(root).as_posix(),
                "coverage": truth.get("coverage", []),
                "facts": truth.get("facts", []),
                "family": truth.get("family"),
                "id": case_id,
                "must_reach": truth.get("must_reach", []),
                "review_expectation": truth.get("review_expectation"),
                "schema_version": truth.get("schema_version"),
                "title": truth.get("title"),
                "topology": topology,
            }
        )

    if len(set(identifiers)) != len(identifiers):
        raise ManifestError("fixture tree contains duplicate case IDs")
    if expected_ids is not None and set(identifiers) != set(expected_ids):
        missing = sorted(set(expected_ids) - set(identifiers))
        unexpected = sorted(set(identifiers) - set(expected_ids))
        details = []
        if missing:
            details.append(f"missing: {', '.join(missing)}")
        if unexpected:
            details.append(f"unexpected: {', '.join(unexpected)}")
        raise ManifestError(f"fixture tree does not match expected case IDs ({'; '.join(details)})")
    return rows


def manifest_text(rows: Iterable[dict[str, Any]]) -> str:
    """Serialize rows as deterministic UTF-8 JSON Lines."""

    return "".join(
        json.dumps(row, ensure_ascii=False, separators=(",", ":"), sort_keys=True) + "\n"
        for row in rows
    )


def write_manifest(
    fixture_root: str | Path,
    output: str | Path | None = None,
    *,
    expected_ids: Iterable[str] | None = None,
) -> list[dict[str, Any]]:
    """Write a deterministic case catalogue and return its parsed rows."""

    root = Path(fixture_root).resolve()
    destination = Path(output) if output is not None else root / "manifest.jsonl"
    rows = case_rows(root, expected_ids=expected_ids)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(manifest_text(rows), encoding="utf-8")
    return rows
