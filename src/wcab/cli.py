"""Command line interface for WCAB."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .adapters.formulafence import FormulaFenceAdapterError, evaluate_reference_suite
from .build import CASE_IDS, build_all
from .manifest import write_manifest
from .validate import validate_all


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Workbook Change Assurance Benchmark")
    subcommands = parser.add_subparsers(dest="command", required=True)
    build = subcommands.add_parser("build", help="generate deterministic paired-workbook fixtures")
    build.add_argument(
        "--output", type=Path, default=Path("fixtures"), help="fixture output directory"
    )
    validate = subcommands.add_parser(
        "validate", help="validate fixture truth against generated workbooks"
    )
    validate.add_argument(
        "--fixtures", type=Path, default=Path("fixtures"), help="fixture directory"
    )
    formulafence = subcommands.add_parser(
        "formulafence", help="run the optional local FormulaFence reference adapter"
    )
    formulafence.add_argument(
        "--fixtures", type=Path, default=Path("fixtures"), help="fixture directory"
    )
    formulafence.add_argument(
        "--executable", default="formulafence", help="FormulaFence executable to invoke"
    )
    formulafence.add_argument(
        "--strict",
        action="store_true",
        help="return nonzero if a mapped fact or lint rule is missed",
    )
    manifest = subcommands.add_parser(
        "manifest", help="write a deterministic, integrity-addressed case catalogue"
    )
    manifest.add_argument(
        "--fixtures", type=Path, default=Path("fixtures"), help="fixture directory"
    )
    manifest.add_argument(
        "--output",
        type=Path,
        default=None,
        help="JSONL destination (default: fixtures/manifest.jsonl)",
    )
    subcommands.add_parser("list", help="list stable case identifiers")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.command == "build":
        manifests = build_all(args.output)
        print(f"built {len(manifests)} WCAB cases in {args.output}")
        return 0
    if args.command == "validate":
        failures = validate_all(args.fixtures)
        if not failures:
            print(f"validated {len(CASE_IDS)} WCAB cases")
            return 0
        for case_id, errors in failures.items():
            for error in errors:
                print(f"{case_id}: {error}", file=sys.stderr)
        return 1
    if args.command == "formulafence":
        try:
            result = evaluate_reference_suite(args.fixtures, executable=args.executable)
        except FormulaFenceAdapterError as error:
            print(f"FormulaFence adapter error: {error}", file=sys.stderr)
            return 2
        print(json.dumps(result, indent=2, sort_keys=True))
        has_miss = bool(result["diff_cases_with_misses"] or result["lint_rule_misses"])
        return 1 if args.strict and has_miss else 0
    if args.command == "manifest":
        rows = write_manifest(args.fixtures, args.output, expected_ids=CASE_IDS)
        destination = args.output or args.fixtures / "manifest.jsonl"
        print(f"wrote {len(rows)} WCAB cases to {destination}")
        return 0
    for case_id in CASE_IDS:
        print(case_id)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
