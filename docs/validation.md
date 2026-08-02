# Validation record

This record describes the WCAB 0.2.0 / schema-version-2 validation run on
2026-08-02. It is reproducible from this repository; no network service or
private workbook is required.

## Fixture integrity

Environment:

- Python 3.12.3 and 3.13.9
- openpyxl 3.1.5
- pytest 9.1.1
- ruff 0.16.1

Commands:

```bash
python -m pip install -e '.[dev]'
wcab build --output fixtures
wcab manifest --fixtures fixtures
ruff check .
ruff format --check .
pytest
wcab validate --fixtures fixtures
```

Results:

- 17 cases: 16 paired-workbook cases and one directory portfolio case.
- 19 observable truth facts: 12 `block` and five `review` dispositions.
- 53 workbook and truth-manifest fixture files, plus one generated JSONL case
  catalogue, all generated from source.
- Eleven unit tests passed under both Python versions, including independent
  regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 17 cases.

## Distribution supplement

The 0.2.0 release builds a one-row-per-case `manifest.jsonl` catalogue. Each
row retains the schema-version-2 truth contract and includes byte counts and
SHA-256 digests for the workbooks it names.

Commands:

```bash
python -m build
twine check dist/workbook_change_benchmark-0.2.0*
python -m venv /tmp/wcab-wheel-test
/tmp/wcab-wheel-test/bin/python -m pip install \
  dist/workbook_change_benchmark-0.2.0-py3-none-any.whl
/tmp/wcab-wheel-test/bin/wcab validate --fixtures fixtures
/tmp/wcab-wheel-test/bin/wcab manifest --fixtures fixtures --output /tmp/manifest.jsonl
cmp fixtures/manifest.jsonl /tmp/manifest.jsonl
```

Results:

- The source distribution and universal wheel passed `twine check`.
- A fresh Python 3.12 wheel installation validated all 17 fixtures and emitted
  byte-identical JSONL output.
- The full eleven-test suite, lint, and format checks passed under the supported
  local Python versions (3.12 and 3.13).

## FormulaFence reference adapter

FormulaFence 0.219.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 18 currently mappable diff/portfolio facts were observed.
- No mapped fact was missed across all 17 cases.
- The schema-version-2 structured Table scope case was observed as a
  `table_definition_changed` diff, even though its summary formula text stays
  unchanged.
- Five targeted lint expectations were observed: copied-formula interruption
  (`FF082`), conditional-aggregate range shape (`FF093`), explicitly unlocked
  formula cell (`FF085`), incomplete manual calculation (`FF086`), and static
  cycle (`FF090`).
- One fact remains intentionally unmapped: the benign structural formula
  rewrite. It records known before/after locations but does not claim that a
  generic tool can prove Excel semantic equivalence from this small fixture.

This is a baseline implementation result, not a claim that FormulaFence or any
other tool has complete Excel coverage.
