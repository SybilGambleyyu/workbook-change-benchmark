# Validation record

This record describes the WCAB 0.4.0 / schema-version-3 validation run on
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
wcab observation-template --fixtures fixtures --output /tmp/observations.json
wcab score --fixtures fixtures --observations /tmp/observations.json
wcab formulafence-observations --fixtures fixtures \
  --executable formulafence --output /tmp/formulafence-observations.json
wcab score --fixtures fixtures --observations /tmp/formulafence-observations.json
ruff check .
ruff format --check .
pytest
wcab validate --fixtures fixtures
```

Results:

- 18 cases: 17 paired-workbook cases and one directory portfolio case.
- 20 observable truth facts across 13 `block` and five `review` cases.
- One scoreable coverage expectation requires an explicit static-dependency
  boundary disclosure for an introduced `INDIRECT` reference.
- 56 workbook and truth-manifest fixture files, plus one generated JSONL case
  catalogue, all generated from source.
- Twenty-three unit tests passed under both Python versions, including independent
  regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 18 cases.

## Distribution supplement

The 0.4.0 release retains the one-row-per-case `manifest.jsonl` catalogue and
upgrades the tool-neutral observation protocol to version 2. Each catalogue
row retains the schema-version-3 truth contract and includes byte counts and
SHA-256 digests for the workbooks it names.

Commands:

```bash
python -m build
twine check dist/workbook_change_benchmark-0.4.0*
python -m venv /tmp/wcab-wheel-test
/tmp/wcab-wheel-test/bin/python -m pip install \
  dist/workbook_change_benchmark-0.4.0-py3-none-any.whl
/tmp/wcab-wheel-test/bin/wcab validate --fixtures fixtures
/tmp/wcab-wheel-test/bin/wcab manifest --fixtures fixtures --output /tmp/manifest.jsonl
cmp fixtures/manifest.jsonl /tmp/manifest.jsonl
/tmp/wcab-wheel-test/bin/wcab observation-template --fixtures fixtures \
  --output /tmp/observations.json
/tmp/wcab-wheel-test/bin/wcab score --fixtures fixtures \
  --observations /tmp/observations.json
```

Results:

- The source distribution and universal wheel passed `twine check`.
- A fresh Python 3.12 wheel installation validated all 18 fixtures and emitted
  byte-identical JSONL output.
- The full twenty-three-test suite, lint, and format checks passed under the supported
  local Python versions (3.12 and 3.13).
- The generated unsupported template scored as zero analyzed coverage and zero
  expected-fact recall and zero coverage-disclosure recall, confirming that
  unsupported cases cannot become a pass.
- The FormulaFence normalizer emitted 19 matched facts, one intentionally
  unmapped fact, one matched coverage declaration, and no invented review
  disposition.

## FormulaFence reference adapter

FormulaFence 0.219.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 19 currently mappable diff/portfolio facts were observed.
- No mapped fact or coverage expectation was missed across all 18 cases.
- The schema-version-2 structured Table scope case was observed as a
  `table_definition_changed` diff, even though its summary formula text stays
  unchanged.
- The schema-version-3 dynamic-reference case was observed as both a
  `dynamic_formula_reference_added` diff and `FF012` coverage warning; the
  adapter preserves that warning as one normalized coverage declaration.
- Five targeted lint expectations were observed: copied-formula interruption
  (`FF082`), conditional-aggregate range shape (`FF093`), explicitly unlocked
  formula cell (`FF085`), incomplete manual calculation (`FF086`), and static
  cycle (`FF090`).
- One fact remains intentionally unmapped: the benign structural formula
  rewrite. It records known before/after locations but does not claim that a
  generic tool can prove Excel semantic equivalence from this small fixture.

This is a baseline implementation result, not a claim that FormulaFence or any
other tool has complete Excel coverage.
