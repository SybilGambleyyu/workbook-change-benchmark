# Validation record

This record describes the WCAB 0.5.0 / schema-version-3 validation run on
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

- 20 cases: 19 paired-workbook cases and one directory portfolio case.
- 22 observable truth facts across 15 `block` and five `review` cases.
- Three scoreable coverage expectations: one newly introduced `INDIRECT`
  boundary and two unchanged-formula selector changes (`INDIRECT` address text
  and `OFFSET` column displacement).
- 63 workbook and truth-manifest fixture files, plus one generated JSONL case
  catalogue, all generated from source.
- Twenty-five unit tests passed under both Python versions, including
  independent regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 20 cases.

## Distribution supplement

The 0.5.0 release retains the one-row-per-case `manifest.jsonl` catalogue and
the tool-neutral observation protocol at version 2. Each catalogue row retains
the schema-version-3 truth contract and includes byte counts and SHA-256
digests for the workbooks it names.

Commands:

```bash
python -m build --outdir /tmp/wcab-v05-dist
twine check /tmp/wcab-v05-dist/*
python -m venv /tmp/wcab-wheel-test
/tmp/wcab-wheel-test/bin/python -m pip install \
  /tmp/wcab-v05-dist/workbook_change_benchmark-0.5.0-py3-none-any.whl
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
- A fresh Python 3.12 wheel installation validated all 20 fixtures and emitted
  byte-identical JSONL output.
- The full twenty-five-test suite, lint, and format checks passed under the
  supported local Python versions (3.12 and 3.13).
- The generated unsupported template scored as zero analyzed coverage, zero
  expected-fact recall, and zero coverage-disclosure recall, confirming that
  unsupported cases cannot become a pass.
- The FormulaFence normalizer emitted 21 matched facts, one intentionally
  unmapped fact, three matched coverage declarations, and no invented review
  disposition.

## FormulaFence reference adapter

FormulaFence 0.219.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 21 currently mappable diff/portfolio facts were observed.
- All three mappable coverage expectations were matched; no mapped fact,
  coverage expectation, or targeted lint rule was missed across all 20 cases.
- The schema-version-2 structured Table scope case was observed as a
  `table_definition_changed` diff, even though its summary formula text stays
  unchanged.
- The schema-version-3 introduced-dynamic case was observed as both a
  `dynamic_formula_reference_added` diff and `FF012` coverage warning.
- The two WCAB 0.5 dynamic-driver cases were observed as an `Inputs!E12`
  `value_changed` diff plus a candidate `dynamic_reference_cells` profile
  feature at `Summary!B2`. The adapter preserves that paired evidence as the
  `dynamic_reference_driver_changed` declaration; it does not claim to
  evaluate either selected target.
- Five targeted lint expectations were observed: copied-formula interruption
  (`FF082`), conditional-aggregate range shape (`FF093`), explicitly unlocked
  formula cell (`FF085`), incomplete manual calculation (`FF086`), and static
  cycle (`FF090`).
- One fact remains intentionally unmapped: the benign structural formula
  rewrite. It records known before/after locations but does not claim that a
  generic tool can prove Excel semantic equivalence from this small fixture.

This is a baseline implementation result, not a claim that FormulaFence or any
other tool has complete Excel coverage.
