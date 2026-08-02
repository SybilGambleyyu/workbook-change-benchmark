# Validation record

This record describes the initial WCAB 0.1 validation run performed on
2026-08-01 and independently repeated on 2026-08-02. It is reproducible from
this repository; no network service or private workbook is required.

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
ruff check .
ruff format --check .
pytest
wcab validate --fixtures fixtures
```

Results:

- 16 cases: 15 paired-workbook cases and one directory portfolio case.
- 18 observable truth facts: 11 `block` and five `review` dispositions.
- 50 committed fixture files, all generated from source.
- Seven unit tests passed under both Python versions, including independent
  regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 16 cases.

## FormulaFence reference adapter

FormulaFence 0.219.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 17 currently mappable diff/portfolio facts were observed.
- No mapped fact was missed across all 16 cases.
- Five targeted lint expectations were observed: copied-formula interruption
  (`FF082`), conditional-aggregate range shape (`FF093`), explicitly unlocked
  formula cell (`FF085`), incomplete manual calculation (`FF086`), and static
  cycle (`FF090`).
- One fact remains intentionally unmapped: the benign structural formula
  rewrite. It records known before/after locations but does not claim that a
  generic tool can prove Excel semantic equivalence from this small fixture.

This is a baseline implementation result, not a claim that FormulaFence or any
other tool has complete Excel coverage.
