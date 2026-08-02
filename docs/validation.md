# Validation record

This record describes the WCAB 0.12.0 / schema-version-3 validation run on
2026-08-03. It is reproducible from this repository; no network service or
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

- 27 cases: 26 paired-workbook cases and one directory portfolio case.
- 29 observable truth facts across 22 `block` and five `review` cases.
- Three scoreable coverage expectations: one newly introduced `INDIRECT`
  boundary and two unchanged-formula selector changes (`INDIRECT` address text
  and `OFFSET` column displacement).
- 84 generated fixture files: 56 workbooks, 27 truth manifests, and one JSONL
  case catalogue, all generated from source.
- Sixty-three unit tests passed locally under Python 3.13, including
  independent regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 27 cases.
- The external-data pair has identical package members except for
  `xl/connections.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its relationship-backed source is a non-routable
  `example.invalid` URL.
- The external-workbook-link policy pair has identical package members except
  for `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its unchanged `LinkedModel!B2` formula names the
  absent synthetic `WCABSource.xlsx`, while `workbookPr/@updateLinks` changes
  exactly from `never` to `always`.
- The iterative-calculation pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain readable
  by openpyxl. Its unchanged `Model!B2` direct self-reference records the same
  explicit 100 / 0.001 bounds in both workbooks while `calcPr/@iterate` changes
  exactly from false to true.
- The precision-as-displayed pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain readable
  by openpyxl. Its stored `Inputs!B2=10.005`, `0.00` format, formula, and local
  downstream reference are unchanged while `calcPr/@fullPrecision` changes
  exactly from true to false. The validation run did not open, calculate, or
  save either workbook.
- The workbook-date-system pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its raw `Inputs!B2=45292` numeric serial,
  `yyyy-mm-dd` style, local `Model!B2` and `Dashboard!B4` formulas, and
  explicit `dateCompatibility=true` control are unchanged while
  `workbookPr/@date1904` changes exactly from false to true. The validation run
  did not calculate, convert a serial, predict a displayed date, open, or save
  either workbook.
- The saved-formula-result pair has identical package members except for
  `xl/worksheets/sheet2.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its direct input, formula expression,
  calculation properties, and local downstream formula are unchanged while
  the raw `Model!B2` numeric `<v>` result changes exactly from `20` to `25`.
  The validation run did not calculate, validate, or interpret either result.
- The array-mode pair preserves the `Model!B1` `=LEN(Inputs!A1:A3)` anchor and
  its stored `B1:B3` range. The candidate adds `xl/metadata.xml`, one
  `sheetMetadata` relationship, one content-type override, and `cm=1` on the
  anchor; both archives pass ZIP integrity checks and remain readable by
  openpyxl. The raw validator classifies the pair exactly as `legacy_cse` then
  `dynamic` without calculating a value or predicting a spill extent.

## Distribution supplement

The 0.12.0 release retains the one-row-per-case `manifest.jsonl` catalogue and
the tool-neutral observation protocol at version 2. Each catalogue row retains
the schema-version-3 truth contract and includes byte counts and SHA-256
digests for the workbooks it names.

Commands:

```bash
python -m build --outdir /tmp/wcab-v012-dist
twine check /tmp/wcab-v012-dist/*
python -m venv /tmp/wcab-wheel-test
/tmp/wcab-wheel-test/bin/python -m pip install \
  /tmp/wcab-v012-dist/workbook_change_benchmark-0.12.0-py3-none-any.whl
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
- Fresh Python 3.13 wheel and source-distribution installations both reported
  version 0.12.0 and validated all 27 fixtures; the wheel emitted
  byte-identical JSONL output.
- The full sixty-three-test suite, lint, and format checks passed locally under
  Python 3.13; hosted release CI runs the same checks under Python 3.10 and
  3.13.
- The generated unsupported template scored as zero analyzed coverage, zero
  expected-fact recall, and zero coverage-disclosure recall, confirming that
  unsupported cases cannot become a pass.
- The FormulaFence normalizer emitted 28 matched facts, one intentionally
  unmapped fact, three matched coverage declarations, and no invented review
  disposition.

## FormulaFence reference adapter

FormulaFence 0.220.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 28 currently mappable diff/portfolio facts were observed.
- All three mappable coverage expectations were matched; no mapped fact,
  coverage expectation, or targeted lint rule was missed across all 27 cases.
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
- The WCAB 0.6 external-data case was observed as the exact
  `external_data_connections_changed` connection-ID-1 `refresh_on_load`
  false-to-true transition and `FF023`. The report omitted the synthetic
  connection name and endpoint; neither tool opened or refreshed it.
- The WCAB 0.8 external-workbook-link policy case was observed as one exact
  `external_data_refresh_settings_changed` transition in `FF023`: only
  `update_links` moved from `never` to `always`; `allow_refresh_query`,
  `refresh_all_connections`, and `save_external_link_values` retained their
  defaults. The report did not expose the synthetic external formula or open
  its absent source workbook.
- The WCAB 0.9 iterative-calculation case was observed as one exact
  `calculation_settings_changed` transition in `FF009`: only `iterate` moved
  from false to true; `iterateCount=100`, `iterateDelta=0.001`, and the other
  stored controls stayed unchanged. FormulaFence's direct-self-reference lint
  surfaced `Model!B2` only while iteration was disabled; neither report claims
  a calculated or converged value.
- The WCAB 0.10 precision-as-displayed case was observed as one exact
  `calculation_settings_changed` transition in `FF009`: only `fullPrecision`
  moved from true to false while the stored input, number format, formula, and
  other calculation controls remained unchanged. Neither report claims a
  rounded stored value, calculated result, or client-side save behavior.
- The WCAB 0.12 workbook-date-system case was observed as one exact
  `workbook_date_system_changed` transition and `FF117`: normalized
  `date1904` moved from false to true while explicit `dateCompatibility=true`
  and the unrecognized-control count of zero remained fixed. FormulaFence
  reported workbook-level evidence only; WCAB's independent raw validation
  established the stable serial, style, formulas, and package boundary. Neither
  report claims a converted or displayed date.
- The WCAB 0.11 saved-formula-result case was observed as one exact
  `formula_cached_result_changed` record and `FF042`: FormulaFence reports
  two formula cells, one numeric saved result, one missing saved result, and
  exactly one unexplained material cache change. It intentionally omits the
  raw result values and formula-cell locations, and neither report claims a
  calculation, stale-cache diagnosis, tampering diagnosis, or displayed value.
- The WCAB 0.7 array-mode case was observed as the exact `Model!B1`
  `legacy_cse`-to-`dynamic` transition with the same stored `B1:B3` range and
  `FF018`. It reported the declared `Dashboard!B2` static impact, emitted no
  parser warning, and did not expose the fixture formula text in the report.
- Five targeted lint expectations were observed: copied-formula interruption
  (`FF082`), conditional-aggregate range shape (`FF093`), explicitly unlocked
  formula cell (`FF085`), incomplete manual calculation (`FF086`), and static
  cycle (`FF090`).
- One fact remains intentionally unmapped: the benign structural formula
  rewrite. It records known before/after locations but does not claim that a
  generic tool can prove Excel semantic equivalence from this small fixture.

This is a baseline implementation result, not a claim that FormulaFence or any
other tool has complete Excel coverage.
