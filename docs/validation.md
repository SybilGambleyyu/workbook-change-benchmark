# Validation record

This record describes the WCAB 0.22.0 / schema-version-3 validation run on
2026-08-03. It is reproducible from this repository; no network service or
private workbook is required.

## Fixture integrity

Environment:

- Python 3.13.9
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

- 37 cases: 36 paired-workbook cases and one directory portfolio case.
- 39 observable truth facts across 31 `block` and six `review` cases.
- Three scoreable coverage expectations: one newly introduced `INDIRECT`
  boundary and two unchanged-formula selector changes (`INDIRECT` address text
  and `OFFSET` column displacement).
- 114 generated fixture files: 76 workbooks, 37 truth manifests, and one JSONL
  case catalogue, all generated from source.
- One hundred thirty-six unit tests passed locally under Python 3.13, including
  independent regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 37 cases.
- The external-data pair has identical package members except for
  `xl/connections.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its relationship-backed source is a non-routable
  `example.invalid` URL.
- The PivotCache pair has identical package members except for
  `xl/pivotCache/pivotCacheDefinition1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its local `Source!A1:B5` worksheet
  binding, `Report!A1:B2` PivotTable location, stored display cells, and direct
  `Dashboard!B4` formula are unchanged while raw `refreshOnLoad` changes from
  false to true. The validation run did not open Excel, refresh a cache,
  calculate or render a PivotTable, or infer a report result.
- The PivotTable aggregation pair has identical package members except for
  `xl/pivotTables/pivotTable1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its local `Source!A1:B5` binding, cache records,
  `Report!A1:B2` location, stored report/dashboard cells, direct
  `Dashboard!B4` formula, refresh control, and calculation properties are
  unchanged while its one `dataFields/dataField/@subtotal` moves from `sum` to
  `average`. The validation run did not open Excel, refresh, calculate, or
  render a PivotTable, infer a changed display value, or claim client behavior.
- The PivotTable Slicer-selection pair has identical package members except
  for `xl/slicerCaches/slicerCache1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its local `Source!A1:B5` binding,
  PivotCache, `Report!A1:B2` location, stored report/dashboard cells, direct
  `Dashboard!B4` formula, refresh control, and calculation properties are
  unchanged while its one selected `Region` cache item moves from index 0
  (`North`) to index 1 (`South`). The validator follows the workbook-to-
  Slicer-cache-to-PivotCache/PivotTable bindings and does not create a visual
  Slicer, apply a filter, refresh, calculate, render, infer a displayed result,
  or claim client behavior.
- The connection-only Power Query pair has identical package members except
  for `customXml/item1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its `Source!A1:B5` `SourceData` Table, worksheet cells,
  table definition, calculation properties, metadata, and permission controls
  are unchanged while the stored `Table.SelectRows` M literal moves from
  `North` to `South`. The validator follows the package-root custom-XML
  relationship and bounded generated Data Mashup envelope; it does not execute
  M, apply the filter, refresh a query, materialize output, calculate a
  workbook, infer returned rows, or claim client behavior.
- The Scenario Manager pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its one selected, locked scenario retains its
  selection/protection, comment/user, summary-reference, second-input, and
  number-format metadata, plus visible `Inputs!B2=0.1`, `Inputs!B3=125`,
  `Inputs!D2=B2*B3`, and `Dashboard!B4=Inputs!$D$2`. Only the raw alternate
  `inputCells/@val` for `B2` moves from `0.08` to `0.16`. The validation run
  did not show or apply a scenario, calculate a workbook, create a Scenario
  Summary, infer an output, or claim client behavior.
- The What-If Data Table pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its raw `Sensitivity!D3` master keeps
  `ref="D3:D5"`, `ca="1"`, column orientation, the input grid, ordinary
  formulas, calculation properties, and saved table results empty while its one
  `r1` input reference moves from `B2` to `B3`. The validation run did not
  substitute inputs, calculate a workbook or Data Table, infer a table output,
  resolve a circular dependency, or claim client behavior.
- The list data-validation source pair has identical package members except
  for `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its one `Inputs!B2` list rule retains the target
  range, error/prompt/dropdown metadata, both `Lists` source columns, current
  `Draft` input, ordinary formulas, and calculation properties while raw
  `formula1` moves exactly from `=Lists!$A$2:$A$4` to `=Lists!$B$2:$B$4`.
  The validation run did not evaluate either source, decide whether a future
  input is valid, accept/reject an entry, calculate a workbook, or claim
  client behavior.
- The conditional-formatting threshold pair has identical package members
  except for `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its one `Operations!B2:B4` `cellIs`
  rule retains its target, priority, `greaterThan` operator, differential red
  fill, and stored `10`, `75`, and `120` metrics while raw `formula` moves from
  `100` to `50`. The validation run did not evaluate the rule, determine a
  rendered cell format, calculate a workbook, or claim client behavior.
- The chart-series pair has identical package members except for
  `xl/charts/chart1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its `Dashboard!D2` anchor, title/category references,
  source cells, and all worksheet parts remain unchanged while raw
  `c:ser/c:val/c:numRef/c:f` moves from `Source!$B$2:$B$4` to
  `Source!$C$2:$C$4`. The validation run did not open Excel, calculate source
  cells, refresh chart data, render a chart, or infer a visual difference.
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
- The active-AutoFilter pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its sole raw `Report!A1:B5` column-0 list value
  changes from `North` to `South`, while its `SUBTOTAL(109,B2:B5)` and
  dashboard formulas remain unchanged. The validation run did not apply a
  filter, calculate a subtotal, infer a visible row set, or open or save either
  workbook.
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

The 0.22.0 release retains the one-row-per-case `manifest.jsonl` catalogue and
the tool-neutral observation protocol at version 2. Each catalogue row retains
the schema-version-3 truth contract and includes byte counts and SHA-256
digests for the workbooks it names.

Commands:

```bash
python -m build --outdir /tmp/wcab-v022-dist
twine check /tmp/wcab-v022-dist/*
python -m venv /tmp/wcab-wheel-test
/tmp/wcab-wheel-test/bin/python -m pip install \
  /tmp/wcab-v022-dist/workbook_change_benchmark-0.22.0-py3-none-any.whl
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
- Final wheel and source-distribution SHA-256 values are recorded alongside
  their uploaded assets in the GitHub release, avoiding a self-referential
  source-distribution checksum in this record.
- Fresh Python 3.13 wheel and source-distribution installations both reported
  version 0.22.0 and validated all 37 fixtures; the wheel emitted
  byte-identical JSONL output.
- The full 136-test suite, lint, and format checks passed locally under Python
  3.13.
- The generated unsupported template scored as zero analyzed coverage, zero
  expected-fact recall, and zero coverage-disclosure recall, confirming that
  unsupported cases cannot become a pass.
- The FormulaFence normalizer emitted 38 matched facts, one intentionally
  unmapped fact, three matched coverage declarations, and no invented review
  disposition.

## FormulaFence reference adapter

FormulaFence 0.220.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 38 currently mappable diff/portfolio facts were observed.
- All three mappable coverage expectations were matched; no mapped fact,
  coverage expectation, or targeted lint rule was missed across all 37 cases.
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
- The WCAB 0.14 PivotCache case was observed as one exact
  `pivot_cache_refresh_controls_changed` record and `FF023`. FormulaFence's
  redacted profile retained one local worksheet cache, ID 1, and all reported
  controls except `refresh_on_load`, which changed from false to true; it
  emitted no parser warning. WCAB's independent raw validation established the
  source and PivotTable binding, stable stored cells, and cache-definition-only
  package change. Neither report refreshed or rendered a PivotTable.
- The WCAB 0.16 PivotTable aggregation case was observed as one exact
  `pivot_table_definitions_changed` record and `FF031`. FormulaFence's redacted
  profile retained one local cache, one PivotTable, one data field, two cache
  fields, four cache records, and no auxiliary material while only its
  PivotTable-layout material changed. It did not expose the source labels,
  selected aggregate function, or a rendered result. WCAB independently
  established the local graph, stable stored cells, exact `sum`-to-`average`
  declaration, and PivotTable-part-only package change. Neither report
  refreshed, calculated, or rendered a PivotTable.
- The WCAB 0.17 PivotTable Slicer-selection case was observed as one exact
  `slicer_timeline_cache_definitions_changed` record and `FF032`. FormulaFence's
  redacted profile retained one Slicer cache, one local PivotCache binding, one
  PivotTable binding, two Slicer items, one selected item, and no timeline or
  auxiliary material while only its Slicer filter-state/definition flag changed.
  It did not expose the Slicer name, selected item/value, or a rendered report.
  WCAB independently established the local graph, stable stored cells, exact
  index-0-`North` to index-1-`South` declaration, and Slicer-cache-part-only
  package change. Neither report created a visual Slicer, applied a filter,
  refreshed, calculated, or rendered a PivotTable.
- The WCAB 0.18 Power Query case was observed as one exact
  `power_query_changed` record and `FF024`. FormulaFence's redacted profile
  retained one parsed mashup, one formula document, three package parts, one
  metadata item, connection-only permission controls, and no opaque or
  embedded content while only `formula_material_changed` changed. It did not
  expose M text, local-table values, or a query result. WCAB independently
  established the local `SourceData` binding, fixed connection-only controls,
  exact `North`-to-`South` M literal, and custom-XML-part-only package change.
  Neither report executed M, refreshed a query, materialized output,
  calculated a workbook, or inferred returned rows.
- The WCAB 0.19 Scenario Manager case was observed as one exact
  `scenario_manager_changed` record and `FF035`. FormulaFence's redacted
  profile retained one scenario-bearing worksheet, one scenario, two stored
  inputs, one selected/locked scenario, one summary reference, one formatted
  input, and no malformed declaration while only
  `scenario_definition_material_changed` changed. It did not expose the
  scenario name, stored values, input references, comment, or user metadata.
  WCAB independently established the `B2` `0.08`-to-`0.16` raw transition,
  fixed metadata and visible cells/formulas, plus the Inputs-worksheet-only
  package change. Neither report showed/applied a scenario, calculated a
  result, or created a Scenario Summary.
- The WCAB 0.20 What-If Data Table case was observed as one exact
  `what_if_data_tables_changed` record and `FF034`. FormulaFence's redacted
  profile retained one Data Table, one column-oriented one-variable table,
  three declared output cells, one recalculation request, no deleted input
  reference, and no unrecognized declaration while only its
  `data_table_definition_material_changed` flag changed. It did not expose the
  output range, local input references, or calculated values. WCAB independently
  established the generated `D3:D5` master, exact `B2`-to-`B3` `r1` transition,
  stable input grid/formulas, ordinary static lower bounds, and
  Sensitivity-worksheet-only package change. Neither report substituted inputs,
  calculated, resolved a circular dependency, or claimed client behavior.
- The WCAB 0.21 list data-validation case was observed as one exact
  `data_validation_changed` record and `FF020`. FormulaFence exposed the one
  `Inputs!B2` list rule's source formula and complete entry-control metadata;
  it did not evaluate either local list or decide whether a future input is
  valid. WCAB independently established the exact `=Lists!$A$2:$A$4` to
  `=Lists!$B$2:$B$4` declaration, stable source values/current input/formulas,
  ordinary static lower bounds, and Inputs-worksheet-only package change.
  Neither report accepted/rejected an input, calculated a workbook, or claimed
  client behavior.
- The WCAB 0.22 conditional-formatting case was observed as one exact
  `conditional_formatting_changed` record and `FF021`. FormulaFence exposed
  the one `Operations!B2:B4` `cellIs` rule and raw `100`-to-`50` formula
  transition but did not render a workbook or determine which cells receive a
  format. WCAB independently established the stable priority, operator,
  differential red fill, metric values, calculation properties, and
  Operations-worksheet-only package change. Neither report evaluated the rule,
  calculated a workbook, or claimed client behavior.
- The WCAB 0.15 chart-series case was observed as one exact
  `chart_definitions_changed` record and `FF030`. FormulaFence's redacted
  profile retained one host sheet, drawing, legacy chart, and series; three
  data references; no chart cache; and no related payloads, while only its
  chart-definition material changed. FormulaFence did not expose the source
  formula or a visual result. WCAB independently established the local
  relationship chain, `D2` anchor, title/category references, exact value
  reference transition, and chart-part-only package change. Neither report
  calculated or rendered a chart.
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
- The WCAB 0.13 active-AutoFilter case was observed as one exact
  `filter_visibility_controls_changed` record and `FF036`: FormulaFence's
  redacted profile retained one worksheet filter, one filter column, one
  criterion, and no other visibility controls while its material-definition
  flag changed. WCAB's independent raw validation established the
  `North`-to-`South` criterion and stable formulas; neither report applies the
  filter, calculates a subtotal, or claims a visible result.
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
