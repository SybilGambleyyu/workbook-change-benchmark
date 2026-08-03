# Validation record

This record describes the WCAB 0.31.0 / schema-version-3 validation run on
2026-08-02. It is reproducible from this repository; no network service or
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

- 46 cases: 45 paired-workbook cases and one directory portfolio case.
- 48 observable truth facts across 36 `block` and 10 `review` cases.
- Three scoreable coverage expectations: one newly introduced `INDIRECT`
  boundary and two unchanged-formula selector changes (`INDIRECT` address text
  and `OFFSET` column displacement).
- 141 generated fixture files: 94 workbooks, 46 truth manifests, and one JSONL
  case catalogue, all generated from source.
- One hundred ninety-nine unit tests passed locally under Python 3.13, including
  independent regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 46 cases.
- The external-data pair has identical package members except for
  `xl/connections.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its relationship-backed source is a non-routable
  `example.invalid` URL.
- The QueryTable pair has identical package members except for
  `xl/queryTables/queryTable1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its direct `ImportedData` worksheet-to-QueryTable
  and workbook-to-connections relationships, non-routable `example.invalid`
  web-query URL, connection-level `refreshOnLoad=false`, remaining QueryTable
  controls, saved `ImportedData!B2=100` cell, and
  `ImportedData!B2 → Summary!B2 → Dashboard!B4` formula context are unchanged
  while raw `queryTable/@refreshOnLoad` changes from false to true. The
  validation run did not open a connection, fetch a URL, refresh a query,
  materialize rows, calculate a workbook, or claim a successful client refresh.
- The cell-hyperlink target pair has identical package members except for
  `xl/worksheets/_rels/sheet1.xml.rels`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its one `Inputs!B2` worksheet
  hyperlink declaration, visible `Open vendor portal` text, fixed relationship
  ID/type/`TargetMode=External`, calculation properties, and
  `Inputs!B2 → Summary!B2 → Dashboard!B4` formula context remain unchanged
  while raw `Relationship/@Target` moves from the reserved
  `approved.example.invalid` URL to `review.example.invalid`. The validation
  run did not resolve, open, fetch, visit, execute, calculate, or otherwise
  interact with either target, or claim that a client follows one.
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
- The custom number-format pair has identical package members except for
  `xl/styles.xml`; both archives pass ZIP integrity checks and remain readable
  by openpyxl. Its one `Operations!B2` style index, raw `0.125` value, and
  neighboring `B3=B2` formula remain unchanged while the referenced custom
  `numFmt/@formatCode` moves from `0.0%;[Red](0.0%);-` to `;;;`. The validator
  compares styles XML after erasing only that code. It does not render a
  number format, resolve locale or column width, determine a displayed value,
  calculate a workbook, or claim client behavior.
- The ignored-error suppression pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its `Operations!B2=10`, `B3=20`, `B4=30`,
  `B5=SUM(B2:B3)`, and `C5=B5` records remain unchanged while one standard
  `ignoredErrors/ignoredError` declaration with `sqref="B5"` and
  `formulaRange="1"` is added. The validator compares worksheet XML after
  removing only that declaration. It does not determine whether Excel would
  show a warning, evaluate the formula, render an indicator, decide whether the
  warning is justified, calculate a workbook, or claim client behavior.
- The workbook-structure-protection pair has identical package members except
  for `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its hidden `ReviewControls` sheet and
  `Inputs!D2=B2*C2` formula remain unchanged while raw
  `workbookProtection/@lockStructure` moves from `1` to `0`. The validator
  compares the workbook XML after removing only that attribute. It did not test
  a password, encryption, authentication, authorization, exposure of a hidden
  sheet, or a particular Excel client's sheet-operation behavior.
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
- The saved Named Sheet View pair has identical package members except for
  `xl/namedSheetViews/namedSheetView1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its `Report!A1:B5` worksheet
  AutoFilter retains no active criterion and its `xr:uid` binding, rows,
  `Report!D2=SUBTOTAL(109,B2:B5)`, and `Dashboard!B4=Report!$D$2` remain
  unchanged while the relationship-backed saved view moves its sole column-0
  list value from `North` to `South`. The validator follows the worksheet
  relationship, confirms the content type and base-filter binding, and does
  not activate, render, or apply the view, calculate a subtotal, infer visible
  rows, or claim a client display or print outcome.
- The XML Map pair has identical package members except for
  `xl/tables/table1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its synthetic local MapInfo/XSD declaration,
  file-binding metadata, `Export!E2` single-cell mapping, table values,
  `Export!D2=SUM(InvoiceLines[Net amount])`, and
  `Dashboard!B4=Export!$D$2` remain unchanged while the mapped table
  column XPath moves from `NetAmount` to `TaxAmount`. The validator follows
  only local workbook and worksheet relationships and does not access a file,
  validate a schema, import/export XML, materialize data, calculate a result,
  or claim client behavior.
- The Office Web Add-in pair has identical package members except for
  `xl/webextensions/webextension1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its workbook-to-taskpane-to-web-
  extension relationship chain, synthetic local FileSystem reference, add-in
  IDs, hidden locked pane, ordinary cells, calculation properties, and
  `Inputs!B2 → Model!B2 → Dashboard!B4` formulas remain unchanged while raw
  `Office.AutoShowTaskpaneWithDocument` moves exactly from false to true. The
  fixture has no manifest payload or external relationship. The validation run
  did not install, load, execute, or fetch an add-in or manifest, or claim
  that a task pane opens or that the add-in accesses a workbook cell.
- The embedded-OLE auto-load pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. It retains one standard direct internal OLE-object
  relationship, the matching content type, fixed opaque ASCII payload bytes, a
  synthetic unregistered ProgID, and unchanged `Inputs!B2 → Model!B2 →
  Dashboard!B4` formula context while raw `oleObject/@autoLoad` moves exactly
  from false to true. It has no linked target, ActiveX control, presentation,
  macro, or external relationship. The validation run did not deserialize,
  open, render, execute, register, or invoke an object server, or claim that
  any object loads successfully.
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

The 0.31.0 release retains the one-row-per-case `manifest.jsonl` catalogue and
the tool-neutral observation protocol at version 2. Each catalogue row retains
the schema-version-3 truth contract and includes byte counts and SHA-256
digests for the workbooks it names.

Commands:

```bash
python -m build --outdir /tmp/wcab-v031-dist
twine check /tmp/wcab-v031-dist/*
python -m venv /tmp/wcab-wheel-test
/tmp/wcab-wheel-test/bin/python -m pip install \
  /tmp/wcab-v031-dist/workbook_change_benchmark-0.31.0-py3-none-any.whl
/tmp/wcab-wheel-test/bin/python -c 'import wcab; print(wcab.__version__)'
/tmp/wcab-wheel-test/bin/wcab validate --fixtures fixtures
/tmp/wcab-wheel-test/bin/wcab manifest --fixtures fixtures --output /tmp/manifest.jsonl
cmp fixtures/manifest.jsonl /tmp/manifest.jsonl
/tmp/wcab-wheel-test/bin/wcab observation-template --fixtures fixtures \
  --output /tmp/observations.json
/tmp/wcab-wheel-test/bin/wcab score --fixtures fixtures \
  --observations /tmp/observations.json
python -m venv /tmp/wcab-sdist-test
/tmp/wcab-sdist-test/bin/python -m pip install \
  /tmp/wcab-v031-dist/workbook_change_benchmark-0.31.0.tar.gz
/tmp/wcab-sdist-test/bin/python -c 'import wcab; print(wcab.__version__)'
/tmp/wcab-sdist-test/bin/wcab validate --fixtures fixtures
```

Results:

- The source distribution and universal wheel passed `twine check`.
- Final wheel and source-distribution SHA-256 values are recorded alongside
  their uploaded assets in the GitHub release, avoiding a self-referential
  source-distribution checksum in this record.
- Fresh Python 3.13 wheel and source-distribution installations both reported
  version 0.31.0 and validated all 46 fixtures; both emitted byte-identical
  JSONL output.
- The full 199-test suite, lint, and format checks passed locally under Python
  3.13.
- The generated unsupported template scored as zero analyzed coverage, zero
  expected-fact recall, and zero coverage-disclosure recall, confirming that
  unsupported cases cannot become a pass.
- The FormulaFence normalizer emitted 47 matched facts, one intentionally
  unmapped fact, three matched coverage declarations, and no invented review
  disposition.

## FormulaFence reference adapter

FormulaFence 0.220.0 was installed from the local checked-out release source
and invoked only against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 47 currently mappable diff/portfolio facts were observed.
- All three mappable coverage expectations were matched; no mapped fact,
  coverage expectation, or targeted lint rule was missed across all 46 cases.
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
- The WCAB 0.30 QueryTable case was observed as one exact
  `query_table_refresh_controls_changed` record and `FF023`: FormulaFence's
  redacted profile retained one `ImportedData` table, connection ID 1, fixed
  background/disable/remove/fill/edit/growth controls, name metadata, and no
  opaque metadata while only `refresh_on_load` moved from false to true. It did
  not expose the endpoint, OOXML part, or result rows. WCAB independently
  established the direct local relationship graph, fixed connection-level
  control, stable saved cells/formulas, and QueryTable-part-only boundary;
  neither report opened a connection, fetched a URL, refreshed a query,
  materialized rows, calculated a workbook, or claimed a client refresh.
- The WCAB 0.31 cell-hyperlink target case was observed as one exact
  `cell_hyperlink_controls_changed` record and `FF047`: FormulaFence's
  redacted profile retained one external worksheet hyperlink binding with no
  location, display, tooltip, or unrecognized declaration while its binding,
  definition material, and relationship material changed. It did not expose
  the target or relationship ID. WCAB independently established the exact
  reserved-target transition, fixed relationship ID/type/mode, stable visible
  text/formulas, and relationship-part-only boundary; neither report resolved,
  opened, fetched, visited, or otherwise interacted with a target, or claimed
  that a client follows it.
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
- The WCAB 0.23 custom number-format case was observed as one exact
  `number_format_controls_changed` record and `FF039`. FormulaFence's
  intentionally redacted before/after profiles each retained one direct-cell
  custom assignment and no default, row, column, built-in, or unrecognized
  control while only `number_format_definition_material_changed` was set. It
  did not expose a format code or cell target. WCAB independently established
  the declared code transition, stable target/style/value/formula context, and
  styles-only package change. Neither report rendered a format, calculated a
  workbook, or claimed client behavior.
- The WCAB 0.24 ignored-error case was observed as one exact
  `ignored_error_controls_changed` record and `FF037`. FormulaFence's redacted
  profile moved from no controls to one standard container, one target range,
  and one formula-range-omission category with no unrecognized controls, while
  only `ignored_error_definition_material_changed` was set. It did not expose
  the target range or formula. WCAB independently established the generated
  `Operations!B5` target, `formulaRange=1` declaration, stable cells/formulas,
  and worksheet-only package change. Neither report determined whether Excel
  would show a warning, evaluated a formula, rendered an indicator, or claimed
  client behavior.
- The WCAB 0.25 workbook-structure case was observed as one exact
  `workbook_protection_changed` record and `FF022`. FormulaFence's non-secret
  profile moved from `lock_structure=true` to all workbook locks false, with no
  credential or opaque metadata on either side. WCAB independently established
  the raw `lockStructure=1`-to-`0` transition, stable hidden-sheet/formula
  context, and workbook-XML-only package change. Neither report tested a
  password, encryption, authentication, authorization, or client action.
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
- The WCAB 0.26 saved Named Sheet View case was observed as one exact
  `named_sheet_views_changed` record and `FF038`: FormulaFence's redacted
  profile retained one worksheet, part, named view, filter, column, and
  criterion; zero sort rules or conditions; and zero unrecognized declarations
  while only its Named Sheet View material-definition flag changed. It did not
  expose the view name, IDs, bound range, or list value. WCAB independently
  established the relationship, base-AutoFilter binding, `North`-to-`South`
  criterion, stable formulas, and Named-Sheet-View-part-only boundary; neither
  report activated, rendered, or applied a view, calculated a subtotal, or
  inferred visible rows.
- The WCAB 0.27 XML Map case was observed as one exact
  `xml_mapping_controls_changed` record and `FF049`: FormulaFence's
  redacted profile retained one map part, schema, map, data binding, file
  binding, table binding, and sheet-level single-cell binding with no
  unrecognized mapping metadata while only its binding-material flag changed.
  It did not expose the schema, map, XPath, table, or cell values. WCAB
  independently established the synthetic `NetAmount`-to-`TaxAmount`
  binding transition, stable declarations/formulas, and table-part-only
  boundary; neither report accessed a file, imported/exported XML,
  materialized data, or inferred a result.
- The WCAB 0.28 Office Web Add-in case was observed as one exact
  `office_web_addins_changed` record and `FF028`: FormulaFence's redacted
  profile retained one declared task-pane part, task pane, web-extension part,
  and local store reference; one locked hidden task pane; no bindings,
  snapshots, external relationships, in-content references, or unrecognized
  parts; and an auto-show count from zero to one. It did not expose add-in IDs,
  store name, or the property value. WCAB independently established the local
  relationship graph, false-to-true auto-show property, stable formula context,
  and web-extension-part-only boundary; neither report installed, loaded,
  executed, or fetched an add-in or manifest, or claimed that a task pane
  opens.
- The WCAB 0.29 embedded-OLE auto-load case was observed as one exact
  `worksheet_embedded_controls_changed` record and `FF029`: FormulaFence's
  redacted profile retained one worksheet, one embedded OLE object, one
  internal payload, no linked/external/ActiveX/VML/presentation material, and
  no unrecognized declaration while its auto-load count moved from zero to one.
  It did not inspect or deserialize the opaque payload. WCAB independently
  established the false-to-true attribute transition, fixed relationship and
  content-type boundary, inert synthetic payload, stable formula context, and
  worksheet-XML-only package change; neither report opened, rendered,
  executed, registered, or invoked an object server, or claimed that an object
  loads successfully.
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
