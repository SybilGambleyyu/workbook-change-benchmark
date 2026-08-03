# Changelog

## 0.35.0 — 2026-08-03

- Add a deterministic Power Pivot/Data Model relationship-definition case. Its
  one `x15:modelRelationship` moves from
  `SalesModel.CalendarKey → CalendarModel.DateKey` to
  `SalesModel.CalendarKey → CalendarModel.FiscalDateKey` while its workbook
  `powerPivotData` binding, content type, local Tables, calculation properties,
  and opaque `xl/model/item.data` payload remain fixed; only
  `xl/workbook.xml` changes.
- Add an exact raw-OOXML validator, package-isolation and corruption regressions,
  and a narrow FormulaFence mapping requiring the exact redacted
  `power_pivot_data_model_changed` profile and high-severity `FF033`.
  WCAB records a relationship declaration only: it does not deserialize the
  Analysis Services payload, evaluate DAX, refresh a model, calculate or render
  a report, infer model-to-cell impact, or claim client behavior.

## 0.34.0 — 2026-08-03

- Add a deterministic workbook-scoped named-LAMBDA definition case. Its sole
  `ScenarioValue` body moves from `=LAMBDA(rate,amount,rate*amount)` to
  `=LAMBDA(rate,amount,rate*(amount+10))` while its input cells, calling
  `Model!B2` formula, direct `Dashboard!B4` consumer, calculation properties,
  workbook relationships, sheet declarations, and every package member except
  `xl/workbook.xml` remain fixed. The compact package has no external-link or
  external-reference declaration.
- Add a deterministic Excel Table calculated-column-master case. Its sole
  third-column raw `calculatedColumnFormula` moves from `A2*B2` to
  `A2*(B2+1)` while the `ScenarioLedger` Table binding/range/headers, ordinary
  `Ledger!C2:C4` formulas, structured-reference dashboard formula, calculation
  properties, and every package member except `xl/tables/table1.xml` remain
  fixed.
- Add exact raw-OOXML validators, package-isolation and corruption regressions,
  plus narrow FormulaFence mappings. WCAB records stored formula definitions
  only: it neither evaluates a LAMBDA nor fills a Table column, calculates a
  structured reference, infers an output, or claims client behavior.

## 0.33.0 — 2026-08-03

- Add a deterministic external defined-name source-change case. Its local
  `ScenarioRate` definition moves between two synthetic qualified external
  workbook expressions while `Model!B2=ScenarioRate*2`, its direct
  `Dashboard!B4` consumer, calculation properties, workbook relationships,
  sheet declarations, and every package member except `xl/workbook.xml` remain
  fixed. The deliberately compact package has no `externalLink` part or
  `externalReferences` declaration.
- Add a deterministic protected-sheet sort-permission case. Its raw
  `sheetProtection/@sort` control moves from `1` (locked) to `0` (permitted)
  while sheet protection remains enabled, all other action locks, cells,
  formulas, styles, calculation properties, and every package member except
  `xl/worksheets/sheet1.xml` remain fixed.
- Add exact `external_defined_name_source_changed` and
  `sheet_protection_sort_permission_enabled` validators, package-isolation
  checks, corruption regressions, and narrow FormulaFence mappings. WCAB reads
  stored local OOXML only: it neither interacts with external sources nor tests
  passwords, authorization, a client sort operation, or a resulting value.

## 0.32.0 — 2026-08-03

- Add a deterministic, relationship-backed external-workbook source-change
  case. Its `LinkedModel!B2` external formula, direct `Dashboard!B4`
  consumer, calculation properties, workbook external-reference binding, and
  `externalLink`/`externalBook` declaration remain fixed while one external
  `externalLinkPath` relationship moves from the reserved
  `approved.example.invalid` source to `review.example.invalid`.
- Add the exact `external_workbook_link_source_changed` fact, a narrow
  raw-OOXML graph validator, externalLink-relationship-only package-isolation
  check, and corruption regressions. WCAB records the stored source target
  only: it does not resolve, open, fetch, authenticate to, trust, refresh,
  calculate, or otherwise interact with either source, and does not claim that
  a client updates a link or returns a value.
- Extend the optional FormulaFence adapter to require its exact redacted
  `external_link_packages_changed` profile and high-severity `FF025`, while
  WCAB independently establishes the relationship IDs, types, target mode,
  target transition, stable formula context, and package boundary.

## 0.31.0 — 2026-08-03

- Add a deterministic, relationship-backed worksheet cell-hyperlink target
  case. Its visible `Inputs!B2` text, local `Inputs!B2 → Summary!B2 →
  Dashboard!B4` formula context, calculation properties, and worksheet XML
  remain fixed while the one external hyperlink relationship moves from the
  reserved `approved.example.invalid` target to `review.example.invalid`.
- Add the exact `cell_hyperlink_target_changed` fact, a narrow raw-OOXML
  relationship validator, relationship-target-only package-isolation check,
  and corruption regressions. WCAB records the stored target only: it does not
  resolve, open, fetch, visit, execute, calculate, or otherwise interact with
  it, and does not claim that a client follows it.
- Extend the optional FormulaFence adapter to require its exact redacted
  `cell_hyperlink_controls_changed` profile and high-severity `FF047`, while
  WCAB independently establishes the raw relationship ID, type, target mode,
  target transition, stable visible cells/formulas, and package boundary.

## 0.30.0 — 2026-08-03

- Add a deterministic, relationship-backed QueryTable refresh-on-open case.
  Its direct worksheet-to-QueryTable and workbook-to-connections relationships,
  fixed non-routable synthetic web-query connection, saved cells, and
  `ImportedData!B2 → Summary!B2 → Dashboard!B4` formula context remain fixed
  while raw `queryTable/@refreshOnLoad` changes from false to true. The
  connection-level `refreshOnLoad` control remains false on both sides.
- Add the exact `query_table_refresh_on_load_changed` fact, a narrow raw-OOXML
  graph validator, QueryTable-part-only package-isolation check, and corruption
  regressions. WCAB records a stored request only: it does not open a
  connection, fetch a URL, refresh a query, materialize rows, calculate a
  workbook, or claim that a client refreshes successfully.
- Extend the optional FormulaFence adapter to require its exact redacted
  `query_table_refresh_controls_changed` profile and critical `FF023`, while
  WCAB independently establishes the local relationship graph, stable
  connection and workbook context, raw transition, and package boundary.

## 0.29.0 — 2026-08-03

- Add a deterministic, relationship-backed worksheet embedded-OLE auto-load
  case. Its one internal relationship, opaque synthetic bytes, synthetic
  unregistered ProgID, ordinary `Inputs!B2 → Model!B2 → Dashboard!B4` formula
  context, and all non-worksheet package members remain fixed while raw
  `oleObject/@autoLoad` changes from false to true.
- Add the exact `ole_object_auto_load_enabled` fact, a narrow raw-OOXML
  relationship-graph validator, worksheet-only package-isolation check, and
  corruption regressions. WCAB records a stored request only: it does not
  deserialize, open, render, execute, register, or invoke an object server,
  and it does not claim that an object loads successfully.
- Extend the optional FormulaFence adapter to require its exact redacted
  `worksheet_embedded_controls_changed` profile and critical `FF029`, while
  WCAB independently establishes the generated local relationship, fixed
  opaque bytes, auto-load transition, stable workbook context, and package
  boundary.

## 0.28.0 — 2026-08-03

- Add a deterministic, relationship-backed Office Web Add-in task-pane case.
  Its synthetic local add-in reference, locked hidden task-pane declaration,
  and ordinary `Inputs!B2 → Model!B2 → Dashboard!B4` formula context remain
  fixed while `Office.AutoShowTaskpaneWithDocument` changes from false to
  true.
- Add the exact `office_web_addin_auto_show_enabled` fact, a narrow
  raw-OOXML relationship-graph validator, web-extension-part-only package
  isolation check, and corruption regressions. WCAB records a stored request
  only: it does not install, load, execute, or fetch an add-in or manifest,
  and it does not claim that a task pane opens.
- Extend the optional FormulaFence adapter to require its exact redacted
  `office_web_addins_changed` profile and high-severity `FF028`, while
  WCAB independently establishes the generated local reference, auto-show
  transition, stable workbook context, and package boundary.

## 0.27.0 — 2026-08-03

- Add a deterministic XML Map table-binding case. Its synthetic invoice
  table's mapped Net amount column changes from the NetAmount XPath to the
  TaxAmount XPath, while the local map/schema/file-binding declarations,
  single-cell mapping, visible cells, table total, and dashboard formula
  remain unchanged.
- Add the exact xml_map_table_column_xpath_retargeted fact, a narrow
  raw-OOXML relationship-graph validator, table-part-only package isolation
  check, and corruption regressions. WCAB records a stored import/export
  mapping contract; it does not open a file, validate a schema, import or
  export XML, materialize data, calculate a result, or claim client behavior.
- Extend the optional FormulaFence adapter to require its exact redacted
  xml_mapping_controls_changed profile and high-severity FF049, while WCAB
  independently establishes the map, schema, table-column XPath, single-cell
  binding, stable formulas, and package boundary.

## 0.26.0 — 2026-08-03

- Add a deterministic, relationship-backed Excel Named Sheet View case. Its
  saved list criterion moves from `North` to `South` while the base worksheet
  AutoFilter, rows, `Report!D2=SUBTOTAL(109,B2:B5)`, and
  `Dashboard!B4=Report!$D$2` remain unchanged.
- Add the exact `named_sheet_view_filter_criterion_changed` fact, a narrow
  raw-OOXML validator, Named Sheet View-part-only package isolation check,
  and corruption regressions. WCAB records a stored alternate review lens;
  it does not activate or render a view, apply a filter, calculate a
  subtotal, infer visible rows, or claim client display or print behavior.
- Extend the optional FormulaFence adapter to require the exact redacted
  `named_sheet_views_changed` profile and high-severity `FF038`, while WCAB
  independently proves the bound base AutoFilter, criterion values, stable
  formulas, and package boundary.

## 0.25.0 — 2026-08-03

- Add a deterministic workbook-structure-protection case. It retains a hidden
  `ReviewControls` sheet and `Inputs!D2=B2*C2` formula while raw
  `workbookProtection/@lockStructure` moves from `1` to `0`; every package
  member except `xl/workbook.xml` remains unchanged.
- Add the exact `workbook_structure_lock_removed` fact, a narrow raw
  workbook-protection validator, workbook-XML-only package isolation check,
  and corruption regressions. WCAB records a stored operational control only:
  it does not test a password, encryption, authentication, authorization, or
  Excel-client sheet-operation behavior.
- Extend the optional FormulaFence adapter to require the exact non-secret
  `workbook_protection_changed` transition and high-severity `FF022`, while
  WCAB independently proves the raw control, stable hidden-sheet/formula
  context, and package boundary.

## 0.24.0 — 2026-08-03

- Add a deterministic stored Excel error-checking suppression case. Its
  `Operations!B5=SUM(B2:B3)` formula, adjacent populated `B4` cell, downstream
  `C5=B5` formula, and calculation properties remain unchanged while one raw
  `ignoredErrors/ignoredError` `formulaRange=1` declaration is added for `B5`.
- Add the exact `ignored_error_rule_added` fact, a narrow raw worksheet
  validator, worksheet-only package isolation check, and corruption regressions.
  WCAB records a stored request to suppress an error-checking category only: it
  does not determine whether Excel would show a warning, evaluate a formula,
  decide whether a warning is justified, render an indicator, calculate a
  workbook, or claim Excel-client behavior.
- Extend the optional FormulaFence adapter to require the exact one-rule,
  redacted `ignored_error_controls_changed` transition and high-severity
  `FF037`, while WCAB independently proves the target, flag, formula context,
  and package boundary.

## 0.23.0 — 2026-08-03

- Add a deterministic custom number-format case. Its one `Operations!B2`
  metric retains its style index, raw numeric value, neighboring formula, and
  calculation properties while the referenced custom `numFmt/@formatCode`
  moves from a percentage display to `;;;`.
- Add the exact `cell_number_format_changed` fact, a narrow raw style/target
  validator, styles-only package isolation check, and corruption regressions.
  WCAB records stored display metadata only: it does not render a format,
  resolve locale or column-width behavior, calculate a workbook, or claim
  Excel-client behavior.
- Extend the optional FormulaFence adapter to require the exact one-custom
  `number_format_controls_changed` transition and high-severity `FF039`, while
  WCAB independently proves the raw format codes, target, style, and package
  boundary.

## 0.22.0 — 2026-08-03

- Add a deterministic conditional-formatting exception-threshold case. Its one
  `Operations!B2:B4` `cellIs` rule retains its target, priority, operator,
  differential red fill, metric values, and calculation properties while raw
  `formula` moves from `100` to `50`.
- Add the exact `conditional_formatting_threshold_changed` fact, a narrow raw
  worksheet/style validator, operations-worksheet-only package isolation
  check, and corruption regressions. WCAB records a stored visual-control
  formula only: it does not evaluate the rule, determine which cells a client
  formats, calculate a workbook, or claim Excel-client behavior.
- Extend the optional FormulaFence adapter to require the exact one-rule
  `conditional_formatting_changed` transition and high-severity `FF021`, while
  WCAB independently proves the raw threshold, stable fill/values, and package
  boundary.

## 0.21.0 — 2026-08-03

- Add a deterministic list data-validation source-retargeting case. Its one
  `Inputs!B2` list rule retains the target range, all entry-control metadata,
  both local source lists, current input, ordinary formulas, and calculation
  properties while raw `formula1` moves from `=Lists!$A$2:$A$4` to
  `=Lists!$B$2:$B$4`.
- Add the exact `data_validation_list_source_changed` fact, a narrow raw
  worksheet validator, Inputs-worksheet-only package isolation check, stable
  static model/dashboard lower bounds, and corruption regressions. WCAB records
  a stored entry-control source only: it does not evaluate the list, decide a
  future input's validity, accept or reject an entry, calculate a workbook, or
  claim Excel-client behavior.
- Extend the optional FormulaFence adapter to require the exact native
  `data_validation_changed` source/control profile and high-severity `FF020`,
  while WCAB independently proves the source transition and stable surrounding
  declarations.

## 0.20.0 — 2026-08-03

- Add a deterministic one-variable What-If Data Table case. Its raw
  `Sensitivity!D3` master retains the output range `D3:D5`, column orientation,
  recalculation request, visible cells, ordinary formula text, calculation
  properties, and saved table results while `f/@r1` switches its local input
  reference from `B2` to `B3`.
- Add the exact `what_if_data_table_input_reference_changed` fact, a narrow
  raw-worksheet Data Table validator, worksheet-part isolation check, stable
  static model/dashboard lower bounds, and corruption regressions. WCAB records
  the stored declaration only: it does not substitute inputs, recalculate a
  table or workbook, infer output values, resolve a circular dependency, or
  claim Excel-client behavior.
- Extend the optional FormulaFence adapter to require the exact redacted
  one-variable `what_if_data_tables_changed` profile and high-severity `FF034`,
  while WCAB independently proves the generated `B2 -> B3` `r1` transition.

## 0.19.0 — 2026-08-03

- Add a deterministic raw Scenario Manager case. One selected, locked
  `WCAB downside` scenario retains its sheet, two input references, second
  stored input, summary reference, selection state, protection, comment/user,
  number-format metadata, visible worksheet cells, formulas, and calculation
  properties while its stored `Inputs!B2` alternate value moves from `0.08` to
  `0.16`.
- Add the exact `scenario_manager_stored_input_value_changed` fact, a narrow
  raw-worksheet Scenario Manager validator, worksheet-part isolation check,
  local formula-path lower bound, and corruption regressions. WCAB records a
  stored alternate input only: it does not show or apply a scenario, calculate
  a model, create a scenario summary, infer an output, or claim client
  behavior.
- Extend the optional FormulaFence adapter to require the exact redacted
  one-scenario `scenario_manager_changed` profile and high-severity `FF035`,
  while WCAB independently proves the stored `0.08 -> 0.16` declaration.

## 0.18.0 — 2026-08-03

- Add a deterministic connection-only Power Query case over a local
  `SourceData` Excel Table. Its stored M `Table.SelectRows` literal moves from
  `North` to `South`; source cells, table definition, calculation properties,
  metadata, permissions, and every package member except
  `customXml/item1.xml` remain unchanged.
- Add the exact `power_query_m_filter_changed` fact, a bounded raw Data Mashup
  envelope validator, custom-XML relationship validation, package-isolation
  check, and corruption regressions. WCAB records only the stored definition:
  it does not execute M, refresh a query, materialize output, calculate a
  workbook, infer returned rows, or claim client behavior.
- Extend the optional FormulaFence adapter to require the exact redacted
  one-query `power_query_changed` profile and high-severity `FF024`, while
  WCAB independently proves the local source and `North -> South` M literal.

## 0.17.0 — 2026-08-03

- Add a deterministic relationship-backed PivotTable Slicer-selection case.
  Its local worksheet source, PivotCache, stored `Report!A1:B2` display cells,
  and `Dashboard!B4=Report!$B$2` consumer remain unchanged while raw
  Slicer-cache selection moves from the `North` item to the `South` item.
- Add the exact `pivot_slicer_selection_changed` fact, raw Slicer-cache to
  PivotCache/PivotTable relationship validator, Slicer-cache-part isolation
  check, and corruption regressions. WCAB records stored cache-item state only:
  it does not create a visual Slicer, apply a filter, refresh, calculate, or
  render a PivotTable, infer a changed report result, or claim client behavior.
- Extend the optional FormulaFence adapter to require its exact redacted
  one-Slicer `slicer_timeline_cache_definitions_changed` profile and
  high-severity `FF032`, while WCAB independently proves the local item-index
  and selected-value transition.

## 0.16.0 — 2026-08-03

- Add a deterministic relationship-backed PivotTable value-field aggregation
  case. Its local worksheet source, cache records, stored `Report!A1:B2`
  display cells, and `Dashboard!B4=Report!$B$2` consumer remain unchanged while
  raw `dataField/@subtotal` switches from `sum` to `average`.
- Add the exact `pivot_data_field_aggregation_changed` fact, raw cache-to-
  PivotTable relationship validator, PivotTable-XML isolation check, and
  corruption regressions. WCAB records the stored aggregate declaration only:
  it does not refresh, calculate, or render a PivotTable, infer a changed
  displayed value, or claim client behavior.
- Extend the optional FormulaFence adapter to require its exact redacted
  one-PivotTable `pivot_table_definitions_changed` profile and high-severity
  `FF031`, while WCAB independently proves the source and `sum -> average`
  declaration transition.

## 0.15.0 — 2026-08-03

- Add a deterministic relationship-backed DrawingML chart case. Its one chart
  at `Dashboard!D2` retains its anchor, title/category references, source
  worksheet cells, and every package member except `xl/charts/chart1.xml`,
  while raw `c:ser/c:val/c:numRef/c:f` switches from
  `Source!$B$2:$B$4` to `Source!$C$2:$C$4`.
- Add the exact `chart_series_value_reference_changed` fact, raw worksheet to
  drawing to chart relationship validator, chart-XML isolation check, and
  corruption regressions. WCAB records the stored source binding only: it does
  not open Excel, calculate values, refresh chart data, render a chart, infer a
  visual difference, or claim client behavior.
- Extend the optional FormulaFence adapter to require its exact redacted
  one-chart `chart_definitions_changed` profile and high-severity `FF030`,
  while WCAB independently proves the chart source-reference transition.

## 0.14.0 — 2026-08-03

- Add a deterministic relationship-backed PivotTable cache case. Its local
  worksheet source, stored `Report!A1:B2` display cells, and
  `Dashboard!B4=Report!$B$2` consumer remain unchanged while raw
  `pivotCacheDefinition/@refreshOnLoad` moves from false to true.
- Add the exact `pivot_cache_refresh_on_load_changed` fact, raw-OOXML graph
  validator, package-member isolation check, and corruption regressions. WCAB
  records a stored open-time request only: it does not open Excel, refresh the
  cache, calculate or render the PivotTable, or claim a changed report value.
- Extend the optional FormulaFence adapter to require its exact redacted
  `pivot_cache_refresh_controls_changed` evidence and high-severity `FF023`,
  while WCAB independently proves the local source and PivotTable binding.

## 0.13.0 — 2026-08-03

- Add a deterministic active-AutoFilter case. Its worksheet `Report!A1:B5`
  filter-column-0 list criterion moves from `North` to `South`, while
  `Report!D2=SUBTOTAL(109,B2:B5)` and its `Dashboard!B4` consumer remain
  unchanged.
- Add the exact `auto_filter_criteria_changed` fact, a raw-OOXML validator,
  and corruption regressions. WCAB records the stored AutoFilter declaration
  and formula/dependency boundary; it does not apply the filter, calculate a
  subtotal, infer a visible row set, or claim a client display or print result.
- Extend the optional FormulaFence adapter to require its matching redacted
  `filter_visibility_controls_changed` evidence and high-severity `FF036`,
  while WCAB independently proves the raw criterion values.

## 0.12.0 — 2026-08-03

- Add a deterministic workbook serial-date-system case. Its raw
  `Inputs!B2=45292` serial, `yyyy-mm-dd` number format, local `Model!B2` and
  `Dashboard!B4` formulas, and explicit `dateCompatibility=true` control stay
  unchanged while raw `workbookPr/@date1904` changes from `false` to `true`.
- Add the exact `workbook_date_system_changed` fact, raw-OOXML validator, and
  corruption regressions. WCAB records the stored controls and package
  boundary; it does not calculate a formula, convert the serial, predict a
  displayed date, or claim behavior for a particular Excel client.
- Extend the optional FormulaFence adapter to require the exact normalized
  `FF117` date-system evidence alongside the mapped fact, rejecting missing or
  broadened date-control observations.

## 0.11.0 — 2026-08-02

- Add a deterministic saved-formula-result case. Its direct input, formula,
  calculation properties, and local downstream formula remain unchanged while
  the raw numeric `<v>` saved beside `Model!B2`'s `<f>` changes from `20` to
  `25`.
- Add the exact `formula_cached_result_changed` fact, relationship-resolved
  raw-OOXML validator, package-member isolation check, and corruption
  regressions. WCAB records a saved result; it does not calculate a formula,
  decide whether either value is current, stale, tampered, or correct, or
  claim what an Excel client displays after opening.
- Extend the optional FormulaFence adapter to require the exact one-result
  `FF042` evidence profile and unexplained-change count, respecting its
  deliberate redaction of cache values and formula-cell locations.

## 0.10.0 — 2026-08-02

- Add a deterministic precision-as-displayed case. Its stored `10.005` input,
  two-decimal number format, formula, and local downstream reference remain
  unchanged while `calcPr/@fullPrecision` changes from `true` to `false`.
- Add the exact `precision_as_displayed_enabled` fact, raw-OOXML validator, and
  corruption regression coverage. WCAB records the stored calculation control;
  it does not open or save an Excel workbook, calculate a formula, claim a
  rounded value, or claim that a client applies the setting.
- Extend the optional FormulaFence adapter to require the exact isolated
  `FF009` full-precision transition rather than accepting an arbitrary
  calculation-setting difference.

## 0.9.0 — 2026-08-02

- Add a deterministic iterative-calculation case. Its direct self-referencing
  formula and local downstream reference stay unchanged while `calcPr/@iterate`
  changes from `false` to `true` with an explicit 100-iteration / 0.001-delta
  bound.
- Add the exact `iterative_calculation_enabled` fact, raw-OOXML validator, and
  corruption regression coverage. WCAB records the stored calculation control;
  it does not calculate the circular model or claim convergence, a cached value,
  or numerical correctness.
- Extend the optional FormulaFence adapter to require the exact isolated
  `FF009` calculation-settings transition rather than accepting an arbitrary
  calculation-setting difference.

## 0.8.0 — 2026-08-02

- Add a deterministic external-workbook link startup-policy case. Its synthetic
  external formula and downstream local reference stay unchanged while
  `workbookPr/@updateLinks` changes from `never` to `always`.
- Add the exact `external_workbook_link_update_policy_changed` fact, raw-OOXML
  validator, and corruption regression coverage. The source workbook is absent
  by design; WCAB does not resolve, refresh, authenticate to, or calculate it.
- Extend the optional FormulaFence adapter to require the exact isolated
  `FF023` external-link policy transition, rejecting a report that combines it
  with another workbook-wide refresh-control change.

## 0.7.0 — 2026-08-02

- Add a deterministic legacy-CSE to dynamic-array formula-mode case. The
  anchor formula and currently stored output range remain unchanged, while
  raw OOXML metadata changes the formula from fixed output semantics to a
  resizable dynamic array.
- Add the exact `array_formula_mode_changed` fact, independent raw-OOXML
  metadata validator, and a corruption regression test so the contract does
  not rely on the truth file alone.
- Extend the optional FormulaFence adapter to require the exact `FF018`
  legacy-CSE-to-dynamic transition and output range rather than accepting a
  generic array-mode diff.

## 0.6.0 — 2026-08-02

- Add a deterministic relationship-backed external-data connection case where
  only `refreshOnLoad` changes from false to true. The endpoint is synthetic
  and non-routable; WCAB neither opens it nor claims a calculated result.
- Add the `external_data_connection_refresh_on_load_changed` fact and raw-OOXML
  validator for the exact workbook-local connection ID and transition.
- Extend the optional FormulaFence adapter to require the matching `FF023`
  control details, rather than treating an arbitrary connection change as the
  fact.

## 0.5.0 — 2026-08-02

- Add deterministic `INDIRECT` address-driver and `OFFSET` displacement-driver
  cases. In both, formula text stays unchanged while the selected effective
  reference can change.
- Add the scoreable `dynamic_reference_driver_changed` coverage expectation
  and validator invariant for a changed literal driver feeding an unchanged
  dynamic formula.
- Extend the optional FormulaFence adapter to pair its native input-change
  record with its candidate-profile dynamic-reference feature, preserving the
  analysis boundary without claiming formula evaluation.

## 0.4.0 — 2026-08-02

- Upgrade the fixture contract to schema version 3 and add a deterministic
  `INDIRECT`-introduction case with an explicit static-dependency boundary.
- Upgrade the normalized observation protocol to version 2, adding scoreable
  coverage declarations, disclosure recall, and strict-mode handling for
  unrecognized declarations.
- Extend the optional FormulaFence adapter to map native `FF012`
  dynamic-reference coverage evidence without imposing an approval decision.

## 0.3.0 — 2026-08-02

- Add a tool-neutral normalized observation protocol, template command, and
  deterministic score report for expected-fact recall, coverage, and
  reference-policy agreement.
- Add an optional FormulaFence normalizer that emits mapped observations while
  leaving approval-policy decisions explicitly unset.
- Keep unrecognized facts explicit rather than calling them false positives
  against WCAB's intentionally targeted oracle.

## 0.2.0 — 2026-08-02

- Upgrade the fixture contract to schema version 2 and add a deterministic
  Excel Table scope-expansion case with an unchanged structured-reference
  formula.
- Validate the Table definition and structured-reference boundary directly,
  and map the case to FormulaFence's table-definition diff evidence.

## 0.1.1 — 2026-08-02

- Add a deterministic JSONL case catalogue with exact workbook byte counts and
  SHA-256 digests.
- Add a `wcab manifest` command and a Hugging Face dataset card for direct
  dataset consumption.

## 0.1.0 — 2026-08-01

- Introduced 16 deterministic synthetic workbook-change cases across finance,
  operations, governance, structural, and multi-workbook portfolio workflows.
- Added a tool-neutral JSON truth contract, a local static-impact lower-bound
  validator, and byte-reproducibility checks.
- Added an optional FormulaFence reference adapter with explicit unmapped-case
  reporting and targeted lint checks.
