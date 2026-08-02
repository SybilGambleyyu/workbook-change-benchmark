# Workbook Change Assurance Benchmark

`WCAB` is an open, reproducible benchmark for tools that review changes to
Excel workbooks. It supplies paired baseline/candidate workbooks, explicit
change facts, dependency-impact lower bounds, scoreable analysis-boundary
disclosures, and a small validator. Its purpose is not to test whether a model
can *write* a formula; it tests whether
a reviewer or CI gate can explain whether a workbook change is safe to accept.

The first release is deliberately synthetic and open by construction.  Every
workbook is generated from source code in this repository, so no private model
or recovered business data is redistributed.

## Why this benchmark exists

Spreadsheets are commonly changed outside conventional version control, while
formula changes can alter a model's behaviour.  Existing public resources are
valuable but target adjacent questions:

- [Modified EUSES](https://spreadsheets.sai.tugraz.at/index.php/corpora-for-benchmarking/euses/)
  supplies injected formula faults and result-cell test decisions.
- [VEnron](https://researchportal.hkust.edu.hk/en/publications/venron-a-versioned-spreadsheet-corpus-and-related-evolution-analy/)
  recovers real workbook version history, but its recovered history does not
  say which changes should block a review gate.
- [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench)
  evaluates spreadsheet manipulation tasks.
- Recent [formula-repair work](https://www.microsoft.com/en-us/research/publication/benchmark-dataset-generation-and-evaluation-for-excel-formula-repair-with-llms/)
  evaluates repair of runtime formula errors.

WCAB fills the change-assurance gap: every case records an observable change
fact, the intended review disposition, and the static impact that a conservative
tool should not miss.  It is useful for spreadsheet-diff tools, CI policy
gates, static analyzers, and agent workflows that propose workbook edits.

## Scope and non-goals

Version `0.20` covers formula-to-value replacements, formula reference drift,
value changes with downstream effects, external formula references, named
ranges, data validation, conditional formatting, sheet visibility, direct cell
protection, calculation settings, static cycles, portfolio dependencies,
formula refactors, 3-D-reference scope changes, and Excel Table scope changes
that leave a structured-reference formula textually unchanged. It also covers
new `INDIRECT` references whose dependency target comes from workbook text,
unchanged `INDIRECT` and `OFFSET` formulas whose address or displacement driver
changes, an external-data connection whose refresh-on-open behavior changes
without any worksheet-cell edit, and an external-workbook link policy that
switches from never to always updating when the workbook opens. It also covers
an unchanged direct circular formula that enables iterative calculation, an
unchanged precision-sensitive input and formula whose calculation switches to
precision as displayed, and an unchanged array formula switching from a fixed
legacy CSE output range to dynamic-array semantics. It also covers a saved
numeric formula result changing while its formula, direct input, calculation
metadata, and local downstream reference remain unchanged.
It also covers a workbook-wide 1900-to-1904 date-system control change while
an explicit compatibility control, a raw numeric serial, its date number
format, and local formulas remain unchanged. It does not infer an Excel-client
display or calculate a converted date.
It also covers an active worksheet AutoFilter criterion moving from `North` to
`South` while `Report!D2=SUBTOTAL(109,B2:B5)` and its `Dashboard!B4` consumer
remain unchanged. It records the stored control only: it does not apply a
filter, calculate a subtotal, infer visible rows, or claim a display or print
outcome.
It also covers a local worksheet-backed PivotTable cache whose raw
`refreshOnLoad` control changes from false to true while its source cells,
stored `Report!A1:B2` display cells, and `Dashboard!B4=Report!$B$2` consumer
remain unchanged. It records only the stored open-time request: it does not
open Excel, refresh the cache, calculate or render the PivotTable, or claim a
changed report result.
It also covers a local worksheet-backed PivotTable value field whose raw
`dataField/@subtotal` moves from `sum` to `average` while its source cells,
cache records, stored `Report!A1:B2` display cells, and
`Dashboard!B4=Report!$B$2` consumer remain unchanged. It records only the
stored aggregation declaration: it does not open Excel, refresh, calculate, or
render the PivotTable, infer a changed displayed value, or claim client
behavior.
It also covers a relationship-backed local PivotTable Slicer cache whose
stored `Region` selection moves from `North` to `South` while its source cells,
PivotCache, stored `Report!A1:B2` display cells, and
`Dashboard!B4=Report!$B$2` consumer remain unchanged. It records Slicer-cache
item state only: it does not create a visual Slicer or drawing, apply the
filter, refresh, calculate, or render the PivotTable, infer a changed report
result, or claim client behavior.
It also covers a connection-only local Power Query whose stored M
`Table.SelectRows` filter moves from `North` to `South` over the generated
`SourceData` Excel Table. Source cells, the table definition, calculation
properties, metadata, and permission controls remain unchanged. It records the
stored M definition only: it does not execute M, apply a filter, refresh a
query, materialize output, calculate formulas, infer returned rows, or claim
client behavior.
It also covers a selected, locked Scenario Manager declaration whose stored
alternate `Inputs!B2` value moves from `0.08` to `0.16` while visible worksheet
cells, formulas, selection state, second input, summary reference, protection,
comment/user, number-format metadata, and calculation properties remain
unchanged. It records only the raw alternate input: it does not show or apply
a scenario, calculate the model, generate a Scenario Summary, infer an output,
or claim client behavior.
It also covers a column-oriented one-variable What-If Data Table whose raw
`Sensitivity!D3` master changes its local `f/@r1` input reference from `B2` to
`B3` while its `D3:D5` output range, orientation, recalculation request,
visible cells, ordinary formula text, calculation properties, and saved table
results remain unchanged. It records the stored declaration only: it does not
substitute input values, calculate the workbook or table, infer an output,
resolve a circular dependency, or claim client behavior.
It also covers a Dashboard DrawingML chart whose raw numeric-series reference
switches from `Source!$B$2:$B$4` to `Source!$C$2:$C$4` while all source cells,
the chart anchor, its title/category references, and every package member
outside the chart part remain unchanged. It records a stored report binding
only: it does not open Excel, calculate values, refresh chart data, render a
chart, infer a visual difference, or claim client behavior.

The benchmark does **not** evaluate formula execution or claim that a
candidate's numerical results are correct.  A case's `review_expectation` is a
benchmark convention, not a universal business policy.  Tools should report
unsupported features instead of treating absent evidence as a pass.

## Quick start

```bash
python -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'

# Regenerate the committed fixtures and validate their ground truth.
wcab build --output fixtures
wcab validate --fixtures fixtures
wcab manifest --fixtures fixtures

# Run the project's own tests.
pytest
```

The fixture tree is intentionally small enough to inspect:

```text
fixtures/
  finance/formula_to_value/
    baseline.xlsx
    candidate.xlsx
    truth.json
```

`fixtures/manifest.jsonl` is a deterministic, one-row-per-case catalogue.
Every row preserves the truth contract and records the relative path, byte
count, and SHA-256 digest for its baseline and candidate workbooks. It lets
benchmark runners enumerate exactly what they consumed without relying on
directory-name conventions.

The release is also mirrored as a
[Hugging Face dataset](https://huggingface.co/datasets/SybilGambleyyu/workbook-change-benchmark),
where `manifest.jsonl` is available at the dataset root for programmatic use.

## Truth contract

Each `truth.json` is schema version 3 and contains:

- `facts`: observable before/after facts, such as a formula becoming a value.
- `must_reach`: formula cells that are statically reachable from a changed
  source in the candidate's local dependency graph.  These are lower bounds,
  not a claim of complete Excel dependency analysis.
- `review_expectation`: `block`, `review`, or `allow` for the benchmark's
  reference policy.
- `coverage`: case-specific boundaries that consumers must preserve.
- `coverage_expectations`: machine-matchable analysis-boundary disclosures;
  current expectations require visible static-dependency coverage evidence
  both when a dynamic reference is introduced and when a pre-existing dynamic
  formula receives a changed selector.

The bundled validator checks the generated workbooks against this contract.  A
tool adapter may map its own output into these facts, but WCAB deliberately
does not prescribe one tool's report format.

The complete fact schema is in [docs/schema.md](docs/schema.md), and the
reproducible validation record is in [docs/validation.md](docs/validation.md).

## Score another tool without adopting its report schema

WCAB supplies a normalized observation protocol for adapters around any local
diff or review tool. Start from an explicitly unsupported template, let the
adapter replace cases with exact observed WCAB facts, then score the report:

```bash
wcab observation-template --fixtures fixtures --output observations.json
# Run your adapter to populate observations.json.
wcab score --fixtures fixtures --observations observations.json --output score.json
```

The score reports expected-fact recall, coverage-disclosure recall, analyzed
coverage, and agreement with the benchmark's reference review convention. It
deliberately lists unrecognized observations instead of calling them false
positives: WCAB facts are targeted change assertions, not a claim to enumerate
every possible workbook difference. See [docs/observations.md](docs/observations.md)
for the complete protocol and strict-mode behavior.

The optional FormulaFence reference adapter can also emit normalized
observations directly. It records mapped change facts and native coverage
warnings, but intentionally leaves review dispositions unset because
FormulaFence is an analyzer, not a universal approval-policy engine:

```bash
wcab formulafence-observations --fixtures fixtures \
  --executable formulafence --output formulafence-observations.json
wcab score --fixtures fixtures --observations formulafence-observations.json
```

## Optional FormulaFence reference adapter

The repository includes a local adapter for FormulaFence to demonstrate how a
real tool can be evaluated without making FormulaFence a dependency or making
its JSON schema the benchmark schema:

```bash
# Install FormulaFence by your normal trusted route, then:
wcab formulafence --fixtures fixtures --strict
```

It runs only local commands against the synthetic fixtures. The adapter maps
the introduced-dynamic case from FormulaFence's native warning, maps the
unchanged-formula driver cases from the observed input change plus the
candidate profile's dynamic-reference feature, and checks the exact
connection-level refresh-on-open transition in FormulaFence's `FF023` evidence.
It separately requires FormulaFence's `pivot_cache_refresh_controls_changed`
record and matching `FF023` for the local worksheet-backed PivotCache case.
FormulaFence safely reports cache-level source/control metadata rather than
source cells or rendered PivotTable output; WCAB's raw validator independently
proves the cache/PivotTable relationship binding, stable stored cells, direct
dashboard edge, and cache-definition-only package change. Neither layer
refreshes or renders the PivotTable.
For the PivotTable aggregation case, it requires FormulaFence's exact redacted
one-PivotTable `pivot_table_definitions_changed` profile,
`pivot_table_layout_material_changed`, and `FF031`. FormulaFence does not
expose the selected aggregate function or a rendered report. WCAB's raw
validator instead proves the local source/cache/PivotTable relationship graph,
stable stored cells, exact `sum`-to-`average` declaration, and
PivotTable-part-only package change. Neither layer refreshes, calculates, or
renders the PivotTable.
For the PivotTable Slicer-selection case, it requires FormulaFence's exact
redacted one-Slicer `slicer_timeline_cache_definitions_changed` profile,
`slicer_filter_state_or_definition_material_changed`, and `FF032`.
FormulaFence does not expose the Slicer name, selected item/value, or a
rendered report. WCAB's raw validator instead proves the local Slicer-cache to
PivotCache/PivotTable graph, stable stored cells, exact `North`-to-`South`
selected-item transition, and Slicer-cache-part-only package change. Neither
layer creates a visual Slicer, applies the filter, refreshes, calculates, or
renders a PivotTable.
For the Power Query M-filter case, it requires FormulaFence's exact redacted
one-query `power_query_changed` profile, `formula_material_changed`, and
`FF024`. FormulaFence does not expose M source text, local-table values, or a
query result. WCAB's raw validator instead proves the package-root custom-XML
relationship, compact Data Mashup envelope, connection-only metadata and
permission controls, exact local table binding, `North`-to-`South` stored M
literal, and custom-XML-part-only package change. Neither layer executes M,
refreshes a query, materializes output, calculates a workbook, or infers
returned rows.
For the Scenario Manager stored-input case, it requires FormulaFence's exact
redacted one-scenario `scenario_manager_changed` profile,
`scenario_definition_material_changed`, and `FF035`. FormulaFence does not
expose the scenario name, stored values, input references, comment, or user.
WCAB's raw validator instead proves the generated `Inputs!B2` `0.08`-to-`0.16`
alternate-value declaration, fixed scenario metadata, stable visible cells and
formula path, and Inputs-worksheet-only package change. Neither layer shows or
applies a scenario, calculates a result, or generates a Scenario Summary.
For the What-If Data Table case, it requires FormulaFence's exact redacted
one-variable `what_if_data_tables_changed` profile,
`data_table_definition_material_changed`, and `FF034`. FormulaFence does not
expose the table output range, local input references, or calculated table
values. WCAB's raw validator instead proves the generated `D3:D5` master
declaration, exact `B2`-to-`B3` `r1` transition, stable input grid and formulas,
direct static model/dashboard lower bounds, and Sensitivity-worksheet-only
package change. Neither layer substitutes inputs, calculates, resolves a
circular dependency, or claims client behavior.
For the chart-series case, it requires FormulaFence's exact redacted one-chart
profile, `chart_definition_material_changed`, and `FF030`; FormulaFence does
not expose the source formula or a visual result. WCAB's raw validator instead
proves the worksheet-to-drawing-to-chart binding, stable anchor/title/category
references, exact value-reference transition, and chart-part-only package
change. Neither layer calculates or renders the chart.
It separately requires the exact isolated external-workbook policy transition
from `never` to `always` in `FF023` evidence. For the array-mode case, it
requires the exact legacy-CSE-to-dynamic transition and output range in
FormulaFence's `FF018` evidence. For the iterative-calculation case, it
requires the exact `FF009` switch from iteration disabled to enabled with the
declared bounds. For the precision-as-displayed case, it requires the exact
isolated `fullPrecision: true -> false` `FF009` control transition while the
stored input, number format, formula, and other calculation controls remain
unchanged. For the saved-formula-result case, it requires FormulaFence's exact
one-result `FF042` evidence: the tool intentionally redacts the result and
formula-cell location, so the adapter also requires its stable numeric cache
profile and unexplained-change count. For the workbook-date-system case, it
requires `workbook_date_system_changed` plus `FF117`: normalized `date1904`
must move from `false` to `true`, explicit `dateCompatibility` must remain
`true`, and no unrecognized date controls may appear. WCAB independently
validates the raw serial, style, formulas, and package-member boundary. It
requires `filter_visibility_controls_changed` and `FF036` for the active
AutoFilter case. FormulaFence intentionally reports its aggregate, redacted
control profile rather than the selected values; WCAB's raw validator
independently establishes the `North`-to-`South` criterion transition, stable
`SUBTOTAL` and dashboard formulas, direct dependency edge, and worksheet-only
package change. Neither layer applies the filter or claims a subtotal result.
The adapter keeps unmapped facts explicit; it does not treat the benign
structural-rewrite annotation as a generic semantic-equivalence claim.

## Reproducibility

`wcab build` fixes workbook metadata and canonicalizes generated XLSX ZIP entry
ordering and timestamps.  Rebuilding the fixture tree should produce identical
bytes on supported Python/openpyxl versions.  `pytest` enforces that property.

## Contributing

New cases need a deterministic generator, an explicit truth contract, an
explanation of the static boundary, and a validation test.  Do not add a
workbook unless its redistribution rights are clear; generated fixtures are
preferred.

## License

MIT.  See [LICENSE](LICENSE).
