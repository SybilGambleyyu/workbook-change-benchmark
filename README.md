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

Version `0.37` covers formula-to-value replacements, formula reference drift,
value changes with downstream effects, external formula references, named
ranges, data validation, conditional formatting, sheet visibility, direct cell
and workbook-structure protection, calculation settings, static cycles,
portfolio dependencies, formula refactors, 3-D-reference scope changes, and Excel Table scope changes
that leave a structured-reference formula textually unchanged. It also covers
one Table-level calculated-column master whose raw
`tableColumn/calculatedColumnFormula` changes while the Table binding, ordinary
row formulas, structured-reference dashboard formula, calculation properties,
and every package member except `xl/tables/table1.xml` remain fixed. It records
the stored master only: it does not fill a column, reconcile a master with row
formulas, evaluate a structured reference, infer a total, or claim client
behavior. It also covers one embedded Power Pivot/Data Model whose sole raw
`x15:modelRelationship/@toColumn` moves from `CalendarModel.DateKey` to
`CalendarModel.FiscalDateKey` while its workbook-to-model relationship, content
type, local Tables, calculation properties, and opaque `xl/model/item.data`
payload remain fixed. It records a stored relationship declaration only: it
does not deserialize the Analysis Services payload, evaluate DAX, refresh a
model, calculate or render a PivotTable or chart, infer model-to-cell impact,
or claim client behavior. It also covers one workbook-scoped reusable `ScenarioValue` named
`LAMBDA` whose stored calculation body changes while input cells, its calling
`Model!B2` formula, dashboard consumer, calculation properties, workbook
relationships, and every package member except `xl/workbook.xml` remain fixed.
It records the formula definition only: it does not calculate a result, infer
Excel-version support, or claim that a client recalculates, spills, or persists
a value. It also covers
new `INDIRECT` references whose dependency target comes from workbook text,
unchanged `INDIRECT` and `OFFSET` formulas whose address or displacement driver
changes, an external-data connection whose refresh-on-open behavior changes
without any worksheet-cell edit, and a separate connection whose raw web-query
URL moves between two reserved `example.invalid` endpoints while its identity,
refresh controls, visible saved cells, and formula context remain fixed. It
records the stored endpoint only: it does not resolve, open, fetch,
authenticate to, trust, refresh, calculate, or otherwise interact with either
source, and does not claim a client returns a value. It also covers an
external-workbook link policy that switches from never to always updating when
the workbook opens. It also covers
an unchanged external-workbook formula whose local `externalLink` package
relationship moves to a different reserved source while its formula text and
local downstream dependency remain fixed. It records that stored relationship
only: it does not resolve, open, fetch, authenticate to, trust, refresh,
calculate, or otherwise interact with either source, and does not claim that a
client updates a link or returns a value. It also covers
one local `ScenarioRate` defined name whose qualified external-workbook source
text changes while `Model!B2=ScenarioRate*2`, its `Dashboard!B4` consumer, and
every package member except `xl/workbook.xml` remain fixed. It is intentionally
not a relationship-backed `externalLink` package: WCAB records only the stored
defined-name text, never opens either synthetic source, and makes no claim that
a client resolves the name or returns a value. It also covers one protected
worksheet whose explicit `sheetProtection/@sort` lock moves from `1` to `0`
while protection remains enabled and all cells, formulas, styles, calculation
properties, and workbook-level protection remain fixed. It records a stored
permission only: it does not test a password, authorization, an editable range,
whether a client permits a sort, or what a sort would change.
It also covers one macro-enabled `.xlsm` workbook whose workbook-scoped
`_xlnm.Auto_Open` binding moves from the very-hidden `Macro Automation!A1`
cell to `Macro Automation!A2` while the XLM macro sheet remains byte-identical
with only two static `HALT()` formulas. Its macro-sheet relationship, content
types, ordinary `Inputs!B2 → Model!B2 → Dashboard!B4` formula context,
calculation properties, and every package member except `xl/workbook.xml`
remain fixed. WCAB records a stored dispatch declaration only: it never opens
Excel, enables or executes XLM code, parses or emulates macro instructions,
resolves a dynamic name, inspects macro-security or trust settings, or claims
client behavior.
It also covers
one relationship-backed QueryTable whose own `refreshOnLoad` request moves from
false to true while its fixed connection-level refresh control remains false,
its synthetic non-routable endpoint, saved cells, and
`ImportedData!B2 → Summary!B2 → Dashboard!B4` formula context remain
unchanged. It records a stored request only: it does not open a connection,
fetch a URL, refresh a query, materialize rows, calculate a workbook, or claim
that a client refreshes successfully.
It also covers one relationship-backed worksheet cell hyperlink whose visible
`Inputs!B2` text, local `Inputs!B2 → Summary!B2 → Dashboard!B4` formula
context, calculation properties, and worksheet XML remain unchanged while its
one external OOXML hyperlink relationship target moves between two reserved
`example.invalid` URLs. It records the stored target only: it does not resolve,
open, fetch, visit, execute, calculate, or otherwise interact with either URL,
and does not claim that a client follows it.
It also covers
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
It also covers a workbook structure lock changing from enabled to disabled
while a hidden `ReviewControls` sheet and `Inputs!D2=B2*C2` formula remain
unchanged. It records only raw `workbookProtection/@lockStructure`: it does not
test a password, encryption, authentication, authorization, or whether a client
will permit a particular sheet operation.
It also covers an active worksheet AutoFilter criterion moving from `North` to
`South` while `Report!D2=SUBTOTAL(109,B2:B5)` and its `Dashboard!B4` consumer
remain unchanged. It records the stored control only: it does not apply a
filter, calculate a subtotal, infer visible rows, or claim a display or print
outcome.
It also covers a relationship-backed saved Named Sheet View whose alternate
list criterion moves from `North` to `South` while its worksheet base
AutoFilter remains untouched and `Report!D2=SUBTOTAL(109,B2:B5)` plus its
`Dashboard!B4` consumer remain unchanged. It records the stored alternate
review lens only: it does not activate or render a view, apply a filter,
calculate a subtotal, infer visible rows, or claim a display or print outcome.
It also covers a local XML Map whose `InvoiceLines[Net amount]` table-column
binding changes from a synthetic invoice `NetAmount` XPath to `TaxAmount`
while its map/schema/file-binding declarations, sheet-level single-cell
mapping, visible cells, table total, and dashboard formula remain unchanged.
It records a stored import/export data contract only: it does not access a
file, validate a schema, import or export XML, materialize data, calculate a
result, or claim client behavior.
It also covers a relationship-backed Office Web Add-in task-pane declaration
whose stored `Office.AutoShowTaskpaneWithDocument` request changes from false
to true while its synthetic local FileSystem reference, hidden locked
task-pane layout, ordinary cells, and `Inputs!B2 → Model!B2 → Dashboard!B4`
formula context remain fixed. It records a stored request only: it does not
install, load, execute, or fetch an add-in or manifest, and it does not claim
that a task pane opens.
It also covers one local worksheet OLE declaration whose stored
`oleObject/@autoLoad` control changes from false to true while its one fixed
internal relationship, opaque synthetic bytes, ordinary cells, calculation
properties, and `Inputs!B2 → Model!B2 → Dashboard!B4` formula context remain
unchanged. It records a raw request only: it does not deserialize, open,
render, execute, register, or invoke an object server, and it does not claim
that an object loads successfully.
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
It also covers a list data-validation rule at `Inputs!B2` whose raw `formula1`
source moves from `=Lists!$A$2:$A$4` to `=Lists!$B$2:$B$4` while the target range,
entry-control metadata, both source lists, current input, ordinary formulas,
and calculation properties remain unchanged. It records the stored input
control only: it does not evaluate the list source, decide whether a future
entry is valid, accept or reject an entry, calculate a workbook, or claim
Excel-client behavior.
It also covers one `Operations!B2:B4` `cellIs` conditional-formatting rule
whose raw threshold formula moves from `100` to `50` while its target, priority,
`greaterThan` operator, differential red fill, metric values, and calculation
properties remain unchanged. It records the stored visual control only: it
does not evaluate the rule, determine which cells a client formats, calculate
a workbook, or claim Excel-client behavior.
It also covers one `Operations!B2` reported margin whose referenced custom
`numFmt/@formatCode` moves from `0.0%;[Red](0.0%);-` to `;;;` while its
style index, raw numeric value, neighboring `=B2` formula, calculation
properties, and every package member except `styles.xml` remain unchanged.
It records stored display metadata only: it does not render a number format,
resolve locale or column-width behavior, decide what a client displays,
calculate a workbook, or claim Excel-client behavior.
It also covers one stored `Operations!B5` Excel error-checking suppression:
the `=SUM(B2:B3)` formula, adjacent populated `B4` cell, downstream `C5=B5`
formula, calculation properties, and all ordinary cells remain unchanged while
a raw `ignoredErrors/ignoredError` `formulaRange=1` declaration is added for
`B5`. It records that stored warning-suppression request only: it does not
determine whether Excel would show a warning, evaluate a formula, decide
whether the warning is justified, render an indicator, change application-level
error-checking options, calculate a workbook, or claim Excel-client behavior.
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
directory-name conventions. A pair case uses exactly one matching
`baseline.xlsx` / `candidate.xlsx` or `baseline.xlsm` / `candidate.xlsm` pair;
the macro-enabled extension is explicit in the catalogue rather than inferred
from a case name.

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
For the QueryTable case, it separately requires FormulaFence's exact redacted
`query_table_refresh_controls_changed` profile and matching `FF023`, including
the isolated QueryTable-level false-to-true request, fixed connection controls,
and no opaque metadata. WCAB's raw validator independently proves the direct
worksheet-to-QueryTable and workbook-to-connections graph, stable saved cells
and direct formula path, and QueryTable-part-only package change. Neither layer
opens a connection, fetches a URL, refreshes a query, materializes rows,
calculates a workbook, or claims that a client refreshes successfully.
For the cell-hyperlink target case, it separately requires FormulaFence's exact
redacted `cell_hyperlink_controls_changed` profile and matching `FF047`,
including one external hyperlink binding and no location, display, tooltip, or
unrecognized declaration. FormulaFence deliberately does not expose the target
or relationship ID; WCAB's raw validator independently proves their exact
transition, stable visible cells/formulas, and relationship-part-only package
change. Neither layer resolves, opens, fetches, visits, or otherwise interacts
with the target, or claims that a client follows it.
For the external-workbook source case, it separately requires FormulaFence's
exact redacted `external_link_packages_changed` profile and matching
high-severity `FF025`, including one external workbook, one declared external
sheet, no DDE/OLE links or cached external data, no opaque metadata, and only
`source_material_changed`. FormulaFence deliberately does not expose the raw
source target or relationship IDs; WCAB's raw validator independently proves
their exact reserved-target transition, fixed external-reference graph, stable
formula context, and externalLink-relationship-part-only package change. A
generic external-relationship diff is deliberately not mapped to this precise
fact. Neither layer resolves, opens, fetches, authenticates to, trusts,
refreshes, calculates, or otherwise interacts with a source, or claims that a
client updates a link or returns a value.
For the external-defined-name source case, it requires FormulaFence's exact
one-surface `external_workbook_link_surfaces_changed` ledger and high-severity
`FF081`, plus its matching `defined_name_changed` record and `FF008`. The
ledger alone is insufficient: the adapter also requires the one `ScenarioRate`
definition to move between the two generated source expressions. WCAB's raw
validator independently establishes the exact stored text, absence of an
`externalLink` package, stable formulas, and workbook-XML-only boundary. Neither
layer resolves, opens, fetches, authenticates to, trusts, refreshes, calculates,
or otherwise interacts with either source, or claims a client result. For the
named-LAMBDA case, it requires FormulaFence's exact `ScenarioValue`
`defined_name_changed` record and matching `FF008`; WCAB independently proves
the LAMBDA body text, one-name workbook boundary, stable inputs/formulas, and
workbook-XML-only change. Neither layer evaluates the function or claims a
client result. For the Table calculated-column case, it requires FormulaFence's
exact redacted `table_definition_changed` profile, its
`calculated_column_formula_material_changed` flag, and `FF013`. FormulaFence
does not expose the formula-master text; WCAB independently proves the exact
`A2*B2` to `A2*(B2+1)` Table declaration, stable row/dashboard formulas, local
relationship binding, and Table-part-only boundary. Neither layer fills a
column, evaluates a structured reference, or claims a client result. For the
Power Pivot/Data Model relationship case, it requires FormulaFence's exact
redacted `power_pivot_data_model_changed` profile and high-severity `FF033`.
The profile retains one internal data part, one workbook binding, one
declaration, two model tables, one model relationship, and no coverage gaps;
only FormulaFence's declaration-change flag may differ. FormulaFence does not
expose model names, relationship keys, DAX, targets, or payload bytes. WCAB's
raw validator independently proves the exact `toColumn` transition, fixed
opaque payload, and workbook-XML-only boundary. Neither layer deserializes a
model, evaluates DAX, refreshes it, renders a report, infers model-to-cell
impact, or claims client behavior. For the
XLM Auto_Open binding case, it requires FormulaFence's exact high-severity
`xlm_automatic_macro_bindings_changed` profile and `FF076`: one automatic
binding and one `Auto_Open` binding remain present, with no
close/activate/deactivate binding, while only its material-change flag differs.
FormulaFence deliberately does not expose macro-sheet names, target cells, XML,
payload bytes, or package-member boundaries. WCAB's raw validator independently
proves the exact `$A$1`-to-`$A$2` stored target transition, fixed byte-identical
two-`HALT()` macro sheet, very-hidden state, fixed relationship/content types,
and workbook-XML-only difference. Neither layer opens Excel, enables or
executes XLM code, parses or emulates macro instructions, resolves a dynamic
name, inspects macro-security or trust state, infers dispatch behavior,
calculates a workbook, or claims client behavior. For the
sheet-protection sort case, it requires FormulaFence's exact
`sheet_protection_changed` profile and high-severity `FF022`: one protected
`Controls` worksheet retains every locked action except `sort`. WCAB's raw
validator independently establishes the `1`-to-`0` stored transition, stable
formula context, and worksheet-only package boundary. Neither layer tests a
password, encryption, authorization, a client sort operation, or a resulting
value.
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
For the list data-validation source case, it requires FormulaFence's exact
`data_validation_changed` source/control transition and `FF020`. FormulaFence
reports the stored source and entry-control metadata but does not evaluate the
source expression or decide whether a future value is permitted. WCAB's raw
validator independently proves the generated `Lists!$A$2:$A$4` to
`Lists!$B$2:$B$4` `formula1` change, stable source lists/current input/formulas,
direct static model/dashboard lower bounds, and Inputs-worksheet-only package
change. Neither layer accepts or rejects an entry, calculates a workbook, or
claims client behavior.
For the conditional-formatting threshold case, it requires FormulaFence's
exact one-rule `conditional_formatting_changed` transition and `FF021`.
FormulaFence exposes stored rule metadata but does not render a workbook or
determine which cells receive a format. WCAB's raw validator independently
proves the raw `100`-to-`50` threshold transition, stable priority, operator,
differential fill, metric values, calculation properties, and
Operations-worksheet-only package change. Neither layer evaluates the rule,
calculates a workbook, or claims client behavior.
For the custom number-format case, it requires FormulaFence's exact
`number_format_controls_changed` structural transition and `FF039`.
FormulaFence deliberately redacts format codes and cell targets from shared
reports, so WCAB's raw validator independently proves the declared
`0.0%;[Red](0.0%);-`-to-`;;;` code transition, stable target/style/value and
formula context, and styles-only package change. Neither layer renders a
format, calculates a workbook, or claims client behavior.
For the stored ignored-error case, it requires FormulaFence's exact redacted
`ignored_error_controls_changed` transition and `FF037`. FormulaFence exposes
only aggregate warning-category counts, so WCAB's raw validator independently
proves the generated `Operations!B5` target, `formulaRange=1` flag, stable
formula context, and worksheet-only package change. Neither layer determines
whether Excel would show a warning, evaluates a formula, renders an indicator,
or claims client behavior.
For the workbook-structure case, it requires FormulaFence's exact
`workbook_protection_changed` transition and `FF022`. FormulaFence reports the
non-secret structure-lock state while omitting verifier material; WCAB's raw
validator independently proves `lockStructure=1` to `0`, stable hidden-sheet
and formula context, and the workbook-XML-only package change. Neither layer
tests a password, encryption, authorization, or Excel-client behavior.
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
For the saved Named Sheet View case, it requires FormulaFence
`named_sheet_views_changed` evidence and `FF038` with one worksheet, part,
view, filter, column, and criterion; no sort rule or condition; no
unrecognized declaration; and only its material-definition flag. FormulaFence
redacts the view name, IDs, bound range, and selected value. The WCAB raw
validator independently proves the `North`-to-`South` criterion, the stable
base-AutoFilter binding and formulas, and the Named-Sheet-View-part-only
package change. Neither layer activates, renders, or applies the view,
calculates a subtotal, or infers visible rows.
For the XML Map case, it requires FormulaFence
`xml_mapping_controls_changed` evidence and `FF049` with one map part,
schema, map, data binding, file binding, table binding, and sheet-level
single-cell binding; no unrecognized mapping metadata; and only its binding
material flag. FormulaFence deliberately redacts schema, map, XPath, table,
and cell values. WCAB's raw validator independently proves the synthetic
`NetAmount`-to-`TaxAmount` XPath transition, stable map/schema/bindings and
formulas, and the table-part-only package change. Neither layer accesses a
file, imports or exports XML, materializes data, or infers a result.
For the Office Web Add-in case, it requires FormulaFence
`office_web_addins_changed` evidence and `FF028` with one declared task-pane
part, task pane, web-extension part, and local store reference; one locked,
hidden task pane; no bindings, snapshots, external relationships, in-content
references, or unrecognized parts; and an auto-show count moving from zero to
one. FormulaFence deliberately redacts the add-in IDs, store name, and property
value. WCAB's raw validator independently proves the generated local
reference, false-to-true stored property transition, stable workbook context,
and web-extension-part-only package change. Neither layer installs, loads,
executes, or fetches an add-in or manifest, or claims that a task pane opens.
For the embedded OLE auto-load case, it requires FormulaFence
`worksheet_embedded_controls_changed` evidence and `FF029` with one
control-bearing worksheet, one internal fingerprinted payload, one OLE object,
no ActiveX, VML, linked object, external relationship, or unrecognized part,
and an auto-load count moving from zero to one. FormulaFence deliberately
redacts the ProgID, relationship target, content type, and bytes. WCAB's raw
validator independently proves the fixed local relationship, inert synthetic
payload, false-to-true stored control, stable formula context, and
worksheet-only package change. Neither layer deserializes, opens, renders,
executes, registers, or invokes an object server, or claims that an object
loads successfully.
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
