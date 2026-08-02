# WCAB truth schema, version 3

Every case has one `truth.json`. It is public metadata for an original,
generated fixture; it never relies on a private workbook or an external data
source.

Version 3 adds `coverage_expectations`: machine-matchable disclosures for
important analysis boundaries. WCAB 0.5 extends that same stable envelope with
cases where an unchanged `INDIRECT` or `OFFSET` formula receives a changed
selector, WCAB 0.6 adds a relationship-backed external-data refresh-on-open
fact, WCAB 0.7 adds an unchanged anchor formula that moves from fixed
legacy-CSE to dynamic-array semantics, and WCAB 0.8 adds an unchanged
external-workbook formula whose stored open-time link-update policy changes.
WCAB 0.9 adds an unchanged direct circular formula whose stored iterative
calculation setting changes, WCAB 0.10 adds a precision-as-displayed
calculation control change without a cell edit, and WCAB 0.11 adds a raw saved
formula-result change without a formula or input edit. WCAB 0.12 adds a raw
workbook serial-date-system control change with unchanged cell content, WCAB
0.13 adds an active worksheet AutoFilter criterion change with unchanged cell
content and formulas, and WCAB 0.14 adds a relationship-backed PivotTable cache
refresh-on-open control change with unchanged stored cells. WCAB 0.15 adds a
relationship-backed DrawingML chart-series source-reference change with
unchanged worksheet cells. WCAB 0.16 adds a relationship-backed PivotTable
value-field aggregation change with unchanged source/cache/report cells.
WCAB 0.17 adds a relationship-backed PivotTable Slicer-cache selection change
with unchanged source/cache/report cells. WCAB 0.18 adds a connection-only
Power Query M filter change over an unchanged local Excel Table. WCAB 0.19
adds a stored Scenario Manager alternate-input change with unchanged visible
worksheet cells, formulas, and calculation properties. WCAB 0.20 adds a
one-variable What-If Data Table input-reference change with unchanged visible
cells, ordinary formulas, calculation properties, and saved table results.
WCAB 0.21 adds a list data-validation source change with unchanged target,
entry-control metadata, source-list values, visible input, ordinary formulas,
and calculation properties. WCAB 0.22 adds a conditional-formatting threshold
change with unchanged target range, priority, operator, differential fill,
metric values, and calculation properties.
Version 2 remains
available in the immutable v0.2.0 and v0.3.0 releases.

```json
{
  "schema_version": 3,
  "id": "family.case_name",
  "title": "Human-readable scenario",
  "family": "finance",
  "topology": "pair",
  "review_expectation": "block",
  "facts": [],
  "must_reach": [],
  "coverage": [],
  "coverage_expectations": []
}
```

`topology` is either `pair` (`baseline.xlsx` and `candidate.xlsx`) or
`portfolio` (`baseline/` and `candidate/` directories). `review_expectation`
is one of `allow`, `review`, or `block`. It is a transparent benchmark
convention, not an assertion that every organization must use the same policy.

## Facts

Facts are observed directly from the fixture files by `wcab validate`.

| Fact kind | Required fields | Observable contract |
| --- | --- | --- |
| `formula_to_value` | `sheet`, `cell` | A formula cell becomes a literal value. |
| `formula_changed` | `sheet`, `cell` | Formula text exists on both sides and differs. |
| `value_changed` | `sheet`, `cell` | Literal values exist on both sides and differ. |
| `external_formula_added` | `sheet`, `cell` | Candidate formula contains an external-workbook reference. The target is never opened. |
| `defined_name_changed` | `name` | A defined name's stored destination differs. |
| `data_validation_count_changed` | `sheet`, `baseline_count`, `candidate_count` | Worksheet data-validation count differs as declared. |
| `data_validation_list_source_changed` | `validation_sheet`, `target_range`, `validation_type`, `baseline_source_formula`, `candidate_source_formula`, `allow_blank`, `dropdown_hidden`, `show_error_message`, `error_style`, `error_title`, `error`, `show_input_message`, `prompt_title`, `prompt`, `source_sheet`, `baseline_source_range`, `candidate_source_range`, `baseline_source_values`, `candidate_source_values`, `input_cell`, `input_value`, `model_sheet`, `model_cell`, `model_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One stored list validation retains its target, rule metadata, both source-list values, current input, ordinary formulas, calculation properties, and every package member except its validation-bearing worksheet while raw `formula1` moves between declared local source ranges. The validator does not evaluate the source, decide a future input's validity, accept/reject an entry, calculate a result, or claim client behavior. |
| `conditional_formatting_count_changed` | `sheet`, `baseline_count`, `candidate_count` | Conditional-formatting range count differs as declared. |
| `conditional_formatting_threshold_changed` | `sheet`, `target_range`, `priority`, `rule_type`, `operator`, `baseline_formula`, `candidate_formula`, `metric_values`, `fill_rgb` | One stored `cellIs` conditional-formatting rule retains its target range, priority, operator, differential fill, worksheet values, calculation properties, and every package member except its worksheet while raw `formula` moves between declared thresholds. The validator does not evaluate the rule, determine which cells a client formats, calculate a workbook, or claim client behavior. |
| `auto_filter_criteria_changed` | `sheet`, `filter_ref`, `filter_column_id`, `baseline_filter_value`, `candidate_filter_value`, `subtotal_cell`, `subtotal_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One raw worksheet AutoFilter list criterion changes while its filter shell, formulas, direct dependency edge, and every package member except the report worksheet remain unchanged. The validator does not apply the filter or calculate a result. |
| `sheet_visibility_changed` | `sheet`, `baseline_state`, `candidate_state` | The stored sheet state changes. |
| `formula_cell_unlocked` | `sheet`, `cell` | A formula cell is explicitly unlocked while its sheet remains protected. |
| `manual_calculation_incomplete` | none | Candidate calculation metadata is `manual` and records incomplete calculation. |
| `iterative_calculation_enabled` | `sheet`, `cell`, `formula`, `baseline_iterate`, `candidate_iterate`, `iteration_count`, `iteration_delta` | The declared direct self-referencing formula remains unchanged while raw `calcPr/@iterate` changes exactly from `false` to `true`; the explicit count and delta remain the declared values, and all non-iteration calculation attributes are unchanged. The validator does not calculate the model. |
| `precision_as_displayed_enabled` | `input_sheet`, `input_cell`, `input_value`, `number_format`, `formula_sheet`, `formula_cell`, `formula`, `baseline_full_precision`, `candidate_full_precision` | The declared stored input, number format, and formula remain unchanged while raw `calcPr/@fullPrecision` changes exactly from `true` to `false`; all other `calcPr` attributes are unchanged. The validator does not calculate, open, or save the workbook. |
| `workbook_date_system_changed` | `baseline_date_1904`, `candidate_date_1904`, `date_compatibility`, `serial_sheet`, `serial_cell`, `serial_value`, `number_format`, `formula_sheet`, `formula_cell`, `formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | Raw `workbookPr/@date1904` changes from `false` to `true` while explicit compatibility remains `true`; the raw numeric serial, custom date format, local formulas, and every package member except `xl/workbook.xml` remain unchanged. The validator does not calculate, convert, or infer displayed dates. |
| `formula_cached_result_changed` | `sheet`, `cell`, `formula`, `input_sheet`, `input_cell`, `input_value`, `result_type`, `baseline_cached_result`, `candidate_cached_result` | One raw numeric formula-cell `<v>` value changes while its raw `<f>` expression, direct input, calculation properties, and every other package member remain unchanged. The validator reads OOXML only; it does not calculate, validate, or interpret the saved result. |
| `external_data_connection_refresh_on_load_changed` | `connection_id`, `baseline_refresh_on_load`, `candidate_refresh_on_load` | The relationship-backed connection with this workbook-local ID explicitly changes `refreshOnLoad` from `false` to `true`. The validator reads raw OOXML only. |
| `pivot_cache_refresh_on_load_changed` | `cache_id`, `source_type`, `source_sheet`, `source_ref`, `pivot_sheet`, `pivot_ref`, `pivot_output_cell`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula`, `baseline_refresh_on_load`, `candidate_refresh_on_load` | One relationship-bound local worksheet PivotCache changes raw `refreshOnLoad` from `false` to `true`; its source binding, PivotTable location, stored report/dashboard cells, calculation properties, and every package member except its cache definition remain unchanged. The validator neither refreshes nor renders a PivotTable. |
| `pivot_data_field_aggregation_changed` | `cache_id`, `source_type`, `source_sheet`, `source_ref`, `pivot_sheet`, `pivot_ref`, `pivot_output_cell`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula`, `data_field_source_index`, `baseline_subtotal`, `candidate_subtotal` | One relationship-bound local worksheet PivotTable retains its source/cache bindings, stored report/dashboard cells, refresh control, and calculation properties while raw `dataFields/dataField/@subtotal` moves between the declared aggregate functions. Every package member except its PivotTable definition remains unchanged. The validator does not refresh, calculate, render, or infer a displayed result. |
| `pivot_slicer_selection_changed` | `cache_id`, `source_type`, `source_sheet`, `source_ref`, `pivot_sheet`, `pivot_ref`, `pivot_output_cell`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula`, `slicer_name`, `slicer_source_name`, `slicer_pivot_table_name`, `slicer_pivot_tab_id`, `item_count`, `baseline_selected_item_index`, `candidate_selected_item_index`, `baseline_selected_value`, `candidate_selected_value` | One relationship-bound local Slicer cache retains its source/PivotCache/PivotTable bindings, stored report/dashboard cells, refresh control, and calculation properties while exactly one selected cache item moves between the declared index/value pairs. Every package member except its Slicer-cache definition remains unchanged. The validator does not create a Slicer drawing, apply a filter, refresh, calculate, render, or infer a displayed result. |
| `power_query_m_filter_changed` | `data_mashup_part`, `source_sheet`, `source_table`, `source_ref`, `query_section`, `query_name`, `filter_column`, `baseline_filter_value`, `candidate_filter_value`, `fill_enabled`, `firewall_enabled`, `future_packages_allowed` | One package-root relationship-bound compact Data Mashup retains its local Excel Table source, metadata, permission controls, calculation properties, and every package member except its custom-XML part while one stored M `Table.SelectRows` literal changes. The validator reads a bounded generated envelope only; it does not execute M, refresh, materialize output, calculate, or infer query results. |
| `scenario_manager_stored_input_value_changed` | `scenario_sheet`, `scenario_name`, `changing_cell`, `stable_input_cell`, `baseline_stored_value`, `candidate_stored_value`, `stable_stored_value`, `input_number_format_id`, `summary_ref`, `worksheet_input_value`, `worksheet_stable_input_value`, `result_cell`, `result_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One raw worksheet Scenario Manager declaration retains its selected/locked scenario metadata, second stored input, summary reference, visible worksheet cells, formula path, calculation properties, and every package member except its scenario-bearing worksheet while one stored `inputCells/@val` value changes. The validator does not show/apply a scenario, calculate a result, create a scenario summary, or infer client behavior. |
| `what_if_data_table_input_reference_changed` | `table_sheet`, `master_cell`, `output_range`, `baseline_input_cell`, `candidate_input_cell`, `orientation`, `recalculation_requested`, `input_value_range`, `input_values`, `primary_input_value`, `alternate_input_value`, `scale_cell`, `scale_value`, `output_formula_cell`, `output_formula`, `model_sheet`, `model_cell`, `model_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One raw column-oriented one-variable Data Table master retains its declared output range, recalculation request, visible input grid, ordinary formulas, calculation properties, and every package member except its table-bearing worksheet while `f/@r1` moves between declared local input cells. The validator does not substitute inputs, calculate, infer table results, resolve a circular dependency, or claim client behavior. |
| `chart_series_value_reference_changed` | `chart_sheet`, `chart_anchor`, `source_sheet`, `series_title_ref`, `category_ref`, `baseline_value_ref`, `candidate_value_ref` | One relationship-bound DrawingML chart retains its host, anchor, title/category references, source worksheet cells, and every package member except its chart part while raw `c:ser/c:val/c:numRef/c:f` moves between declared local value ranges. The validator does not calculate, refresh, or render a chart. |
| `external_workbook_link_update_policy_changed` | `sheet`, `cell`, `formula`, `baseline_update_links`, `candidate_update_links` | The declared external-workbook formula remains unchanged while raw `workbookPr/@updateLinks` changes exactly from `never` to `always`; all other stored `workbookPr` attributes are unchanged. The validator does not resolve the source workbook. |
| `array_formula_mode_changed` | `sheet`, `cell`, `formula`, `baseline_mode`, `candidate_mode`, `baseline_output_range`, `candidate_output_range` | The declared unchanged array anchor moves from `legacy_cse` to `dynamic`, with its stored formula text and output range exactly as declared. The validator reads the raw OOXML cell-metadata binding. |
| `static_cycle_introduced` | `cells` | Every declared direct A1 cell reaches itself in the local static dependency graph. |
| `three_d_scope_changed` | `formula_sheet`, `formula_cell`, `inserted_sheet`, `after_sheet`, `before_sheet` | Formula text remains unchanged while a sheet is inserted inside the declared tab span. |
| `structural_formula_rewrite` | `baseline`, `candidate` | Declared before/after formula locations and text exist. This is an annotation, not a general proof of equivalence. |
| `portfolio_value_changed` | `workbook`, `sheet`, `cell` | A literal value differs between paired portfolio members. |
| `portfolio_external_reference` | `workbook`, `sheet`, `cell`, `target_workbook` | The local portfolio model contains the declared external workbook reference. |
| `structured_table_scope_changed` | `table_sheet`, `table`, `baseline_ref`, `candidate_ref`, `formula_sheet`, `formula_cell` | A stored Excel Table range changes while the declared formula remains textually unchanged and retains a reference to that table. |
| `dynamic_formula_reference_added` | `sheet`, `cell`, `functions` | A previously direct formula changes to one containing the declared introduced dynamic-reference functions. |

## List data-validation source

Microsoft's [Excel data-validation API guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-data-validation)
describes a validation rule's `formula1` and `formula2` fields, including a
list source. The [MS-XLSX data-validation formula
specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/b71087ce-6f73-461b-b23e-fcd4ece396aa)
likewise identifies `formula1` as the stored validation formula. That makes a
list-source retarget review-material even when a current input and ordinary
formulas remain unchanged.

WCAB's pair has one list validation on `Inputs!B2`. Its baseline raw
`formula1` is `=Lists!$A$2:$A$4`; its candidate is
`=Lists!$B$2:$B$4`. The two lists, `Inputs!B2="Draft"`, all other validation
attributes, `Model!B2=Inputs!$B$2`, and `Dashboard!B4=Model!$B$2` remain
unchanged. The validator accepts only this one-container/one-rule shape,
compares the Inputs worksheet after erasing its sole `formula1` text, and
requires that worksheet to be the only changed package member. It does not
evaluate either source formula, test a proposed entry, accept/reject an input,
calculate a result, or claim Excel-client behavior. The direct input-to-model
paths are ordinary static lower bounds only if a user later enters a value.

## Conditional-formatting threshold

Microsoft's [Open XML conditional-formatting guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-conditional-formatting)
shows that a worksheet `cfRule` can hold a `cellIs` operator and a `formula`
threshold. WCAB's pair has one `Operations!B2:B4` rule whose baseline raw
formula is `100` and candidate formula is `50`. Its priority, `greaterThan`
operator, differential red fill, and stored `10`, `75`, and `120` metrics stay
unchanged. The validator accepts only the generated one-control/one-rule shape,
compares the Operations worksheet after erasing the threshold formula, and
requires that worksheet to be the only changed package member. It does not
evaluate a rule, infer a rendered format, calculate a workbook, or claim
Excel-client behavior.

## Static impact lower bounds

`must_reach` entries have one `source` plus one or more `targets`, each with a
sheet and cell. The bundled validator uses only direct, ordinary local A1
references and walks the resulting graph. The targets are therefore lower
bounds: no result implies a complete Excel dependency calculation, a formula
evaluation, dynamic-reference resolution, or a claim about cached values.

## External-data refresh controls

Excel documents that a connection can [refresh external data when the workbook
opens](https://support.microsoft.com/en-us/excel/connection-properties), which
can update data outside a normal worksheet formula edit. WCAB's connection case
uses a synthetic, non-routable endpoint and validates the stored
`refreshOnLoad` attribute through its relationship-backed `connections.xml`
part. It does not open a connection, test credentials or Trust Center policy,
retrieve data, calculate formulas, or infer a downstream result. Connection
paths, URLs, names, commands, and credentials are deliberately not truth
fields.

## PivotTable cache refresh-on-open

Microsoft documents a PivotTable option to [refresh data when the workbook
opens](https://support.microsoft.com/en-us/excel/refresh-pivottable-data), and
the Open XML [`pivotCacheDefinition` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivotcachedefinition?view=openxml-3.0.1)
defines `refreshOnLoad` as the cache-level stored control. A cache can be shared
by PivotTables, so a change to this request is review-material even when no
worksheet cell or formula text changes.

WCAB's original fixture has one local worksheet cache: `Source!A1:B5` binds to
a PivotTable at `Report!A1:B2`, whose stored `B2` display cell is directly
referenced by `Dashboard!B4`. The validator follows only that local OOXML
relationship graph, verifies the cache definition changes solely at
`refreshOnLoad`, and checks the stored cells and direct dashboard formula stay
fixed. It does not open a workbook in Excel, refresh a cache, calculate or
render the PivotTable, infer a visible result, or claim whether any client will
honor the request.

## PivotTable value aggregation

Excel's [PivotTable layout guidance](https://support.microsoft.com/en-US/Excel/design-the-layout-and-format-of-a-pivottable)
documents placing fields in the Values area and changing their settings. In
SpreadsheetML, [`dataField/@subtotal`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datafield.subtotal?view=openxml-3.0.1)
stores the value field's data-consolidate function. That declaration is
review-material even if neither source cells nor stored report cells change.

WCAB's aggregation fixture retains the same one-local-cache graph:
`Source!A1:B5` binds to a PivotTable at `Report!A1:B2`, whose stored `B2`
display cell is directly referenced by `Dashboard!B4`. Baseline and candidate
retain every cache record, source binding, relationship, cell, formula, and
calculation property; the only raw difference is its sole
`dataFields/dataField/@subtotal` from `sum` to `average`. The validator follows
that local graph and compares the PivotTable XML after removing exactly that
attribute. It does not open Excel, refresh, calculate, render, infer a changed
displayed value, or claim client behavior.

## PivotTable Slicer selection

Microsoft's [PivotTable filtering guidance](https://support.microsoft.com/en-us/excel/get-started/filter-data-in-a-pivottable)
describes Slicers as controls that filter a PivotTable and convey its filtering
state. The Office Open XML [Slicer Cache Part
specification](https://learn.microsoft.com/en-us/openspecs/office_standards/MS-XLSX/e7eda20c-c65e-45ed-9540-de59c4a07b7d)
stores item indices as `x` and selected items as `s=1`, so this state can be
material even when no worksheet cell changes.

WCAB's Slicer fixture retains the same local worksheet cache graph:
`Source!A1:B5` binds to a PivotTable at `Report!A1:B2`, whose stored `B2`
display cell is directly referenced by `Dashboard!B4`. The local Slicer cache
targets that PivotTable and cache, has two `Region` items, and changes the one
selected item from index 0 (`North`) to index 1 (`South`). The validator follows
the local package relationships, maps selected item indices through the cache's
shared items, and compares the Slicer XML after removing only `s` attributes.
It creates no visual Slicer or drawing and does not apply the filter, refresh,
calculate, render, infer a displayed value, or claim client behavior.

## Power Query M local-table filter

Microsoft's [Power Query overview](https://support.microsoft.com/en-us/excel/about-power-query-in-excel)
documents recorded transformations and their M source, while its
[query-management guidance](https://support.microsoft.com/en-us/excel/manage-queries-power-query)
allows queries to load to the Data Model or remain connection-only. The
[Power Query M reference](https://learn.microsoft.com/en-us/powerquery-m/)
defines the language, and [`Excel.CurrentWorkbook`](https://learn.microsoft.com/en-us/powerquery-m/excel-currentworkbook)
returns local workbook Tables, named ranges, and dynamic arrays. A stored M
definition can therefore be review-material even when ordinary worksheet cells
and saved outputs do not change.

WCAB's pair has one generated `Source!A1:B5` Table named `SourceData` and one
connection-only Data Mashup carried by a package-root `customXml` relationship.
The only semantic change is the M `Table.SelectRows` literal from `North` to
`South`; `FillEnabled=false`, a local-table source, metadata, permission
controls, and all ordinary package parts remain fixed. The validator follows
only the generated relationship and bounded compact envelope: one nested
three-part package, one formula document, one metadata item, and explicit
firewall/future-package controls. It does not claim to be a general M parser,
execute M, apply the filter, refresh a query, materialize output, calculate a
workbook, infer returned rows, or predict any client behavior.

## Scenario Manager stored alternate input

Microsoft's [Scenario Manager guidance](https://support.microsoft.com/en-us/excel/switch-between-various-sets-of-values-by-using-scenarios)
describes scenarios as saved value sets for declared changing cells, with
formula result cells and optional scenario protection. Its Open XML
[`inputCells` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.inputcells?view=openxml-3.0.1)
records each stored input's cell reference, value, deletion state, and
display-only number-format ID. That creates a review surface outside ordinary
visible worksheet-cell values.

WCAB's pair stores one selected, locked `WCAB downside` scenario in
`Inputs` with two raw input records. It preserves the visible `Inputs!B2=0.1`,
`Inputs!B3=125`, `Inputs!D2=B2*B3`, and
`Dashboard!B4=Inputs!$D$2` formula path while only the alternate `B2` stored
value moves from `0.08` to `0.16`. The raw validator accepts only this compact
one-scenario/two-input declaration, checks the scenario selection, protection,
comment/user, summary-reference, second-input, and number-format metadata,
compares the worksheet after erasing the sole mutable `@val`, and requires that
the Inputs worksheet part is the only changed package member. It does not show
or apply a scenario, calculate a workbook, generate a Scenario Summary, infer
an output, or claim a particular client will act on the declaration.

## What-If Data Table input reference

Microsoft's [Data Table guidance](https://support.microsoft.com/en-us/excel/calculate-multiple-results-by-using-a-data-table)
describes one- and two-variable What-If Data Tables and their row/column input
cells. The Open XML
[`CellFormula` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1)
describes a Data Table master with `t="dataTable"`, output `ref`, input
references, and orientation flags. A stored input reference can therefore be
material even when ordinary formula text and visible cells are unchanged.

WCAB's pair has one raw `Sensitivity!D3` master with `ref="D3:D5"`, `ca="1"`,
and a column-oriented one-variable `r1` reference. The baseline declares
`r1="B2"`; the candidate declares `r1="B3"`. Its `C3:C5` input grid,
`Sensitivity!D2=Model!$B$2`, both possible input cells,
`Model!B2=Sensitivity!$B$2*Sensitivity!$B$3*Sensitivity!$B$4`, and
`Dashboard!B4=Model!$B$2` remain unchanged. The raw validator accepts only one
master with the fixed generated attributes, erases only `r1` before comparing
worksheet XML, and requires that worksheet to be the sole changed package
member. It does not substitute temporary Data Table inputs, calculate a model
or result, infer a table output, resolve a circular dependency, or claim
Excel-client behavior. The declared ordinary static paths from both possible
inputs are lower bounds, not an evaluation claim.

## External-workbook link update policy

Open XML defines [`workbookPr/@updateLinks`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1)
as the behavior for updating external links when the workbook opens. This is a
workbook-wide control, distinct from the connection-level `refreshOnLoad`
attribute described above. Microsoft also documents the user-facing external
workbook-link startup choices in its
[workbook-link guidance](https://support.microsoft.com/en-us/excel/manage-workbook-links).

WCAB's policy case keeps `LinkedModel!B2` and its local `Dashboard!B4`
consumer unchanged while changing the explicit stored policy from `never` to
`always`. `WCABSource.xlsx` is synthetic and absent by design. The raw validator
requires the exact policy values, matching formula text, and equal non-policy
`workbookPr` attributes. It does not open, resolve, authenticate to, trust,
refresh, or calculate the external source, and it does not claim an updated
cached result or a successful workbook recalculation.

## Workbook serial-date system

Excel documents two workbook date systems, 1900 and 1904, with a 1,462-day
difference for the same stored serial; it also documents the workbook-level
setting in its [date-system guidance](https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487).
In SpreadsheetML, the `date1904` and `dateCompatibility` controls live on
[`workbookPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1): their Open XML defaults are respectively `false` and `true`.

WCAB's date-system pair makes both controls explicit. It preserves the raw
numeric `Inputs!B2` serial `45292`, its custom `yyyy-mm-dd` number format,
`Model!B2`'s local formula, and `Dashboard!B4`'s local consumer while only
`date1904` changes from `false` to `true`; `dateCompatibility` stays `true`.
The validator resolves the raw cell style through `styles.xml`, compares the
formula expressions and direct dependency edges, and compares `workbook.xml`
after removing only those two date-control attributes. It does not execute a
formula, convert the serial, predict a displayed calendar date, open or save a
workbook, or claim behavior for a particular client.

## Active AutoFilter criteria

Microsoft's [filter guidance](https://support.microsoft.com/en-us/excel/get-started/filter-data-in-a-range-or-table-in-excel)
documents that filters show rows meeting the selected criteria and hide the
rest. Its [`SUBTOTAL` reference](https://support.microsoft.com/en-us/excel/functions/subtotal-function)
states that rows excluded by a filter are always excluded, so a stored
criterion can be review-material even when formulas stay textually stable.

WCAB's active-filter pair keeps the `Report!A1:B5` AutoFilter shell, its
column ID, `Report!D2=SUBTOTAL(109,B2:B5)`, and
`Dashboard!B4=Report!$D$2` unchanged while the sole column-0 list value moves
from `North` to `South`. The validator resolves the worksheet through its
workbook relationship, checks one raw `<autoFilter>`, `<filterColumn>`,
`<filters>`, and `<filter>` declaration on each side, compares the worksheet
after removing that sole criterion container, checks the formula texts and
direct dependency edge, and requires the report worksheet to be the only
changed package member. It does not execute Excel's filter logic, calculate
the subtotal, infer a visible row set, or claim a display, copy, chart, or
print result.

## DrawingML chart-series source references

Microsoft's [series-editing guidance](https://support.microsoft.com/en-US/Excel/rename-a-data-series)
allows a user to change a chart's series values without changing the worksheet
data. The Open XML [`NumberReference` definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberreference?view=openxml-3.0.1)
contains the stored formula reference used by a numeric chart series. That
creates a report-review surface outside ordinary cell diffs.

WCAB's chart pair starts with one `Dashboard!D2` chart bound through the
worksheet drawing relationship to one DrawingML chart part. Its title and
category references remain `Source!B1` and `Source!$A$2:$A$4`; only the one
numeric series value reference moves from `Source!$B$2:$B$4` to
`Source!$C$2:$C$4`. Both source columns, all worksheet cells, the drawing
anchor, calculation properties, and every package member outside
`xl/charts/chart1.xml` remain fixed. The validator follows only that generated
worksheet-to-drawing-to-chart relationship chain, compares the chart XML after
erasing the one declared reference, and never evaluates a chart formula,
calculates cells, refreshes data, renders a chart, infers a visual difference,
or claims client behavior.

## Iterative calculation

Open XML stores calculation controls in
[`calcPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties?view=openxml-3.0.1),
including `iterate`, `iterateCount`, and `iterateDelta`. Microsoft documents
that iterative calculation can intentionally allow circular references and
that its maximum-iteration and maximum-change settings bound the repeated
calculation process in its
[circular-reference guidance](https://support.microsoft.com/en-US/Excel/remove-or-allow-a-circular-reference-in-excel).

WCAB's iterative case keeps `Model!B2`'s direct self-reference and its local
`Dashboard!B4` consumer unchanged. The raw validator requires explicit
baseline `iterate=false`, candidate `iterate=true`, shared declared count and
delta, equal non-iteration `calcPr` attributes, and the unchanged direct
self-reference. It does not calculate either workbook, establish convergence,
count recalculations, predict a terminal value, or assert Excel-client
compatibility.

## Precision as displayed

Excel normally calculates from stored values rather than the values a number
format displays. Microsoft documents that its
[precision-as-displayed setting](https://support.microsoft.com/en-US/Excel/change-formula-recalculation-iteration-or-precision-in-excel)
uses displayed values for calculation and permanently changes stored values;
its [rounding-precision guidance](https://support.microsoft.com/en-us/excel/set-rounding-precision)
warns that the control can have cumulative calculation effects. Open XML stores
the corresponding `fullPrecision` control on
[`calcPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties?view=openxml-3.0.1).

WCAB's precision case preserves an explicitly stored `Inputs!B2=10.005`, its
`0.00` number format, `Model!B2`'s `=Inputs!$B$2*2` formula, and the local
`Dashboard!B4` consumer. The raw validator requires explicit baseline
`fullPrecision=true`, candidate `fullPrecision=false`, and equal remaining
calculation attributes. It never opens or saves either workbook, executes a
formula, alters the stored 10.005 value, predicts a rounded value, or claims
that a particular Excel client will apply or persist the setting.

## Saved formula results

SpreadsheetML stores a formula expression in [`<f>`](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas)
and its last calculated result in the neighboring `<v>` element. The Open XML
[`CellValue` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-3.0.1)
likewise describes formula-cell `<v>` content as the last calculated result.

WCAB's saved-result case keeps `Inputs!B2=10`, `Model!B2`'s
`=Inputs!$B$2*2` expression, calculation properties, and the local
`Dashboard!B4` consumer unchanged. It changes only the raw numeric `<v>` next
to `Model!B2`'s `<f>` from `20` to `25`. The validator resolves the worksheet
through the workbook relationship, checks the raw expression and numeric
result text, and proves that the selected worksheet is identical once that
one result text is erased. It does not execute the formula, determine whether
either saved value is current or correct, infer volatile or external inputs,
claim that a client will display either value, or assert a downstream result.

## Array-formula mode

Microsoft distinguishes dynamic arrays from legacy Ctrl+Shift+Enter (CSE)
arrays: dynamic arrays can resize their spill output, while CSE arrays have a
fixed output range. Its [dynamic-array versus legacy-CSE guidance](https://support.microsoft.com/en-US/Excel/dynamic-array-formulas-vs-legacy-cse-array-formulas)
also notes that the same formula can be moved from an entire CSE range to one
dynamic anchor cell.

WCAB's array-mode case consequently preserves `=LEN(Inputs!A1:A3)` and its
currently stored `B1:B3` output range. The candidate adds only the
relationship-backed metadata binding that marks `Model!B1` as dynamic. The
validator checks the anchor's raw array formula, `cm` cell-metadata index, and
`XLDAPR` `fDynamic=true` record; it does not execute the formula, predict a
future spill extent, locate a blocker, or claim Excel-client compatibility.

## Scoreable coverage expectations

`coverage_expectations` is an array of exact, validated objects. It does not
ask a tool to claim universal support for an Excel construct. Instead, it asks
whether an adapter can point to native evidence that the relevant analysis
boundary was visible.

| Expectation kind | Required fields | Observable contract |
| --- | --- | --- |
| `dynamic_reference_static_coverage` | `sheet`, `cell`, `functions` | Candidate formula contains the declared newly introduced `INDIRECT` or `OFFSET` function while the baseline formula has none. A matching observation declaration means the tool surfaced the static-dependency coverage boundary. |
| `dynamic_reference_driver_changed` | `driver` (`sheet`, `cell`), `formula` (`sheet`, `cell`), `functions` | A literal driver changes, the declared dynamic formula remains textually unchanged on both sides, and the candidate formula directly reads that driver. A matching observation declaration means the tool surfaced the pre-existing dynamic-reference boundary alongside the input change. |

Excel documents that [INDIRECT returns a reference specified by a text
string](https://support.microsoft.com/en-us/excel/functions/indirect-function)
and that [OFFSET returns a reference displaced from another
reference](https://support.microsoft.com/en-us/excel/functions/offset-function).
WCAB consequently does not treat a simple static graph as a complete statement
of all possible dependencies once either function is present. The driver cases
separate a formula-text diff from an effective-target change: `INDIRECT` reads
an address from text, while `OFFSET` reads a displacement, so each can change
which cell is selected even when its formula text is unchanged.

The observation protocol's `coverage.declarations` items wrap an exact
expectation plus optional native evidence. See
[docs/observations.md](observations.md) for matching and scoring rules.

## Coverage text

Each case also contains an explicit free-text `coverage` list. Consumers must
preserve these boundaries in reports and must not turn unsupported constructs
into silent passes. The machine-matchable expectations complement rather than
replace that prose: a case can document broader caveats than a score can fairly
reduce to one metric.

## JSONL case catalogue

`fixtures/manifest.jsonl` contains one deterministic JSON object per case. It
copies the truth fields (`id`, `title`, `family`, `topology`,
`review_expectation`, `facts`, `must_reach`, `coverage`, and
`coverage_expectations`) and adds:

| Field | Meaning |
| --- | --- |
| `case_path` | Case directory relative to the fixture root. |
| `baseline_files` | One or more baseline file records. |
| `candidate_files` | One or more candidate file records. |

Each file record has a root-relative `path`, exact `bytes`, and SHA-256
`sha256`. Pair cases have one file on each side; portfolio cases can have more.
The catalogue is generated by `wcab build` or `wcab manifest`; consumers may
use it to verify the exact bytes they evaluated. It does not add any formula
execution or semantic-equivalence claim beyond the underlying truth contract.
