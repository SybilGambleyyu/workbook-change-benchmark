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
metric values, and calculation properties. WCAB 0.23 adds a custom
number-format transition with unchanged target style index, raw numeric value,
formula context, calculation properties, and every package member except
`styles.xml`. WCAB 0.24 adds a worksheet-local ignored-error rule with
unchanged ordinary cells, formula context, calculation properties, and every
package member except its worksheet XML. WCAB 0.25 adds a raw workbook
structure-lock transition with unchanged hidden-sheet/formula context and every
package member except `xl/workbook.xml`. WCAB 0.26 adds a relationship-backed
Named Sheet View list-criterion transition with an unchanged base AutoFilter,
formulas, and every package member except its Named Sheet View part.
WCAB 0.27 adds a relationship-backed XML Map table-column XPath transition
with an unchanged map/schema/file-binding declaration, single-cell mapping,
cells, formulas, and every package member except its table part.
WCAB 0.28 adds a relationship-backed Office Web Add-in task-pane auto-show
property transition with an unchanged synthetic local reference, task-pane
layout, cells, formulas, and every package member except its web-extension
part.
WCAB 0.29 adds a worksheet embedded-OLE auto-load transition with an unchanged
internal relationship, inert synthetic bytes, cells, formulas, and every
package member except its worksheet part.
WCAB 0.30 adds a relationship-backed QueryTable refresh-on-open transition with
an unchanged internal connection, saved cells, formulas, and every package
member except its QueryTable part.
WCAB 0.31 adds a relationship-backed worksheet cell-hyperlink target transition
with unchanged visible text, formulas, and every package member except its
worksheet relationship part.
WCAB 0.32 adds a relationship-backed external-workbook source transition with
unchanged formula text, source-sheet declaration, local downstream formula,
and every package member except its externalLink relationship part.
WCAB 0.33 adds a local defined-name external-source transition with unchanged
formula cells and every package member except `xl/workbook.xml`, plus a
protected-sheet sort permission transition with unchanged cells, formulas,
styles, calculation properties, and every package member except its worksheet.
WCAB 0.34 adds a workbook-scoped named-LAMBDA body transition with unchanged
inputs and ordinary formulas, plus a Table calculated-column master transition
with unchanged row and dashboard formulas. Each isolates one raw formula
definition in its owning package part. WCAB 0.35 adds an embedded Power
Pivot/Data Model relationship-key transition with fixed workbook binding,
content type, local Tables, and opaque model payload. It isolates one raw
`x15:modelRelationship` declaration without claiming model execution. WCAB
0.36 adds a macro-enabled XLM automatic-macro binding transition with a fixed,
very-hidden macro sheet and fixed package shape; it records only the stored
`_xlnm.Auto_Open` dispatch declaration without claiming macro execution.
WCAB 0.37 adds a relationship-backed external-data web-query source transition
with unchanged connection controls, saved cells, formulas, and every package
member except `xl/connections.xml`; it records only the stored endpoint
declaration without contacting it. WCAB 0.38 adds a structurally shaped OPC
package-signature `Object/Manifest` direct-part transition with fixed
relationship graph and ordinary workbook context. WCAB 0.39 adds a
complementary root-relationships selector retarget with a Relationships
Transform immediately followed by C14N; it deliberately keeps every published
aggregate selector count equal. Each records only a declared structural package
scope, never cryptographic verification or a trust decision.
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

`topology` is either `pair` (exactly one same-extension `baseline.xlsx` /
`candidate.xlsx` or `baseline.xlsm` / `candidate.xlsm` pair) or `portfolio`
(`baseline/` and `candidate/` directories). `review_expectation` is one of
`allow`, `review`, or `block`. It is a transparent benchmark convention, not
an assertion that every organization must use the same policy.

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
| `cell_number_format_changed` | `sheet`, `cell`, `value`, `custom_number_format_id`, `baseline_format`, `candidate_format`, `formula_cell`, `formula` | One direct-cell custom number-format declaration moves between declared codes while the target style index, raw numeric text, neighboring formula, calculation properties, and every package member except `xl/styles.xml` remain unchanged. The validator does not render a format, resolve locale or column-width behavior, calculate a workbook, or claim client behavior. |
| `ignored_error_rule_added` | `sheet`, `target_range`, `warning_flag`, `formula`, `adjacent_populated_cell`, `adjacent_populated_value`, `downstream_formula_cell`, `downstream_formula` | One standard worksheet `ignoredErrors/ignoredError` declaration is added with the declared target and enabled warning flag while ordinary cells, formulas, calculation properties, and every package member except its worksheet remain unchanged. The validator does not determine whether a client would show a warning, evaluate a formula, judge a warning's justification, render an indicator, or claim client behavior. |
| `auto_filter_criteria_changed` | `sheet`, `filter_ref`, `filter_column_id`, `baseline_filter_value`, `candidate_filter_value`, `subtotal_cell`, `subtotal_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One raw worksheet AutoFilter list criterion changes while its filter shell, formulas, direct dependency edge, and every package member except the report worksheet remain unchanged. The validator does not apply the filter or calculate a result. |
| `named_sheet_view_filter_criterion_changed` | `sheet`, `view_member`, `base_filter_ref`, `filter_column_id`, `baseline_filter_value`, `candidate_filter_value`, `subtotal_cell`, `subtotal_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One relationship-backed Named Sheet View list criterion changes while its worksheet AutoFilter binding, formulas, direct dependency edge, and every package member except its Named Sheet View part remain unchanged. The validator does not activate, render, or apply the view, calculate a subtotal, or infer visible rows. |
| `xml_map_table_column_xpath_retargeted` | `sheet`, `table_member`, `table_name`, `table_ref`, `mapped_column_id`, `mapped_column_name`, `map_member`, `map_id`, `schema_id`, `connection_id`, `baseline_xpath`, `candidate_xpath`, `single_cell_member`, `single_cell`, `single_cell_xpath`, `total_cell`, `total_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One relationship-backed XML Map table-column binding changes its raw XPath while map/schema/file-binding declarations, a sheet-level single-cell mapping, table cells, formulas, calculation properties, and every package member except its table part remain unchanged. The validator does not access a file, validate a schema, import/export XML, materialize data, calculate, or infer a client result. |
| `office_web_addin_auto_show_enabled` | `taskpane_member`, `web_extension_member`, `addin_id`, `reference_id`, `reference_version`, `store`, `store_type`, `baseline_auto_show`, `candidate_auto_show`, `input_sheet`, `input_cell`, `input_value`, `model_sheet`, `model_cell`, `model_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One workbook-to-taskpane-to-web-extension relationship graph retains its synthetic local reference, task-pane layout, ordinary cells, formulas, calculation properties, and every package member except its web-extension part while raw `Office.AutoShowTaskpaneWithDocument` changes from false to true. The validator does not install, load, execute, or fetch an add-in or manifest, or claim that a task pane opens. |
| `ole_object_auto_load_enabled` | `sheet`, `worksheet_member`, `worksheet_relationships_member`, `relationship_id`, `relationship_type`, `target`, `embedded_object_member`, `content_type`, `prog_id`, `dv_aspect`, `shape_id`, `baseline_auto_load`, `candidate_auto_load`, `input_sheet`, `input_cell`, `input_value`, `model_sheet`, `model_cell`, `model_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One worksheet OLE declaration retains its fixed internal relationship, opaque synthetic bytes, ordinary cells, formulas, calculation properties, and every package member except its worksheet while raw `oleObject/@autoLoad` changes from false to true. The validator does not deserialize, open, render, execute, register, or invoke an object server, or claim that an object loads successfully. |
| `sheet_visibility_changed` | `sheet`, `baseline_state`, `candidate_state` | The stored sheet state changes. |
| `formula_cell_unlocked` | `sheet`, `cell` | A formula cell is explicitly unlocked while its sheet remains protected. |
| `sheet_protection_sort_permission_enabled` | `sheet`, `worksheet_member`, `baseline_sort_locked`, `candidate_sort_locked`, `formula_cell`, `formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One protected worksheet retains its stored protection and every other action lock while raw `sheetProtection/@sort` changes exactly from `1` (locked) to `0` (permitted). Its formula context, styles, calculation properties, and every package member except the declared worksheet remain unchanged. The validator does not test a password, encryption, authorization, editable ranges, a client sort operation, or a resulting value. |
| `workbook_structure_lock_removed` | `baseline_lock_structure`, `candidate_lock_structure`, `hidden_sheet`, `hidden_sheet_state`, `formula_sheet`, `formula_cell`, `formula` | Raw `workbookProtection/@lockStructure` changes exactly from `true` to `false` while the declared hidden-sheet state, formula, calculation properties, and every package member except `xl/workbook.xml` remain unchanged. The validator does not test a password, encryption, authorization, or client behavior. |
| `manual_calculation_incomplete` | none | Candidate calculation metadata is `manual` and records incomplete calculation. |
| `iterative_calculation_enabled` | `sheet`, `cell`, `formula`, `baseline_iterate`, `candidate_iterate`, `iteration_count`, `iteration_delta` | The declared direct self-referencing formula remains unchanged while raw `calcPr/@iterate` changes exactly from `false` to `true`; the explicit count and delta remain the declared values, and all non-iteration calculation attributes are unchanged. The validator does not calculate the model. |
| `precision_as_displayed_enabled` | `input_sheet`, `input_cell`, `input_value`, `number_format`, `formula_sheet`, `formula_cell`, `formula`, `baseline_full_precision`, `candidate_full_precision` | The declared stored input, number format, and formula remain unchanged while raw `calcPr/@fullPrecision` changes exactly from `true` to `false`; all other `calcPr` attributes are unchanged. The validator does not calculate, open, or save the workbook. |
| `workbook_date_system_changed` | `baseline_date_1904`, `candidate_date_1904`, `date_compatibility`, `serial_sheet`, `serial_cell`, `serial_value`, `number_format`, `formula_sheet`, `formula_cell`, `formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | Raw `workbookPr/@date1904` changes from `false` to `true` while explicit compatibility remains `true`; the raw numeric serial, custom date format, local formulas, and every package member except `xl/workbook.xml` remain unchanged. The validator does not calculate, convert, or infer displayed dates. |
| `formula_cached_result_changed` | `sheet`, `cell`, `formula`, `input_sheet`, `input_cell`, `input_value`, `result_type`, `baseline_cached_result`, `candidate_cached_result` | One raw numeric formula-cell `<v>` value changes while its raw `<f>` expression, direct input, calculation properties, and every other package member remain unchanged. The validator reads OOXML only; it does not calculate, validate, or interpret the saved result. |
| `external_data_connection_refresh_on_load_changed` | `connection_id`, `baseline_refresh_on_load`, `candidate_refresh_on_load` | The relationship-backed connection with this workbook-local ID explicitly changes `refreshOnLoad` from `false` to `true`. The validator reads raw OOXML only. |
| `external_data_connection_web_query_url_changed` | `connection_id`, `connection_member`, `connection_content_type`, `workbook_relationships_member`, `relationship_id`, `relationship_type`, `refresh_on_load`, `baseline_url`, `candidate_url`, `saved_value_sheet`, `saved_value_cell`, `saved_value`, `summary_sheet`, `summary_cell`, `summary_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One relationship-backed web-query connection retains its ID, type, refresh controls, relationship, content type, stored imported cell, ordinary formulas, calculation properties, and every package member except `xl/connections.xml` while raw `webPr/@url` moves between declared reserved URLs. The validator reads local OOXML only; it does not resolve, open, fetch, authenticate to, trust, refresh, calculate, or claim a client result. |
| `package_signature_manifest_direct_part_retargeted` | `root_relationships_member`, `origin_member`, `origin_relationships_member`, `origin_relationship_id`, `origin_relationship_type`, `signature_member`, `signature_relationship_id`, `signature_relationship_type`, `signature_content_type`, `signed_info_reference_uri`, `baseline_manifest_uri`, `candidate_manifest_uri`, `baseline_direct_part_class`, `candidate_direct_part_class`, `stable_sheet`, `stable_value_cell`, `stable_value`, `stable_formula_cell`, `stable_formula` | One structurally shaped OPC digital-signature envelope retains its origin/signature relationship graph, content types, `SignedInfo` local-object reference, ordinary cells/formula context, calculation properties, and every package member except `_xmlsignatures/sig1.xml` while its one `Object/Manifest` direct-part URI moves from the workbook part to a worksheet part. The fixture's digest/signature values are synthetic; the validator does not verify cryptography, transforms, certificates, identity, trust, or a consumer decision. |
| `package_signature_manifest_relationship_selector_retargeted` | `root_relationships_member`, `origin_member`, `origin_relationships_member`, `origin_relationship_id`, `origin_relationship_type`, `signature_member`, `signature_relationship_id`, `signature_relationship_type`, `signature_content_type`, `signed_info_reference_uri`, `manifest_uri`, `relationship_transform_algorithm`, `canonicalization_algorithm`, `office_document_relationship_id`, `office_document_relationship_type`, `office_document_relationship_target`, `baseline_selector_source_id`, `candidate_selector_source_id`, `baseline_selected_relationship_type`, `baseline_selected_relationship_target`, `candidate_selected_relationship_type`, `candidate_selected_relationship_target`, `stable_sheet`, `stable_value_cell`, `stable_value`, `stable_formula_cell`, `stable_formula` | One structurally shaped OPC digital-signature envelope retains its package graph, Manifest URI, Relationships Transform followed immediately by C14N, ordinary cells/formula context, calculation properties, and every package member except `_xmlsignatures/sig1.xml` while one root-relationship selector `SourceId` changes. The selector identifies a relationship entry, not proof its target part was signed. The fixture's digest/signature values are synthetic; the validator neither executes transforms nor verifies cryptography, certificates, identity, trust, or a consumer decision. |
| `query_table_refresh_on_load_changed` | `sheet`, `connection_id`, `connection_member`, `connection_url`, `query_table_member`, `worksheet_member`, `worksheet_relationships_member`, `relationship_id`, `relationship_type`, `baseline_refresh_on_load`, `candidate_refresh_on_load`, `background_refresh`, `refresh_disabled`, `remove_data_on_save`, `fill_formulas`, `connection_edit_disabled`, `growth_behavior`, `saved_value_cell`, `saved_value`, `summary_sheet`, `summary_cell`, `summary_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One direct internal worksheet QueryTable relationship and one fixed internal workbook connection retain their controls, saved cells, formulas, calculation properties, and every package member except the QueryTable part while raw `queryTable/@refreshOnLoad` changes from `false` to `true`. The validator does not open a connection, fetch a URL, refresh, materialize data, calculate, or claim a client result. |
| `cell_hyperlink_target_changed` | `sheet`, `cell`, `cell_value`, `worksheet_member`, `worksheet_relationships_member`, `relationship_id`, `relationship_type`, `target_mode`, `baseline_target`, `candidate_target`, `summary_sheet`, `summary_cell`, `summary_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One worksheet `hyperlink` declaration retains its visible cell value, relationship ID/type/mode, no local location/display/tooltip, formulas, calculation properties, and every package member except its worksheet relationship part while the one external relationship `Target` moves between the declared reserved URLs. The validator does not resolve, open, fetch, visit, execute, calculate, or claim that a client follows a target. |
| `pivot_cache_refresh_on_load_changed` | `cache_id`, `source_type`, `source_sheet`, `source_ref`, `pivot_sheet`, `pivot_ref`, `pivot_output_cell`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula`, `baseline_refresh_on_load`, `candidate_refresh_on_load` | One relationship-bound local worksheet PivotCache changes raw `refreshOnLoad` from `false` to `true`; its source binding, PivotTable location, stored report/dashboard cells, calculation properties, and every package member except its cache definition remain unchanged. The validator neither refreshes nor renders a PivotTable. |
| `pivot_data_field_aggregation_changed` | `cache_id`, `source_type`, `source_sheet`, `source_ref`, `pivot_sheet`, `pivot_ref`, `pivot_output_cell`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula`, `data_field_source_index`, `baseline_subtotal`, `candidate_subtotal` | One relationship-bound local worksheet PivotTable retains its source/cache bindings, stored report/dashboard cells, refresh control, and calculation properties while raw `dataFields/dataField/@subtotal` moves between the declared aggregate functions. Every package member except its PivotTable definition remains unchanged. The validator does not refresh, calculate, render, or infer a displayed result. |
| `pivot_slicer_selection_changed` | `cache_id`, `source_type`, `source_sheet`, `source_ref`, `pivot_sheet`, `pivot_ref`, `pivot_output_cell`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula`, `slicer_name`, `slicer_source_name`, `slicer_pivot_table_name`, `slicer_pivot_tab_id`, `item_count`, `baseline_selected_item_index`, `candidate_selected_item_index`, `baseline_selected_value`, `candidate_selected_value` | One relationship-bound local Slicer cache retains its source/PivotCache/PivotTable bindings, stored report/dashboard cells, refresh control, and calculation properties while exactly one selected cache item moves between the declared index/value pairs. Every package member except its Slicer-cache definition remains unchanged. The validator does not create a Slicer drawing, apply a filter, refresh, calculate, render, or infer a displayed result. |
| `power_query_m_filter_changed` | `data_mashup_part`, `source_sheet`, `source_table`, `source_ref`, `query_section`, `query_name`, `filter_column`, `baseline_filter_value`, `candidate_filter_value`, `fill_enabled`, `firewall_enabled`, `future_packages_allowed` | One package-root relationship-bound compact Data Mashup retains its local Excel Table source, metadata, permission controls, calculation properties, and every package member except its custom-XML part while one stored M `Table.SelectRows` literal changes. The validator reads a bounded generated envelope only; it does not execute M, refresh, materialize output, calculate, or infer query results. |
| `scenario_manager_stored_input_value_changed` | `scenario_sheet`, `scenario_name`, `changing_cell`, `stable_input_cell`, `baseline_stored_value`, `candidate_stored_value`, `stable_stored_value`, `input_number_format_id`, `summary_ref`, `worksheet_input_value`, `worksheet_stable_input_value`, `result_cell`, `result_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One raw worksheet Scenario Manager declaration retains its selected/locked scenario metadata, second stored input, summary reference, visible worksheet cells, formula path, calculation properties, and every package member except its scenario-bearing worksheet while one stored `inputCells/@val` value changes. The validator does not show/apply a scenario, calculate a result, create a scenario summary, or infer client behavior. |
| `what_if_data_table_input_reference_changed` | `table_sheet`, `master_cell`, `output_range`, `baseline_input_cell`, `candidate_input_cell`, `orientation`, `recalculation_requested`, `input_value_range`, `input_values`, `primary_input_value`, `alternate_input_value`, `scale_cell`, `scale_value`, `output_formula_cell`, `output_formula`, `model_sheet`, `model_cell`, `model_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One raw column-oriented one-variable Data Table master retains its declared output range, recalculation request, visible input grid, ordinary formulas, calculation properties, and every package member except its table-bearing worksheet while `f/@r1` moves between declared local input cells. The validator does not substitute inputs, calculate, infer table results, resolve a circular dependency, or claim client behavior. |
| `chart_series_value_reference_changed` | `chart_sheet`, `chart_anchor`, `source_sheet`, `series_title_ref`, `category_ref`, `baseline_value_ref`, `candidate_value_ref` | One relationship-bound DrawingML chart retains its host, anchor, title/category references, source worksheet cells, and every package member except its chart part while raw `c:ser/c:val/c:numRef/c:f` moves between declared local value ranges. The validator does not calculate, refresh, or render a chart. |
| `external_workbook_link_update_policy_changed` | `sheet`, `cell`, `formula`, `baseline_update_links`, `candidate_update_links` | The declared external-workbook formula remains unchanged while raw `workbookPr/@updateLinks` changes exactly from `never` to `always`; all other stored `workbookPr` attributes are unchanged. The validator does not resolve the source workbook. |
| `external_workbook_link_source_changed` | `sheet`, `cell`, `formula`, `workbook_member`, `workbook_relationships_member`, `workbook_relationship_id`, `workbook_relationship_type`, `external_link_member`, `external_link_relationships_member`, `external_link_content_type`, `external_link_relationship_id`, `external_link_relationship_type`, `external_sheet`, `target_mode`, `baseline_target`, `candidate_target`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One external-workbook formula, workbook external-reference relationship, `externalLink`/`externalBook` declaration, content type, source-sheet name, local downstream formula, and calculation properties remain unchanged while one external `externalLinkPath` relationship `Target` moves between declared reserved URLs. The validator reads local OOXML only; it does not resolve, open, fetch, authenticate to, trust, refresh, calculate, or claim that a client updates a link or returns a value. |
| `external_defined_name_source_changed` | `name`, `workbook_member`, `baseline_refers_to`, `candidate_refers_to`, `formula_sheet`, `formula_cell`, `formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One local `definedName` text moves between declared qualified external-workbook expressions while the name, local formula context, calculation properties, sheet declarations, workbook relationships, and every package member except `xl/workbook.xml` remain unchanged. The compact package deliberately has no `externalLink` part or `externalReferences` declaration. The validator does not resolve, open, fetch, authenticate to, trust, refresh, calculate, or claim a client result. |
| `named_lambda_definition_changed` | `name`, `workbook_member`, `parameters`, `baseline_refers_to`, `candidate_refers_to`, `input_sheet`, `rate_cell`, `rate_value`, `amount_cell`, `amount_value`, `formula_sheet`, `formula_cell`, `formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One workbook-scoped named `LAMBDA` body changes while its declared inputs, calling formula, dashboard consumer, calculation properties, sheet declarations, workbook relationships, and every package member except `xl/workbook.xml` remain unchanged. The compact package has no `externalLink` part or `externalReferences` declaration. The validator does not evaluate a LAMBDA, calculate a result, infer client/version support, or claim a client result. |
| `power_pivot_data_model_relationship_changed` | `workbook_member`, `workbook_relationships_member`, `data_model_member`, `workbook_relationship_id`, `workbook_relationship_type`, `workbook_relationship_target`, `data_model_content_type`, `extension_uri`, `min_version_load`, `model_tables`, `from_table`, `from_column`, `to_table`, `baseline_to_column`, `candidate_to_column`, `data_model_payload_sha256`, `data_model_payload_size` | One workbook-level `x15:modelRelationship` target key changes while the generated `powerPivotData` binding, `.data` content type, local Table context, calculation properties, and opaque `xl/model/item.data` payload remain unchanged; `xl/workbook.xml` is the only changed package member. The validator reads the declaration, local package binding, and payload digest only. It does not deserialize an Analysis Services payload, evaluate DAX, refresh a model, calculate or render a report, infer model-to-cell impact, or claim client behavior. |
| `xlm_auto_open_binding_retargeted` | `workbook_member`, `workbook_relationships_member`, `macro_sheet_member`, `macro_sheet_relationship_id`, `macro_sheet_relationship_type`, `macro_sheet_relationship_target`, `macro_sheet_content_type`, `workbook_content_type`, `macro_sheet_name`, `macro_sheet_sheet_id`, `macro_sheet_state`, `automatic_macro_name`, `automatic_macro_event`, `baseline_target`, `candidate_target`, `macro_sheet_formula`, `macro_sheet_formula_cells`, `macro_sheet_sha256`, `macro_sheet_size`, `input_sheet`, `input_cell`, `input_value`, `model_sheet`, `model_cell`, `model_formula`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One macro-enabled `.xlsm` workbook-scoped `_xlnm.Auto_Open` defined-name target changes between two declared cells on one very-hidden XLM macro sheet while the raw macro sheet, macro relationship, content types, ordinary formula context, calculation properties, and every package member except `xl/workbook.xml` remain unchanged. The validator reads the stored dispatch declaration and bounded static package shape only: it does not open Excel, enable or execute XLM code, parse or emulate instructions, resolve a dynamic name, inspect security or trust state, infer dispatch behavior, calculate a workbook, or claim client behavior. |
| `array_formula_mode_changed` | `sheet`, `cell`, `formula`, `baseline_mode`, `candidate_mode`, `baseline_output_range`, `candidate_output_range` | The declared unchanged array anchor moves from `legacy_cse` to `dynamic`, with its stored formula text and output range exactly as declared. The validator reads the raw OOXML cell-metadata binding. |
| `static_cycle_introduced` | `cells` | Every declared direct A1 cell reaches itself in the local static dependency graph. |
| `three_d_scope_changed` | `formula_sheet`, `formula_cell`, `inserted_sheet`, `after_sheet`, `before_sheet` | Formula text remains unchanged while a sheet is inserted inside the declared tab span. |
| `structural_formula_rewrite` | `baseline`, `candidate` | Declared before/after formula locations and text exist. This is an annotation, not a general proof of equivalence. |
| `portfolio_value_changed` | `workbook`, `sheet`, `cell` | A literal value differs between paired portfolio members. |
| `portfolio_external_reference` | `workbook`, `sheet`, `cell`, `target_workbook` | The local portfolio model contains the declared external workbook reference. |
| `structured_table_scope_changed` | `table_sheet`, `table`, `baseline_ref`, `candidate_ref`, `formula_sheet`, `formula_cell` | A stored Excel Table range changes while the declared formula remains textually unchanged and retains a reference to that table. |
| `table_calculated_column_formula_changed` | `table_sheet`, `table_member`, `table`, `table_ref`, `calculated_column_id`, `calculated_column_name`, `baseline_formula`, `candidate_formula`, `stable_formula_cells`, `dashboard_sheet`, `dashboard_cell`, `dashboard_formula` | One local Table calculated-column formula master changes while the Table binding/range/headers, ordinary row formulas, structured-reference dashboard formula, calculation properties, and every package member except its Table definition remain unchanged. The validator does not fill a column, reconcile master and row formulas, calculate a structured reference, infer a total, or claim client behavior. |
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

## Custom number format

Microsoft's [hide/display guidance](https://support.microsoft.com/en-us/excel/hide-or-display-cell-values)
documents `;;;` as a custom format that hides worksheet values, while its
[percentage guidance](https://support.microsoft.com/en-us/excel/format-numbers-as-percentages-in-excel)
documents the distinct display behavior of percentage formats. WCAB's pair has
one `Operations!B2` cell whose custom format code moves from
`0.0%;[Red](0.0%);-` to `;;;`. Its raw `0.125` numeric text, cell style
index, neighboring `Operations!B3=B2` formula, and calculation properties
stay unchanged.

The validator follows the direct cell style to one custom `numFmt` with ID
`164`, compares styles XML after erasing only that format code, and requires
`xl/styles.xml` to be the only changed package member. It records stored
display metadata only: it does not render a format, resolve locale or
column-width behavior, decide what a client displays, calculate a workbook, or
claim Excel-client behavior.

## Ignored error-checking rule

Microsoft's [formula-error guidance](https://support.microsoft.com/en-us/excel/detect-formula-errors-in-excel)
explains that a selected ignored error no longer appears in later error checks
until errors are reset, and documents the formula-range-omission checking rule.
The Office 2010 [MS-XLSX `ignoredErrors` specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/0d164d85-23bf-4d43-87c5-9fcde148aabe)
records the related extension form. A stored per-range suppression is therefore
a review surface even when no ordinary formula or cell changes.

WCAB's pair keeps `Operations!B2=10`, `B3=20`, `B4=30`,
`B5=SUM(B2:B3)`, and `C5=B5` unchanged. The candidate adds exactly one
standard `ignoredErrors/ignoredError` child with `sqref="B5"` and
`formulaRange="1"`. The validator accepts only that compact raw shape,
compares the Operations worksheet after removing the declaration, and requires
that worksheet to be the only changed package member. It does not determine
whether Excel would show a warning, evaluate the formula, infer a displayed
indicator, decide whether the warning is warranted, change application-level
settings, calculate a workbook, or claim Excel-client behavior.

## Workbook-structure protection

Microsoft's [Protect a workbook guidance](https://support.microsoft.com/en-us/office/protect-a-workbook-7e365a4d-3e89-4616-84ca-1931257c1517)
distinguishes structural protection from file and worksheet protection, and
states that a protected structure restricts adding, moving, deleting, hiding,
unhiding, and renaming worksheets. It also makes the password optional. The
Office Open XML [`workbookProtection` specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/ae53f189-04e6-45e6-acb2-7f61fecabee4)
defines the stored workbook-level protection element. This makes the raw
structure control review-material even if ordinary worksheet cells stay fixed.

WCAB 0.25 keeps the hidden `ReviewControls` sheet and
`Inputs!D2=B2*C2` formula unchanged. Its baseline has exactly
`workbookProtection/@lockStructure="1"`; the candidate has exactly `"0"`.
The validator compares raw workbook XML after removing only that attribute and
requires `xl/workbook.xml` to be the only changed package member. It does not
test a password, encryption, authentication, authorization, whether a hidden
sheet is exposed, or whether a particular client enables a sheet operation.

## Sheet-protection sort permission

Microsoft's [Protect a worksheet guidance](https://support.microsoft.com/en-us/excel/protect-a-worksheet)
distinguishes worksheet protection from file protection and lists **Sort** as
one of the actions a protected worksheet can permit. The Open XML
[`SheetProtection` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetprotection?view=openxml-3.0.1)
defines the stored worksheet-protection element and its action controls. That
makes a stored action permission review-material even when no cells change.

WCAB 0.33 keeps `Controls!D2=B2*C2`, its direct
`Dashboard!B4=Controls!$D$2` consumer, worksheet protection, every other
protection action lock, styles, calculation properties, and all non-worksheet
package members fixed. The raw `sheetProtection/@sort` control moves exactly
from `1` (locked) to `0` (permitted), and `xl/worksheets/sheet1.xml` is the
only changed package member. The validator records that stored declaration
only: it does not test a password, encryption, authentication, authorization,
editable ranges, an actual sort operation, a client permission decision, or a
resulting value.

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

## QueryTable refresh-on-open

Microsoft's [QueryTable.RefreshOnFileOpen reference](https://learn.microsoft.com/en-us/office/vba/api/excel.querytable.refreshonfileopen)
defines a table-level control for automatically refreshing a QueryTable when a
workbook opens. The Open XML [`QueryTable` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.querytable?view=openxml-3.0.1)
models this as `refreshOnLoad` alongside the table's `connectionId`. It is
separate from the connection-level control: a reviewer needs to see the stored
QueryTable request even when the connection metadata itself is unchanged.

WCAB's original fixture has one `ImportedData` worksheet QueryTable relationship
and one fixed internal workbook connection to a reserved `example.invalid` URL.
Only `xl/queryTables/queryTable1.xml` changes: raw
`queryTable/@refreshOnLoad` moves from `false` to `true`; the connection's own
`refreshOnLoad=false`, all other QueryTable controls, relationships, content
types, saved `ImportedData!B2=100` cell, and
`ImportedData!B2 → Summary!B2 → Dashboard!B4` formula context stay fixed. The
validator follows local OOXML relationships only and does not open a
connection, fetch a URL, refresh a query, materialize rows, calculate a
workbook, or claim that any client refreshes successfully.

## Worksheet cell hyperlink target

Microsoft's [Hyperlink.Address reference](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.address)
defines the target-document address, and the Open XML
[`Hyperlink` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1)
defines worksheet `x:hyperlink/@r:id` as the relationship binding that expresses
that target. A visible cell value can remain unchanged while that relationship's
external destination changes.

WCAB's original fixture has one `Inputs!B2` external cell hyperlink. Its visible
text is `Open vendor portal`; its worksheet declaration has only `ref` and
`r:id`, and its one relationship retains a fixed ID, standard hyperlink type,
and `TargetMode=External`. Only `xl/worksheets/_rels/sheet1.xml.rels` changes:
the raw relationship `Target` moves from `approved.example.invalid` to
`review.example.invalid`; the worksheet XML, cell text, calculation properties,
and `Inputs!B2 → Summary!B2 → Dashboard!B4` formula context stay fixed. The
validator reads local OOXML only, compares the relationship part after erasing
`Target`, and does not resolve, open, fetch, visit, execute, calculate, or
claim that any client follows either URL.

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

## External-workbook link source

Microsoft's [workbook-link guidance](https://support.microsoft.com/en-us/excel/manage-workbook-links)
documents **Change source** as a way to point existing links at another
workbook. The Open XML [`ExternalBook` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.externalbook?view=openxml-3.0.1)
defines `externalBook` as an external workbook supplying data to the current
workbook and exposes its relationship to the supporting book path. A source
change can therefore be review-material even when the external formula text
does not change.

WCAB's source pair keeps `LinkedModel!B2` as
`='[WCABSource.xlsx]Inputs'!$B$2` and keeps
`Dashboard!B4=LinkedModel!$B$2` unchanged. Each package has exactly one
workbook `<externalReferences>` binding, one externalLink relationship, one
`externalLink/externalBook` declaration with the `Inputs` sheet name, and one
externalLinkPath relationship with `TargetMode=External`. Only
`xl/externalLinks/_rels/externalLink1.xml.rels` changes: its raw `Target`
moves from the reserved `approved.example.invalid` URL to
`review.example.invalid`. The validator follows only those local package parts,
compares the relationship part after erasing `Target`, and requires it to be
the sole package difference. It does not resolve, open, fetch, authenticate
to, trust, refresh, calculate, or otherwise interact with either source, and
does not claim a client updates a link or returns a value.

## External defined-name source

Microsoft's [workbook-link guidance](https://support.microsoft.com/en-us/excel/manage-workbook-links)
calls out defined names as a place workbook links can be used, including its
workflow for finding links in defined names. The Open XML
[`DefinedNames` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.definednames?view=openxml-3.0.1)
defines the workbook-level collection that stores those expressions. A source
can therefore move inside a local name even when neither an ordinary formula
cell nor an `externalLink` package part changes.

WCAB 0.33 keeps local `Model!B2=ScenarioRate*2` and its direct
`Dashboard!B4=Model!$B$2` consumer unchanged while the one workbook-level
`ScenarioRate` definition moves from
`'[WCABApprovedSource.xlsx]Inputs'!$B$2` to
`'[WCABReviewSource.xlsx]Inputs'!$B$2`. The deliberately compact package has
no `<externalReferences>` declaration and no `xl/externalLinks/` member. The
raw validator requires the exact name/text transition, unchanged calculation
properties and sheet declarations, and `xl/workbook.xml` as the only package
difference. It does not resolve, open, fetch, authenticate to, trust, refresh,
calculate, or otherwise interact with either synthetic source, or claim that a
client resolves the name or returns a value.

## Named LAMBDA definition

Microsoft's [LAMBDA guidance](https://support.microsoft.com/en-us/excel/functions/lambda-function)
describes assigning a named LAMBDA in Name Manager so it is callable as a
reusable custom function throughout the workbook. Its [defined-names
guidance](https://support.microsoft.com/en-US/Excel/get-started/define-and-use-names-in-formulas)
also identifies names as reusable formula definitions. The name's stored body
is therefore a review surface even when each caller's ordinary cell formula is
unchanged.

WCAB 0.34 keeps `Inputs!B2=0.08`, `Inputs!B3=100`,
`Model!B2=ScenarioValue(Inputs!B2,Inputs!B3)`, and
`Dashboard!B4=Model!$B$2` fixed while the sole workbook-level `ScenarioValue`
definition moves exactly from `=LAMBDA(rate,amount,rate*amount)` to
`=LAMBDA(rate,amount,rate*(amount+10))`. The validator requires one bare
workbook-scoped name, no external-reference declarations, identical raw input
and ordinary formula context, equal calculation properties, and
`xl/workbook.xml` as the sole changed member. It reads the stored body only; it
does not evaluate a LAMBDA, resolve all named-formula dependencies, calculate a
result, infer Excel-version support, or claim client behavior.

## Table calculated-column formula master

Microsoft's [calculated-column guidance](https://support.microsoft.com/en-US/Excel/use-calculated-columns-in-an-excel-table)
documents that one formula in an Excel Table can be filled throughout a column.
The Open XML [`calculatedColumnFormula` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculatedcolumnformula?view=openxml-3.0.1)
defines the `tableColumn` child that stores that formula. This is a Table-level
formula definition, separate from a worksheet formula-cell element.

WCAB 0.34 keeps the local `ScenarioLedger` binding at `Ledger!A1:C4`, its
headers, `Ledger!C2:C4` ordinary formulas, and
`Dashboard!B4=SUM(ScenarioLedger[Calculated amount])` fixed. The third Table
column's sole raw `calculatedColumnFormula` text moves from `A2*B2` to
`A2*(B2+1)`, with `xl/tables/table1.xml` as the sole package difference. The
validator follows the worksheet's one local Table relationship, checks the
compact Table shape and formula-master text, and compares the Table part after
erasing only that text. It does not fill a column, reconcile row formulas to a
master, calculate a structured reference, infer a total, add rows, open Excel,
or claim client behavior.

## Power Pivot/Data Model relationship

Microsoft's [PowerPivot Model object guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-the-powerpivot-model-object-in-excel)
explains that relationships connect model tables and enable multi-table
PivotTable or PivotChart filtering. The Open XML
[`x15:dataModel` declaration](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/eaea0fe6-3e3c-401d-a3a0-2d2cbb9fce00)
contains `modelRelationship` entries with `fromTable`, `fromColumn`, `toTable`,
and `toColumn` attributes. That declaration is a review surface outside the
ordinary A1 formula graph.

WCAB 0.35 keeps the generated `SalesModel` and `CalendarModel` local Tables,
one `powerPivotData` workbook binding, the `.data` content type, calculation
properties, and fixed opaque `xl/model/item.data` bytes unchanged. Its sole
`x15:modelRelationship` changes `toColumn` from `DateKey` to `FiscalDateKey`.
The raw validator checks the exact relationship, package binding, model-table
metadata, and payload SHA-256, then compares workbook XML with just that target
key erased. It does not deserialize an Analysis Services stream, evaluate DAX,
load or refresh a model, calculate or render a report, infer model-to-cell
impact, or claim client behavior.

## External-data web-query source declarations

External-data connections can store provider, server, authentication, command,
and refresh settings outside worksheet cells. For a web query, the Open XML
[`webPr` element](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.webqueryproperties?view=openxml-3.0.0)
holds the stored URL and request controls. Those source declarations are a
review surface even when imported cells and ordinary formulas remain fixed.

WCAB 0.37 isolates one relationship-backed `xl/connections.xml` web-query
connection. Only raw `webPr/@url` moves from the reserved
`approved.example.invalid` endpoint to `review.example.invalid`; the connection
ID/type/name, refresh controls, relationship, content type, saved
`ImportedData!B2=100` cell, `ImportedData!B2 → Summary!B2 → Dashboard!B4`
formula context, calculation properties, and every other package member remain
unchanged. The raw validator checks the local package graph and compares the
connection part after removing only the URL. It does not resolve, open, fetch,
authenticate to, trust, refresh, calculate, or otherwise interact with either
endpoint, or claim a client result.

## OPC package-signature Manifest declarations

An OPC package signature has package relationships that locate a signature
origin and XML signature part. Within an XMLDSIG envelope, a `SignedInfo`
reference can name a local `Object`; an `Object/Manifest/Reference` is the
separate declaration that can name a package part. Those two scopes are not
interchangeable review evidence.

WCAB 0.38 keeps one root-to-origin relationship, one origin-to-signature
relationship, the required origin/signature content types, one `SignedInfo`
reference to `#idWCABPackageObject`, ordinary `Controls!B10=12` and
`Controls!D10=B10*C10` context, calculation properties, and all package
members except `_xmlsignatures/sig1.xml` fixed. Its one
`Object/Manifest/Reference/@URI` moves from the direct workbook part to the
direct worksheet part. The raw validator verifies that bounded structural
shape, compares the signature XML after erasing only that URI, and never emits
the raw URI through an adapter contract.

WCAB 0.39 uses the same bounded signature graph, local-object reference, and
ordinary workbook context but fixes the Manifest URI to the root relationships
part. Its Manifest has exactly one OPC Relationships Transform followed
immediately by XML C14N. Only the transform's
`RelationshipReference/@SourceId` changes, from the root office-document
relationship to the root signature-origin relationship. The validator compares
the signature XML after erasing only that source ID. The selected relationship
is a declaration of an entry in `_rels/.rels`, not evidence that the entry's
target part was signed; an adapter sees only the redacted equal-count aggregate
and the separate Manifest-coverage-change flag.

The generated digest and signature values are deliberately synthetic. WCAB
does not validate a digest, signature, canonicalization, transform, certificate,
identity, revocation, trust chain, timestamp, policy, or a package consumer's
decision. It records only the declared structural package scope.

## XLM Auto_Open binding

WCAB 0.36 provides a real macro-enabled `.xlsm` pair rather than relabeling an
ordinary workbook. Each archive has one very-hidden `Macro Automation`
`xlMacrosheet` relationship and one workbook-scoped `_xlnm.Auto_Open` defined
name. Only that name's raw target changes, from `'Macro Automation'!$A$1` to
`'Macro Automation'!$A$2`; the fixed macro-sheet part contains exactly the two
static formula cells `A1=HALT()` and `A2=HALT()`. The raw macro-sheet bytes,
workbook-to-sheet relationship, macro-enabled workbook content type,
macro-sheet content type, normal worksheet cells and formulas, calculation
properties, and all other package members remain unchanged.

The validator compares the stored OOXML declaration and bounded local package
shape. It neither opens Excel nor enables, executes, parses, emulates, or
otherwise interprets XLM instructions. It does not resolve a dynamic name,
inspect macro-security or trust settings, infer a dispatch outcome, calculate
the workbook, or claim client behavior. The stable
`Inputs!B2 → Model!B2 → Dashboard!B4` path is ordinary local review context,
not a macro dependency claim.

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

## Named Sheet Views

Microsoft's [Sheet View guidance](https://support.microsoft.com/en-us/excel/create-and-manage-sheet-views-in-excel)
documents saved, customized filters and sorts for collaboration. The
[MS-XLSX Named Sheet Views specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/78cf20a1-2551-45c6-86bf-f1e92bd5fc39)
defines the stored collection as sort/filter settings associated with
AutoFilters on a worksheet. A saved view is therefore a distinct review
surface from the worksheet active filter.

The WCAB Named Sheet View pair keeps a no-criterion worksheet
`Report!A1:B5` AutoFilter, rows, `Report!D2=SUBTOTAL(109,B2:B5)`, and
`Dashboard!B4=Report!$D$2` fixed. A worksheet relationship binds the base
AutoFilter `xr:uid` to one `xl/namedSheetViews/namedSheetView1.xml` part;
only its stored column-0 list value moves from `North` to `South`. The raw
validator follows that relationship, reconciles the stable filter ID and
range, compares the view part after erasing only the terminal list value,
and requires that part to be the sole package difference. It does not
activate, render, or apply a view, calculate a subtotal, infer visible rows,
or claim a client display or print outcome.

## XML Maps

Microsoft's [XML overview](https://support.microsoft.com/en-US/Excel/overview-of-xml-in-excel)
describes XML Maps as the relationship between worksheet mapped cells or XML
tables and XML schema elements, used by Excel import/export workflows. The
[XmlMap API](https://learn.microsoft.com/en-us/office/vba/api/excel.xmlmap)
exposes corresponding Import and Export operations. A table-column XPath is
therefore a reviewable data-contract declaration even when its current cells
and formulas remain unchanged.

The WCAB XML Map pair keeps the synthetic local MapInfo/XSD declaration,
file-binding metadata, `Export!E2` single-cell mapping, table values,
`Export!D2=SUM(InvoiceLines[Net amount])`, and
`Dashboard!B4=Export!$D$2` fixed. Its only difference is the `InvoiceLines`
table's `Net amount` `xmlColumnPr/@xpath`: `NetAmount` becomes
`TaxAmount`. The raw
validator follows the local workbook-to-map and worksheet-to-table/single-cell
relationships, compares the mapped table part with only the XPath erased,
and requires `xl/tables/table1.xml` to be the sole differing package member.
It does not access a file, validate the schema, import/export XML, materialize
data, calculate a result, or claim a client outcome.

## Office Web Add-ins

Microsoft's [workbook auto-open guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/pnp-open-in-excel)
documents the package parts needed to associate a workbook with an Office
Add-in and notes that the add-in must already be installed, sideloaded, or
deployed before an application can honor the association. The
[MS-OWEMXML specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-owemxml/4b94ac5d-a8df-44f4-8433-81d43e35a2d7)
defines `Office.AutoShowTaskpaneWithDocument` as an extension property. That
makes the stored request reviewable independently of any installed client
state or executable add-in payload.

The WCAB Office Web Add-in pair has one workbook-to-`taskpanes.xml` relation
and one taskpane-to-`webextension1.xml` relation. It keeps the synthetic
local FileSystem reference, add-in/reference IDs, hidden locked task-pane
layout, `Inputs!B2=10`, `Model!B2=Inputs!$B$2*2`, and
`Dashboard!B4=Model!$B$2` fixed. Only the web-extension property changes:
`Office.AutoShowTaskpaneWithDocument` is `false` in the baseline and
`true` in the candidate. The raw validator follows only those local
relationships, erases only that property value for comparison, and requires
`xl/webextensions/webextension1.xml` to be the sole changed package member.
There is no manifest payload or external relationship. The benchmark does not
install, load, execute, or fetch an add-in or manifest, and it does not claim
that a task pane opens or that the add-in reads, writes, calculates, or
displays workbook cells.

## Worksheet embedded OLE auto-load

Microsoft's [OLEObject.AutoLoad reference](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobject.autoload)
defines the property as whether an OLE object is automatically loaded when the
containing workbook opens, documents false as the default for new OLE objects,
and notes that ActiveX controls ignore this property. The Open XML
[`OleObject.AutoLoad` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.oleobject.autoload?view=openxml-3.0.1)
models `autoLoad` as a boolean attribute, and its
[`OleObject` element reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.oleobject?view=openxml-3.0.1)
defines the relationship identifier that points to object-persistence data.

The WCAB embedded-OLE pair has one `Inputs` worksheet `oleObject` declaration,
one direct internal `oleObject` relationship, and one fixed opaque ASCII
payload marked with the standard embedded-object content type. It keeps the
synthetic unregistered ProgID, `DVASPECT_CONTENT`, shape ID, relationship,
bytes, `Inputs!B2=10`, `Model!B2=Inputs!$B$2*2`, and
`Dashboard!B4=Model!$B$2` fixed. Only `oleObject/@autoLoad` changes: `false`
in the baseline and `true` in the candidate. The raw validator follows only
the local relationship, verifies the fixed bytes without deserializing them,
erases only `autoLoad` for comparison, and requires
`xl/worksheets/sheet1.xml` to be the sole changed package member. There is no
linked target, ActiveX control, presentation, macro, or external relationship.
The benchmark does not deserialize, open, render, execute, register, or invoke
an object server, and does not claim successful loading.

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
