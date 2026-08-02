# WCAB truth schema, version 1

Every case has one `truth.json`. It is public metadata for an original,
generated fixture; it never relies on a private workbook or an external data
source.

```json
{
  "schema_version": 1,
  "id": "family.case_name",
  "title": "Human-readable scenario",
  "family": "finance",
  "topology": "pair",
  "review_expectation": "block",
  "facts": [],
  "must_reach": [],
  "coverage": []
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
| `conditional_formatting_count_changed` | `sheet`, `baseline_count`, `candidate_count` | Conditional-formatting range count differs as declared. |
| `sheet_visibility_changed` | `sheet`, `baseline_state`, `candidate_state` | The stored sheet state changes. |
| `formula_cell_unlocked` | `sheet`, `cell` | A formula cell is explicitly unlocked while its sheet remains protected. |
| `manual_calculation_incomplete` | none | Candidate calculation metadata is `manual` and records incomplete calculation. |
| `static_cycle_introduced` | `cells` | Every declared direct A1 cell reaches itself in the local static dependency graph. |
| `three_d_scope_changed` | `formula_sheet`, `formula_cell`, `inserted_sheet`, `after_sheet`, `before_sheet` | Formula text remains unchanged while a sheet is inserted inside the declared tab span. |
| `structural_formula_rewrite` | `baseline`, `candidate` | Declared before/after formula locations and text exist. This is an annotation, not a general proof of equivalence. |
| `portfolio_value_changed` | `workbook`, `sheet`, `cell` | A literal value differs between paired portfolio members. |
| `portfolio_external_reference` | `workbook`, `sheet`, `cell`, `target_workbook` | The local portfolio model contains the declared external workbook reference. |

## Static impact lower bounds

`must_reach` entries have one `source` plus one or more `targets`, each with a
sheet and cell. The bundled validator uses only direct, ordinary local A1
references and walks the resulting graph. The targets are therefore lower
bounds: no result implies a complete Excel dependency calculation, a formula
evaluation, dynamic-reference resolution, or a claim about cached values.

## Coverage text

Each case contains an explicit `coverage` list. Consumers must preserve these
boundaries in reports and must not turn unsupported constructs into silent
passes.
