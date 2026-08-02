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
calculation setting changes. Version 2 remains available in the immutable
v0.2.0 and v0.3.0 releases.

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
| `conditional_formatting_count_changed` | `sheet`, `baseline_count`, `candidate_count` | Conditional-formatting range count differs as declared. |
| `sheet_visibility_changed` | `sheet`, `baseline_state`, `candidate_state` | The stored sheet state changes. |
| `formula_cell_unlocked` | `sheet`, `cell` | A formula cell is explicitly unlocked while its sheet remains protected. |
| `manual_calculation_incomplete` | none | Candidate calculation metadata is `manual` and records incomplete calculation. |
| `iterative_calculation_enabled` | `sheet`, `cell`, `formula`, `baseline_iterate`, `candidate_iterate`, `iteration_count`, `iteration_delta` | The declared direct self-referencing formula remains unchanged while raw `calcPr/@iterate` changes exactly from `false` to `true`; the explicit count and delta remain the declared values, and all non-iteration calculation attributes are unchanged. The validator does not calculate the model. |
| `external_data_connection_refresh_on_load_changed` | `connection_id`, `baseline_refresh_on_load`, `candidate_refresh_on_load` | The relationship-backed connection with this workbook-local ID explicitly changes `refreshOnLoad` from `false` to `true`. The validator reads raw OOXML only. |
| `external_workbook_link_update_policy_changed` | `sheet`, `cell`, `formula`, `baseline_update_links`, `candidate_update_links` | The declared external-workbook formula remains unchanged while raw `workbookPr/@updateLinks` changes exactly from `never` to `always`; all other stored `workbookPr` attributes are unchanged. The validator does not resolve the source workbook. |
| `array_formula_mode_changed` | `sheet`, `cell`, `formula`, `baseline_mode`, `candidate_mode`, `baseline_output_range`, `candidate_output_range` | The declared unchanged array anchor moves from `legacy_cse` to `dynamic`, with its stored formula text and output range exactly as declared. The validator reads the raw OOXML cell-metadata binding. |
| `static_cycle_introduced` | `cells` | Every declared direct A1 cell reaches itself in the local static dependency graph. |
| `three_d_scope_changed` | `formula_sheet`, `formula_cell`, `inserted_sheet`, `after_sheet`, `before_sheet` | Formula text remains unchanged while a sheet is inserted inside the declared tab span. |
| `structural_formula_rewrite` | `baseline`, `candidate` | Declared before/after formula locations and text exist. This is an annotation, not a general proof of equivalence. |
| `portfolio_value_changed` | `workbook`, `sheet`, `cell` | A literal value differs between paired portfolio members. |
| `portfolio_external_reference` | `workbook`, `sheet`, `cell`, `target_workbook` | The local portfolio model contains the declared external workbook reference. |
| `structured_table_scope_changed` | `table_sheet`, `table`, `baseline_ref`, `candidate_ref`, `formula_sheet`, `formula_cell` | A stored Excel Table range changes while the declared formula remains textually unchanged and retains a reference to that table. |
| `dynamic_formula_reference_added` | `sheet`, `cell`, `functions` | A previously direct formula changes to one containing the declared introduced dynamic-reference functions. |

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
