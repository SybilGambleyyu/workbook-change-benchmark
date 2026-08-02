# WCAB truth schema, version 3

Every case has one `truth.json`. It is public metadata for an original,
generated fixture; it never relies on a private workbook or an external data
source.

Version 3 adds `coverage_expectations`: machine-matchable disclosures for
important analysis boundaries. WCAB 0.5 extends that same stable envelope with
cases where an unchanged `INDIRECT` or `OFFSET` formula receives a changed
selector. Version 2 remains available in the immutable v0.2.0 and v0.3.0
releases.

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
