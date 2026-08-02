# WCAB normalized observation protocol, version 2

WCAB evaluates tools that inspect a baseline/candidate workbook change. Native
tool reports differ, so an adapter should translate its result into this small
JSON protocol before invoking `wcab score`. The protocol is deliberately
separate from the fixture truth schema: a tool may keep its own report format,
but its benchmark adapter has one portable target.

Version 2 adds scoreable **coverage declarations**. They let the benchmark
separately ask whether a tool surfaced a known static-analysis boundary rather
than merely finding the surrounding formula edit. Version 1 remains available
with the immutable WCAB 0.3.0 release.

## Start with a template

```bash
wcab observation-template --fixtures fixtures --output observations.json
# Run an adapter that replaces unsupported entries with observed evidence.
wcab score --fixtures fixtures --observations observations.json \
  --output score.json
```

The template is valid and marks every case `unsupported`. That is intentional:
unsupported coverage is visible in the score rather than being treated as a
pass.

## Document shape

```json
{
  "schema_version": 2,
  "benchmark": {"fixture_schema_versions": [3]},
  "tool": {"name": "example-tool", "version": "1.4.0"},
  "cases": [
    {
      "id": "structural.dynamic_reference_introduced",
      "status": "analyzed",
      "facts": [
        {
          "fact": {
            "kind": "dynamic_formula_reference_added",
            "sheet": "Summary",
            "cell": "B2",
            "functions": ["INDIRECT"]
          },
          "evidence": {"native_rule": "dynamic-reference-added"}
        }
      ],
      "review": "block",
      "coverage": {
        "declarations": [
          {
            "expectation": {
              "kind": "dynamic_reference_static_coverage",
              "sheet": "Summary",
              "cell": "B2",
              "functions": ["INDIRECT"]
            },
            "evidence": {"native_rule": "dependency-coverage-warning"}
          }
        ],
        "notes": []
      }
    }
  ]
}
```

`schema_version` is the observation protocol version.
`fixture_schema_versions` must name the truth-schema versions the adapter
analyzed. `tool.name` is required; `tool.version` is optional.

Each `facts` item wraps an exact WCAB truth fact in `fact`. The optional
`evidence` object is adapter-owned metadata and is never scored; it can record
a native rule ID, report location, or another reproducible pointer without
making one tool's report schema normative.

## Coverage declarations

`coverage.declarations` contains zero or more objects with an exact truth
`expectation` plus optional adapter-owned `evidence`. A declaration means that
the native tool visibly surfaced that specific analysis boundary. Its matching
is exact and deterministic, just like fact matching.

For example, the introduced-dynamic fixture requires the
`dynamic_reference_static_coverage` expectation. `INDIRECT` turns text into a
reference, so a static dependency reviewer must not silently imply complete
target or impact coverage. WCAB 0.5 also includes
`dynamic_reference_driver_changed`: its exact expectation names both the
changed driver and the unchanged dynamic formula. An adapter can use a native
input-change record plus a candidate-profile feature as evidence; it need not
and must not invent a resolved target value. A tool may have richer evaluation
capabilities, but the declaration measures whether the boundary was made
visible—not an untestable claim that every possible Excel evaluation is
resolved.

For a changed `INDIRECT` driver, an adapter declaration can therefore look
like this (the evidence keys remain adapter-owned):

```json
{
  "expectation": {
    "kind": "dynamic_reference_driver_changed",
    "driver": {"sheet": "Inputs", "cell": "E12"},
    "formula": {"sheet": "Summary", "cell": "B2"},
    "functions": ["INDIRECT"]
  },
  "evidence": {
    "native_change_kind": "value_changed",
    "native_profile_feature": "dynamic_reference_cells"
  }
}
```

`coverage.notes` is an optional array of human-readable adapter notes. It is
not scored and is the right place to explain an intentionally unavailable
mapping. Unsupported or errored cases may include notes but cannot include
facts, reviews, or coverage declarations.

## Case status and review

| `status` | Required behavior | Score treatment |
| --- | --- | --- |
| `analyzed` | May report facts, coverage declarations, and one optional `review` disposition. | Its recognized observations can contribute to recall. |
| `unsupported` | Must not report facts, a review, or coverage declarations. | Explicitly counted as unsupported, never as a pass. |
| `error` | Must include a non-empty `error` string and must not report facts, a review, or coverage declarations. | Counted separately from unsupported coverage. |

`review` is `allow`, `review`, `block`, or `null`. A non-null disposition is
compared with WCAB's reference convention and reported as **reference-policy
agreement**. It is not a claim that every organization must use the same
approval policy.

Omitting a known case is permitted for partial runs, but it is reported as
`not_reported` and cannot contribute to coverage or agreement.

## What the score means

The scorer returns deterministic per-case details plus these aggregate metrics:

| Metric | Meaning |
| --- | --- |
| `fact_recall` | Matched declared facts divided by all declared benchmark facts. |
| `coverage_disclosure_recall` | Matched required coverage declarations divided by all declared coverage expectations. |
| `analysis_coverage` | Cases marked `analyzed` divided by all WCAB cases. |
| `review_agreement` | Cases whose analyzed review matches the reference convention divided by all WCAB cases. |
| `complete_case_count` | Analyzed cases with every fact and coverage expectation matched, no unrecognized observation, and matching reference review. |
| `unrecognized_fact_count` | Reported normalized facts that are not one of the case's declared facts. |
| `unrecognized_coverage_declaration_count` | Reported coverage expectations that are not declared for that case. |

WCAB intentionally declares targeted observable facts and explicit boundaries;
it does **not** claim an exhaustive inventory of every workbook difference.
For that reason, the scorer does not call unrecognized facts or declarations
false positives and does not publish a precision metric. They remain reviewable
evidence rather than a silent score adjustment. This follows the basic
distinction between precision and recall used in the [NIST static-analysis
study methodology](https://www.nist.gov/system/files/documents/2021/03/24/CAS%202012%20Static%20Analysis%20Tool%20Study%20Methodology.pdf),
while respecting WCAB's deliberately partial oracle.

Use `--strict` when a complete reference-policy run is required. It exits
nonzero unless every case is complete and no unrecognized normalized fact or
coverage declaration was reported.

## Optional FormulaFence reference adapter

FormulaFence remains optional and is not a package dependency. If a trusted
local executable is installed, WCAB can emit a normalized report from its
existing adapter mapping:

```bash
wcab formulafence-observations --fixtures fixtures \
  --executable formulafence --output formulafence-observations.json
wcab score --fixtures fixtures --observations formulafence-observations.json
```

The adapter copies a WCAB fact into the normalized report only after it matches
the documented native FormulaFence change category. It copies an introduced
dynamic-reference expectation only after FormulaFence emits `FF012`; for a
pre-existing dynamic formula with a changed driver, it requires both the
native `value_changed` record and the candidate profile's
`dynamic_reference_cells` feature. For the external-data refresh case, it
requires FormulaFence's `external_data_connections_changed` details to contain
the exact connection ID and `refresh_on_load` false-to-true transition; the
stored control is a fact, not a coverage declaration. For the PivotCache case,
it requires one exact
`pivot_cache_refresh_controls_changed` record and matching `FF023`: its
redacted cache-level profile must retain the worksheet source type, cache ID,
and all controls except `refresh_on_load`, which moves from `false` to `true`.
WCAB's raw validator—not the adapter—then establishes the `Source!A1:B5`
binding, `Report!A1:B2` PivotTable location, stable stored report/dashboard
cells, direct dashboard edge, and cache-definition-only package change.
Neither report refreshes or renders a PivotTable. For the external-workbook
link policy case, it requires the exact `external_data_refresh_settings_changed`
details: `update_links` moves from `never` to `always`, while the three other
workbook-wide refresh controls retain their defaults. For the array-mode case,
it requires `array_formula_mode_changed` at the exact anchor with the declared
legacy-CSE and dynamic modes plus the declared output range; a generic array
change is not enough. For the iterative-calculation case, it requires the
exact `calculation_settings_changed` details: `iterate` moves from `false` to
`true`, while the explicit 100 / 0.001 bounds and other stored calculation
controls remain unchanged. For the precision-as-displayed case, it requires
the exact `calculation_settings_changed` details: `fullPrecision` moves from
`true` to `false`, while the declared input, number format, formula, and all
other stored calculation controls remain unchanged. For the saved-formula-result
case, it requires one `formula_cached_result_changed` record with FormulaFence's
stable two-formula/one-numeric-cache profile, exactly one unexplained cache
change, and material-change evidence. FormulaFence deliberately redacts raw
cache values and formula-cell locations from that record; WCAB's independent
raw validator establishes the declared `Model!B2` location and `20`-to-`25`
value change. For the workbook-date-system case, it requires one
`workbook_date_system_changed` record and `FF117` with normalized
`date1904=false` to `true`, explicit `dateCompatibility=true` on both sides,
and zero unrecognized date controls. WCAB's raw validator—not the adapter—then
establishes the stable serial, style, formulas, direct dependency edges, and
workbook.xml-only package change; neither report calculates a date or infers a
client display. For the active-AutoFilter case, it requires one exact
`filter_visibility_controls_changed` record and `FF036` with FormulaFence's
redacted one-filter-column/one-criterion profile and material-definition flag.
WCAB's raw validator—not the adapter—then establishes the `North`-to-`South`
criterion, stable `SUBTOTAL` and dashboard formulas, direct dependency edge,
and report-worksheet-only package change. Neither report applies the filter,
calculates a subtotal, or infers a visible row set. The adapter leaves `review`
as `null`: FormulaFence surfaces evidence but does not impose WCAB's reference
policy as a universal approval decision. Any intentionally unmapped fact or coverage expectation
becomes an adapter note and remains missing from its respective recall metric.
