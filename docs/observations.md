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
`dynamic_reference_cells` feature. The adapter leaves `review` as `null`:
FormulaFence surfaces evidence but does not impose WCAB's reference policy as
a universal approval decision. Any intentionally unmapped fact or coverage
expectation becomes an adapter note and remains missing from its respective
recall metric.
