# WCAB normalized observation protocol, version 1

WCAB evaluates tools that inspect a baseline/candidate workbook change. Native
tool reports differ, so an adapter should translate its result into this small
JSON protocol before invoking `wcab score`. The protocol is deliberately
separate from the fixture truth schema: a tool may keep its own report format,
but its benchmark adapter has one portable target.

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
  "schema_version": 1,
  "benchmark": {"fixture_schema_versions": [2]},
  "tool": {"name": "example-tool", "version": "1.4.0"},
  "cases": [
    {
      "id": "finance.formula_to_value",
      "status": "analyzed",
      "facts": [
        {
          "fact": {
            "kind": "formula_to_value",
            "sheet": "Revenue",
            "cell": "C8"
          },
          "evidence": {"native_rule": "hardcode-replacement"}
        }
      ],
      "review": "block",
      "coverage": []
    }
  ]
}
```

`schema_version` is the observation protocol version. `fixture_schema_versions`
must name the truth-schema versions the adapter analyzed. `tool.name` is
required; `tool.version` is optional.

Each `facts` item wraps an exact WCAB truth fact in `fact`. The optional
`evidence` object is adapter-owned metadata and is never scored; it can record a
native rule ID, report location, or another reproducible pointer without making
one tool's report schema normative.

## Case status and review

| `status` | Required behavior | Score treatment |
| --- | --- | --- |
| `analyzed` | May report facts and one optional `review` disposition. | Its recognized facts can contribute to expected-fact recall. |
| `unsupported` | Must not report facts or a review. | Explicitly counted as unsupported, never as a pass. |
| `error` | Must include a non-empty `error` string and must not report facts or a review. | Counted separately from unsupported coverage. |

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
| `analysis_coverage` | Cases marked `analyzed` divided by all WCAB cases. |
| `review_agreement` | Cases whose analyzed review matches the reference convention divided by all WCAB cases. |
| `complete_case_count` | Analyzed cases with every declared fact matched, no unrecognized fact, and matching reference review. |
| `unrecognized_fact_count` | Reported normalized facts that are not one of the case's declared facts. |

WCAB intentionally declares targeted observable facts and explicit boundaries;
it does **not** claim an exhaustive inventory of every workbook difference.
For that reason, the scorer does not call unrecognized facts false positives
and does not publish a precision metric. They remain reviewable evidence rather
than a silent score adjustment. This follows the basic distinction between
precision and recall used in the [NIST static-analysis study methodology](https://www.nist.gov/system/files/documents/2021/03/24/CAS%202012%20Static%20Analysis%20Tool%20Study%20Methodology.pdf),
while respecting WCAB's deliberately partial oracle.

Use `--strict` when a complete reference-policy run is required. It exits
nonzero unless every case is complete and no unrecognized normalized fact was
reported.

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
the documented native FormulaFence change/rule category, which is preserved in
the optional `evidence` object. It leaves `review` as `null`: FormulaFence
surfaces evidence but does not impose WCAB's reference policy as a universal
approval decision. Any intentionally unmapped WCAB fact becomes adapter
coverage text and remains missing from expected-fact recall.
