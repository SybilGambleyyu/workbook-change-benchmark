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

Version `0.7` covers formula-to-value replacements, formula reference drift,
value changes with downstream effects, external formula references, named
ranges, data validation, conditional formatting, sheet visibility, direct cell
protection, calculation settings, static cycles, portfolio dependencies,
formula refactors, 3-D-reference scope changes, and Excel Table scope changes
that leave a structured-reference formula textually unchanged. It also covers
new `INDIRECT` references whose dependency target comes from workbook text,
unchanged `INDIRECT` and `OFFSET` formulas whose address or displacement driver
changes, and an external-data connection whose refresh-on-open behavior changes
without any worksheet-cell edit. It also covers an unchanged array formula
switching from a fixed legacy CSE output range to dynamic-array semantics.

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
For the array-mode case, it requires the exact legacy-CSE-to-dynamic transition
and output range in FormulaFence's `FF018` evidence. It keeps unmapped facts
explicit; it does not treat the benign structural-rewrite
annotation as a generic semantic-equivalence claim.

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
