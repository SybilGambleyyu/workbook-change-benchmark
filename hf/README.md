---
pretty_name: Workbook Change Assurance Benchmark
license: mit
language:
- en
tags:
- excel
- spreadsheet
- benchmark
- audit
- testing
- ci
size_categories:
- n<1K
configs:
- config_name: default
  data_files:
  - split: train
    path: manifest.jsonl
---

# Workbook Change Assurance Benchmark (WCAB)

WCAB is an open, deterministic benchmark for tools that review changes to
Excel workbooks. Each of its 26 synthetic cases supplies a baseline workbook,
a candidate workbook, explicit observable change facts, a reference review
disposition, documented coverage boundaries, and—in relevant cases—a
machine-matchable coverage expectation.

The dataset's machine-readable entry point is `manifest.jsonl`: one JSON
record per case. Every record includes the full truth contract plus exact
relative workbook paths, byte counts, and SHA-256 digests. The corresponding
binary fixtures live under `fixtures/`.

## What this evaluates

WCAB targets spreadsheet-change assurance: whether a diff tool, policy gate,
static analyzer, or editing agent preserves material review evidence when a
workbook changes. The current cases include formula-to-value replacements,
reference drift, input propagation, external references, named ranges,
controls, calculation state, cycles, structural changes (including an Excel
Table scope expansion with unchanged structured-reference text, a newly
introduced `INDIRECT` reference, unchanged `INDIRECT`/`OFFSET` formulas whose
dynamic selectors change, and an external-data connection that begins
refreshing when the workbook opens), an unchanged external-workbook formula
whose `never`-to-`always` open-time update policy changes, an unchanged
direct circular formula whose iterative-calculation setting becomes enabled, an
unchanged precision-sensitive input and formula whose calculation switches to
precision as displayed, an unchanged legacy-CSE array formula that switches to
dynamic-array semantics, a saved numeric formula result that changes despite
an unchanged formula and direct input, and a small multi-workbook portfolio.

The benchmark does **not** evaluate Excel formula execution or claim that
candidate numerical results are correct. `review_expectation` is a transparent
benchmark convention, not a universal business policy. Consumers should retain
each case's `coverage` boundary, disclose applicable `coverage_expectations`,
and report unsupported features explicitly.

## Use

Download `manifest.jsonl` to enumerate the cases, then retrieve the referenced
fixture files. Verify each downloaded workbook with its SHA-256 digest before
running an evaluator. The source repository includes a local validator and a
deterministic fixture generator:

```bash
git clone https://github.com/SybilGambleyyu/workbook-change-benchmark
cd workbook-change-benchmark
python -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
wcab validate --fixtures fixtures
```

See the [source repository](https://github.com/SybilGambleyyu/workbook-change-benchmark)
for the truth schema, tool-neutral observation protocol, validation contract,
and releases. The dataset is MIT licensed.
