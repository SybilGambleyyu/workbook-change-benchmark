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
Excel workbooks. Each of its 41 synthetic cases supplies baseline/candidate
fixtures, explicit observable change facts, a reference review
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
dynamic selectors change, an external-data connection that begins refreshing
when the workbook opens, a local worksheet-backed PivotTable cache whose
stored `refreshOnLoad` request becomes enabled while its source, stored report
cells, and direct dashboard formula remain fixed, a local PivotTable value
field whose stored `subtotal` changes from `sum` to `average` while its source,
cache, and stored report cells remain fixed, a local PivotTable Slicer cache
whose selected `Region` item moves from `North` to `South` while its source,
cache, and stored report cells remain fixed, a connection-only Power Query M
definition over a local `SourceData` Table whose stored filter literal moves
from `North` to `South` while its cells, table, metadata, and permissions
remain fixed, a selected locked Scenario Manager declaration whose stored
alternate `Inputs!B2` value moves from `0.08` to `0.16` while its visible
worksheet cells, formulas, second input, selection/protection metadata, and
calculation properties remain fixed, a column-oriented one-variable What-If
Data Table whose raw master changes its local input reference from `B2` to `B3`
while its output range, input grid, ordinary formulas, and saved table results
remain fixed, a list data-validation rule whose stored source switches between
two local status ranges while its target, metadata, source values, and ordinary
formulas remain fixed, a `cellIs` conditional-formatting rule whose stored
exception threshold moves while its target, priority, operator, differential
fill, and worksheet values remain fixed, a custom number-format declaration
whose code moves while its target style, raw value, and formula context remain
fixed, a stored Excel error-checking declaration whose formula-range suppression
is added while its ordinary cells and formula context remain fixed, and a
Dashboard DrawingML chart
whose stored numeric-series reference changes while its source cells, anchor,
and other chart references remain fixed), an unchanged external-workbook formula
whose `never`-to-`always` open-time update policy changes, an unchanged
direct circular formula whose iterative-calculation setting becomes enabled, an
unchanged precision-sensitive input and formula whose calculation switches to
precision as displayed, an unchanged legacy-CSE array formula that switches to
dynamic-array semantics, a saved numeric formula result that changes despite
an unchanged formula and direct input, a workbook-wide 1900-to-1904
serial-date-system setting change with a stable numeric serial, date format,
and local formulas, an active worksheet AutoFilter criterion change with a
stable `SUBTOTAL` and downstream formula, a relationship-backed saved Excel
Named Sheet View whose alternate list criterion changes while its base
AutoFilter, cells, formulas, and dashboard dependency stay fixed, and a small
multi-workbook
portfolio.

The benchmark does **not** evaluate Excel formula execution or claim that
candidate numerical results are correct. `review_expectation` is a transparent
benchmark convention, not a universal business policy. Consumers should retain
each case's `coverage` boundary, disclose applicable `coverage_expectations`,
and report unsupported features explicitly.

The Named Sheet View case records stored OOXML only: it does not activate,
render, or apply the view; calculate a subtotal; infer visible rows; or claim
a client display or print outcome.

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
