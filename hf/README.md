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
Excel workbooks. Each of its 57 synthetic cases supplies baseline/candidate
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
Table scope expansion with unchanged structured-reference text, a Table
calculated-column master whose raw formula changes while worksheet cells and
the structured-reference dashboard formula remain fixed, an embedded Power
Pivot/Data Model relationship whose raw target key changes while its binding,
local Tables, and opaque payload remain fixed, a newly
introduced `INDIRECT` reference, unchanged `INDIRECT`/`OFFSET` formulas whose
dynamic selectors change, an external-data connection that begins refreshing
when the workbook opens, a separate relationship-backed external-data
web-query connection whose stored source URL moves between reserved
`example.invalid` endpoints while its connection identity, refresh controls,
saved cells, and formula context remain fixed, a structurally shaped OPC
package-signature `Object/Manifest` declaration whose one direct-part URI moves
from the workbook part to a worksheet part while its bounded signature graph
and ordinary workbook context remain fixed (with deliberately synthetic digest
and signature values, so this is not cryptographic or trust validation), a
second package-signature Manifest whose root-relationships selector moves
between two relationship entries while its URI, Relationships Transform plus
C14N sequence, and all published selector counts remain fixed (also no
transform execution or trust validation), a worksheet-associated modern
threaded-comment thread whose top-level stored resolution state moves from
unresolved to resolved while its synthetic text, person record, relationships,
content types, ordinary cells, formulas, and every other package member remain
fixed (a stored state only: not proof of review, approval, notification,
identity, authorization, or workflow completion), a
legacy shared-workbook revision log whose one synthetic historic value changes
while its relationship-backed workbook/header/log graph, content types,
tracking/retention controls, record shape, ordinary cells, formulas, and every
other package member remain fixed (stored audit-trail material only: not proof
of provenance, identity, conflict resolution, review, approval, authorization,
or workflow completion), a
local worksheet-backed PivotTable cache whose
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
external-workbook formula whose local `externalLink` package source target
moves between reserved `example.invalid` URLs while its formula text and local
dependency remain fixed, a local `ScenarioRate` defined name whose qualified
external-workbook source text changes while its `Model!B2` formula and local
dashboard consumer remain fixed, a workbook-scoped `ScenarioValue` named
LAMBDA whose stored body changes while its inputs, calling formula, and local
dashboard consumer remain fixed, a protected `Controls` worksheet whose stored
sort permission moves from locked to permitted while its cells and formula
context remain fixed, an unchanged
direct circular formula whose iterative-calculation setting becomes enabled, an
unchanged precision-sensitive input and formula whose calculation switches to
precision as displayed, an unchanged legacy-CSE array formula that switches to
dynamic-array semantics, a saved numeric formula result that changes despite
an unchanged formula and direct input, a workbook-wide 1900-to-1904
serial-date-system setting change with a stable numeric serial, date format,
and local formulas, an active worksheet AutoFilter criterion change with a
stable `SUBTOTAL` and downstream formula, a relationship-backed saved Excel
Named Sheet View whose alternate list criterion changes while its base
AutoFilter, cells, formulas, and dashboard dependency stay fixed, a local XML
Map whose mapped invoice-table column XPath changes while its map/schema/file
binding declarations, single-cell mapping, cells, formulas, and dashboard
dependency stay fixed, and a relationship-backed Office Web Add-in task-pane
declaration whose stored `Office.AutoShowTaskpaneWithDocument` request changes
from false to true while its synthetic local reference, locked hidden pane,
ordinary cells, and formula context stay fixed, and one local worksheet OLE
declaration whose raw `autoLoad` request moves from false to true while its
internal relationship, opaque synthetic bytes, ordinary cells, and formula
context stay fixed, a relationship-backed QueryTable whose own raw
`refreshOnLoad` request moves from false to true while its fixed internal
connection control, non-routable synthetic endpoint, saved cells, and formula
context stay fixed, a relationship-backed worksheet cell hyperlink whose
visible text and local formula context stay fixed while its one external OOXML
relationship target moves between reserved `example.invalid` URLs, and a small
multi-workbook
portfolio.

It also includes a real macro-enabled `.xlsm` pair whose workbook-scoped
`_xlnm.Auto_Open` binding moves from one cell to another on a fixed
very-hidden XLM macro sheet. The raw macro-sheet payload, macro relationship,
content types, ordinary formula context, calculation properties, and every
other package member remain fixed, so the case isolates the stored dispatch
declaration rather than macro content.

The benchmark does **not** evaluate Excel formula execution or claim that
candidate numerical results are correct. `review_expectation` is a transparent
benchmark convention, not a universal business policy. Consumers should retain
each case's `coverage` boundary, disclose applicable `coverage_expectations`,
and report unsupported features explicitly.

The Named Sheet View case records stored OOXML only: it does not activate,
render, or apply the view; calculate a subtotal; infer visible rows; or claim
a client display or print outcome.

The XML Map case records a stored import/export binding only: it does not
access a file, validate a schema, import or export XML, materialize data,
calculate a result, or claim client behavior.

The Office Web Add-in case records a stored package property only: it does not
install, load, execute, or fetch an add-in or manifest, and it does not claim
that a task pane opens or that an add-in accesses workbook cells.

The embedded OLE case records a stored package property only: it does not
deserialize, open, render, execute, register, or invoke an object server, and
it does not claim that the opaque synthetic object loads successfully.

The QueryTable case records a stored package property only: it does not open a
connection, fetch a URL, refresh a query, materialize rows, calculate a
workbook, or claim that a client refreshes successfully.

The cell-hyperlink case records a stored package relationship only: it does not
resolve, open, fetch, visit, execute, calculate, or otherwise interact with a
target, and does not claim that a client follows it.

The external-workbook source case records a stored package relationship only:
it does not resolve, open, fetch, authenticate to, trust, refresh, calculate,
or otherwise interact with a source, and does not claim that a client updates a
link or returns a value.

The external-defined-name source case records local `definedName` text only:
it has no `externalLink` package, does not resolve, open, fetch, authenticate
to, trust, refresh, calculate, or otherwise interact with either synthetic
source, and does not claim that a client resolves the name or returns a value.

The named-LAMBDA case records local `definedName` formula text only: it does
not evaluate the function, calculate a result, infer Excel-version support, or
claim that a client recalculates, spills, or persists a value.

The Table calculated-column case records a Table-level formula master only: it
does not fill a column, reconcile that master with row formulas, calculate a
structured reference, infer a total, or claim client behavior.

The Power Pivot/Data Model case records a raw relationship declaration only:
it does not deserialize the Analysis Services payload, evaluate DAX, refresh a
model, calculate or render a report, infer model-to-cell impact, or claim
client behavior.

The XLM Auto_Open case records a stored dispatch declaration only: it does not
open Excel, enable or execute XLM code, parse or emulate macro instructions,
resolve a dynamic name, inspect macro-security or trust settings, infer a
dispatch result, calculate a workbook, or claim client behavior.

The sheet-protection sort case records a stored action permission only: it does
not test a password, encryption, authorization, editable ranges, an actual
client sort operation, or a resulting value.

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
