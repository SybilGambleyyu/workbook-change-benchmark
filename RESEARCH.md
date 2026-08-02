# Research basis

WCAB was selected after examining change-assurance needs across spreadsheet
audit, versioning, repair, and research corpora on 2026-08-01.

## What is already available

| Resource | What it makes measurable | Why it is not sufficient for change assurance |
| --- | --- | --- |
| [Modified EUSES](https://spreadsheets.sai.tugraz.at/index.php/corpora-for-benchmarking/euses/) | Injected formula faults and output-cell test decisions | It is fault-localization/testing data, not an explicit before/after review policy or dependency-impact contract. |
| [VEnron](https://researchportal.hkust.edu.hk/en/publications/venron-a-versioned-spreadsheet-corpus-and-related-evolution-analy/) | Recovered spreadsheet version history | The original work says version information is fragmented and uses recovered evolution groups; it does not publish a CI-oriented accept/review/block oracle for changes. |
| [Enron Error Corpus discussion](https://web-ainf.aau.at/pub/jannach/files/Conference_VL_HCC_2016.pdf) | A small collection of real faults and repairs | It is important evidence that real historical changes can be faults, but it targets fault recovery rather than broad change-review controls. |
| [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench) | Real-world spreadsheet manipulation tasks | It evaluates producing or editing workbooks, not deciding whether a candidate change should pass a review gate. |
| [Formula repair benchmark](https://www.microsoft.com/en-us/research/publication/benchmark-dataset-generation-and-evaluation-for-excel-formula-repair-with-llms/) | Formula runtime-error repair | It evaluates repair generation and execution checks, not version comparison or conservative static impact. |

## Evidence of the workflow problem

Microsoft's [Spreadsheet Inquire documentation](https://support.microsoft.com/en-US/Excel/compare-workbooks-using-spreadsheet-inquire)
describes cell-by-cell comparison, workbook analysis, and relationship diagrams,
but the workflow is a desktop inspection rather than a portable policy gate.
Commercial tools such as [xltrail](https://www.xltrail.com/) provide version
history and component diffs.  Their existence is useful confirmation that
formula-level history matters, while it also leaves room for an open,
reproducible way to evaluate policy and impact claims.

The spreadsheet-evolution literature describes a central review problem: a
structural edit can rewrite many formulas even when intent is preserved, while a
single formula or value change can alter model behaviour.  The paper surfaced
through [this TU Delft copy](https://pure.tudelft.nl/ws/portalfiles/portal/47903754/00_Spreadsheet_Evolution.pdf)
explicitly calls out both effects.  A benchmark must therefore include both
clear regressions and benign-looking structural changes.

## Design decision

The first release uses original generated workbooks instead of rehosting public
corporate corpora.  This provides unambiguous redistribution rights, directly
observable ground truth, deterministic regeneration, and cases that include
modern OOXML controls often absent from older `.xls` research datasets.

WCAB's result is intentionally narrower than a claim of Excel semantic
equivalence.  It provides checkable change facts and static dependency lower
bounds.  A tool that cannot handle a case must report a coverage gap; it cannot
turn a missing observation into an "allow" result.

## Structured-reference scope expansion

Microsoft documents that a structured reference combines a Table and column
name, and that the reference adjusts when data is added to or removed from the
Table. Its [structured-reference guidance](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables)
also shows formulas outside a Table that refer to Table data. This creates a
change-review case that ordinary formula text comparison misses: the stored
formula can remain unchanged while the Table's stored range grows. WCAB schema
version 2 therefore adds a generated Table-expansion case that checks those
observable facts without evaluating a formula or asserting a result value.

## External-workbook link startup policy

Microsoft's [workbook-link guidance](https://support.microsoft.com/en-us/excel/manage-workbook-links)
documents startup choices for external workbook links, including always
refreshing them, never refreshing them, and asking the user. The Open XML
[`workbookPr` specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1)
identifies `updateLinks` as the stored behavior for updating external links
when a workbook opens. This is distinct from a query, connection, or pivot
refresh switch: an unchanged formula that refers to a separate workbook can
start attempting an update solely because this workbook-wide policy changes.

WCAB 0.8 therefore preserves one original synthetic external-link formula and
its local downstream consumer while changing only `workbookPr/@updateLinks`
from `never` to `always`. The named source workbook is deliberately absent;
the validator reads local OOXML and formula text only. It does not resolve the
link, test source availability or trust, retrieve a value, or claim that Excel
successfully recalculates the workbook.

## Iterative calculation and intentional circular models

Microsoft's [circular-reference guidance](https://support.microsoft.com/en-US/Excel/remove-or-allow-a-circular-reference-in-excel)
explains that iterative calculation can intentionally allow circular references
in financial or engineering models, and that maximum iterations and maximum
change bound that repeated calculation. Its [calculation guidance](https://support.microsoft.com/en-US/Excel/change-formula-recalculation-iteration-or-precision-in-excel)
likewise distinguishes a direct or indirect self-reference from a workbook
where iteration has been enabled to calculate it. The Open XML
[`calcPr` specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties?view=openxml-3.0.1)
defines the stored calculation-properties element and its `iterate`,
`iterateCount`, and `iterateDelta` attributes.

WCAB 0.9 uses one small original direct self-reference whose formula text,
input, and downstream local consumer do not change. The candidate changes only
the stored `iterate` flag from false to true while retaining explicit 100 and
0.001 bounds. The fixture records a calculation-control change—not a numerical
outcome. The validator does not calculate the circular model, assume that it
converges, predict how many iterations occur, or assert a cached or
business-correct result.

## Array-formula semantics

Microsoft's [dynamic-array versus legacy-CSE guidance](https://support.microsoft.com/en-US/Excel/dynamic-array-formulas-vs-legacy-cse-array-formulas)
distinguishes a dynamic array entered in one anchor cell from a legacy CSE
formula entered across a fixed range. It also documents that dynamic arrays
resize as source data changes while legacy CSE ranges do not. That creates a
change-review concern beyond formula text: retaining the same array-capable
formula and currently stored range can still change editability, resizing, and
spill-blocker behavior.

WCAB 0.7 therefore uses a small original `LEN` array formula rather than a
private workbook or a calculated result. The candidate adds the compact OOXML
cell-metadata binding used by dynamic-array writers, while the validator checks
only the stored mode, formula text, and range. It deliberately does not
calculate the array, predict a future spill extent, or assert client-version
compatibility.

## Evidence scoring without a false-positive fiction

The [NIST static-analysis methodology](https://www.nist.gov/system/files/documents/2021/03/24/CAS%202012%20Static%20Analysis%20Tool%20Study%20Methodology.pdf)
defines precision as reported true findings divided by all reported findings,
and recall as reported true findings divided by all known findings. WCAB's
truth is intentionally not a complete list of every difference in a workbook:
it declares review-relevant facts plus coverage boundaries. Calling every
unlisted tool observation a false positive would therefore misrepresent the
oracle. Version 0.3 adds a normalized adapter protocol that measures recall of
declared facts, analyzed coverage, and reference-policy agreement, while
preserving unrecognized observations for review instead of forcing a precision
claim.
