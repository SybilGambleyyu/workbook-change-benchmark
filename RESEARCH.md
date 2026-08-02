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
