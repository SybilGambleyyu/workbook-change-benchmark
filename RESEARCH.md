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

## Precision as displayed and irreversible stored values

Microsoft's [calculation and precision guidance](https://support.microsoft.com/en-US/Excel/change-formula-recalculation-iteration-or-precision-in-excel)
distinguishes the values Excel stores from the values a number format displays.
It warns that enabling calculation using displayed values permanently changes
stored values to that displayed precision and cannot restore the prior
underlying values. Its separate [rounding-precision guidance](https://support.microsoft.com/en-us/excel/set-rounding-precision)
also warns of cumulative calculation effects. The Open XML
[`calcPr` specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties?view=openxml-3.0.1)
identifies `fullPrecision` as the stored calculation control.

That creates a review boundary that a formula or cell-value diff can miss: a
package can retain a stored `10.005` input, a `0.00` format, and unchanged
formulas while its calculation control moves from full precision to displayed
precision. WCAB 0.10 therefore isolates raw `fullPrecision=true` to
`fullPrecision=false` in an original synthetic pair. The validator confirms
the stored metadata, input value, format, and formulas only. It does not open
or save either workbook, emulate Excel, calculate a result, round a stored
value, or assert any client will apply or persist the control.

## Saved formula results without a visible precedent edit

Microsoft's [SpreadsheetML formula guidance](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas)
separates a formula's `<f>` expression from the `<v>` cached value based on
the last calculation. Its [`CellValue` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-3.0.1)
also states that a formula-cell `<v>` stores the last calculated result. This
is a distinct review surface: an ordinary formula or literal-value diff can
stay silent when the saved result evidence has changed.

WCAB 0.11 therefore retains an original `Inputs!B2=10`, the direct
`Model!B2 = Inputs!$B$2*2` formula, stable calculation properties, and a local
`Dashboard!B4` consumer, while changing only `Model!B2`'s raw numeric `<v>`
from `20` to `25`. The pair is deliberately not a claim that either value is
correct, stale, tampered, or will be displayed after opening: volatile,
external, and client-specific calculation behavior are out of scope. Its
purpose is to make the saved-result discrepancy observable and reviewable.

## Workbook serial-date systems without a cell edit

Microsoft documents that Excel workbooks use either a 1900 or 1904 date
system, and that the same stored serial differs by 1,462 days between those
systems in its [date-system guidance](https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487). The Open XML
[`workbookPr` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1)
stores `date1904` and `dateCompatibility` as workbook controls, with defaults
of `false` and `true` respectively. That makes an omitted control materially
different from an arbitrary boolean parser: a comparator should normalize the
specified defaults before deciding that a date-system change occurred.

WCAB 0.12 isolates that surface in an original synthetic pair. It makes both
controls explicit, preserves a raw `45292` numeric serial and its
`yyyy-mm-dd` custom format, and changes only `date1904=false` to `true` while
`dateCompatibility=true` remains fixed. The validator reads local OOXML and
style metadata only. It neither converts the serial nor asserts what date a
client will display, and it does not calculate, open, or save a workbook.

## Active AutoFilter criteria without a cell edit

Microsoft's [filter guidance](https://support.microsoft.com/en-us/excel/get-started/filter-data-in-a-range-or-table-in-excel)
explains that AutoFilter shows the requested subset and hides the rest; a
filtered subset can then be copied, charted, or printed. Its
[`SUBTOTAL` reference](https://support.microsoft.com/en-us/excel/functions/subtotal-function)
states that rows excluded by a filter are always excluded. Consequently, an
active criterion can be review-material even when no formula text or stored
cell value changes.

WCAB 0.13 records that control boundary in one original synthetic pair. The
sole list criterion in `Report!A1:B5` moves from `North` to `South`, while
`Report!D2=SUBTOTAL(109,B2:B5)` and `Dashboard!B4=Report!$D$2` remain unchanged.
The raw validator checks the stored `<autoFilter>`, `<filterColumn>`,
`<filters>`, and `<filter>` declaration plus the formula/dependency and
package-member boundary. It does not apply a filter, calculate the subtotal,
infer which rows a client shows, or claim a displayed, copied, charted, or
printed result.

## PivotTable cache refresh requests without a cell edit

Microsoft documents an option to [refresh PivotTable data when a workbook
opens](https://support.microsoft.com/en-us/excel/refresh-pivottable-data), and
its [PivotTable overview](https://support.microsoft.com/en-us/excel/overview-of-pivottables-and-pivotcharts)
explains that PivotTables keep a cache that can be shared. The Open XML
[`PivotCacheDefinition` reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivotcachedefinition?view=openxml-3.0.1)
identifies `refreshOnLoad` as the stored instruction for whether an application
refreshes that cache when the workbook opens. This leaves a material review
surface outside worksheet-cell diffs: the request can change without a formula,
source value, or stored display value changing.

WCAB 0.14 isolates this control in an original synthetic package with one local
worksheet cache. `Source!A1:B5` binds through a raw PivotCache/PivotTable
relationship graph to `Report!A1:B2`; its stored `Report!B2` is directly
referenced by `Dashboard!B4`. The pair changes only
`pivotCacheDefinition/@refreshOnLoad` from false to true. The raw validator
checks the relationship binding, source declaration, stored report/dashboard
cells, and package-member isolation. It does not open Excel, refresh data,
calculate or render the PivotTable, infer a new report result, or claim that a
client honors the setting.

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
