# Validation record

This record describes the WCAB 0.44.1 / schema-version-3 validation run. It is
reproducible from this repository; no network service or private workbook is
required.

## Fixture integrity

Environment:

- Python 3.13.9
- openpyxl 3.1.5
- pytest 9.1.1
- ruff 0.16.1

Commands:

```bash
python -m pip install -e '.[dev]'
wcab build --output fixtures
wcab manifest --fixtures fixtures
wcab observation-template --fixtures fixtures --output /tmp/observations.json
wcab score --fixtures fixtures --observations /tmp/observations.json
wcab formulafence-observations --fixtures fixtures \
  --executable formulafence --output /tmp/formulafence-observations.json
wcab score --fixtures fixtures --observations /tmp/formulafence-observations.json
ruff check .
ruff format --check .
pytest
wcab validate --fixtures fixtures
```

Results:

- 61 cases: 60 paired-workbook cases and one directory portfolio case.
- 63 observable truth facts across 51 `block` and 10 `review` cases.
- Three scoreable coverage expectations: one newly introduced `INDIRECT`
  boundary and two unchanged-formula selector changes (`INDIRECT` address text
  and `OFFSET` column displacement).
- 186 generated fixture files: 124 workbooks, 61 truth manifests, and one JSONL
  case catalogue, all generated from source.
- Two hundred ninety-four unit tests passed locally under Python 3.13, including
  independent regeneration and byte-for-byte fixture-tree equality.
- The fixture validator accepted all 61 cases.
- The external-data refresh pair has identical package members except for
  `xl/connections.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its relationship-backed source is a non-routable
  `example.invalid` URL.
- The external-data web-query source pair has identical package members except
  for `xl/connections.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its one relationship-backed connection retains ID 1,
  type 4, name, `refreshOnLoad=false`, `refreshedVersion=1`, workbook
  relationship, content type, saved `ImportedData!B2=100` cell, and
  `ImportedData!B2 → Summary!B2 → Dashboard!B4` formula context while raw
  `webPr/@url` moves from the reserved `approved.example.invalid` endpoint to
  `review.example.invalid`. The validator compares the connection part after
  removing only that URL and does not resolve, open, fetch, authenticate to,
  trust, refresh, calculate, or otherwise interact with either endpoint, or
  claim a client result.
- The OPC package-signature Manifest pair has identical package members except
  `_xmlsignatures/sig1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its root-to-origin and origin-to-signature relationships,
  origin/signature content types, one `SignedInfo` local-object reference,
  stable `Controls!B10=12` value and `Controls!D10=B10*C10` formula, and
  calculation properties remain fixed while the one
  `Object/Manifest/Reference/@URI` moves from the workbook part to the first
  worksheet part. The generated digest and signature values are synthetic: the
  validation run does not verify a digest, signature, transform, certificate,
  identity, trust chain, or consumer decision.
- The OPC package-signature relationship-selector pair likewise has identical
  package members except `_xmlsignatures/sig1.xml`; both archives pass ZIP
  integrity checks and remain readable by openpyxl. Its root-to-origin and
  origin-to-signature relationships, content types, `SignedInfo` local-object
  reference, Manifest URI, one Relationships Transform immediately followed by
  XML C14N, stable `Controls!B10=12` value and `Controls!D10=B10*C10` formula,
  and calculation properties remain fixed while the sole
  `RelationshipReference/@SourceId` selects the root office-document
  relationship before and the root signature-origin relationship after. The
  selected relationship is not proof its target part was signed; the validation
  run does not execute a transform or verify a digest, signature, certificate,
  identity, trust chain, or consumer decision.
- The modern threaded-comment resolution pair has identical package members
  except `xl/threadedComments/threadedComment1.xml`; both archives pass ZIP
  integrity checks and remain readable by openpyxl. Its one top-level comment
  keeps its synthetic text, timestamp, identifiers, cell binding, linked person
  record, content types, worksheet/workbook relationships, stable
  `Controls!B10=12` value, `Controls!D10=B10*C10` formula, and calculation
  properties fixed while raw `threadedComment/@done` changes from `0` to `1`.
  The validator reads only the bounded local OOXML shape and compares the
  comment part after erasing that token. It does not expose or assess comment
  content or identity, prove review or approval, send notifications,
  authenticate or authorize a person, open a client, or claim workflow
  completion.
- The legacy shared-workbook revision-log pair has identical package members
  except `xl/revisions/revisionLog1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its workbook-to-header-to-log graph,
  content types, tracking/history/retention/protection controls, synthetic
  header metadata, revision-record shape, stable `Controls!B10=12` value,
  `Controls!D10=B10*C10` formula, and calculation properties remain fixed while
  one private historic value changes. The validator reads only that bounded
  local OOXML shape and compares the log after erasing the private value. It
  does not expose or validate historic values, locations, author identity,
  timestamps, GUIDs, or relationship IDs; verify provenance; replay history;
  resolve a conflict; prove review/approval; authenticate or authorize a
  person; or claim workflow/client behavior.
- The QueryTable pair has identical package members except for
  `xl/queryTables/queryTable1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its direct `ImportedData` worksheet-to-QueryTable
  and workbook-to-connections relationships, non-routable `example.invalid`
  web-query URL, connection-level `refreshOnLoad=false`, remaining QueryTable
  controls, saved `ImportedData!B2=100` cell, and
  `ImportedData!B2 → Summary!B2 → Dashboard!B4` formula context are unchanged
  while raw `queryTable/@refreshOnLoad` changes from false to true. The
  validation run did not open a connection, fetch a URL, refresh a query,
  materialize rows, calculate a workbook, or claim a successful client refresh.
- The cell-hyperlink target pair has identical package members except for
  `xl/worksheets/_rels/sheet1.xml.rels`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its one `Inputs!B2` worksheet
  hyperlink declaration, visible `Open vendor portal` text, fixed relationship
  ID/type/`TargetMode=External`, calculation properties, and
  `Inputs!B2 → Summary!B2 → Dashboard!B4` formula context remain unchanged
  while raw `Relationship/@Target` moves from the reserved
  `approved.example.invalid` URL to `review.example.invalid`. The validation
  run did not resolve, open, fetch, visit, execute, calculate, or otherwise
  interact with either target, or claim that a client follows one.
- The external-workbook source pair has identical package members except for
  `xl/externalLinks/_rels/externalLink1.xml.rels`; both archives pass ZIP
  integrity checks and remain readable by openpyxl. Its `LinkedModel!B2`
  formula, direct `Dashboard!B4` consumer, calculation properties, workbook
  external-reference binding, externalLink/externalBook declaration, source
  sheet name, relationship IDs/types, content type, and `TargetMode=External`
  remain unchanged while its one externalLinkPath `Relationship/@Target` moves
  from the reserved `approved.example.invalid` source to
  `review.example.invalid`. The validation run read local package parts only:
  it did not resolve, open, fetch, authenticate to, trust, refresh, calculate,
  or otherwise interact with either source, or claim a client updates a link or
  returns a value.
- The external-defined-name source pair has identical package members except
  for `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its one local `ScenarioRate` definition moves exactly
  from `'[WCABApprovedSource.xlsx]Inputs'!$B$2` to
  `'[WCABReviewSource.xlsx]Inputs'!$B$2`, while `Model!B2=ScenarioRate*2`, its
  direct `Dashboard!B4` consumer, calculation properties, sheet declarations,
  and workbook relationships remain unchanged. Neither package has an
  `externalReferences` declaration or an `xl/externalLinks/` member. The
  validation run read local OOXML only: it did not resolve, open, fetch,
  authenticate to, trust, refresh, calculate, or otherwise interact with a
  source, or claim a client resolves the name or returns a value.
- The named-LAMBDA definition pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its one workbook-scoped `ScenarioValue` definition
  moves exactly from `=LAMBDA(rate,amount,rate*amount)` to
  `=LAMBDA(rate,amount,rate*(amount+10))`, while `Inputs!B2=0.08`,
  `Inputs!B3=100`, `Model!B2=ScenarioValue(Inputs!B2,Inputs!B3)`, its direct
  `Dashboard!B4` consumer, calculation properties, sheet declarations, and
  workbook relationships remain unchanged. Neither package has an
  `externalReferences` declaration or an `xl/externalLinks/` member. The
  validation run read local OOXML only: it did not evaluate the LAMBDA body,
  calculate a workbook, infer Excel-version support, or claim a recalculated,
  spilled, or persisted result.
- The Table calculated-column formula pair has identical package members except
  for `xl/tables/table1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its one `Ledger!A1:C4` `ScenarioLedger` Table,
  worksheet-to-Table relationship, headers, `tableColumn` IDs, `autoFilter`,
  raw `C2:C4` formulas, calculation properties, and
  `Dashboard!B4=SUM(ScenarioLedger[Calculated amount])` remain unchanged while
  the raw ID-3 `calculatedColumnFormula` moves exactly from `A2*B2` to
  `A2*(B2+1)`. The validation run did not fill a calculated column, calculate
  a workbook, infer stored results, or claim client behavior.
- The Power Pivot/Data Model relationship pair has identical package members
  except for `xl/workbook.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its one `x15:modelRelationship` moves exactly
  from `SalesModel.CalendarKey → CalendarModel.DateKey` to
  `SalesModel.CalendarKey → CalendarModel.FiscalDateKey`, while its two local
  Tables, `powerPivotData` workbook relationship, `.data` content type,
  calculation properties, and fixed opaque `xl/model/item.data` payload remain
  unchanged. The validation run read the local declaration and payload digest
  only: it did not deserialize a model, evaluate DAX, refresh it, calculate or
  render a report, infer model-to-cell impact, or claim client behavior.
- The XLM Auto_Open binding pair is a real macro-enabled `.xlsm` pair with
  identical package members except for `xl/workbook.xml`; both archives pass
  ZIP integrity checks and remain readable by openpyxl. Its one
  workbook-scoped `_xlnm.Auto_Open` defined name moves exactly from
  `'Macro Automation'!$A$1` to `'Macro Automation'!$A$2`, while its one
  very-hidden `Macro Automation` macro-sheet declaration, workbook
  relationship, macro-enabled workbook and macro-sheet content types, raw
  234-byte `xl/macrosheets/sheet1.xml` payload, and ordinary
  `Inputs!B2 → Model!B2 → Dashboard!B4` formula context remain unchanged. The
  payload contains only `A1=HALT()` and `A2=HALT()`. The validation run read
  local OOXML only: it did not open Excel, enable or execute XLM code, parse or
  emulate macro instructions, resolve a dynamic name, inspect macro-security
  or trust state, infer an automatic dispatch result, calculate a workbook, or
  claim client behavior.
- The sheet-protection sort-permission pair has identical package members
  except for `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its `Controls` worksheet remains
  protected, `Controls!D2=B2*C2` and its direct `Dashboard!B4` consumer remain
  unchanged, and every other protection action lock, style, and calculation
  property remains fixed while raw `sheetProtection/@sort` moves from `1`
  (locked) to `0` (permitted). The validation run did not test a password,
  encryption, authentication, authorization, editable ranges, a client sort
  operation, or a resulting value.
- The protected-range descriptor pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its protected `Controls!B2:B2` target remains locked,
  the range name/reference and legacy verifier stay fixed, and
  `Controls!D2=B2*C2` plus its direct `Dashboard!B4` consumer remain unchanged
  while one standard nested `protectedRange/securityDescriptor` text changes.
  The validator privately reads that compact generated element and compares the
  worksheet after erasing only the descriptor text. Public truth excludes the
  descriptor, range name, and verifier; the validation run does not test a
  password or verifier, encryption, identity, authentication, authorization,
  editable-range enforcement, spreadsheet-client behavior, or a resulting
  value.
- The sensitivity-label metadata pair has identical package members except for
  `docProps/custom.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its document custom-property relationship, root
  `classificationlabels` relationship, LabelInfo `labelList` part, ordinary
  `Controls!B2=12` input, `Controls!D2=B2*C2` formula, direct `Dashboard!B4`
  consumer, and calculation properties remain fixed while one private MIP
  custom-property value changes. The validator reads the exact generated
  relationship/content-type/property envelope privately and compares custom
  properties after erasing only that value. Public truth exposes only counts
  and safe package members; the validation run did not resolve a label, contact
  a policy service, determine effective classification, inspect encryption or
  permissions, infer access, authenticate an identity, enforce policy, or
  claim Office, SharePoint, OneDrive, or storage-service behavior.
- The relationship-bound worksheet-control macro-assignment pair has identical
  package members except for `xl/worksheets/sheet1.xml`; both archives pass ZIP
  integrity checks and remain readable by openpyxl. Its one control declaration,
  one control-properties relationship and Office 2010 properties part,
  content-type binding, ordinary `Controls!B2=12` input,
  `Controls!D2=B2*C2` formula, direct `Dashboard!B4` consumer, and calculation
  properties remain fixed while one private inline macro assignment changes.
  The validator compares the worksheet after erasing only that assignment.
  Public truth exposes safe counts and member paths only; the validation run did
  not load a control, inspect or execute VBA, resolve a macro, authenticate a
  user, evaluate permissions, invoke an Office client, or claim an event result.
- The PivotCache pair has identical package members except for
  `xl/pivotCache/pivotCacheDefinition1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its local `Source!A1:B5` worksheet
  binding, `Report!A1:B2` PivotTable location, stored display cells, and direct
  `Dashboard!B4` formula are unchanged while raw `refreshOnLoad` changes from
  false to true. The validation run did not open Excel, refresh a cache,
  calculate or render a PivotTable, or infer a report result.
- The PivotTable aggregation pair has identical package members except for
  `xl/pivotTables/pivotTable1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its local `Source!A1:B5` binding, cache records,
  `Report!A1:B2` location, stored report/dashboard cells, direct
  `Dashboard!B4` formula, refresh control, and calculation properties are
  unchanged while its one `dataFields/dataField/@subtotal` moves from `sum` to
  `average`. The validation run did not open Excel, refresh, calculate, or
  render a PivotTable, infer a changed display value, or claim client behavior.
- The PivotTable Slicer-selection pair has identical package members except
  for `xl/slicerCaches/slicerCache1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its local `Source!A1:B5` binding,
  PivotCache, `Report!A1:B2` location, stored report/dashboard cells, direct
  `Dashboard!B4` formula, refresh control, and calculation properties are
  unchanged while its one selected `Region` cache item moves from index 0
  (`North`) to index 1 (`South`). The validator follows the workbook-to-
  Slicer-cache-to-PivotCache/PivotTable bindings and does not create a visual
  Slicer, apply a filter, refresh, calculate, render, infer a displayed result,
  or claim client behavior.
- The connection-only Power Query pair has identical package members except
  for `customXml/item1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its `Source!A1:B5` `SourceData` Table, worksheet cells,
  table definition, calculation properties, metadata, and permission controls
  are unchanged while the stored `Table.SelectRows` M literal moves from
  `North` to `South`. The validator follows the package-root custom-XML
  relationship and bounded generated Data Mashup envelope; it does not execute
  M, apply the filter, refresh a query, materialize output, calculate a
  workbook, infer returned rows, or claim client behavior.
- The Scenario Manager pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its one selected, locked scenario retains its
  selection/protection, comment/user, summary-reference, second-input, and
  number-format metadata, plus visible `Inputs!B2=0.1`, `Inputs!B3=125`,
  `Inputs!D2=B2*B3`, and `Dashboard!B4=Inputs!$D$2`. Only the raw alternate
  `inputCells/@val` for `B2` moves from `0.08` to `0.16`. The validation run
  did not show or apply a scenario, calculate a workbook, create a Scenario
  Summary, infer an output, or claim client behavior.
- The What-If Data Table pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its raw `Sensitivity!D3` master keeps
  `ref="D3:D5"`, `ca="1"`, column orientation, the input grid, ordinary
  formulas, calculation properties, and saved table results empty while its one
  `r1` input reference moves from `B2` to `B3`. The validation run did not
  substitute inputs, calculate a workbook or Data Table, infer a table output,
  resolve a circular dependency, or claim client behavior.
- The list data-validation source pair has identical package members except
  for `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its one `Inputs!B2` list rule retains the target
  range, error/prompt/dropdown metadata, both `Lists` source columns, current
  `Draft` input, ordinary formulas, and calculation properties while raw
  `formula1` moves exactly from `=Lists!$A$2:$A$4` to `=Lists!$B$2:$B$4`.
  The validation run did not evaluate either source, decide whether a future
  input is valid, accept/reject an entry, calculate a workbook, or claim
  client behavior.
- The conditional-formatting threshold pair has identical package members
  except for `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its one `Operations!B2:B4` `cellIs`
  rule retains its target, priority, `greaterThan` operator, differential red
  fill, and stored `10`, `75`, and `120` metrics while raw `formula` moves from
  `100` to `50`. The validation run did not evaluate the rule, determine a
  rendered cell format, calculate a workbook, or claim client behavior.
- The custom number-format pair has identical package members except for
  `xl/styles.xml`; both archives pass ZIP integrity checks and remain readable
  by openpyxl. Its one `Operations!B2` style index, raw `0.125` value, and
  neighboring `B3=B2` formula remain unchanged while the referenced custom
  `numFmt/@formatCode` moves from `0.0%;[Red](0.0%);-` to `;;;`. The validator
  compares styles XML after erasing only that code. It does not render a
  number format, resolve locale or column width, determine a displayed value,
  calculate a workbook, or claim client behavior.
- The ignored-error suppression pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its `Operations!B2=10`, `B3=20`, `B4=30`,
  `B5=SUM(B2:B3)`, and `C5=B5` records remain unchanged while one standard
  `ignoredErrors/ignoredError` declaration with `sqref="B5"` and
  `formulaRange="1"` is added. The validator compares worksheet XML after
  removing only that declaration. It does not determine whether Excel would
  show a warning, evaluate the formula, render an indicator, decide whether the
  warning is justified, calculate a workbook, or claim client behavior.
- The workbook-structure-protection pair has identical package members except
  for `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its hidden `ReviewControls` sheet and
  `Inputs!D2=B2*C2` formula remain unchanged while raw
  `workbookProtection/@lockStructure` moves from `1` to `0`. The validator
  compares the workbook XML after removing only that attribute. It did not test
  a password, encryption, authentication, authorization, exposure of a hidden
  sheet, or a particular Excel client's sheet-operation behavior.
- The chart-series pair has identical package members except for
  `xl/charts/chart1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its `Dashboard!D2` anchor, title/category references,
  source cells, and all worksheet parts remain unchanged while raw
  `c:ser/c:val/c:numRef/c:f` moves from `Source!$B$2:$B$4` to
  `Source!$C$2:$C$4`. The validation run did not open Excel, calculate source
  cells, refresh chart data, render a chart, or infer a visual difference.
- The external-workbook-link policy pair has identical package members except
  for `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its unchanged `LinkedModel!B2` formula names the
  absent synthetic `WCABSource.xlsx`, while `workbookPr/@updateLinks` changes
  exactly from `never` to `always`.
- The iterative-calculation pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain readable
  by openpyxl. Its unchanged `Model!B2` direct self-reference records the same
  explicit 100 / 0.001 bounds in both workbooks while `calcPr/@iterate` changes
  exactly from false to true.
- The precision-as-displayed pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain readable
  by openpyxl. Its stored `Inputs!B2=10.005`, `0.00` format, formula, and local
  downstream reference are unchanged while `calcPr/@fullPrecision` changes
  exactly from true to false. The validation run did not open, calculate, or
  save either workbook.
- The workbook-date-system pair has identical package members except for
  `xl/workbook.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its raw `Inputs!B2=45292` numeric serial,
  `yyyy-mm-dd` style, local `Model!B2` and `Dashboard!B4` formulas, and
  explicit `dateCompatibility=true` control are unchanged while
  `workbookPr/@date1904` changes exactly from false to true. The validation run
  did not calculate, convert a serial, predict a displayed date, open, or save
  either workbook.
- The active-AutoFilter pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its sole raw `Report!A1:B5` column-0 list value
  changes from `North` to `South`, while its `SUBTOTAL(109,B2:B5)` and
  dashboard formulas remain unchanged. The validation run did not apply a
  filter, calculate a subtotal, infer a visible row set, or open or save either
  workbook.
- The saved Named Sheet View pair has identical package members except for
  `xl/namedSheetViews/namedSheetView1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its `Report!A1:B5` worksheet
  AutoFilter retains no active criterion and its `xr:uid` binding, rows,
  `Report!D2=SUBTOTAL(109,B2:B5)`, and `Dashboard!B4=Report!$D$2` remain
  unchanged while the relationship-backed saved view moves its sole column-0
  list value from `North` to `South`. The validator follows the worksheet
  relationship, confirms the content type and base-filter binding, and does
  not activate, render, or apply the view, calculate a subtotal, infer visible
  rows, or claim a client display or print outcome.
- The XML Map pair has identical package members except for
  `xl/tables/table1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. Its synthetic local MapInfo/XSD declaration,
  file-binding metadata, `Export!E2` single-cell mapping, table values,
  `Export!D2=SUM(InvoiceLines[Net amount])`, and
  `Dashboard!B4=Export!$D$2` remain unchanged while the mapped table
  column XPath moves from `NetAmount` to `TaxAmount`. The validator follows
  only local workbook and worksheet relationships and does not access a file,
  validate a schema, import/export XML, materialize data, calculate a result,
  or claim client behavior.
- The Office Web Add-in pair has identical package members except for
  `xl/webextensions/webextension1.xml`; both archives pass ZIP integrity
  checks and remain readable by openpyxl. Its workbook-to-taskpane-to-web-
  extension relationship chain, synthetic local FileSystem reference, add-in
  IDs, hidden locked pane, ordinary cells, calculation properties, and
  `Inputs!B2 → Model!B2 → Dashboard!B4` formulas remain unchanged while raw
  `Office.AutoShowTaskpaneWithDocument` moves exactly from false to true. The
  fixture has no manifest payload or external relationship. The validation run
  did not install, load, execute, or fetch an add-in or manifest, or claim
  that a task pane opens or that the add-in accesses a workbook cell.
- The embedded-OLE auto-load pair has identical package members except for
  `xl/worksheets/sheet1.xml`; both archives pass ZIP integrity checks and remain
  readable by openpyxl. It retains one standard direct internal OLE-object
  relationship, the matching content type, fixed opaque ASCII payload bytes, a
  synthetic unregistered ProgID, and unchanged `Inputs!B2 → Model!B2 →
  Dashboard!B4` formula context while raw `oleObject/@autoLoad` moves exactly
  from false to true. It has no linked target, ActiveX control, presentation,
  macro, or external relationship. The validation run did not deserialize,
  open, render, execute, register, or invoke an object server, or claim that
  any object loads successfully.
- The saved-formula-result pair has identical package members except for
  `xl/worksheets/sheet2.xml`; both archives pass ZIP integrity checks and
  remain readable by openpyxl. Its direct input, formula expression,
  calculation properties, and local downstream formula are unchanged while
  the raw `Model!B2` numeric `<v>` result changes exactly from `20` to `25`.
  The validation run did not calculate, validate, or interpret either result.
- The array-mode pair preserves the `Model!B1` `=LEN(Inputs!A1:A3)` anchor and
  its stored `B1:B3` range. The candidate adds `xl/metadata.xml`, one
  `sheetMetadata` relationship, one content-type override, and `cm=1` on the
  anchor; both archives pass ZIP integrity checks and remain readable by
  openpyxl. The raw validator classifies the pair exactly as `legacy_cse` then
  `dynamic` without calculating a value or predicting a spill extent.

## Distribution supplement

The 0.44.1 release retains the one-row-per-case `manifest.jsonl` catalogue and
the tool-neutral observation protocol at version 2. Each catalogue row retains
the schema-version-3 truth contract and includes byte counts and SHA-256
digests for the workbooks it names.

Commands:

```bash
python -m build --outdir /tmp/wcab-v0441-dist
twine check /tmp/wcab-v0441-dist/*
python -m venv /tmp/wcab-v0441-wheel-test
/tmp/wcab-v0441-wheel-test/bin/python -m pip install \
  /tmp/wcab-v0441-dist/workbook_change_benchmark-0.44.1-py3-none-any.whl
/tmp/wcab-v0441-wheel-test/bin/python -c 'import wcab; print(wcab.__version__)'
/tmp/wcab-v0441-wheel-test/bin/wcab validate --fixtures fixtures
/tmp/wcab-v0441-wheel-test/bin/wcab manifest --fixtures fixtures --output /tmp/wcab-v0441-wheel-manifest.jsonl
cmp fixtures/manifest.jsonl /tmp/wcab-v0441-wheel-manifest.jsonl
/tmp/wcab-v0441-wheel-test/bin/wcab observation-template --fixtures fixtures \
  --output /tmp/wcab-v0441-wheel-observations.json
/tmp/wcab-v0441-wheel-test/bin/wcab score --fixtures fixtures \
  --observations /tmp/wcab-v0441-wheel-observations.json
python -m venv /tmp/wcab-v0441-sdist-test
/tmp/wcab-v0441-sdist-test/bin/python -m pip install \
  /tmp/wcab-v0441-dist/workbook_change_benchmark-0.44.1.tar.gz
/tmp/wcab-v0441-sdist-test/bin/python -c 'import wcab; print(wcab.__version__)'
/tmp/wcab-v0441-sdist-test/bin/wcab validate --fixtures fixtures
/tmp/wcab-v0441-sdist-test/bin/wcab manifest --fixtures fixtures --output /tmp/wcab-v0441-sdist-manifest.jsonl
cmp fixtures/manifest.jsonl /tmp/wcab-v0441-sdist-manifest.jsonl
```

Results:

- The source distribution and universal wheel passed `twine check`.
- Final wheel and source-distribution SHA-256 values are recorded alongside
  their uploaded assets in the GitHub release, avoiding a self-referential
  source-distribution checksum in this record.
- Fresh Python 3.13 wheel and source-distribution installations both reported
  version 0.44.1 and validated all 61 fixtures; both emitted byte-identical
  JSONL output.
- The full 294-test suite, lint, and format checks passed locally under Python
  3.13.
- The generated unsupported template scored as zero analyzed coverage, zero
  expected-fact recall, and zero coverage-disclosure recall, confirming that
  unsupported cases cannot become a pass.
- The FormulaFence normalizer emitted 62 matched facts, one intentionally
  unmapped fact, three matched coverage declarations, and no invented review
  disposition.

## FormulaFence reference adapter

FormulaFence 0.225.0 was installed from its released wheel and invoked only
against WCAB's generated files:

```bash
wcab formulafence --fixtures fixtures --strict
```

Results:

- All 62 currently mappable diff/portfolio facts were observed.
- All three mappable coverage expectations were matched; no mapped fact,
  coverage expectation, or targeted lint rule was missed across all 61 cases.
- The schema-version-2 structured Table scope case was observed as a
  `table_definition_changed` diff, even though its summary formula text stays
  unchanged.
- The WCAB 0.34 named-LAMBDA case was observed as one exact
  `ScenarioValue` `defined_name_changed` record and `FF008`, including the
  generated before/after LAMBDA texts. FormulaFence did not expose the
  package-member boundary or evaluate the name; WCAB independently established
  the workbook-XML-only difference, stable inputs/formulas, and absence of
  external-link declarations. Neither report calculated a workbook or claimed
  a result.
- The WCAB 0.34 Table calculated-column formula case was observed as one exact
  `table_definition_changed` record and high-severity `FF013`. FormulaFence's
  redacted profile retained the `ScenarioLedger` identity, `Ledger!A1:C4`
  scope, headers, and row-count controls while reporting that its calculated
  column formula materially changed; it did not expose the master text or a
  package-member boundary. WCAB independently established the exact raw
  ID-3 formula transition, stable row formulas/dashboard formula, and
  Table-part-only difference. Neither report filled a column, calculated a
  workbook, inferred results, or claimed client behavior.
- The WCAB 0.35 Power Pivot/Data Model relationship case was observed as one
  exact high-severity `power_pivot_data_model_changed` record and `FF033`.
  FormulaFence's redacted profile retained one internal data part, one workbook
  binding, one declaration, two model tables, one model relationship, one
  fingerprinted payload, and no coverage gaps; only its declaration-change flag
  differed. FormulaFence did not expose model names, relationship keys, DAX,
  targets, XML, or payload bytes. WCAB independently established the exact raw
  `DateKey`-to-`FiscalDateKey` transition, fixed opaque payload, and
  workbook-XML-only difference. Neither report deserialized a model, evaluated
  DAX, refreshed it, rendered a report, inferred model-to-cell impact, or
  claimed client behavior.
- The WCAB 0.36 XLM Auto_Open binding case was observed as one exact
  high-severity `xlm_automatic_macro_bindings_changed` record and `FF076`.
  FormulaFence's redacted profile retained one automatic-macro binding and one
  `Auto_Open` binding, with no close/activate/deactivate binding; only its
  material-change flag differed. It did not expose macro-sheet names, target
  cells, XML, byte payloads, or package-member boundaries. WCAB independently
  established the exact `$A$1`-to-`$A$2` stored target transition, fixed
  byte-identical two-`HALT()` macro sheet, very-hidden state, fixed
  relationship/content types, and workbook-XML-only difference. Neither report
  opened Excel, enabled or executed XLM code, parsed or emulated macro
  instructions, resolved a dynamic name, inferred dispatch behavior,
  calculated a workbook, or claimed client behavior.
- The schema-version-3 introduced-dynamic case was observed as both a
  `dynamic_formula_reference_added` diff and `FF012` coverage warning.
- The two WCAB 0.5 dynamic-driver cases were observed as an `Inputs!E12`
  `value_changed` diff plus a candidate `dynamic_reference_cells` profile
  feature at `Summary!B2`. The adapter preserves that paired evidence as the
  `dynamic_reference_driver_changed` declaration; it does not claim to
  evaluate either selected target.
- The WCAB 0.6 external-data case was observed as the exact
  `external_data_connections_changed` connection-ID-1 `refresh_on_load`
  false-to-true transition and `FF023`. The report omitted the synthetic
  connection name and endpoint; neither tool opened or refreshed it.
- The WCAB 0.37 external-data web-query source case was observed as one exact
  high-severity `external_data_connections_changed` record and `FF023`.
  FormulaFence's before/after profiles retained the same safe one-web-query
  connection description while reporting
  `source_configuration_material_changed=true` and exactly
  `source_material_change_categories: ["web_query_url"]`. It exposed neither
  endpoint, connection string, command, parameter, XML, fingerprint, or
  package-member boundary. WCAB independently established the reserved URL
  transition, fixed connection graph/controls/saved cells/formulas, and
  `xl/connections.xml`-only boundary; neither report opened, fetched,
  authenticated to, trusted, refreshed, calculated, or otherwise interacted
  with a source, or claimed a client result.
- The WCAB 0.38 OPC package-signature Manifest case was observed as one exact
  high-severity `digital_signature_controls_changed` record and `FF050`.
  FormulaFence's redacted before/after profiles retained one signature origin,
  one XML signature, one Manifest direct-part reference, no certificates, no
  VBA-project signatures, no selectors, and no unrecognized signature evidence
  while the aggregate direct-part category moved from one workbook part to one
  worksheet part. It exposed no URI, selector, digest, signature value,
  certificate, identity, or trust assertion. WCAB independently established
  the synthetic direct-part URI transition, fixed relationship graph/content
  types/local-object reference/cells/formula, and
  `_xmlsignatures/sig1.xml`-only boundary; neither report verified a digest,
  signature, transform, certificate, trust chain, or consumer decision.
- The WCAB 0.39 package-signature relationship-selector case was observed as
  one exact high-severity `digital_signature_controls_changed` record and
  `FF050`. FormulaFence's redacted before/after profiles were exactly equal:
  one signature origin, one XML signature, one Manifest relationship reference,
  zero direct-part references, no certificates, no VBA-project signatures, and
  no unrecognized signature evidence. Its distinct
  `package_signature_manifest_coverage_changed=true` signal still exposed the
  material retarget. It exposed no URI, selector, digest, signature value,
  certificate, identity, or trust assertion. WCAB independently established
  the synthetic root-relationship source-ID transition, required
  Relationships-Transform-plus-C14N sequence, fixed graph/cells/formula, and
  `_xmlsignatures/sig1.xml`-only boundary; neither report executed a transform
  or verified a digest, signature, certificate, trust chain, or consumer
  decision.
- The WCAB 0.40 modern threaded-comment resolution case was observed as one
  exact high-severity `threaded_comment_controls_changed` record and `FF045`.
  FormulaFence's redacted profile retained one worksheet-bound comment part,
  one thread/comment/person, no replies or mentions, two internal bindings,
  and no external or unrecognized metadata while `resolved_comment_count`
  moved from `0` to `1` and only its definition-material flag differed. It did
  not expose text, comment-cell reference, timestamp, relationship ID, comment ID,
  person ID, or identity data. WCAB independently established the bounded
  synthetic package graph, stable ordinary cell/formula context, and
  `xl/threadedComments/threadedComment1.xml`-only boundary. Neither layer
  proves review or approval, sends a notification, authenticates or authorizes
  a person, opens a client, or claims workflow completion.
- The WCAB 0.41 legacy shared-workbook revision-log case was observed as one
  exact high-severity `shared_workbook_revisions_changed` record and `FF062`.
  FormulaFence's redacted before/after profile retained one header part/header,
  one revision-log part, three log records, enabled shared/tracking/history/
  retention/protection controls, and no unrecognized metadata while only
  `revision_log_material_changed=true` differed. It did not expose historic
  values, cell locations, author identity, timestamps, GUIDs, or relationship
  IDs. WCAB independently established the bounded synthetic package graph,
  stable ordinary cell/formula context, and
  `xl/revisions/revisionLog1.xml`-only boundary. Neither layer verifies
  provenance or identity, replays a revision, resolves a conflict, proves
  review/approval, authenticates or authorizes a person, or claims workflow
  completion.
- The WCAB 0.30 QueryTable case was observed as one exact
  `query_table_refresh_controls_changed` record and `FF023`: FormulaFence's
  redacted profile retained one `ImportedData` table, connection ID 1, fixed
  background/disable/remove/fill/edit/growth controls, name metadata, and no
  opaque metadata while only `refresh_on_load` moved from false to true. It did
  not expose the endpoint, OOXML part, or result rows. WCAB independently
  established the direct local relationship graph, fixed connection-level
  control, stable saved cells/formulas, and QueryTable-part-only boundary;
  neither report opened a connection, fetched a URL, refreshed a query,
  materialized rows, calculated a workbook, or claimed a client refresh.
- The WCAB 0.31 cell-hyperlink target case was observed as one exact
  `cell_hyperlink_controls_changed` record and `FF047`: FormulaFence's
  redacted profile retained one external worksheet hyperlink binding with no
  location, display, tooltip, or unrecognized declaration while its binding,
  definition material, and relationship material changed. It did not expose
  the target or relationship ID. WCAB independently established the exact
  reserved-target transition, fixed relationship ID/type/mode, stable visible
  text/formulas, and relationship-part-only boundary; neither report resolved,
  opened, fetched, visited, or otherwise interacted with a target, or claimed
  that a client follows it.
- The WCAB 0.32 external-workbook source case was observed as one exact
  `external_link_packages_changed` record and `FF025`: FormulaFence's redacted
  profile retained one external workbook and source sheet, no DDE/OLE link, no
  cached external data, and no opaque metadata while only
  `source_material_changed` was set. It did not expose the source target or
  relationship IDs. WCAB independently established the exact reserved-target
  transition, fixed external-reference graph, stable formula context, and
  externalLink-relationship-part-only boundary; neither report resolved,
  opened, fetched, authenticated to, trusted, refreshed, calculated, or
  otherwise interacted with a source, or claimed a client updates a link or
  returns a value. FormulaFence also emitted its generic
  `external_relationships_changed` / `FF063` diagnostic; the adapter leaves it
  unmapped because it is not sufficient evidence for this narrower fact.
- The WCAB 0.33 external-defined-name source case was observed only when both
  FormulaFence's exact one-surface
  `external_workbook_link_surfaces_changed` / `FF081` evidence and its exact
  `ScenarioRate` `defined_name_changed` / `FF008` evidence agreed on the two
  generated qualified source expressions. The surface ledger alone is not
  sufficient. FormulaFence did not expose a package-member boundary; WCAB
  independently established the exact definition text, absence of an
  `externalLink` package and `externalReferences` declaration, stable formula
  context, and workbook-XML-only difference. Neither report resolved, opened,
  fetched, authenticated to, trusted, refreshed, calculated, or otherwise
  interacted with a source, or claimed a client result.
- The WCAB 0.33 sheet-protection sort-permission case was observed as one exact
  `sheet_protection_changed` record and `FF022`: its protected `Controls`
  profile retained every reported action lock except `sort`, which moved from
  locked to permitted. FormulaFence did not expose the raw OOXML attribute or
  package-member boundary; WCAB independently established the exact
  `sheetProtection/@sort` `1`-to-`0` transition, stable formulas, and
  worksheet-only difference. Neither report tested a password, encryption,
  authorization, an actual client sort operation, or a resulting value.
- The WCAB 0.42 protected-range descriptor case was observed as one exact
  high-severity `protected_range_permissions_changed` record and `FF022`.
  FormulaFence's equal redacted before/after profiles retained one named
  range, one legacy verifier, one standard descriptor, and no opaque metadata
  while only `security_descriptor_material_changed=true` differed. It exposed
  no descriptor, identity, range name, or verifier. WCAB independently
  established the compact nested XML shape, stable locked input/formula path,
  and worksheet-only package boundary. Neither layer tested a password,
  encryption, authentication, authorization, editable-range enforcement, a
  spreadsheet client, or a resulting value.
- The WCAB 0.43 sensitivity-label metadata case was observed as one exact
  high-severity `sensitivity_label_metadata_changed` record and `FF118`.
  FormulaFence's equal redacted before/after profiles retained one relevant
  custom-property part, one standard label-ID property, seven MIP properties,
  one label ID, one LabelInfo part, one internal LabelInfo relationship, and no
  malformed or external metadata while only the private custom-property
  material-change flags differed. It exposed no label ID/name, action/site ID,
  timestamp, property name/value, XML, relationship ID, or target. WCAB
  independently established the exact generated package envelope, stable
  input/formula path, and `docProps/custom.xml`-only boundary. Neither layer
  resolved a label, evaluated policy, determined effective classification,
  inspected encryption or permissions, inferred access, enforced policy, or
  claimed Office, SharePoint, OneDrive, or storage-service behavior.
- The WCAB 0.44.1 worksheet-control macro-assignment case was observed as one
  exact critical `worksheet_embedded_controls_changed` record and `FF029`.
  FormulaFence's equal redacted before/after profiles retained one control
  sheet, one worksheet control, one form-control-properties part, one control
  macro assignment, one related relationship, and no ActiveX, OLE, legacy VML,
  external, or unrecognized surface while only
  `worksheet_control_definition_material_changed=true` differed. It exposed no
  control identity, shape ID, macro name, relationship ID, or raw XML. WCAB
  independently established the private assignment transition, stable
  input/formula path, and worksheet-XML-only package boundary. Neither layer
  loaded a control, inspected or executed VBA, resolved a macro, evaluated
  permissions, or claimed Office-client behavior.
- The WCAB 0.14 PivotCache case was observed as one exact
  `pivot_cache_refresh_controls_changed` record and `FF023`. FormulaFence's
  redacted profile retained one local worksheet cache, ID 1, and all reported
  controls except `refresh_on_load`, which changed from false to true; it
  emitted no parser warning. WCAB's independent raw validation established the
  source and PivotTable binding, stable stored cells, and cache-definition-only
  package change. Neither report refreshed or rendered a PivotTable.
- The WCAB 0.16 PivotTable aggregation case was observed as one exact
  `pivot_table_definitions_changed` record and `FF031`. FormulaFence's redacted
  profile retained one local cache, one PivotTable, one data field, two cache
  fields, four cache records, and no auxiliary material while only its
  PivotTable-layout material changed. It did not expose the source labels,
  selected aggregate function, or a rendered result. WCAB independently
  established the local graph, stable stored cells, exact `sum`-to-`average`
  declaration, and PivotTable-part-only package change. Neither report
  refreshed, calculated, or rendered a PivotTable.
- The WCAB 0.17 PivotTable Slicer-selection case was observed as one exact
  `slicer_timeline_cache_definitions_changed` record and `FF032`. FormulaFence's
  redacted profile retained one Slicer cache, one local PivotCache binding, one
  PivotTable binding, two Slicer items, one selected item, and no timeline or
  auxiliary material while only its Slicer filter-state/definition flag changed.
  It did not expose the Slicer name, selected item/value, or a rendered report.
  WCAB independently established the local graph, stable stored cells, exact
  index-0-`North` to index-1-`South` declaration, and Slicer-cache-part-only
  package change. Neither report created a visual Slicer, applied a filter,
  refreshed, calculated, or rendered a PivotTable.
- The WCAB 0.18 Power Query case was observed as one exact
  `power_query_changed` record and `FF024`. FormulaFence's redacted profile
  retained one parsed mashup, one formula document, three package parts, one
  metadata item, connection-only permission controls, and no opaque or
  embedded content while only `formula_material_changed` changed. It did not
  expose M text, local-table values, or a query result. WCAB independently
  established the local `SourceData` binding, fixed connection-only controls,
  exact `North`-to-`South` M literal, and custom-XML-part-only package change.
  Neither report executed M, refreshed a query, materialized output,
  calculated a workbook, or inferred returned rows.
- The WCAB 0.19 Scenario Manager case was observed as one exact
  `scenario_manager_changed` record and `FF035`. FormulaFence's redacted
  profile retained one scenario-bearing worksheet, one scenario, two stored
  inputs, one selected/locked scenario, one summary reference, one formatted
  input, and no malformed declaration while only
  `scenario_definition_material_changed` changed. It did not expose the
  scenario name, stored values, input references, comment, or user metadata.
  WCAB independently established the `B2` `0.08`-to-`0.16` raw transition,
  fixed metadata and visible cells/formulas, plus the Inputs-worksheet-only
  package change. Neither report showed/applied a scenario, calculated a
  result, or created a Scenario Summary.
- The WCAB 0.20 What-If Data Table case was observed as one exact
  `what_if_data_tables_changed` record and `FF034`. FormulaFence's redacted
  profile retained one Data Table, one column-oriented one-variable table,
  three declared output cells, one recalculation request, no deleted input
  reference, and no unrecognized declaration while only its
  `data_table_definition_material_changed` flag changed. It did not expose the
  output range, local input references, or calculated values. WCAB independently
  established the generated `D3:D5` master, exact `B2`-to-`B3` `r1` transition,
  stable input grid/formulas, ordinary static lower bounds, and
  Sensitivity-worksheet-only package change. Neither report substituted inputs,
  calculated, resolved a circular dependency, or claimed client behavior.
- The WCAB 0.21 list data-validation case was observed as one exact
  `data_validation_changed` record and `FF020`. FormulaFence exposed the one
  `Inputs!B2` list rule's source formula and complete entry-control metadata;
  it did not evaluate either local list or decide whether a future input is
  valid. WCAB independently established the exact `=Lists!$A$2:$A$4` to
  `=Lists!$B$2:$B$4` declaration, stable source values/current input/formulas,
  ordinary static lower bounds, and Inputs-worksheet-only package change.
  Neither report accepted/rejected an input, calculated a workbook, or claimed
  client behavior.
- The WCAB 0.22 conditional-formatting case was observed as one exact
  `conditional_formatting_changed` record and `FF021`. FormulaFence exposed
  the one `Operations!B2:B4` `cellIs` rule and raw `100`-to-`50` formula
  transition but did not render a workbook or determine which cells receive a
  format. WCAB independently established the stable priority, operator,
  differential red fill, metric values, calculation properties, and
  Operations-worksheet-only package change. Neither report evaluated the rule,
  calculated a workbook, or claimed client behavior.
- The WCAB 0.23 custom number-format case was observed as one exact
  `number_format_controls_changed` record and `FF039`. FormulaFence's
  intentionally redacted before/after profiles each retained one direct-cell
  custom assignment and no default, row, column, built-in, or unrecognized
  control while only `number_format_definition_material_changed` was set. It
  did not expose a format code or cell target. WCAB independently established
  the declared code transition, stable target/style/value/formula context, and
  styles-only package change. Neither report rendered a format, calculated a
  workbook, or claimed client behavior.
- The WCAB 0.24 ignored-error case was observed as one exact
  `ignored_error_controls_changed` record and `FF037`. FormulaFence's redacted
  profile moved from no controls to one standard container, one target range,
  and one formula-range-omission category with no unrecognized controls, while
  only `ignored_error_definition_material_changed` was set. It did not expose
  the target range or formula. WCAB independently established the generated
  `Operations!B5` target, `formulaRange=1` declaration, stable cells/formulas,
  and worksheet-only package change. Neither report determined whether Excel
  would show a warning, evaluated a formula, rendered an indicator, or claimed
  client behavior.
- The WCAB 0.25 workbook-structure case was observed as one exact
  `workbook_protection_changed` record and `FF022`. FormulaFence's non-secret
  profile moved from `lock_structure=true` to all workbook locks false, with no
  credential or opaque metadata on either side. WCAB independently established
  the raw `lockStructure=1`-to-`0` transition, stable hidden-sheet/formula
  context, and workbook-XML-only package change. Neither report tested a
  password, encryption, authentication, authorization, or client action.
- The WCAB 0.15 chart-series case was observed as one exact
  `chart_definitions_changed` record and `FF030`. FormulaFence's redacted
  profile retained one host sheet, drawing, legacy chart, and series; three
  data references; no chart cache; and no related payloads, while only its
  chart-definition material changed. FormulaFence did not expose the source
  formula or a visual result. WCAB independently established the local
  relationship chain, `D2` anchor, title/category references, exact value
  reference transition, and chart-part-only package change. Neither report
  calculated or rendered a chart.
- The WCAB 0.8 external-workbook-link policy case was observed as one exact
  `external_data_refresh_settings_changed` transition in `FF023`: only
  `update_links` moved from `never` to `always`; `allow_refresh_query`,
  `refresh_all_connections`, and `save_external_link_values` retained their
  defaults. The report did not expose the synthetic external formula or open
  its absent source workbook.
- The WCAB 0.9 iterative-calculation case was observed as one exact
  `calculation_settings_changed` transition in `FF009`: only `iterate` moved
  from false to true; `iterateCount=100`, `iterateDelta=0.001`, and the other
  stored controls stayed unchanged. FormulaFence's direct-self-reference lint
  surfaced `Model!B2` only while iteration was disabled; neither report claims
  a calculated or converged value.
- The WCAB 0.10 precision-as-displayed case was observed as one exact
  `calculation_settings_changed` transition in `FF009`: only `fullPrecision`
  moved from true to false while the stored input, number format, formula, and
  other calculation controls remained unchanged. Neither report claims a
  rounded stored value, calculated result, or client-side save behavior.
- The WCAB 0.12 workbook-date-system case was observed as one exact
  `workbook_date_system_changed` transition and `FF117`: normalized
  `date1904` moved from false to true while explicit `dateCompatibility=true`
  and the unrecognized-control count of zero remained fixed. FormulaFence
  reported workbook-level evidence only; WCAB's independent raw validation
  established the stable serial, style, formulas, and package boundary. Neither
  report claims a converted or displayed date.
- The WCAB 0.13 active-AutoFilter case was observed as one exact
  `filter_visibility_controls_changed` record and `FF036`: FormulaFence's
  redacted profile retained one worksheet filter, one filter column, one
  criterion, and no other visibility controls while its material-definition
  flag changed. WCAB's independent raw validation established the
  `North`-to-`South` criterion and stable formulas; neither report applies the
  filter, calculates a subtotal, or claims a visible result.
- The WCAB 0.26 saved Named Sheet View case was observed as one exact
  `named_sheet_views_changed` record and `FF038`: FormulaFence's redacted
  profile retained one worksheet, part, named view, filter, column, and
  criterion; zero sort rules or conditions; and zero unrecognized declarations
  while only its Named Sheet View material-definition flag changed. It did not
  expose the view name, IDs, bound range, or list value. WCAB independently
  established the relationship, base-AutoFilter binding, `North`-to-`South`
  criterion, stable formulas, and Named-Sheet-View-part-only boundary; neither
  report activated, rendered, or applied a view, calculated a subtotal, or
  inferred visible rows.
- The WCAB 0.27 XML Map case was observed as one exact
  `xml_mapping_controls_changed` record and `FF049`: FormulaFence's
  redacted profile retained one map part, schema, map, data binding, file
  binding, table binding, and sheet-level single-cell binding with no
  unrecognized mapping metadata while only its binding-material flag changed.
  It did not expose the schema, map, XPath, table, or cell values. WCAB
  independently established the synthetic `NetAmount`-to-`TaxAmount`
  binding transition, stable declarations/formulas, and table-part-only
  boundary; neither report accessed a file, imported/exported XML,
  materialized data, or inferred a result.
- The WCAB 0.28 Office Web Add-in case was observed as one exact
  `office_web_addins_changed` record and `FF028`: FormulaFence's redacted
  profile retained one declared task-pane part, task pane, web-extension part,
  and local store reference; one locked hidden task pane; no bindings,
  snapshots, external relationships, in-content references, or unrecognized
  parts; and an auto-show count from zero to one. It did not expose add-in IDs,
  store name, or the property value. WCAB independently established the local
  relationship graph, false-to-true auto-show property, stable formula context,
  and web-extension-part-only boundary; neither report installed, loaded,
  executed, or fetched an add-in or manifest, or claimed that a task pane
  opens.
- The WCAB 0.29 embedded-OLE auto-load case was observed as one exact
  `worksheet_embedded_controls_changed` record and `FF029`: FormulaFence's
  redacted profile retained one worksheet, one embedded OLE object, one
  internal payload, no linked/external/ActiveX/VML/presentation material, and
  no unrecognized declaration while its auto-load count moved from zero to one.
  It did not inspect or deserialize the opaque payload. WCAB independently
  established the false-to-true attribute transition, fixed relationship and
  content-type boundary, inert synthetic payload, stable formula context, and
  worksheet-XML-only package change; neither report opened, rendered,
  executed, registered, or invoked an object server, or claimed that an object
  loads successfully.
- The WCAB 0.11 saved-formula-result case was observed as one exact
  `formula_cached_result_changed` record and `FF042`: FormulaFence reports
  two formula cells, one numeric saved result, one missing saved result, and
  exactly one unexplained material cache change. It intentionally omits the
  raw result values and formula-cell locations, and neither report claims a
  calculation, stale-cache diagnosis, tampering diagnosis, or displayed value.
- The WCAB 0.7 array-mode case was observed as the exact `Model!B1`
  `legacy_cse`-to-`dynamic` transition with the same stored `B1:B3` range and
  `FF018`. It reported the declared `Dashboard!B2` static impact, emitted no
  parser warning, and did not expose the fixture formula text in the report.
- Five targeted lint expectations were observed: copied-formula interruption
  (`FF082`), conditional-aggregate range shape (`FF093`), explicitly unlocked
  formula cell (`FF085`), incomplete manual calculation (`FF086`), and static
  cycle (`FF090`).
- One fact remains intentionally unmapped: the benign structural formula
  rewrite. It records known before/after locations but does not claim that a
  generic tool can prove Excel semantic equivalence from this small fixture.

This is a baseline implementation result, not a claim that FormulaFence or any
other tool has complete Excel coverage.
