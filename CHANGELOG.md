# Changelog

## 0.5.0 — 2026-08-02

- Add deterministic `INDIRECT` address-driver and `OFFSET` displacement-driver
  cases. In both, formula text stays unchanged while the selected effective
  reference can change.
- Add the scoreable `dynamic_reference_driver_changed` coverage expectation
  and validator invariant for a changed literal driver feeding an unchanged
  dynamic formula.
- Extend the optional FormulaFence adapter to pair its native input-change
  record with its candidate-profile dynamic-reference feature, preserving the
  analysis boundary without claiming formula evaluation.

## 0.4.0 — 2026-08-02

- Upgrade the fixture contract to schema version 3 and add a deterministic
  `INDIRECT`-introduction case with an explicit static-dependency boundary.
- Upgrade the normalized observation protocol to version 2, adding scoreable
  coverage declarations, disclosure recall, and strict-mode handling for
  unrecognized declarations.
- Extend the optional FormulaFence adapter to map native `FF012`
  dynamic-reference coverage evidence without imposing an approval decision.

## 0.3.0 — 2026-08-02

- Add a tool-neutral normalized observation protocol, template command, and
  deterministic score report for expected-fact recall, coverage, and
  reference-policy agreement.
- Add an optional FormulaFence normalizer that emits mapped observations while
  leaving approval-policy decisions explicitly unset.
- Keep unrecognized facts explicit rather than calling them false positives
  against WCAB's intentionally targeted oracle.

## 0.2.0 — 2026-08-02

- Upgrade the fixture contract to schema version 2 and add a deterministic
  Excel Table scope-expansion case with an unchanged structured-reference
  formula.
- Validate the Table definition and structured-reference boundary directly,
  and map the case to FormulaFence's table-definition diff evidence.

## 0.1.1 — 2026-08-02

- Add a deterministic JSONL case catalogue with exact workbook byte counts and
  SHA-256 digests.
- Add a `wcab manifest` command and a Hugging Face dataset card for direct
  dataset consumption.

## 0.1.0 — 2026-08-01

- Introduced 16 deterministic synthetic workbook-change cases across finance,
  operations, governance, structural, and multi-workbook portfolio workflows.
- Added a tool-neutral JSON truth contract, a local static-impact lower-bound
  validator, and byte-reproducibility checks.
- Added an optional FormulaFence reference adapter with explicit unmapped-case
  reporting and targeted lint checks.
