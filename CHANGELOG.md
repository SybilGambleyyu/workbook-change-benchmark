# Changelog

## 0.11.0 — 2026-08-02

- Add a deterministic saved-formula-result case. Its direct input, formula,
  calculation properties, and local downstream formula remain unchanged while
  the raw numeric `<v>` saved beside `Model!B2`'s `<f>` changes from `20` to
  `25`.
- Add the exact `formula_cached_result_changed` fact, relationship-resolved
  raw-OOXML validator, package-member isolation check, and corruption
  regressions. WCAB records a saved result; it does not calculate a formula,
  decide whether either value is current, stale, tampered, or correct, or
  claim what an Excel client displays after opening.
- Extend the optional FormulaFence adapter to require the exact one-result
  `FF042` evidence profile and unexplained-change count, respecting its
  deliberate redaction of cache values and formula-cell locations.

## 0.10.0 — 2026-08-02

- Add a deterministic precision-as-displayed case. Its stored `10.005` input,
  two-decimal number format, formula, and local downstream reference remain
  unchanged while `calcPr/@fullPrecision` changes from `true` to `false`.
- Add the exact `precision_as_displayed_enabled` fact, raw-OOXML validator, and
  corruption regression coverage. WCAB records the stored calculation control;
  it does not open or save an Excel workbook, calculate a formula, claim a
  rounded value, or claim that a client applies the setting.
- Extend the optional FormulaFence adapter to require the exact isolated
  `FF009` full-precision transition rather than accepting an arbitrary
  calculation-setting difference.

## 0.9.0 — 2026-08-02

- Add a deterministic iterative-calculation case. Its direct self-referencing
  formula and local downstream reference stay unchanged while `calcPr/@iterate`
  changes from `false` to `true` with an explicit 100-iteration / 0.001-delta
  bound.
- Add the exact `iterative_calculation_enabled` fact, raw-OOXML validator, and
  corruption regression coverage. WCAB records the stored calculation control;
  it does not calculate the circular model or claim convergence, a cached value,
  or numerical correctness.
- Extend the optional FormulaFence adapter to require the exact isolated
  `FF009` calculation-settings transition rather than accepting an arbitrary
  calculation-setting difference.

## 0.8.0 — 2026-08-02

- Add a deterministic external-workbook link startup-policy case. Its synthetic
  external formula and downstream local reference stay unchanged while
  `workbookPr/@updateLinks` changes from `never` to `always`.
- Add the exact `external_workbook_link_update_policy_changed` fact, raw-OOXML
  validator, and corruption regression coverage. The source workbook is absent
  by design; WCAB does not resolve, refresh, authenticate to, or calculate it.
- Extend the optional FormulaFence adapter to require the exact isolated
  `FF023` external-link policy transition, rejecting a report that combines it
  with another workbook-wide refresh-control change.

## 0.7.0 — 2026-08-02

- Add a deterministic legacy-CSE to dynamic-array formula-mode case. The
  anchor formula and currently stored output range remain unchanged, while
  raw OOXML metadata changes the formula from fixed output semantics to a
  resizable dynamic array.
- Add the exact `array_formula_mode_changed` fact, independent raw-OOXML
  metadata validator, and a corruption regression test so the contract does
  not rely on the truth file alone.
- Extend the optional FormulaFence adapter to require the exact `FF018`
  legacy-CSE-to-dynamic transition and output range rather than accepting a
  generic array-mode diff.

## 0.6.0 — 2026-08-02

- Add a deterministic relationship-backed external-data connection case where
  only `refreshOnLoad` changes from false to true. The endpoint is synthetic
  and non-routable; WCAB neither opens it nor claims a calculated result.
- Add the `external_data_connection_refresh_on_load_changed` fact and raw-OOXML
  validator for the exact workbook-local connection ID and transition.
- Extend the optional FormulaFence adapter to require the matching `FF023`
  control details, rather than treating an arbitrary connection change as the
  fact.

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
