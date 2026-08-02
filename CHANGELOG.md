# Changelog

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
