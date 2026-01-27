# Changelog

## Unreleased

- Remove in-memory/pickle input flow; materialize multi-sheet CSVs to temporary XLSX.
- Add deterministic per-MM sampling and CSV expansion parity with Excel flow.
- Remove pickle-dependent debug tools and temporary generated inputs.
- Add basic unit tests for CSV multi-sheet parsing and distribution spec parsing.
