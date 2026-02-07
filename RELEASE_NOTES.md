Release notes — 2026-01-25
=================================

Summary
-------

Fixes an issue where workbooks generated from "Variable ..." `MM_PSF` standard
presets sometimes retained the template numeric values instead of sampling
per-MM instances from the preset-specified distributions. The sampling logic
now forces per-index deterministic sampling for any preset that contains
"Variable" or a percent token, and parses percent/alpha tokens before the
sampling-decision is evaluated.

Change details
--------------

- File changed: `sensitivity/sensitivity_run.py`
- Key changes:
  - Moved detection of "force-gamma" (Variable/percent) presets and parsing
    of percent/alpha tokens to execute prior to the sampling-decision check.
  - Construct explicit gamma definitions from parsed percent/alpha tokens
    and the standard preset baseline parameters before sampling.
  - Ensures sampled values overwrite the left-side per-MM numeric columns
    while preserving the right-hand template tail (gamma(...) placeholders).
  - Keeps deterministic seeding: base seed derived from hash(filename + preset),
    per-index seed = base + index.

Why this matters
-----------------

Before this change, some generated inputs (notably presets like
"50% Variable Pseudo-Voigt ...") occasionally wrote the template numeric
value across all 600 micro-mirror (MM) rows instead of producing 600
deterministic random draws from the target distribution. That behavior could
produce misleading inputs and incorrect downstream sensitivity results. The
reorder guarantees Variable presets drive sampling as intended.

Verification
------------

- Ran a full generate-only sweep with `--persist` and `--generate-only` to
  regenerate inputs. All produced workbooks for Variable presets now contain
  numeric, varied per-MM `sigma_rad`/`sigma_azi` columns.
- Ran `sensitivity/validate_mm_psf.py` to aggregate per-workbook sampling
 diagnostics; `sensitivity/input/validation_report.csv` was produced.

Impact & compatibility
----------------------

This is a backwards-compatible behavioral fix that affects only generation of
new input workbooks. Existing generated workbooks (created before this change)
are unchanged. Re-running a generation for the same baseline and combo will
produce identical per-index samples (deterministic seeding) to any other run
using the same filename/preset and index numbering.

Next steps
----------

- Optionally tag a release and/or create a GitHub release for this fix.
- Optionally run an integration sweep on the compute cluster to reconfirm
  end-to-end sensitivities.

Commit: b69a765 (already pushed)

Contact
-------
If you want this note expanded into a formal CHANGELOG entry, release note,
or PR description, tell me which format you prefer and I will prepare it.

## Release v5 (2026-02-07)

- GUI A_eff export: GUI-selected standard presets are evaluated per-MM and
  written numerically into column B of `A_eff`; GUI export clears column C.
- Percent-variable presets are synthesized to explicit gaussian forms
  (e.g. "Variable 10% 1 keV" → `gaussian(L,10%*L)`) to ensure correct
  deterministic per-index sampling.
- Preset energy tokens like `1 keV` are parsed and written into `C2` of
  vignetting sheets at export time.

