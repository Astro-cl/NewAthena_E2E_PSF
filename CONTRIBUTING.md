Contributing
============

Thank you for contributing to this repository. This project is focused on
simulating and assembling per-MM PSF distributions and applying instrument
perturbations (alignment, gravity offload, thermal) plus polar vignetting
factors.

Quick guidelines
----------------
- Run unit tests (pytest) before submitting changes.
- Keep changes small and focused;
- Preserve public function signatures when possible; prefer adding new
  helper functions instead of changing existing contracts.

Local development
-----------------
- Create a virtualenv and install dependencies from `requirements.txt`.
- Launch the GUI (interactive) with:

```bash
source .venv/bin/activate
python3 gui_distributions.py
```

- Run a headless generation and save a plot with:

```bash
python3 main.py -f Distributions/NewTest_Distribution.xlsx -o /tmp/out.png
```

Reporting issues
----------------
Open an issue with a short reproduction (commands, sample workbook) and any
stack traces. Prefer attaching a small Excel file that reproduces the problem.
