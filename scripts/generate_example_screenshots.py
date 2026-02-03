#!/usr/bin/env python3
"""Generate example screenshots (non-interactive) from Test_Distribution.xlsx.

Creates PNGs in the project's `Figures/` directory:
 - gui_load.png
 - aeff_apply.png
 - export_dialog.png
 - vignetting_copy.png

The script tries to mimic the GUI's visible summaries (tables/text) so the
images can be included in docs without manual screenshots.
"""
import sys
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
FIG_DIR = ROOT / 'Figures'
FIG_DIR.mkdir(exist_ok=True)

WB = ROOT / 'Distributions' / 'Test_Distribution.xlsx'
if not WB.exists():
    print('Workbook not found:', WB)
    sys.exit(1)

# Load MM configuration
try:
    mm_cfg = pd.read_excel(WB, sheet_name='MM configuration', engine='openpyxl')
except Exception:
    mm_cfg = None

# Load A_eff raw
try:
    aeff_raw = pd.read_excel(WB, sheet_name='A_eff', engine='openpyxl', header=None)
except Exception:
    aeff_raw = None

# 1) gui_load.png: simple summary
fig, ax = plt.subplots(figsize=(8, 2))
ax.axis('off')
if mm_cfg is not None:
    n = len(mm_cfg)
    ax.text(0, 0.6, f'Loaded: {WB.name}', fontsize=12, weight='bold')
    ax.text(0, 0.1, f'Found {n} MMs in "MM configuration" sheet.', fontsize=10)
else:
    ax.text(0, 0.1, 'Could not read MM configuration', fontsize=10, color='red')
fig.tight_layout()
fig.savefig(FIG_DIR / 'gui_load.png', dpi=150)
plt.close(fig)

# 2) aeff_apply.png: show first 12 MM rows and applied preset (simulate selecting first preset column)
if aeff_raw is not None:
    # attempt to detect preset column: prefer header-like text in row 0 at cols >=3
    preset_col = None
    ncols = aeff_raw.shape[1]
    for c in range(3, min(ncols, 12)):
        v = aeff_raw.iloc[0, c]
        if pd.notna(v) and isinstance(v, str) and any(ch.isalpha() for ch in v):
            preset_col = c
            break
    if preset_col is None:
        # fallback: use column 1 (index 1)
        preset_col = 1
    # build table data
    rows = []
    for r in range(1, min(20, aeff_raw.shape[0])):
        mm = aeff_raw.iloc[r, 0] if aeff_raw.shape[1] > 0 else None
        val = aeff_raw.iloc[r, preset_col] if preset_col < aeff_raw.shape[1] else None
        try:
            mm_i = int(float(mm))
        except Exception:
            continue
        try:
            w = float(val)
        except Exception:
            w = None
        rows.append((mm_i, w))
    df_a = pd.DataFrame(rows, columns=['MM #', 'A_eff (applied)'])
    # render as table
    fig, ax = plt.subplots(figsize=(6, max(2, 0.3 * len(df_a))))
    ax.axis('off')
    tbl = ax.table(cellText=df_a.values, colLabels=df_a.columns, loc='center')
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(10)
    tbl.scale(1, 1.2)
    fig.tight_layout()
    fig.savefig(FIG_DIR / 'aeff_apply.png', dpi=150)
    plt.close(fig)
else:
    print('A_eff sheet not found; skipping aeff_apply.png')

# 3) export_dialog.png: textual preview built from MM config and first enabled data type
fig, ax = plt.subplots(figsize=(8, 4))
ax.axis('off')
lines = []
lines.append(f'Preview — {WB.name}')
if mm_cfg is not None:
    lines.append('MM Configuration (first 8 rows):')
    lines.extend(mm_cfg.head(8).to_string().splitlines())
else:
    lines.append('No MM configuration loaded')

# add MM_PSF preview if present
try:
    mm_psf = pd.read_excel(WB, sheet_name='MM_PSF', engine='openpyxl')
    lines.append('\nMM_PSF (first 8 rows):')
    lines.extend(mm_psf.head(8).to_string().splitlines())
except Exception:
    pass

ax.text(0, 1, '\n'.join(lines), fontsize=8, family='monospace', va='top')
fig.tight_layout()
fig.savefig(FIG_DIR / 'export_dialog.png', dpi=150)
plt.close(fig)

# 4) vignetting_copy.png: show header and two columns, and simulated copy into B (if sheet exists)
try:
    vig = pd.read_excel(WB, sheet_name='Vignetting rotazi', engine='openpyxl', header=0)
    # pick a column to simulate as preset (prefer 2nd column)
    if vig.shape[1] >= 2:
        src_col = vig.columns[1]
        simulated = vig.copy()
        # copy src_col into 'factor' (B)
        simulated.insert(1, 'factor_copy', simulated[src_col].values)
        sim = simulated.iloc[:12, :min(4, simulated.shape[1])]
        fig, ax = plt.subplots(figsize=(6, max(2, 0.3 * len(sim))))
        ax.axis('off')
        tbl = ax.table(cellText=sim.values, colLabels=sim.columns, loc='center')
        tbl.auto_set_font_size(False)
        tbl.set_fontsize(9)
        tbl.scale(1, 1.2)
        fig.tight_layout()
        fig.savefig(FIG_DIR / 'vignetting_copy.png', dpi=150)
        plt.close(fig)
except Exception:
    print('Vignetting sheet not found or unreadable; skipping vignetting_copy.png')

print('Generated example images in', FIG_DIR)
