# -*- coding: utf-8 -*-
"""
Coverage Debt at Scale - re-render Fig. 1 for print
===================================================
Fig. 1 went into the submission at ~158 dpi with sans-serif tick labels around
4-5 pt, which is under the IEEE floor. This re-renders it at 600 dpi with serif
labels at 8 pt.

A figure is content, so the re-render is gated on proof that nothing changed:
the 10x18 presence grid is rebuilt from outputs/embodiment_skill_matrix.csv and
asserted, cell by cell, against the grid read pixel-wise out of the figure that
is actually embedded in the paper. If any of the 180 cells disagree the script
refuses to write.

Unchanged: colours, row/column order, axis captions, no internal title.
Changed:   typeface, label size, resolution, and the aspect ratio (it was
           embedded at 2.208 against a true 2.070, i.e. stretched ~7%).

Env: Python 3.12, Windows (I:/ROBO).
"""

import numpy as np
import pandas as pd
from PIL import Image
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap

CSV = "../analysis/outputs/embodiment_skill_matrix.csv"
REF = "figures/fig1_heatmap.png"
OUT = "figures/fig1_heatmap_600dpi.png"
W_IN = 6.9                      # printed width in the two-column layout

EMB = ["franka", "widowx", "tienkung", "fanuc", "ur5", "google_robot",
       "hello_stretch", "xarm", "jaco", "dlr_edan"]
SK = ["pick", "close", "open", "place", "pour", "move", "push", "press", "wipe",
      "wrap", "stack", "insert", "rotate", "pull", "cut", "fold", "screw", "sort"]

# dataset/view-specific ids -> robot family, same normalisation as Sec. III-A
FAMILY = {"franka": "franka", "h5_franka_3rgb": "franka",
          "h5_franka_fr3_dual": "franka", "widowx": "widowx",
          "h5_tienkung_gello_1rgb": "tienkung",
          "h5_tienkung_prod1_gello_1rgb": "tienkung",
          "h5_tienkung_xsens_1rgb": "tienkung", "fanuc_mate": "fanuc",
          "ur5": "ur5", "h5_ur_1rgb": "ur5", "google_robot": "google_robot",
          "hello_stretch": "hello_stretch", "xarm": "xarm", "jaco_2": "jaco",
          "dlr_edan": "dlr_edan"}

COVERED, EMPTY = "#2a7d4f", "#f2f2f2"


def from_csv():
    df = pd.read_csv(CSV, index_col=0)
    df = df[df.index.isin(FAMILY)].rename(index=FAMILY).groupby(level=0).sum()
    return (df.reindex(index=EMB, columns=SK).fillna(0).values > 0).astype(int)


def from_embedded_png():
    a = np.array(Image.open(REF).convert("RGB")).astype(int)
    green = (a[:, :, 1] > a[:, :, 0] + 30) & (a[:, :, 1] > a[:, :, 2] + 20)
    x0, x1, y0, y1 = 134, 1080, 0, 446          # plot area of the embedded png
    out = np.zeros((len(EMB), len(SK)), int)
    for r in range(len(EMB)):
        for c in range(len(SK)):
            cy = int(y0 + (r + .5) * (y1 - y0) / len(EMB))
            cx = int(x0 + (c + .5) * (x1 - x0) / len(SK))
            out[r, c] = green[cy - 6:cy + 6, cx - 6:cx + 6].mean() > .5
    return out


grid, ref = from_csv(), from_embedded_png()
bad = int((grid != ref).sum())
print("cells compared: %d   mismatches: %d" % (grid.size, bad))
print("row sums (csv): %s" % grid.sum(1).tolist())
print("row sums (png): %s" % ref.sum(1).tolist())
if bad:
    raise SystemExit("ABORT: content differs from the embedded figure in %d cells" % bad)

plt.rcParams.update({"font.family": "serif",
                     "font.serif": ["Times New Roman", "Nimbus Roman", "DejaVu Serif"],
                     "mathtext.fontset": "stix"})

fig, ax = plt.subplots(figsize=(W_IN, 3.35), dpi=600)
ax.imshow(grid, cmap=ListedColormap([EMPTY, COVERED]), vmin=0, vmax=1,
          aspect="auto", interpolation="nearest")

ax.set_xticks(np.arange(len(SK)))
ax.set_xticklabels(SK, rotation=45, ha="right", fontsize=8)
ax.set_yticks(np.arange(len(EMB)))
ax.set_yticklabels(EMB, fontsize=8)
ax.set_xlabel("canonical skill  (common -> rare)", fontsize=8.5)
ax.set_ylabel("embodiment (broad -> narrow)", fontsize=8.5)
ax.tick_params(length=2, width=.5, pad=2)

# white separators, as in the original
ax.set_xticks(np.arange(-.5, len(SK), 1), minor=True)
ax.set_yticks(np.arange(-.5, len(EMB), 1), minor=True)
ax.grid(which="minor", color="white", linewidth=.8)
ax.tick_params(which="minor", length=0)
for s in ax.spines.values():
    s.set_visible(False)

fig.tight_layout(pad=0.2)
fig.savefig(OUT, dpi=600, facecolor="white")
print("wrote %s  (%.0f dpi at %.1f in, labels 8 pt)"
      % (OUT, Image.open(OUT).size[0] / W_IN, W_IN))
