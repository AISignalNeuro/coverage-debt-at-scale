# Coverage Debt at Scale

Quantifying long-tail and safety-critical skill representation across
cross-embodiment robot-learning datasets.

Two things live here:

- **`analysis/`** — the audit itself: the embodiment×skill coverage matrix, the
  rarefaction control on language-diversity metrics, the ontology κ check, and
  the CSVs and figures the paper reports.
- **`paper-format/`** — the toolchain that rebuilds the manuscript on the
  official RA-L Word template, plus a format gate that has to go green before
  the paper is submitted.

Submission 26-4993, IEEE Robotics and Automation Letters. Currently private
while the paper is under review.

---

## What the paper claims

Pooling demonstrations across robots is the default recipe for generalist
manipulation policies, but the field measures what a corpus *covers* with
single diversity scalars computed over language instructions. Two findings:

1. **Those scalars do not measure coverage.** Rarefy DROID, AgiBot and RoboMIND
   to a common instruction budget (N = 479, B = 500 bootstrap resamples) and
   their first-verb entropy intervals overlap. The Gini coefficients do spread,
   but in an order set by annotation style — AgiBot lowest with its controlled
   vocabulary, Bridge highest with its templated pick-and-place phrasing — and
   the two free-language corpora still sit ~0.07 apart. The headline diversity
   gap is mostly sample size and labelling protocol.

2. **There is a structural long tail.** Measured on the one axis those confounds
   do not touch — the embodiment — coverage concentrates hard. Across ten real
   embodiments the Franka family covers all 18 canonical skills and WidowX 16,
   while eight of ten cover at most eleven. Structural Coverage Debt is
   **SCD = 0.567** (jackknife SE 0.10 across embodiments), rising to 0.656 under
   an episode-support threshold of K = 5. The tail is contact-rich: `insert`,
   `stack`, `fold`, `screw`, `cut`, `sort`, each in only two or three
   embodiments.

The ontology is validated against AgiBot's independent expert labels over
1.85×10⁵ action segments: Cohen's κ = 0.61 overall, 0.68 on core skills.

---

## `analysis/`

```
analysis/
  notebooks/project.ipynb    the pipeline, phase 0 through the reviewer close-out
  notebooks/Figure.ipynb     the OXE/LeRobot figure fallback path
  outputs/                   every CSV and figure the paper draws on
```

Key outputs:

| file | what it holds |
|---|---|
| `embodiment_skill_matrix.csv` | raw embodiment × canonical-skill counts, before family normalisation |
| `structural_coverage_debt.csv` | per-skill `d(s)`, embodiments covering, risk weight |
| `backbone_rarefied_ci.csv` | rarefied entropy / Gini / vocab size with 95% bootstrap intervals |
| `master_metrics_final.csv` | full-dataset (un-rarefied) metrics, for the contrast in §IV-A |
| `robomind_failure_by_task_v2.csv` | RoboMIND failure-cause labels behind §IV-D |
| `task_strings.csv` | the raw instruction text the ontology maps from |

The notebooks read dataset metadata from the HuggingFace LeRobot mirrors and
the RoboMIND repository layout — task tables and `meta/info.json` only, no video
downloads. You need your own HF read token; the notebook prompts for it via
`getpass` and nothing is stored in the repo.

### Reproducing the audit

```bash
pip install -r requirements.txt
jupyter lab analysis/notebooks/project.ipynb
```

Cells are ordered by phase. The embodiment×skill matrix comes out of the
"Patha embodiment skill" cell; the four-dataset rarefaction and the M2/M3
close-out are the last two cells.

---

## `paper-format/`

The manuscript came back from the editorial office on formatting. The cause was
not a mis-set margin — the file had never been built on an IEEE template at
all. Its `styles.xml` carried only Word defaults, 115 of its 121 paragraphs were
direct-formatted with no style, and three of its four `sectPr` blocks specified
**A4**. Exported, it produced a PDF whose first page was US Letter and whose
remaining six were A4, which fails a format check on its own.

So this rebuilds it *from* `ieeeconf_letter.dot` rather than restyling it to
match:

```
paper-format/
  make_base.ps1        Word: ieeeconf_letter.dot -> base.docx (+ template_probe.dotx)
  build_paper.py       pours the text into the template's own named styles
  export_pdf.ps1       Word: repaginate, save, export PDF with fonts embedded
  scrub_metadata.py    strips author metadata Word stamps back in on every save
  verify_format.py     13 checks; non-zero exit if any fails
  content_diff.py      proves the rebuild changed no words
  regen_fig1.py        re-renders Fig. 1 at 600 dpi, gated on a content assertion
  FORMAT_REPORT.md     what changed, what was decided, what is still open
```

### Running it

```bash
cd paper-format
powershell -File make_base.ps1
python regen_fig1.py
python build_paper.py
powershell -File export_pdf.ps1
python scrub_metadata.py
python verify_format.py
python content_diff.py
```

Needs Word (COM) on Windows. `ieeeconf_letter.dot` is a binary Word 97-2003
template, so python-docx cannot open it directly — `make_base.ps1` is the bridge.

`scrub_metadata.py` is not optional. Word stamps Author and Last author from the
install's registered user on every save and ignores attempts to blank them
through COM — on this machine that is `RePack by Diakov`, and the template adds a
stale subject of `IEEE Transactions on Magnetics`. Both ride into the PDF, and
author metadata is exactly what de-anonymises a blind submission.

### Result

| | before | after |
|---|---|---|
| page size | p1 Letter, p2–7 A4 | US Letter throughout |
| pages | 7 | **6** |
| body font | 9.5 pt | 10 pt |
| line spacing | squeezed to 0.79×–0.96× | template default |
| named IEEE styles | 3 / 14 | **14 / 14** |
| unstyled paragraphs | 115 / 121 | **0 / 121** |
| equations | plain text, Unicode subscripts | 2 × real OMML |
| structural whitespace | ~1 full page | none |

Margins, columns and text width were measured out of the `.dot` itself rather
than assumed: 0.75″ margins on all four sides, two columns with a 0.2″ gutter,
7.0″ text block, 3.4″ columns.

### The content guarantee

Formatting work on a paper under review has exactly one hard rule: don't change
the words. That is checked mechanically, not by eye. `content_diff.py` compares
the source `.docx` text against the rendered PDF text as word multisets, after
normalising what a PDF renderer legitimately changes (unicode spaces,
hyphenation at line breaks, quote glyphs, small caps):

```
source tokens: 4499   pdf tokens: 4498
in SOURCE but not in PDF: {'c': 1, 'se': 1, 'd': 1, 'ss': 1}
in PDF but not in SOURCE: {'no': 1, 'sec': 1, 'ssd': 1}
```

All three residuals are artefacts of the comparison script, not the document —
`Σₑ c(e,s)` gets set tight by Word's math layout, and the script's own
table-cell de-duplication collapses the two identical `No` cells on the source
side. Everything else is byte-for-byte.

This check earns its keep. It caught a real regression during a refactor: the
source separates `[1]` from the reference text with an **em space**, not a plain
space, so a `^\[\d+\] ` strip silently missed and the template's auto-numbering
produced `[1] [1] Open X-Embodiment…`. The regexes now use `\s+` and assert that
they actually stripped something.

Only two strings were deleted, both under instruction: the `Preprint — submitted
to IEEE Robotics and Automation Letters` banner and the
`(author affiliations withheld for review)` line.

### Fig. 1

Fig. 1 was the one figure that genuinely needed re-rendering — it went in at
~158 dpi with sans-serif tick labels around 4–5 pt, under the IEEE floor. But a
figure is content, so `regen_fig1.py` rebuilds the 10×18 presence grid from
`embodiment_skill_matrix.csv` and asserts it cell by cell against the grid read
pixel-wise out of the figure actually embedded in the paper. All 180 cells
match; the script refuses to write if they don't. Only typeface, label size,
resolution and a ~7% aspect-ratio distortion changed.

Figures 2 and 3 were deliberately **not** regenerated. The CSVs in `outputs/`
are stale relative to them — `backbone_rarefied_ci.csv` holds three datasets
while Fig. 2 shows four — so re-rendering from the repository would have changed
what they show. Both already clear 8 pt and 280 dpi, so they are kept
byte-identical.

---

## Open items

- **Byline.** Still the template placeholder. Whether RA-L review is currently
  double-anonymous was never confirmed — check the PaperPlaza submission page
  before filling it in.
- **`[INSERT REPOSITORY URL]`** in the paper's data-availability statement points
  here and still needs filling once this repo goes public.
- **LaTeX route.** `ieeeconf.cls` is fetchable from
  `ras.papercept.net/conferences/support/files/ieeeconf.cls`, but no TeX
  distribution was installed, so nothing was compiled or verified. The Word
  route is complete and passing.

---

## Requirements

Python 3.12 on Windows. `pip install -r requirements.txt`. Word is needed only
for `make_base.ps1` and `export_pdf.ps1`; everything else is pure Python.
