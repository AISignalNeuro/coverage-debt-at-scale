# Format rebuild — submission 26-4993

Returned by the editorial office on formatting. Rebuilt on `ieeeconf_letter.dot`.
6 pages, 13/13 checks green, no words changed.

## What was actually wrong

The review said margins and font size. Those were real, but they were symptoms.
The file had never been built on an IEEE template at all:

- `styles.xml` held only Word defaults. Of the 14 named styles the template
  defines, 3 were present, and those 3 (`Title`, `Heading1`, `Heading2`) are
  Word built-ins that happen to share the names.
- 115 of 121 paragraphs carried direct formatting with no `pStyle`.
- There were **four** `sectPr` blocks, and three of them said A4
  (11906 × 16838 twips) with 1.0″ margins.

That last one is the fatal one. Exporting the returned file gives a PDF where
page 1 is US Letter and pages 2–7 are A4 — see `src_reference.pdf` in the
handover zip if you want to see it. A mixed-page-size PDF fails a format check
on its own, regardless of anything else.

## Template specs, measured

`ieeeconf_letter.dot` is a binary Word 97-2003 template (OLE2, `D0CF11E0`), so it
was opened through Word and converted to `template_probe.dotx` to read its real
`sectPr` and `styles.xml`.

| | |
|---|---|
| `w:pgSz` | 12240 × 15840, `w:code="1"` — US Letter |
| `w:pgMar` | 1080 twips top/bottom/left/right — **0.75″ all four sides** |
| header / footer | 432 twips (0.3″) |
| `w:cols` | `num="2" space="288"` — 0.2″ gutter |
| text block | 10080 twips = 7.0″ |
| column | 4896 twips = 3.4″ |

This confirms 0.75″/0.2″/7.0″/3.4″ rather than the 0.625/0.75/1.0 figures in the
first review. The practical consequence also held: with the correct template
margins the 6.6″ figures and 6.9″ table were never overruns — they only
overflowed because the file was sitting on 1.0″ margins.

A few style names in the work order were guesses and didn't match. The real ones:

| assumed | actual (styleId) |
|---|---|
| `Heading` | `Title` (`Title`) — 16 pt bold, page-width frame |
| `figurecaption1` | `Figure Caption` (`FigureCaption0`) — 8 pt |
| `TableTitle` | `Table Title` (`TableTitle`) — 8 pt small caps |
| `equation1` | `Equation` (`Equation`) — right tab at 4810 twips |

`Abstract`, `IndexTerms`, `Authors`, `References`, `sponsors`, `Heading1`,
`Heading2`, `BodyText`, `table head/copy/col head` all matched.

## Method

`base.docx` is `Documents.Add(ieeeconf_letter.dot)`, so it arrives with the
template's `styles.xml`, `numbering.xml`, `settings.xml`, `fontTable.xml`,
`header1.xml` and `sectPr` already in place. The body is emptied, the `sectPr`
kept, and the text poured back into the template's own named styles. Text is
lifted run-by-run out of the source `.docx` — nothing is retyped, so runs keep
their bold/italic and their exact whitespace.

No page-layout number is hard-coded in `build_paper.py`. The numbers in the
table above are used only by `verify_format.py`, and even there they're re-read
from the template at run time.

Two things about this template that cost time:

- `Heading 1` / `Heading 2` auto-number (`I.`, `II.` / `A.`, `B.`) and
  `References` auto-numbers `[1]`…`[40]`, so the literal numerals have to be
  stripped from the text or they double up.
- `Title` and `Authors` are page-width **frames**. The template anchors them on
  a 1 pt `Text` paragraph and overrides `framePr` on each. Skip that and the two
  frames collide — the byline renders *above* the title.

## Layout

`Table I`, `Fig. 1` and `Fig. 2` are column-spanning floats pinned to the top of
their page: a floating table (`tblpPr`) for the table, anchored text boxes with
`wrapTopAndBottom` for the figures. `Fig. 3` is single column, which recovered
about half a page on its own.

No section breaks, no manual page breaks, no spacer paragraphs. Trailing
whitespace measured on the exported PDF by ink projection:

| page | trailing | left col | right col |
|---|---|---|---|
| 1 | 0.6% | 0.6% | 2.8% |
| 2 | 0.0% | 1.5% | 0.0% |
| 3 | 1.2% | 1.2% | 7.6% |
| 4 | 0.4% | 1.5% | 0.4% |
| 5 | 0.3% | 0.3% | 2.9% |
| 6 | 0.6% | 0.6% | 82.5% (last page) |

`Fig. 3` sits at the head of §IV-B rather than after the paragraph that cites it.
Putting it after left ~2″ of white at the foot of the left column, because at
4.07″ tall it couldn't fit in what remained and got pushed to the next column.

## Equations

Both numbered equations are real OMML now, not text. Unicode `Σₑ` / `Σₛ` became
`<m:sSub>`; variables italic, numerals and delimiters upright via
`<m:sty m:val="p"/>`. Numbered with a right tab at the column edge — checked in
the PDF, both `(1)` and `(2)` end at x = 556.3 pt against a column edge of 558.0.

Math is set in Cambria Math, which is what Word's equation editor produces and
what embeds cleanly.

## Table

Horizontal rules only — above the header, below the header, below the last row.
Every vertical border gone (`insideV`: 1 → 0). Caption is `TABLE I` on its own
line with the descriptive title beneath, both in `Table Title`. Column widths set
explicitly (1500/2560/2560/3220 twips); Word's autofit produced a badly
unbalanced grid where "ABot-M0 [5]" wrapped over four lines.

## Content

Two strings deleted, both under instruction: the `Preprint — submitted to IEEE
Robotics and Automation Letters` banner and `(author affiliations withheld for
review)`. Nothing else.

Verified with `content_diff.py` rather than by eye — source `.docx` text against
rendered PDF text, as word multisets, normalised for unicode spaces, hyphenation
at line breaks, quote glyphs and small caps:

```
source tokens: 4499   pdf tokens: 4498
in SOURCE but not in PDF: {'c': 1, 'se': 1, 'd': 1, 'ss': 1}
in PDF but not in SOURCE: {'no': 1, 'sec': 1, 'ssd': 1}
```

`se`+`c` → `sec` and `ss`+`d` → `ssd` are Word setting the sigma tight against
`c(e,s)`. `no` is the script's own table-cell de-duplication collapsing the two
identical `No` cells on the source side — the document has both. Nothing else
differs.

Worth keeping: this check caught a real regression during a later refactor. The
source separates `[1]` from the reference text with an em space, not a plain
space, so `^\[\d+\] ` silently failed to strip and the auto-numbering produced
`[1] [1] Open X-Embodiment…`. Same bug hit the section numerals and the Table I
anchor lookup. The regexes now use `\s+` and assert they stripped something.

Fonts in the output PDF: Times New Roman (regular/bold/italic/bold-italic) and
Cambria Math, all embedded.

## Figures

| | placement | width | dpi | labels | |
|---|---|---|---|---|---|
| Fig. 1 | spanning float | 6.9″ | 600 (was 158) | 8 pt (was ~4–5) | re-rendered |
| Fig. 2 | spanning float | 6.9″ | 283 | ~8.4 pt | kept |
| Fig. 3 | single column | 3.25″ | 307 | ~9.4 pt | kept |

Fig. 2 is confirmed to be the four-dataset version — the embedded image is
byte-identical (MD5 `d029fb58089d0a6f1e112aa3c8a25017`) to
`outputs/fig2_rarefaction_4ds.png`, and it shows DROID / AgiBot / RoboMIND /
Bridge. No substitution needed.

Fig. 1 was the only real violation. Since a figure is content, the re-render is
gated: `regen_fig1.py` rebuilds the grid from `embodiment_skill_matrix.csv` and
asserts it against the grid read pixel-wise out of the embedded image. 180/180
cells match, row sums 18/16/11/8/8/6/5/3/2/1 — which is also what the paper
claims in §IV-B. The script won't write if they disagree. It also fixed a ~7%
vertical stretch: the image is 1089×526 (ratio 2.070) but was embedded at
6.625″×3.000″ (ratio 2.208).

Figures 2 and 3 were deliberately left alone. The CSVs in `outputs/` are stale
relative to them — `backbone_rarefied_ci.csv` has three datasets while Fig. 2
shows four, and the notebook's `figure3_per_skill_debt` cell makes a landscape
chart with a different tie ordering than the portrait figure in the paper.
Re-rendering from the repo would have changed what they show. Both already clear
8 pt and 280 dpi.

## Decisions

- **Title 16 pt, not the 20 pt in the source.** That's what the template's
  `Title` style says, and the template wins.
- **`Abstract` / `Index Terms` labels italic.** The template's own abstract
  italicises the leading word; the source had it bold.
- **Table I stayed full width.** Single column was tried first — four columns of
  prose at 3.4″ is unreadable.

## Still open

- **Byline** is the template placeholder. Whether RA-L review is currently
  double-anonymous was never confirmed — check PaperPlaza. No anonymisation
  decision was made here.
- **`[INSERT REPOSITORY URL]`** left untouched, as instructed. Needs filling, and
  note that reviewer item 8 (the data-availability contradiction) is still
  outstanding.
- **LaTeX route not built.** `ieeeconf.cls` is fetchable (HTTP 200, 202,166 bytes
  from `ras.papercept.net/conferences/support/files/ieeeconf.cls`) but no TeX
  distribution is installed here, so nothing could be compiled or verified.
  Shipping untested `.tex` seemed worse than not shipping it.
- **Cosmetic.** The PDF has 54 space glyphs (U+0020) set in Arial rather than
  Times. They're the separator Word emits for the template's own *legacy*
  auto-numbering on `Heading 1`, `Heading 2` and `References`. Invisible, 2–3 pt
  wide, and embedded. Adding `w:rFonts` to the numbering levels doesn't override
  Word's internal fallback; the only fix is to drop the template's auto-numbering,
  which is a worse trade. Noted so nobody re-discovers it.

## Not touched

None of the ten scientific review items — claim scope, nestedness, the
systematic-bias discussion, the rarefaction curve, `0.567`, SCD(K), §IV-D and the
SIMPLER probe, the data-availability contradiction, the controlled experiment, or
the reference venues. Formatting only.
