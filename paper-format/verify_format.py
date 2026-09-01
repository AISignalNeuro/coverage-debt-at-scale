# -*- coding: utf-8 -*-
"""
Coverage Debt at Scale - format gate for the RA-L resubmission
==============================================================
Unzips the built .docx, reads the raw XML and asserts the 13 things that got
the paper returned the first time. Prints PASS/FAIL and exits non-zero if any
check fails, so it can sit in front of a submission.

Reference values are read back out of the template (template_probe.dotx) at
run time rather than trusted as literals; the constants below are only the
fallback for when that file is missing.

Usage:
  python verify_format.py [paper.docx] [paper.pdf]

Env: Python 3.12. pymupdf only for the page count (check 11 degrades without it).
"""

import os
import re
import sys
import zipfile

DOCX = sys.argv[1] if len(sys.argv) > 1 else "Coverage_Debt_at_Scale_ieeeconf.docx"
PDF = sys.argv[2] if len(sys.argv) > 2 else "Coverage_Debt_at_Scale_ieeeconf.pdf"
TPL = "template_probe.dotx"

EMU_IN = 914400
TWIP_IN = 1440

results = []


def check(n, name, ok, detail=""):
    results.append((n, name, bool(ok), detail))


def xml_parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n).decode("utf-8", "replace")
                for n in z.namelist() if n.endswith(".xml")}


# --- reference: prefer the template itself ---
ref = {"pgSz": ("12240", "15840"),
       "pgMar": {"top": "1080", "bottom": "1080", "left": "1080", "right": "1080"},
       "cols": {"num": "2", "space": "288"}}
ref_from = "fallback literals"
if os.path.exists(TPL):
    t = xml_parts(TPL)["word/document.xml"]
    s = re.search(r"<w:sectPr\b.*?</w:sectPr>", t, re.S).group(0)
    m = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', s)
    ref["pgSz"] = (m.group(1), m.group(2))
    mar = re.search(r"<w:pgMar [^/]*/>", s).group(0)
    for k in ref["pgMar"]:
        ref["pgMar"][k] = re.search(r'w:%s="(\d+)"' % k, mar).group(1)
    cols = re.search(r"<w:cols [^/]*/>", s).group(0)
    ref["cols"]["num"] = re.search(r'w:num="(\d+)"', cols).group(1)
    ref["cols"]["space"] = re.search(r'w:space="(\d+)"', cols).group(1)
    ref_from = TPL

P = xml_parts(DOCX)
doc = P["word/document.xml"]
sty = P["word/styles.xml"]
sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S).group(0)

# [1] page size
m = re.search(r'<w:pgSz w:w="(\d+)" w:h="(\d+)"', sect)
got = (m.group(1), m.group(2)) if m else ("?", "?")
check(1, "pgSz == %s x %s (US Letter)" % ref["pgSz"],
      got == tuple(ref["pgSz"]), "got %s x %s" % got)

# [2] margins
mar = re.search(r"<w:pgMar [^/]*/>", sect).group(0)
have = {k: re.search(r'w:%s="(-?\d+)"' % k, mar).group(1) for k in ref["pgMar"]}
check(2, "pgMar top/bottom/left/right == template",
      all(have[k] == v for k, v in ref["pgMar"].items()),
      "got %s | want %s" % (have, ref["pgMar"]))

# [3] columns
cols = re.search(r"<w:cols [^/]*/>", sect).group(0)
cn = re.search(r'w:num="(\d+)"', cols)
cs = re.search(r'w:space="(\d+)"', cols)
ceq = re.search(r'w:equalWidth="(\w+)"', cols)
check(3, "cols num==%s space==%s equalWidth" % (ref["cols"]["num"], ref["cols"]["space"]),
      cn and cn.group(1) == ref["cols"]["num"]
      and (cs is None or cs.group(1) == ref["cols"]["space"])
      and (ceq is None or ceq.group(1) in ("1", "true", "on")), cols)

# [4] 10pt body, and no leftover 9.5pt runs
dd = re.search(r"<w:docDefaults>.*?</w:docDefaults>", sty, re.S).group(0)
normal = re.search(r'<w:style [^>]*w:styleId="Normal".*?</w:style>', sty, re.S).group(0)
sz = re.search(r'<w:sz w:val="(\d+)"/>', normal) or re.search(r'<w:sz w:val="(\d+)"/>', dd)
base = int(sz.group(1)) if sz else 20
sz19 = re.findall(r'<w:sz w:val="19"/>', doc)
check(4, "base font == 10pt (sz 20) and no sz=19 runs", base == 20 and not sz19,
      "base sz=%d, sz19 runs=%d" % (base, len(sz19)))

# [5] no line compression
tight = [v for v in re.findall(r'<w:spacing [^/]*w:line="(\d+)"[^/]*/>', doc) if int(v) < 240]
check(5, "no direct w:line < 240 (no line compression)", not tight,
      "found %d: %s" % (len(tight), sorted(set(tight))))

# [6] the template's named styles survived
need = {"Abstract": "Abstract", "Authors": "Authors", "Table Title": "TableTitle",
        "Figure Caption": "FigureCaption0", "References": "References",
        "heading 1": "Heading1", "heading 2": "Heading2", "Body Text": "BodyText",
        "Equation": "Equation", "IndexTerms": "IndexTerms", "table copy": "tablecopy",
        "table col head": "tablecolhead", "sponsors": "sponsors", "Title": "Title"}
have_ids = set(re.findall(r'<w:style [^>]*w:styleId="([^"]+)"', sty))
missing = [k for k, v in need.items() if v not in have_ids]
check(6, "styles.xml carries the template's named styles", not missing,
      "missing: %s" % missing if missing else "all %d present" % len(need))

# [7] no paragraph left on direct formatting
body = doc[doc.index("<w:body>"):]
paras = re.findall(r"<w:p(?: [^>]*)?>.*?</w:p>|<w:p(?: [^>]*)?/>", body, re.S)
unstyled = [p for p in paras if "<w:pStyle" not in p and re.search(r"<w:t[ >]", p)]
check(7, "every text paragraph has a pStyle", not unstyled,
      "%d unstyled of %d paragraphs" % (len(unstyled), len(paras)))

# [8] equations are objects, not text
omath = re.findall(r"<m:oMath[ >]", doc)
check(8, "at least two <m:oMath> equation objects", len(omath) >= 2, "found %d" % len(omath))

# [9] nothing wider than the 7.0in text block
imgs = [int(c) / EMU_IN for c in re.findall(r'<wp:extent cx="(\d+)"', doc)]
tbls = [int(w) / TWIP_IN for w in re.findall(r'<w:tblW w:w="(\d+)" w:type="dxa"/>', doc)]
over = [round(x, 3) for x in imgs + tbls if x > 7.001]
check(9, "every image and table <= 7.0in wide", not over,
      "max image %.3f in, max table %.3f in%s"
      % (max(imgs or [0]), max(tbls or [0]), "; OVER: %s" % over if over else ""))

# [10] preprint banner gone
check(10, 'no "Preprint" banner in document.xml', "Preprint" not in doc,
      "clean" if "Preprint" not in doc else "FOUND")

# [11] page budget
try:
    import pymupdf
    pages = len(pymupdf.open(PDF))
except Exception:
    pages = None
if pages is None:
    check(11, "PDF page count <= 6 (hard cap 8)", False, "could not open %s" % PDF)
else:
    check(11, "PDF page count <= 6 (hard cap 8)", pages <= 8,
          "%d pages%s" % (pages, "" if pages <= 6 else "  <-- over 6, costs page fees"))

# [12] no hand-placed page breaks
brk = re.findall(r'<w:br w:type="page"/>', doc)
check(12, "no manual page breaks", not brk, "found %d" % len(brk))

# [13] every reference actually cited
nrefs = len(re.findall(r'<w:pStyle w:val="References"/>', doc))
text = re.sub(r"<[^>]+>", "", body)
cited = set()
for m in re.finditer(r"\[(\d+)\](?:[–-]\[(\d+)\])?", text):
    a = int(m.group(1))
    b = int(m.group(2)) if m.group(2) else a
    cited.update(range(a, b + 1))
uncited = [i for i in range(1, nrefs + 1) if i not in cited]
check(13, "all %d references are cited in the body" % nrefs, not uncited,
      "uncited: %s" % uncited if uncited else "%d/%d cited" % (nrefs, nrefs))

# --- report ---
w = max(len(n) for _, n, _, _ in results)
print("verify_format.py  --  %s" % os.path.basename(DOCX))
print("reference values from: %s" % ref_from)
print("-" * (w + 34))
for n, name, ok, detail in results:
    print("[%2d] %-*s  %s   %s" % (n, w, name, "PASS" if ok else "FAIL", detail))
print("-" * (w + 34))
failed = sum(1 for _, _, ok, _ in results if not ok)
print("%d/%d checks passed" % (len(results) - failed, len(results)))
sys.exit(1 if failed else 0)
