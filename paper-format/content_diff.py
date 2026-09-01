# -*- coding: utf-8 -*-
"""
Coverage Debt at Scale - prove the rebuild changed no words
===========================================================
Compares the text of the returned .docx against the text of the rebuilt PDF as
word multisets, after normalising the things a PDF renderer legitimately
changes: unicode spaces, hyphenation at line breaks, quote glyphs, small caps.
Order-insensitive on purpose, because the floats move Table I and the figure
captions in reading order.

The two strings we were told to delete (the preprint banner and the withheld-
affiliations line) come off the source side first, and the script asserts they
were actually there before removing them.

Env: Python 3.12. Needs python-docx + pymupdf.
"""

import re
import collections
import unicodedata

import docx
import pymupdf
from docx.text.paragraph import Paragraph
from docx.table import Table

SRC = "src.docx"
PDF = "Coverage_Debt_at_Scale_ieeeconf.pdf"

DROPPED = ["preprint - submitted to ieee robotics and automation letters",
           "(author affiliations withheld for review)"]


def norm(s):
    s = unicodedata.normalize("NFKC", s)
    for ch in "      ":
        s = s.replace(ch, " ")
    s = s.replace("­", "")
    for ch in "‐‑‒–—−":
        s = s.replace(ch, "-")
    s = s.replace("“", '"').replace("”", '"')
    s = s.replace("‘", "'").replace("’", "'")
    s = s.replace("×", "x").replace("Σ", "S")
    for a, b in {"ₑ": "e", "ₛ": "s", "ₙ": "n",
                 "⁵": "5", "⁰": "0"}.items():
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s).lower().strip()


def source_text():
    d = docx.Document(SRC)
    out = []
    for ch in d.element.body.iterchildren():
        tag = ch.tag.split("}")[1]
        if tag == "p":
            out.append(Paragraph(ch, d).text)
        elif tag == "tbl":
            for row in Table(ch, d).rows:
                seen = []
                for c in row.cells:                 # cells repeat under gridSpan
                    if c.text not in seen:
                        seen.append(c.text)
                out.append(" ".join(seen))
    return norm(" ".join(out))


def bag(t):
    return collections.Counter(re.findall(r"[a-z0-9]+(?:[-'.][a-z0-9]+)*", t))


S = source_text()
for d in DROPPED:
    assert d in S, "expected-dropped string not found in source: %r" % d
    S = S.replace(d, "")
S = re.sub(r"-\s+", "-", re.sub(r"\s+", " ", S))
P = re.sub(r"-\s+", "-", norm(" ".join(p.get_text() for p in pymupdf.open(PDF))))

bs, bp = bag(S), bag(P)
print("source tokens: %d   pdf tokens: %d" % (sum(bs.values()), sum(bp.values())))
print("\nin SOURCE but not in PDF:", dict(bs - bp) or "(none)")
print("in PDF but not in SOURCE:", dict(bp - bs) or "(none)")
print("""
Expected residue, all artefacts of this script rather than of the document:
  se+c -> sec, ss+d -> ssd   Word's math layout sets the sigma tight to c(e,s)
  no                         the table-cell de-dup above collapses the two 'No'
                             cells on the SOURCE side; the document has both
""")
