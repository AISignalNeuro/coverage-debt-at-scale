# -*- coding: utf-8 -*-
"""
Coverage Debt at Scale - strip author metadata from the built files
===================================================================
Word re-stamps Author and Last author from the install's registered user on
every save, and it ignores attempts to blank them through the COM properties.
This machine's Word writes "RePack by Diakov", and the template leaves a stale
subject of "IEEE Transactions on Magnetics"; both ride into the exported PDF.

Author metadata is what de-anonymises a blind submission, so this runs last,
straight on the zip and the PDF, after Word has had its say. Blank rather than
the real name -- the byline is still an open question.

Run after export_pdf.ps1, before verify_format.py.

Env: Python 3.12. Needs pymupdf.
"""

import os
import re
import shutil
import zipfile

import pymupdf

DOCX = "Coverage_Debt_at_Scale_ieeeconf.docx"
PDF = "Coverage_Debt_at_Scale_ieeeconf.pdf"

# core.xml elements and app.xml elements that can carry a name or an org
CORE = ["dc:creator", "cp:lastModifiedBy", "dc:title", "dc:subject",
        "dc:description", "cp:keywords", "cp:category", "cp:contentStatus",
        "cp:lastPrinted"]
APP = ["Company", "Manager", "Template"]


def blank(xml, tags):
    for t in tags:
        xml = re.sub(r"<%s(\s[^>]*)?>.*?</%s>" % (re.escape(t), re.escape(t)),
                     "<%s></%s>" % (t, t), xml, flags=re.S)
    return xml


def scrub_docx(path):
    tmp = path + ".tmp"
    changed = []
    with zipfile.ZipFile(path) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                new = blank(data.decode("utf-8"), CORE).encode("utf-8")
                if new != data:
                    changed.append("core.xml")
                data = new
            elif item.filename == "docProps/app.xml":
                new = blank(data.decode("utf-8"), APP).encode("utf-8")
                if new != data:
                    changed.append("app.xml")
                data = new
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return changed


def scrub_pdf(path):
    doc = pymupdf.open(path)
    doc.set_metadata({"title": "", "author": "", "subject": "", "keywords": "",
                      "creator": "", "producer": ""})
    tmp = path + ".tmp"
    doc.save(tmp, garbage=3, deflate=True)
    doc.close()
    shutil.move(tmp, path)


def report():
    with zipfile.ZipFile(DOCX) as z:
        core = z.read("docProps/core.xml").decode("utf-8")
    left = [(t, v) for t, v in re.findall(r"<(dc:creator|cp:lastModifiedBy)>([^<]*)<", core) if v.strip()]
    meta = pymupdf.open(PDF).metadata
    dirty = {k: v for k, v in meta.items()
             if (v or "").strip() and k not in ("format", "creationDate", "modDate", "encryption")}
    print("docx author fields left:", left or "none")
    print("pdf metadata left:", dirty or "none")
    return not left and not dirty


for f in (DOCX, PDF):
    if not os.path.exists(f):
        raise SystemExit("missing %s -- run build_paper.py and export_pdf.ps1 first" % f)

print("docx parts rewritten:", scrub_docx(DOCX) or "none")
scrub_pdf(PDF)
print("clean:", report())
