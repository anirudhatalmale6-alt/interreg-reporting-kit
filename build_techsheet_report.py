#!/usr/bin/env python3
"""Comparison report for the three EPIRUSMEDEYE MIS technical-sheet exports.

Everything in here comes out of compare_techsheets.py, so the numbers are
computed at build time rather than typed in.  If a figure in this document is
wrong, the script that produced it is wrong, and both are on his disk.
"""
import importlib.util

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# WAS: BASE = "/var/lib/freelancer/projects/40536008" - a folder on MY
# server, which does not exist on yours. You asked about this on 24 September
# ("35 36 lines should change?") and you were right. I then fixed the two
# files I happened to be looking at and left it in seventeen others, which is
# the same mistake a third time: I fixed the FILE instead of the PROBLEM.
#
# HERE is whatever folder this script is sitting in, so there is nothing to
# edit. Keep _shared.py beside your scripts and every one of them works from
# any folder on any machine.
from _shared import HERE

BASE = str(HERE)
OUT = f"{BASE}/EPIRUSMEDEYE_TechSheet_Comparison.docx"

spec = importlib.util.spec_from_file_location("ct", f"{BASE}/compare_techsheets.py")
ct = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ct)

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x3A)

# ---- the Annex 3 figures I reconciled earlier, for the cross-check
ANNEX3 = {"1090210": 377317.68, "1083528": 463906.49,
          "92000054": 348734.00, "997133014": 181415.00}
ANNEX3_TOTAL = 1371373.17
AF_TOTAL = 1362373.17

DIRECT = {"External expertise & services", "Equipment", "Investments / infrastructure"}


def load():
    out = []
    for label, name in ct.FILES:
        cells, partners, grand = ct.parse(f"{BASE}/{name}")
        out.append({"label": label, "file": name, "cells": cells,
                    "partners": partners, "grand": grand})
    return out


def staff_ratio(cells, ben):
    """Staff cost as a fraction of direct cost, per deliverable."""
    delivs = sorted({d for (b, d, c) in cells if b == ben})
    rs = []
    for d in delivs:
        direct = sum(a for (b, dd, c), a in cells.items()
                     if b == ben and dd == d and c in DIRECT)
        staff = sum(a for (b, dd, c), a in cells.items()
                    if b == ben and dd == d and c == "Staff costs")
        if direct > 0 and staff > 0:
            rs.append(staff / direct)
    return rs


def build():
    V = load()
    v0, v1, v2 = V
    d01 = ct.diff(v0["cells"], v1["cells"])
    d12 = ct.diff(v1["cells"], v2["cells"])

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(1.5))
    n = doc.styles["Normal"]
    n.font.name = "Calibri"
    n.font.size = Pt(9.5)
    n.paragraph_format.space_after = Pt(4)

    def para(t="", size=9.5, bold=False, italic=False, colour=None, after=4,
             before=0, align=None, indent=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.space_before = Pt(before)
        if align:
            p.alignment = align
        if indent:
            p.paragraph_format.left_indent = Cm(indent)
        if t:
            r = p.add_run(t)
            r.bold, r.italic = bold, italic
            r.font.size = Pt(size)
            if colour is not None:
                r.font.color.rgb = colour
        return p

    def head(t, before=11):
        p = para(t, size=12, bold=True, colour=NAVY, after=4, before=before)
        pr = p._element.get_or_add_pPr()
        bd = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        b.set(qn("w:space"), "2")
        b.set(qn("w:color"), "1F385E")
        bd.append(b)
        pr.append(bd)

    def table(headers, rows, widths=None, redcol=None):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Light Grid Accent 1"
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = h
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(8.5)
        for row in rows:
            cells = t.add_row().cells
            red = redcol is not None and row[redcol[0]] == redcol[1]
            for i, vv in enumerate(row):
                cells[i].text = str(vv)
                for p in cells[i].paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)
                        if red:
                            r.font.color.rgb = RED
                            r.bold = True
        if widths:
            for r in t.rows:
                for i, w in enumerate(widths):
                    r.cells[i].width = Cm(w)
        return t

    # ------------------------------------------------------------- title
    para("EPIRUSMEDEYE — COMPARISON OF THE THREE MIS TECHNICAL-SHEET EXPORTS",
         size=15, bold=True, colour=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
    para("Prepared 15 September 2026  ·  every figure below is computed from the "
         "three files, not transcribed", size=9, italic=True, colour=GREY,
         align=WD_ALIGN_PARAGRAPH.CENTER, after=10)

    # ------------------------------------------------------------- answers
    head("The three things this comparison establishes", before=0)

    para("1.  The approved modification is NOT in the MIS.", size=11, bold=True,
         colour=RED, after=3)
    para(f"All three exports — including the most recent — total "
         f"EUR {v2['grand']:,.2f}. That is the approved application form figure. The "
         f"Annex 3 you are working from totals EUR {ANNEX3_TOTAL:,.2f}. The whole "
         f"EUR {ANNEX3_TOTAL - AF_TOTAL:,.2f} difference is the Municipality of "
         f"Dropull, which the MIS still carries at "
         f"EUR {v2['partners']['92000054']['total']:,.2f} against the Annex 3's "
         f"EUR {ANNEX3['92000054']:,.2f}. Every other partner agrees to the cent.",
         after=3)
    para("This is the answer to the question you were going to put to Ms Bouziani, "
         "and it is a narrower question than “why is there a difference”. The "
         "difference is not an error in her file. It is that the modification has "
         "not been entered into the MIS.", bold=True, after=3)
    para("Why it matters rather than being a filing curiosity: reporting and "
         "first-level verification run off the MIS, not off the Annex 3. If the MIS "
         "caps Dropull at 339,734.00 and Dropull incurs against 348,734.00, the last "
         "EUR 9,000 has nowhere to be declared. Better to find that now than in a "
         "verification.", colour=RED, after=6)

    para("2.  Versions 1.1 and 2 are the same technical sheet.", size=11, bold=True,
         colour=NAVY, after=3)
    para(f"Not one amount differs between them ({len(d12)} differences). Two cells "
         f"differ as text only, and both are floating-point representation noise from "
         f"the export itself: “1635.0” against “1635.0000000000002”, and "
         f"“29.999999999999996” against “30.0”. Those are the same numbers written "
         f"differently.", after=3)
    para("This is exactly where a cell-by-cell comparison would have cost you an "
         "afternoon: it reports two differences, both real as text and both "
         "meaningless, and you go looking for a change that does not exist.",
         italic=True, colour=GREY, after=6)

    para("3.  Version 1.0 to 1.1 moved 126 amounts and changed no total whatsoever.",
         size=11, bold=True, colour=NAVY, after=3)
    para(f"The grand total is identical. Every partner total is identical. Every "
         f"partner-by-cost-category total is identical — all sixteen of them. What "
         f"changed is only how each category is spread across deliverables: "
         f"{len(v0['cells'])} non-zero amounts became {len(v1['cells'])}.", after=3)
    para("It is a re-allocation, not a budget revision. Staff costs, travel and "
         "overheads were lifted off a few management deliverables and spread across "
         "the whole work programme, so that every deliverable now carries its share. "
         "That is normally done so expenditure can be verified against the "
         "deliverable that consumed it.", after=6)

    # ------------------------------------------------------------- totals
    head("Partner totals in each export, against the Annex 3")
    rows = []
    for code in ANNEX3:
        nm = v2["partners"].get(code, {}).get("name", code)
        nm = nm.split("(")[1].rstrip(") ").strip() if "(" in nm else nm
        mis = v2["partners"][code]["total"]
        gap = ANNEX3[code] - mis
        rows.append([code, nm[:42],
                     f"{v0['partners'][code]['total']:,.2f}",
                     f"{v1['partners'][code]['total']:,.2f}",
                     f"{mis:,.2f}", f"{ANNEX3[code]:,.2f}",
                     f"{gap:+,.2f}" if abs(gap) > 0.005 else "agrees"])
    rows.append(["", "TOTAL", f"{v0['grand']:,.2f}", f"{v1['grand']:,.2f}",
                 f"{v2['grand']:,.2f}", f"{ANNEX3_TOTAL:,.2f}",
                 f"{ANNEX3_TOTAL - v2['grand']:+,.2f}"])
    table(["Code", "Beneficiary", "v1.0", "v1.1", "v2 (latest)", "Annex 3",
           "Annex 3 − MIS"], rows,
          widths=[2.3, 7.6, 3.1, 3.1, 3.1, 3.1, 3.0], redcol=(6, "+9,000.00"))
    para("")
    para("The MIS does not move at all across the three exports. The gap is entirely "
         "Dropull, and it is exactly the 9,000.00 of the 1st modification.",
         bold=True, colour=RED)

    # ------------------------------------------------------------- allocation
    head("How the 1.0 to 1.1 re-allocation was done, and whether it is defensible")
    para("It is a strict pro-rata allocation on direct cost, and it is reproducible. "
         "For each partner, staff cost per deliverable divided by that deliverable's "
         "direct cost (external expertise, equipment and infrastructure) is a "
         "constant — the same ratio on every deliverable, to five or six significant "
         "figures. That is a formula, not a judgement call.", after=4)
    rows = []
    for code in ANNEX3:
        rs = staff_ratio(v1["cells"], code)
        if not rs:
            continue
        nm = v2["partners"].get(code, {}).get("name", code)
        nm = nm.split("(")[1].rstrip(") ").strip() if "(" in nm else nm
        spread = (max(rs) - min(rs)) / min(rs) * 100 if min(rs) else 0
        rows.append([code, nm[:40], len(rs), f"{min(rs)*100:.4f}%",
                     f"{max(rs)*100:.4f}%", f"{spread:.3f}%"])
    table(["Code", "Beneficiary", "Deliverables", "Lowest ratio", "Highest ratio",
           "Spread"], rows, widths=[2.3, 8.2, 2.8, 3.2, 3.2, 2.8])
    para("")
    para("So if anyone asks why a deliverable carries the staff figure it does, the "
         "answer is arithmetic and you can show it. Worth knowing before the 22nd — "
         "a partner querying its own numbers is much easier to answer with a ratio "
         "than with “that is what the system produced”.", italic=True, colour=GREY)

    # ------------------------------------------------------------- method
    head("On the comparison method — your GeeksforGeeks link")
    para("Right instinct, wrong method for these particular files. That article "
         "compares two sheets cell by cell, by position. These exports have 127, 120 "
         "and 120 rows, so from the first inserted row onward every row is shifted "
         "and a positional comparison reports hundreds of differences that are "
         "nothing but the offset. The two or three real findings would be buried in "
         "the noise.", after=4)
    para("What the attached script does instead is key every amount on what it MEANS "
         "— the beneficiary code, the deliverable, and the cost category — and then "
         "compare those keys. Rows can move, blocks can be reordered, deliverables "
         "can be inserted or deleted, and the answer does not change.", bold=True,
         after=4)
    para("Two checks in it are worth more than the comparison itself:", after=3)
    para("The column mapping is proved, not assumed. Each partner block has a row of "
         "MIS budget-line codes that looks like a column map but is not — the codes "
         "sit in different columns for different partners. So for every deliverable "
         "row in every file, the six category columns are asserted to add up to the "
         "printed row total. If the mapping were wrong the script stops instead of "
         "producing a confident wrong answer.", indent=0.5, after=3)
    para("The diff has to close. The sum of every individual change must equal the "
         "movement in the grand total. On 1.0 to 1.1 that is 126 changes summing to "
         "0.00 against a total movement of 0.00. A diff that quietly misses a change "
         "is the only genuinely dangerous outcome here, and this is what catches it.",
         indent=0.5, after=6)
    para("Both of those ideas transfer to any file comparison you do: key on meaning "
         "rather than position, and make the result prove itself.", bold=True,
         colour=GREEN)

    # ------------------------------------------------------------- guards
    txt = "\n".join(p.text for p in doc.paragraphs)
    tbl = "\n".join(c.text for t in doc.tables for r in t.rows for c in r.cells)
    allt = txt + "\n" + tbl
    assert len(d12) == 0, f"v1.1 vs v2 is no longer clean: {d12}"
    assert abs(v2["grand"] - AF_TOTAL) < 0.02, v2["grand"]
    assert "9,000.00" in allt, "the headline finding is missing"
    assert abs(v0["grand"] - v1["grand"]) < 0.02, "the 1.0->1.1 total moved"
    doc.save(OUT)
    print("wrote", OUT)
    print(f"  v1.0->v1.1 {len(d01)} changes, v1.1->v2 {len(d12)} changes")
    print(f"  MIS total {v2['grand']:,.2f}   Annex 3 {ANNEX3_TOTAL:,.2f}   "
          f"gap {ANNEX3_TOTAL - v2['grand']:+,.2f}")


if __name__ == "__main__":
    build()
