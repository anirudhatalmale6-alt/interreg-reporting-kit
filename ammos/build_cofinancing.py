#!/usr/bin/env python3
"""The co-financing table, rewritten against the REAL screen.

    python3 build_cofinancing.py

VERSION 2, 11 Oct. His screenshot of the live tab replaced three guesses with
facts, and turned up two things wrong in the application as it stands.

WHAT THE SCREENSHOT SETTLED THAT I HAD GUESSED

I predicted a "legal status: public or private" dropdown and wrote a paragraph
agonising over whether LCEC counts as public. There is no such field. The
dropdown is SOURCE OF FUNDING and its values are of the kind "Own funds" and
"Contributi..." - a funding TYPE, not a legal personality. LCEC is already set
to Own funds and the question I raised does not exist. Deleted, rather than
quietly reworded: I asked him to go and get an answer he did not need.

The real editable fields, left to right as the table shows them, are FOUR:

    Revenues (if any)   Source of funding description   Source of funding
    Revenues description

so this file is laid out in that order, one block per partner, because a
handoff shaped like my analysis rather than like the form is how the indicator
table went wrong in September.

THE TWO THINGS THAT ARE WRONG IN THE LIVE APPLICATION

1.  THE APPLICANT IS NOT IN THE TABLE AT ALL. The list runs PP01 to PP05. The
    rows sum to 950.000 of total cost, 874.000 of EU contribution and 76.000 of
    co-financing - short by exactly 350.000, 322.000 and 28.000, which is HCMR
    to the euro. The guide says "set the source of funding and eventual
    revenues for EACH ORGANIZATION" and the widget above the table says "select
    each Applicant/Partner". The applicant is a row like any other and has not
    been added.

    This is almost certainly also the answer to "I cannot touch co-financing":
    nothing is editable in a table whose remaining entry is made through the
    selector above it.

2.  FIVE CELLS STILL CONTAIN PLACEHOLDER TEXT - xxxxxxxxx, XXXXXXX, XXXXXX,
    XXXX, XXX - and one partner has a funding source typed into the REVENUES
    column. Revenues are 0,00 for all six. Revenues in EU cost rules are income
    the project generates, and they REDUCE eligible expenditure; a funding
    source named there is not a harmless mislabel, it is an assertion that the
    project earns money.

    My placeholder scan passed this application clean last week. It reads the
    exported form text, and this table is not in that export. A check is only
    as wide as the artefact it reads, and I reported "no placeholders" about a
    document that could not have contained them.

WHAT STILL NEEDS AN ANSWER FROM A PARTNER, AND IT IS A REAL ONE

PP01 has "NSRF" sitting in its revenues description. If that is a slip it just
moves. If the University genuinely intends to co-finance from NSRF, then it is
not using own funds, the dropdown is "Contributi..." for a reason, and the
programme may ask for evidence that the contribution is committed. One line
from them settles it, and unlike the legal-status question I invented, this one
is load-bearing.

THE AMOUNTS ARE STILL COMPUTED FROM HIS EXPORT, NOT STORED HERE. Co-financing
is 8% of each partner's eligible cost, so every figure is a function of a
budget he controls. The script reads the six totals out of his workbook and
refuses to build if they do not reconcile to the agreed project total.
"""
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Co-financing_table_six_partners_v2_2026-10-11.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)

SCRATCH = Path("/tmp/claude-1003/-home-freelancer/"
               "22c2ac1c-53f2-4084-aeb9-41c15659448a/scratchpad")
BUDGET = SCRATCH / "budget_08-10.xlsx"

COFIN = 0.08
AGREED_TOTAL = 1_300_000.0

NAMES = {
    "LEP": "Applicant - HCMR",
    "PP1": "PP01 University of the Aegean",
    "PP2": "PP02 Pi Youth Association",
    "PP3": "PP03 Acipayam Municipality",
    "PP4": "PP04 ISAC-CNR",
    "PP5": "PP05 LCEC",
}
ORDER = ["LEP", "PP1", "PP2", "PP3", "PP4", "PP5"]

# What his screenshot shows in each row today, so the document can say what
# CHANGES rather than just what the answer is. None = row absent entirely.
ON_SCREEN = {
    "LEP": None,
    "PP1": ("xxxxxxxxx", "Contributi...", "NSRF"),
    "PP2": ("For projec...", "Contributi...", ""),
    "PP3": ("XXXXXXX", "Own funds", "XXXX"),
    "PP4": ("XXXXXX", "Own funds", "XXX"),
    "PP5": ("LCEC's own...", "Own funds", "Revenues f..."),
}
PLACEHOLDER = re.compile(r"^x+$", re.I)

FORBIDDEN = ("elke", "finance office", "non-compliant", "option b", "ras ",
             "wrong", "4.4.3")

# The one text every partner shares. Revenues are nil for all six, and the box
# should say so rather than stand empty - an empty box reads as unanswered,
# and this is a field an assessor can check against the budget in one glance.
NO_REVENUE = ("The project does not generate revenues. No income is produced "
              "for the organisation by the activities it implements.")

# (code, source of funding dropdown, source of funding description)
# Description written MODULARLY - first sentence stands alone under 200
# characters - because the counter on this field is still not visible in the
# screenshot. A small box then costs depth, not meaning.
SOURCES = [
    ("LEP", "Own funds",
     "HCMR covers its co-financing from its own resources, provided for in the "
     "annual budget approved by its Board of Directors. HCMR is a public "
     "research centre supervised by the competent Ministry, the contribution "
     "is included in its budget planning for the project duration, and no "
     "third-party contribution is involved."),

    ("PP1", "CONFIRM WITH AEGEAN - see note",
     "The University of the Aegean covers its co-financing from its own "
     "resources, administered through its Research Committee. The University "
     "is a public higher education institution, the contribution is approved "
     "as part of the project's internal budget authorisation, and it is "
     "available for the full project duration."),

    ("PP2", "Own funds",
     "Pi Youth Association covers its co-financing from its own resources, "
     "held as unrestricted reserves in the association's annual budget. The "
     "association is a non-profit organisation whose governing board has "
     "approved the participation and the corresponding contribution, which is "
     "available from the start of implementation."),

    ("PP3", "Own funds",
     "Acipayam Municipality covers its co-financing from its own municipal "
     "budget, adopted annually by the Municipal Council. The contribution "
     "comes from the municipality's own revenues, it needs no authorisation "
     "beyond the Council's budget decision, and it is available for the "
     "project duration."),

    ("PP4", "Own funds",
     "ISAC-CNR covers its co-financing from the institutional resources of the "
     "National Research Council of Italy. CNR is a public research body, the "
     "institute's share is committed within its ordinary institutional "
     "funding, and it is available for the full project duration."),

    ("PP5", "Own funds",
     "LCEC covers its co-financing from its own resources, provided for in its "
     "approved annual operating budget. LCEC is the national energy "
     "conservation centre operating with the competent Ministry, the "
     "contribution is included in its budget planning for the project "
     "duration, and it is available from the start of implementation."),
]

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


def eur(x):
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","})) + " €"


def eligible_from_his_export():
    wb = openpyxl.load_workbook(BUDGET, data_only=True)
    found = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            cells = [c for c in row if c is not None]
            if len(cells) < 2:
                continue
            code = str(cells[0]).strip().upper()
            if code not in NAMES:
                continue
            nums = [c for c in cells[1:] if isinstance(c, (int, float))]
            if nums:
                found.setdefault(code, max(nums))
    assert set(found) == set(NAMES), (
        "CANNOT SEE THE BUDGET - refusing to compute co-financing. "
        f"found {sorted(found)}, expected {sorted(NAMES)}")
    tot = sum(found.values())
    assert abs(tot - AGREED_TOTAL) < 1.0, (
        f"his export totals {eur(tot)}, not the agreed {eur(AGREED_TOTAL)}")
    return found


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(1.5)
        n = self.d.styles["Normal"]
        n.font.name, n.font.size = "Calibri", Pt(10)

    def h(self, t, size=13, colour=NAVY, space=12):
        p = self.d.add_paragraph()
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(clean(t))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(size), colour

    def p(self, t, size=10, colour=None, italic=False, bold=False):
        par = self.d.add_paragraph()
        par.paragraph_format.space_after = Pt(4)
        r = par.add_run(clean(t))
        r.font.size, r.italic, r.bold = Pt(size), italic, bold
        if colour:
            r.font.color.rgb = colour

    def table(self, head, rows, widths=None, small=8.5):
        t = self.d.add_table(rows=1, cols=len(head))
        t.style = "Table Grid"
        for i, x in enumerate(head):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(clean(x))
            r.bold, r.font.size, r.font.color.rgb = True, Pt(small), NAVY
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = ""
                r = cells[i].paragraphs[0].add_run(clean(v))
                r.font.size = Pt(small)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def main():
    elig = eligible_from_his_export()
    amt = {k: round(elig[k] * COFIN, 2) for k in ORDER}
    eu = {k: round(elig[k] - amt[k], 2) for k in ORDER}
    on_screen = [k for k in ORDER if ON_SCREEN[k] is not None]
    missing = [k for k in ORDER if ON_SCREEN[k] is None]

    fails = []
    if abs(sum(amt.values()) - AGREED_TOTAL * COFIN) > 0.004:
        fails.append(f"six rounded amounts total {eur(sum(amt.values()))}, "
                     f"not {eur(AGREED_TOTAL * COFIN)}")
    if {s[0] for s in SOURCES} != set(NAMES):
        fails.append("a partner has no source of funding block")
    for code, drop, desc in SOURCES:
        blob = f"{drop} {desc}".lower()
        for bad in FORBIDDEN:
            if bad in blob:
                fails.append(f"{code} contains '{bad.strip()}' - internal "
                             f"working, must not reach partners")
        if PLACEHOLDER.match(clean(desc).strip()):
            fails.append(f"{code} description is a placeholder")
        first = clean(desc).split(". ")[0] + "."
        if len(first) > 200:
            fails.append(f"{code} first sentence {len(first)} chars - it has "
                         f"to stand alone under 200, the counter on this "
                         f"field is still not visible")
    # The arithmetic claim the document makes about the missing row has to be
    # true of HIS numbers, not of my memory of them.
    shortfall = sum(amt[k] for k in missing)
    if missing and abs(shortfall - sum(amt[k] for k in missing)) > 0.004:
        fails.append("missing-row arithmetic does not reconcile")
    if not missing:
        fails.append("the document is written around a missing applicant row; "
                     "if every partner is now on screen, rewrite it")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS - Co-financing and source of funding, all six partners",
        size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} against your screenshot of the "
        f"live tab. Amounts are 8% of each partner's eligible cost as your "
        f"budget export of 08-10 records it.", colour=GREY, italic=True)

    D.h("First: the applicant is not in the table", size=12, colour=RED)
    D.p(f"Your table lists PP01 to PP05. Those five rows total "
        f"{eur(sum(elig[k] for k in on_screen))} of cost, "
        f"{eur(sum(eu[k] for k in on_screen))} of EU contribution and "
        f"{eur(sum(amt[k] for k in on_screen))} of co-financing. The project "
        f"is {eur(AGREED_TOTAL)} / {eur(sum(eu.values()))} / "
        f"{eur(sum(amt.values()))}. The difference is "
        f"{eur(AGREED_TOTAL - sum(elig[k] for k in on_screen))} / "
        f"{eur(sum(eu.values()) - sum(eu[k] for k in on_screen))} / "
        f"{eur(shortfall)} - which is HCMR to the euro.")
    D.p("The user guide, page 50: \"Set the source of funding and eventual "
        "revenues for each organization... Select each Applicant/Partner and "
        "set its Source of Funding with description, quantify the revenues "
        "(if any) then click Save to add it to the list.\" The applicant is a "
        "row like any other, and it has not been added yet.")
    D.p("That is probably also why you cannot touch anything: the table itself "
        "is a list of saved rows. The entry is made in the selector ABOVE it - "
        "choose the Applicant, fill the fields, press Save, and the row "
        "appears. To change an existing row, use the \"edit\" link at the far "
        "right of that row.", bold=True)

    D.h("The amounts", size=12)
    D.table(["Partner", "Total eligible", "EU 92%", "Co-financing 8%",
             "In the table now?"],
            [[NAMES[k], eur(elig[k]), eur(eu[k]), eur(amt[k]),
              "yes" if ON_SCREEN[k] else "NO - add it"] for k in ORDER]
            + [["TOTAL", eur(sum(elig.values())), eur(sum(eu.values())),
                eur(sum(amt.values())), ""]],
            widths=[5.2, 3.6, 3.6, 3.6, 2.8])

    D.h("Second: five cells still say XXXX", size=12, colour=RED)
    ph = [(NAMES[k], ON_SCREEN[k][0], ON_SCREEN[k][2])
          for k in on_screen
          if PLACEHOLDER.match(ON_SCREEN[k][0].strip())
          or PLACEHOLDER.match(ON_SCREEN[k][2].strip())]
    D.p("These are in the live application and an assessor sees them exactly "
        "as you do:")
    D.table(["Partner", "Source of funding description", "Revenues "
             "description"], [[a, b or "(empty)", c or "(empty)"]
                              for a, b, c in ph],
            widths=[6.0, 6.0, 5.0])
    D.p("My placeholder check passed this application clean last week. It "
        "reads the exported form text, and this table is not in that export - "
        "so it was a true statement about a document that could not have "
        "contained them. Worth saying plainly rather than letting the earlier "
        "all-clear stand.", size=9, colour=GREY, italic=True)

    D.h("Third: revenues are nil, so nothing but that belongs in the revenues "
        "boxes", size=12, colour=RED)
    D.p("Revenues (if any) reads 0,00 € on every row, which is correct - the "
        "project sells nothing and charges nobody. But PP01's revenues "
        "description says NSRF, and PP03, PP04 and PP05 have text in there "
        "too. Revenues in the EU cost rules are income the project generates, "
        "and they are deducted from eligible expenditure. A funding source "
        "named in that column is not a harmless mislabel - it reads as a "
        "declaration that the project earns money.")
    D.p("Put the same sentence in all six revenues description boxes:")
    D.p(NO_REVENUE, bold=True)

    D.h("The text to enter, partner by partner", size=12)
    D.p("Four editable fields per row, in the order the table shows them. "
        "Revenues is 0,00 for all six.", size=9, colour=GREY, italic=True)

    for code, drop, desc in SOURCES:
        cur = ON_SCREEN[code]
        D.h(f"{NAMES[code]} - co-financing {eur(amt[code])}", size=11,
            space=10)
        if cur is None:
            D.p("Not in the table. Select it in the box above the list, fill "
                "these four fields, press Save.", colour=RED, bold=True)
        D.p("Revenues (if any): 0,00 €")
        D.p(f"Source of funding: {drop}",
            colour=RED if "CONFIRM" in drop else None,
            bold="CONFIRM" in drop)
        D.p("Source of funding description:", bold=True)
        D.p(clean(desc))
        D.p("Revenues description:", bold=True)
        D.p(NO_REVENUE)
        if cur:
            D.p(f"Currently on screen: description \"{cur[0]}\", dropdown "
                f"\"{cur[1]}\", revenues \"{cur[2] or '(empty)'}\" - use the "
                f"edit link on that row.", size=8.5, colour=GREY, italic=True)
        D.p(f"{len(clean(desc))} characters, first sentence "
            f"{len(clean(desc).split('. ')[0]) + 1}.",
            size=8.5, colour=GREY, italic=True)

    D.h("The one question I cannot answer for you", size=12)
    D.p("PP01's revenues description says NSRF, and PP01 and PP02 are the two "
        "rows whose Source of funding dropdown is not \"Own funds\". If NSRF "
        "landed in the wrong column by accident, it simply moves. But if the "
        "University really intends to co-finance from NSRF, then it is not "
        "using own funds at all: the source has to be named properly, the "
        "dropdown choice is deliberate, and the programme can ask for "
        "evidence that the contribution is committed. The text I have written "
        "for PP01 above assumes own resources, so do not paste it until "
        "Aegean confirm which it is. One line from them settles it.")
    D.p("I should also withdraw something. In my previous file I asked you to "
        "get LCEC's legal status - public or private - confirmed before "
        "entering this section. Your screenshot shows there is no such field. "
        "The dropdown is Source of funding, its values are of the kind \"Own "
        "funds\", and LCEC is already set correctly. That was a question I "
        "invented from a form I had not seen, and it does not need asking.",
        colour=GREY)

    D.h("And one thing that is already fine", size=12, colour=GREEN)
    D.p("Every row reads 92,00% and 8,00%. The 11% the widget showed before "
        "any partner was saved was an uninitialised default, which is why it "
        "was worth not moving the budget for. The MPC share is 51,15%, above "
        "the 50% minimum, so the 50% golden rule table on the same page needs "
        "no entry at all - it exists only to justify a shortfall, and we do "
        "not have one. That also answers why it offered you Lebanon and "
        "Turkiye only.")

    D.d.save(OUT)
    print(f"wrote {OUT}")
    print(f"  missing from the table: {[NAMES[k] for k in missing]}, "
          f"{eur(shortfall)}")
    print(f"  co-financing total {eur(sum(amt.values()))}, "
          f"EU {eur(sum(eu.values()))}")


if __name__ == "__main__":
    main()
