#!/usr/bin/env python3
"""The six co-financing sources, and the amounts taken from HIS closed budget.

    python3 build_cofinancing.py

THE BUDGET IS CLOSED, SO THIS SECTION IS NOW SAFE TO FILL - AND NOT BEFORE.

Co-financing is 8% of each partner's total eligible cost, fixed by the
programme, so every one of these six amounts is a FUNCTION of the budget. Had I
written them a week ago against the figures I was carrying, four of the six
would have been wrong and would have needed re-entering after each correction.
His export of 08-10 totals exactly 1.300.000,00 EUR and reconciles partner by
partner and category by category, so the derived amounts are now stable.

WHICH IS WHY THIS SCRIPT DOES NOT CONTAIN THE AMOUNTS. It reads them out of his
workbook and computes 8%. If his export and the amounts below ever disagree, it
refuses to produce the document. I have been bitten twice on this project by my
own constants drifting away from his platform while both looked plausible, and
the fix both times was to stop holding a second copy of the truth.

THE CHECK THAT MATTERS: SUM TO THE PROJECT FIGURE, NOT JUST INDIVIDUALLY RIGHT

Eight per cent of each partner rounds to whole euros in six places, and six
roundings do not have to add up to eight per cent of the total. They do here -
104.000,00 against 1.196.000,00 of EU contribution - but that is checked, not
assumed. The form validates the project-level pair, and a set of individually
defensible partner figures that misses the total by three euros would be
rejected with no indication of which line caused it.

WHAT I AM NOT CLAIMING

The legal status of the source - public or private - is a dropdown, and for
five partners it is obvious. For LCEC it is not: it is the national energy
agency, it sits with the Ministry of Energy and Water, and it was established
through a UNDP project, so whether the platform wants "public" or "private" is
a question for LCEC's own finance people rather than for me. That one is marked
to confirm rather than answered, because a wrong declaration of a funding
source's legal status is the kind of thing that surfaces at contract stage.

Likewise the wording below describes each partner's OWN resources, which is
what every partner has told us they are using. If any partner is in fact
bringing a third-party contribution - a regional grant, a foundation, a
sponsor - that partner's box has to name it instead, and the programme will ask
for evidence that it is committed. Worth one line in the email that goes with
this.
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
OUT = f"{BASE}/AMMOS_Co-financing_sources_six_partners_2026-10-11.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

SCRATCH = Path("/tmp/claude-1003/-home-freelancer/"
               "22c2ac1c-53f2-4084-aeb9-41c15659448a/scratchpad")
BUDGET = SCRATCH / "budget_08-10.xlsx"

COFIN = 0.08
AGREED_TOTAL = 1_300_000.0

# His sheets use template codes. Order is the order the form lists partners.
NAMES = {
    "LEP": "HCMR (Applicant)",
    "PP1": "University of the Aegean",
    "PP2": "Pi Youth Association",
    "PP3": "Acipayam Municipality",
    "PP4": "ISAC-CNR",
    "PP5": "LCEC",
}
ORDER = ["LEP", "PP1", "PP2", "PP3", "PP4", "PP5"]

# Strings that must never reach a partner-facing file. The internal names of
# a partner's finance machinery are ours to know and not ours to publish, and
# two of these have appeared in my working notes.
FORBIDDEN = ("elke", "finance office", "non-compliant", "option b", "ras ",
             "wrong", "4.4.3")

# (code, source name <= ~100 chars, legal status, the box text)
# The box text is written MODULARLY because I have not seen the character
# counter on this field: the first sentence stands alone and is under 200
# characters, so if the box turns out to be small the second sentence is what
# gets dropped and nothing loses its meaning. Same approach that saved the
# relevance sections when their limits arrived later than the text did.
SOURCES = [
    ("LEP",
     "Own resources of HCMR, approved in its annual institutional budget",
     "Public",
     "HCMR covers its co-financing from its own resources, provided for in the "
     "annual budget approved by its Board of Directors. HCMR is a public "
     "research centre supervised by the competent Ministry, its budget is "
     "adopted annually, and the co-financing of this project is included in "
     "the budget planning for the project duration. No third-party "
     "contribution is involved."),

    ("PP1",
     "Own resources of the University of the Aegean, through its Research "
     "Committee",
     "Public",
     "The University of the Aegean covers its co-financing from its own "
     "resources, administered through its Research Committee. The University "
     "is a public higher education institution, the contribution is approved "
     "as part of the project's internal budget authorisation, and it is "
     "available for the full project duration. No third-party contribution is "
     "involved."),

    ("PP2",
     "Own resources of Pi Youth Association",
     "Private",
     "Pi Youth Association covers its co-financing from its own resources, "
     "held as unrestricted reserves in the association's annual budget. The "
     "association is a non-profit organisation whose governing board has "
     "approved the participation and the corresponding contribution, which is "
     "available from the start of implementation. No third-party contribution "
     "is involved."),

    ("PP3",
     "Own resources of Acipayam Municipality, from its approved municipal "
     "budget",
     "Public",
     "Acipayam Municipality covers its co-financing from its own municipal "
     "budget, which is adopted annually by the Municipal Council. The "
     "contribution is provided for under the municipality's own revenues, it "
     "requires no further authorisation beyond the Council's budget decision, "
     "and it is available for the project duration. No third-party "
     "contribution is involved."),

    ("PP4",
     "Own institutional resources of the National Research Council of Italy "
     "(CNR)",
     "Public",
     "ISAC-CNR covers its co-financing from the institutional resources of the "
     "National Research Council of Italy. CNR is a public research body, the "
     "institute's share is committed within its ordinary institutional "
     "funding, and it is available for the full project duration. No "
     "third-party contribution is involved."),

    ("PP5",
     "Own resources of LCEC, within its approved annual operating budget",
     "TO CONFIRM WITH LCEC - public or private",
     "LCEC covers its co-financing from its own resources, provided for in its "
     "approved annual operating budget. LCEC is the national energy "
     "conservation centre operating with the competent Ministry, the "
     "contribution is included in its budget planning for the project "
     "duration, and it is available from the start of implementation. No "
     "third-party contribution is involved."),
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
    """The eight per cent has to come off HIS numbers, not mine."""
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
            if not nums:
                continue
            # The partner's eligible total is the largest figure on its row:
            # the row carries the four cost categories and then the total.
            found.setdefault(code, max(nums))
    assert set(found) == set(NAMES), (
        "CANNOT SEE THE BUDGET - refusing to compute co-financing. "
        f"found {sorted(found)}, expected {sorted(NAMES)}")
    tot = sum(found.values())
    assert abs(tot - AGREED_TOTAL) < 1.0, (
        f"his export totals {eur(tot)}, not the agreed {eur(AGREED_TOTAL)} - "
        "co-financing is a function of the budget, so it cannot be written "
        "until the budget closes")
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
    amounts = {k: round(elig[k] * COFIN, 2) for k in ORDER}
    eu = {k: round(elig[k] - amounts[k], 2) for k in ORDER}

    fails = []
    if abs(sum(amounts.values()) - AGREED_TOTAL * COFIN) > 0.004:
        fails.append(
            f"the six rounded co-financing amounts total "
            f"{eur(sum(amounts.values()))}, but 8% of the project is "
            f"{eur(AGREED_TOTAL * COFIN)} - the form validates the project "
            f"pair, so this must reconcile")
    if abs(sum(eu.values()) - AGREED_TOTAL * (1 - COFIN)) > 0.004:
        fails.append("the EU shares do not total 92% of the project")
    if {s[0] for s in SOURCES} != set(NAMES):
        fails.append("a partner has no co-financing source, or one is named "
                     "that is not in the partnership")
    for code, name, status, text in SOURCES:
        blob = f"{name} {text}".lower()
        for bad in FORBIDDEN:
            if bad in blob:
                fails.append(f"{code} contains '{bad.strip()}' - internal "
                             f"working, must not go to partners")
        if len(clean(name)) > 100:
            fails.append(f"{code} source name {len(clean(name))} > 100 - "
                         f"cut {len(clean(name)) - 100}")
        first = clean(text).split(". ")[0] + "."
        if len(first) > 200:
            fails.append(
                f"{code} first sentence is {len(first)} characters. It has to "
                f"stand alone under 200 in case the box is small - the limit "
                f"on this field has not been seen yet")
        if "third-party" not in clean(text).lower():
            fails.append(f"{code} does not state whether a third party is "
                         f"contributing, which is what the box is asking")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS - Source of co-financing, all six partners", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} from your budget export of "
        f"08-10, which closes at {eur(AGREED_TOTAL)} exactly. Every amount "
        f"below is 8% of that partner's total eligible cost as the export "
        f"records it, not as my model predicts it - so if you correct a "
        f"partner's budget after today, that partner's co-financing figure "
        f"changes too and this file has to be regenerated.",
        colour=GREY, italic=True)

    D.h("The amounts", size=12)
    D.table(
        ["Partner", "Total eligible", "EU contribution 92%",
         "Co-financing 8%"],
        [[NAMES[k], eur(elig[k]), eur(eu[k]), eur(amounts[k])]
         for k in ORDER]
        + [["TOTAL", eur(sum(elig.values())), eur(sum(eu.values())),
            eur(sum(amounts.values()))]],
        widths=[5.0, 4.2, 4.6, 4.2])
    D.p("The six rounded amounts reconcile to the project figure to the cent. "
        "That is checked in the generator rather than eyeballed, because six "
        "individually correct roundings are not obliged to add up and the form "
        "validates the project-level pair.", size=9, colour=GREY, italic=True)

    D.h("The text, partner by partner", size=12)
    D.p("Three fields per partner: the name of the source, its legal status, "
        "and the explanation. The explanation is written so that the FIRST "
        "SENTENCE stands alone - I have not seen the character counter on this "
        "field, so if the box is smaller than it looks, delete from the end "
        "and nothing loses its meaning.", size=9, colour=GREY, italic=True)

    for code, name, status, text in SOURCES:
        D.h(f"{NAMES[code]} - {eur(amounts[code])}", size=11, space=10)
        D.p(f"Source of co-financing: {name}", bold=True)
        if status.startswith("TO CONFIRM"):
            D.p(f"Legal status: {status}", colour=RED, bold=True)
        else:
            D.p(f"Legal status: {status}")
        D.p(clean(text))
        D.p(f"{len(clean(text))} characters, first sentence "
            f"{len(clean(text).split('. ')[0]) + 1}.",
            size=8.5, colour=GREY, italic=True)

    D.h("Two things to settle before this is entered", size=12)
    D.p("1. LCEC's legal status. Every other partner is plainly public or "
        "plainly private. LCEC is the national energy conservation centre, it "
        "works with the competent Ministry, and it was set up through a UNDP "
        "project, so I am not going to guess which box the platform wants. "
        "Their finance contact will know in one line, and a wrong declaration "
        "of a funding source's legal status is the kind of thing that "
        "resurfaces at contracting.")
    D.p("2. Whether anyone is actually bringing third-party money. All six "
        "texts say the partner uses its OWN resources, because that is what "
        "each of them has told us. If any partner is in fact counting on a "
        "regional grant, a foundation or a sponsor, that partner's box has to "
        "name the source instead - and the programme can ask for proof that "
        "it is committed. One line in the next partner email settles it for "
        "all six.")

    D.d.save(OUT)
    print(f"wrote {OUT}")
    print(f"  co-financing total {eur(sum(amounts.values()))}, "
          f"EU {eur(sum(eu.values()))}")
    for k in ORDER:
        print(f"  {NAMES[k]:28s} {eur(elig[k]):>16s}  ->  "
              f"{eur(amounts[k]):>13s}")


if __name__ == "__main__":
    main()
