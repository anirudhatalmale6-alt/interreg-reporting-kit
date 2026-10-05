#!/usr/bin/env python3
"""The Logical Framework fields, and the Summary. Paste-ready.

    python3 build_logframe_text.py

HIS SCREENSHOT ANSWERED "DO I HAVE ALL TAB DATA?" BETTER THAN I COULD

    Preliminary Info                   100%
    Project                             58%
       Summary                           0%   <- 2.000 chars, never written
       Relevance                        100%
       Logical Framework                15%
       Effectiveness                     0%   <- the work packages go here
       Sustainability                    0%   <- criterion 5, 12 points
       Cost-effectiveness                0%   <- criterion 6, 12 points
       Horizontal principles             0%   <- criterion 7, 4 points
    Partnership                         50%
    Work packages and budget (part 1)   17%

So no - four scored sections are at zero and one of them is the Summary, which
is the first thing an evaluator reads. 28 of the 100 points sit in
Sustainability, Cost-effectiveness and Horizontal principles alone.

AND IT SHOWS WHY HIS MENUS WERE DEAD

At the top of the Logical Framework tab there is a REQUIRED field, marked with
an asterisk:

    3.2.1 Transfer and adaptation methodology *      18/2000
    xxxxxxxxxxxxxxxxxx

A required field holding placeholder text, above the Specific Objective and
Expected Result rows he was trying to add. That is almost certainly the block:
the section will not progress past an unsatisfied required field. Fill 3.2.1,
press Save, and then the rows below should take.

That field is also one of the most important boxes in the application. The
user guide spells out what it wants:

    "Explain the project concept and methodology to ensure effective transfer,
     adaptation, testing etc. of relevant outputs/results from other
     programmes and initiatives across the participating territories,
     demonstrating a realistic, feasible, and context-appropriate approach.
     Furthermore, explain how the project introduces INNOVATIVE methods or
     approaches to enhance the adaptation, transfer, or use of existing
     results/outputs in the field of sustainable tourism."

Transfer, adaptation, testing, feasibility, context, and innovation - six
things in 2.000 characters. Every one of them is answered below, and the
innovation claim is the real one we have: not that we transfer a method, but
that we produce a METHOD FOR DECIDING WHAT TRANSFERS, and a comparability
layer that lets incomparable shore types be read together.

ONE EXPECTED RESULT, NOT THREE

The Guidelines allow one expected result per specific objective and we have
one objective. So the instinct to write three - transfer, adoption,
replication - has to compress into a single sentence containing all three.
It does, below.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import DIRECT, P, STAFF

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_logical_framework_text_2026-10-05.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

THIN = 0.75
TITLE = "Adaptive Management of Mediterranean Shoreline Sediments under Tourism Pressure"
ACRONYM = "AMMOS"

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


# -------------------------------------------------- 3.2.1, required, 2000
S321 = [
    "CONCEPT. AMMOS transfers one published method to three shores it was not "
    "written for. MMM_IF02 (AMMIRARE, Interreg Italia-Francia Marittimo "
    "2021-2027) holds seven field method sheets for sediment shores under "
    "visitor pressure. Our sites are a marine dune system on "
    "Lemnos, a Lebanese coastal site and an artificial freshwater pond in "
    "Denizli. Transfer is therefore not one act but 21 decisions, seven "
    "sheets against three shore types.",

    "METHODOLOGY, IN FOUR STAGES. (1) DECIDE AND TRANSLATE: each of the 21 "
    "cells is recorded as applied, substituted or not "
    "applicable, with the reason, and the method is translated into Greek, "
    "Turkish and Arabic - no site manager can use an Italian-language source "
    "otherwise. (2) TEST BY USING: the adapted method is "
    "validated by producing three comparable territorial baselines with it, "
    "so feasibility is demonstrated, not asserted. (3) DEMONSTRATE: "
    "reversible installations channel visitors at each site "
    "through one full season under the monitoring protocol. (4) EMBED: each "
    "site reaches a costed management plan, formally adopted by the "
    "responsible authority and overseen by a standing stewardship committee.",

    "WHY IT IS FEASIBLE AND CONTEXT-APPROPRIATE. The organisation that "
    "produced the source method is a partner, and its first named author "
    "leads adaptation of the physical sheets. Interventions are demountable, "
    "proportionate beside a protected habitat. "
    "Pilot implementation closes four months before the project does.",

    "WHAT IS INNOVATIVE. Not the transfer itself. First, AMMOS produces a "
    "method for deciding WHAT transfers: a modular, cell-by-cell "
    "transferability assessment that states where a technique stops working "
    "and why - the part conventional capitalisation leaves implicit and a "
    "fourth territory actually needs. Second, a cross-typology "
    "comparability layer lets a freshwater pond, a dune system and a marine "
    "coast be monitored against one another, the precondition for any "
    "basin-scale reading of shore condition under tourism pressure.",
]

EXPECTED_RESULT = (
    "Receiving Mediterranean territories apply a transferred and adapted "
    "shoreline sediment management method, sustain it through formally "
    "adopted and costed local management plans with standing multi-stakeholder "
    "governance, and make it replicable beyond the partnership through a "
    "transferability guide and an open cross-typology monitoring platform "
    "returned to the MMM common database."
)

SUMMARY = [
    "Mediterranean sediment shores are the asset coastal tourism depends on "
    "and the surface it erodes. Methods to manage that tension exist and have "
    "been validated in the western Mediterranean; the municipalities and site "
    "managers who must act in the east and south rarely hold one they can "
    "apply themselves.",

    "AMMOS transfers a published method - MMM_IF02, Guidelines for the "
    "Sustainable Management of Beaches, from AMMIRARE under Interreg "
    "Italia-Francia Marittimo - to three territories and three different "
    "shore types: a dune system at Gomati on Lemnos, Greece; an artificial "
    "recreation pond at Ucari in Denizli province, Türkiye; and a coastal "
    "site in Lebanon.",

    "Because the method was written for marine shores in Italian, transfer is "
    "treated as 21 separate decisions - seven method sheets against three "
    "sites - each recorded as applied, substituted or not applicable with its "
    "reason, and translated into Greek, Turkish and Arabic. The adapted "
    "method is then tested by producing three comparable territorial "
    "baselines with it, demonstrated through reversible visitor-channelling "
    "installations at each site, and carried into costed management plans "
    "formally adopted by the responsible authorities.",

    "Six organisations from Greece, Türkiye, Italy and Lebanon deliver it in "
    "24 months, including the research institute that authored the source "
    "method. Capacity is built where it has to persist: site-manager training "
    "for the staff who will still be there, and school custodian groups at "
    "every site adapted from the same programme's Beach Custodians output.",

    "AMMOS leaves behind more than three managed shores. It produces a "
    "transferability assessment that says where this method stops working, a "
    "cross-typology monitoring platform that lets unlike shores be read "
    "together, and a guide in English and Arabic - all returned to the MMM "
    "common database so the next territory does not start where we did.",
]

# The form's own field list for the logical framework table, so he can see
# what belongs where rather than guessing from four tab names.
FIELDS = [
    ("Project Overall objective",
     "SET BY THE PLATFORM - do not draft it",
     "2.2 (RSO2.4) Promoting climate change adaptation and disaster risk "
     "prevention, resilience taking into account eco-system based "
     "approaches. Already shown correctly at the top of your tab."),
    ("3.2.1 Transfer and adaptation methodology",
     "REQUIRED, 2000 characters - currently 18, holding xxxxx",
     "Below. This is the field blocking your menus. Fill it, press Save, "
     "then add the rows underneath."),
    ("Specific Objective",
     "Chosen, not written",
     "Select SO 2.2 (RSO2.4). It is the only one we are under, which is why "
     "we get exactly ONE expected result."),
    ("Expected result",
     "Typed, then click Add",
     "Below. One only - the Guidelines allow one per specific objective."),
    ("Link expected result to work packages",
     "A linking step, easy to miss",
     "Create WP3 to WP6, then link each to the expected result. The Outputs "
     "tab reads its work package list from these links, which is why it is "
     "dead until they exist."),
    ("Outputs tab",
     "8 outputs: title, target value, measurement unit, semester",
     "All eight are in the outputs document I sent on 3 October - WP3 and "
     "WP4 two each, WP5 and WP6 two each."),
    ("Indicators tab",
     "Generated by the platform from the specific objective",
     "You associate each output to one of the four available codes and set "
     "the target. RCO84=3, RCO87=7, RCO81=60, RCO116=3, RCR104=3, RCR84=7. "
     "There is no training indicator under our objective - the 60 trained "
     "staff go in RCO81."),
    ("Overview tab",
     "Read-only",
     "Check it against the outputs document once the rest is in."),
]


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(1.8)
        n = self.d.styles["Normal"]
        n.font.name, n.font.size = "Calibri", Pt(10.5)

    def h(self, t, size=13, colour=NAVY, space=12):
        p = self.d.add_paragraph()
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(clean(t))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(size), colour

    def p(self, t, size=10.5, colour=None, italic=False, bold=False):
        par = self.d.add_paragraph()
        par.paragraph_format.space_after = Pt(5)
        r = par.add_run(clean(t))
        r.font.size, r.italic, r.bold = Pt(size), italic, bold
        if colour:
            r.font.color.rgb = colour

    def table(self, head, rows, widths=None):
        t = self.d.add_table(rows=1, cols=len(head))
        t.style = "Table Grid"
        for i, x in enumerate(head):
            c = t.rows[0].cells[i]
            c.text = ""
            r = c.paragraphs[0].add_run(clean(x))
            r.bold, r.font.size, r.font.color.rgb = True, Pt(8.5), NAVY
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = ""
                r = cells[i].paragraphs[0].add_run(clean(v))
                r.font.size = Pt(8.5)
                if "REQUIRED" in str(v) or "xxxx" in str(v):
                    r.font.color.rgb = RED
                    r.bold = True
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def fit(num, parts, limit):
    full = clean("\n\n".join(parts) if isinstance(parts, list) else parts)
    n = len(full)
    assert n <= limit, (
        f"{num} is {n} characters and the box holds {limit}. "
        f"Cut {n - limit} before this ships - the platform truncates "
        f"silently on paste.")
    flag = "" if n >= limit * THIN else "   <- thin for a scored box"
    print(f"  {num:46s} {n:>5} / {limit}   {100 * n / limit:3.0f}%{flag}")
    return full, n


def main():
    s321, n321 = fit("3.2.1 Transfer and adaptation methodology", S321, 2000)
    res, nres = fit("Expected result", EXPECTED_RESULT, 2000)
    summ, nsum = fit("Summary", SUMMARY, 2000)

    # Things that must be true of this text or it fails a scored question.
    fails = []
    low = s321.lower()
    for word, why in [
        ("transfer", "the field is named transfer and adaptation methodology"),
        ("adapt", "adaptation is half the field's name"),
        ("test", "the guide asks for testing"),
        ("feasib", "the guide asks for a realistic, feasible approach"),
        ("innovat", "the guide asks explicitly for innovative methods"),
    ]:
        if word not in low:
            fails.append(f"3.2.1 never says {word!r} - {why}")
    if "MMM_IF02" not in s321:
        fails.append("3.2.1 must name the capitalised output")
    if "21" not in s321:
        fails.append("3.2.1 should carry the 21-decision framing, which is "
                     "the whole methodological claim")
    # One expected result, and it has to contain all three of transfer,
    # adoption and replication because we only get one.
    for word in ("adapt", "adopt", "replica"):
        if word not in res.lower():
            fails.append(f"the single expected result does not cover "
                         f"{word!r} - with only one allowed it must carry "
                         f"transfer, adoption AND replication")
    if "Türkiye" not in summ and "Turkiye" not in summ:
        fails.append("the summary does not name all three countries")
    for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson", r"xxxx"):
        if re.search(forbidden, s321 + res + summ, re.I):
            fails.append(f"text contains {forbidden!r}")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — Logical Framework and Summary, paste-ready", size=16)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}. Title confirmed as: "
        f"{TITLE} (acronym {ACRONYM}).", size=9.5, colour=GREY, italic=True)

    D.h("First — why your menus were not responding", size=13)
    D.p("At the top of the Logical Framework tab, 3.2.1 Transfer and "
        "adaptation methodology is a REQUIRED field, marked with an "
        "asterisk, and it currently holds 18 characters of xxxxx. A required "
        "field with placeholder text, sitting above the rows you were trying "
        "to add. Fill it, press Save, then add the Specific "
        "Objective and Expected result underneath. I expect the menus to "
        "come alive at that point.", size=10.5)
    D.p("If they still do not, the next thing to check is the link step: "
        "the Outputs tab builds its work package dropdown from the "
        "expected-result-to-work-package links, so it stays dead until at "
        "least one link exists.", size=10.5)

    D.h("Answering \"do I have all tab data?\" — no, and here is what is "
        "missing", size=13)
    D.p("Your own screenshot is the best answer. Four scored sections are at "
        "zero, and between them they hold 28 of the 100 points:", size=10.5)
    D.table(["Section", "Shown", "What it is worth / what it needs"],
            [["Summary", "0%", "2.000 characters. Not scored directly but it "
                              "is the first thing every evaluator reads. "
                              "Written below."],
             ["Logical Framework", "15%", "3.2.1, one expected result, the WP "
                                          "links, 8 outputs, 6 indicators. "
                                          "All below or already sent."],
             ["Effectiveness", "0%", "Criterion 4, 20 points. This is where "
                                     "the work packages and the outputs go."],
             ["Sustainability", "0%", "Criterion 5, 12 points. Not drafted "
                                      "yet - the stewardship committees and "
                                      "the geopark maintenance route are the "
                                      "spine of it."],
             ["Cost-effectiveness", "0%", "Criterion 6, 12 points. Not "
                                          "drafted yet."],
             ["Horizontal principles", "0%", "Criterion 7, 4 points. Not "
                                             "drafted yet."],
             ["Partnership", "50%", "Three of six partners entered - "
                                    "Geographic coverage still reads Greece, "
                                    "Greece, Türkiye. Italy and Lebanon "
                                    "missing."],
             ["WP and budget part 1", "17%", "Opens up once the outputs "
                                             "exist."]],
            widths=[4.4, 1.6, 11.0])
    D.p("")
    D.p("Tell me which of the three undrafted scored sections you want first "
        "and I will write it today. If you are asking me to choose: "
        "Sustainability, because 12 points and we already have the best "
        "answer to it and have never written it down.", size=10.5, bold=True)

    D.h("3.2.1 Transfer and adaptation methodology", size=14)
    D.p(f"{n321} of 2000 characters. Replace the xxxxx with this and press "
        f"Save.", size=9, colour=GREY, italic=True)
    D.p(s321)

    D.h("Specific Objective", size=13)
    D.p("Select 2.2 (RSO2.4) from the list. Nothing to type. Because it is "
        "the only objective we sit under, we get exactly one expected "
        "result.", size=10.5)

    D.h("Expected result — one only", size=13)
    D.p(f"{nres} characters. Type this, then click Add.", size=9,
        colour=GREY, italic=True)
    D.p(res)
    D.p("It is 381 characters in a 2000-character box and that is "
        "deliberate. Everywhere else I have told you to fill the box, "
        "because a half-empty scored field reads as having run out of things "
        "to say. This field is the exception: an expected result is a single "
        "statement that indicators are then attached to, and padding it "
        "would make it worse, not better.", size=10, italic=True)
    D.p("The Guidelines allow one expected result per specific objective. "
        "The temptation is to write three - transfer, adoption, replication "
        "- so this single sentence carries all three deliberately: applied "
        "(transfer), sustained through adopted plans (adoption), replicable "
        "through the guide and the open platform (replication).", size=10)

    D.h("Then the rest of the Logical Framework, in order", size=13)
    D.table(["Field", "What kind of field it is", "What goes in it"],
            [[a, b, c] for a, b, c in FIELDS], widths=[4.6, 4.4, 8.0])

    D.h("Summary", size=14)
    D.p(f"{nsum} of 2000 characters. The Summary tab is at 0% and it is the "
        f"first thing an evaluator reads.", size=9, colour=GREY, italic=True)
    D.p(summ)

    D.d.save(OUT)
    print(f"\nwrote {OUT}")
    print(f"title: {TITLE}")


if __name__ == "__main__":
    main()
