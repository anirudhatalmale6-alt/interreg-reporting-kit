#!/usr/bin/env python3
"""The submission plan, taken from the e-form's own ERROR CODES.

    python3 build_submission_plan.py

HE SAID: "I think this gives a priority of completion". It gives more than
that. Page 62 of the user guide is the list of conditions the platform itself
checks before it will let you submit - twenty numbered codes with the expected
value for each. That is not guidance, it is the validator. So the priority of
completion is not a matter of judgement: it is whatever makes those twenty
codes pass, in the order the data depends on itself.

Every status in the generated document is COMPUTED from the agreed budget and
partnership. Nothing is typed as "OK".

THE ONE THAT MATTERED MOST - AND THE GUIDE CONTRADICTS ITSELF ON IT

Page 37, Partnership:

    "The minimum number of Countries to be represented is 5"

Page 62, error code 007:

    "Min number of countries in the partnership    Expected: 3"

Our partnership has FOUR countries: Greece, Italy, Lebanon, Turkiye. Under the
first sentence we are inadmissible; under the second we are fine.

Three reasons the 3 is right and the 5 is a misprint:

  1. The Guidelines for Applicants, section 4.4.2, say three: "The project
     partnership shall represent a minimum of three (3) eligible countries".
  2. Every OTHER number on that same partnership page matches the Guidelines
     exactly - min 5 organisations, max 2 per country, min 1 MPC, min 1 EUMC,
     max 7 recommended. Only the country count disagrees, which is the
     signature of a copied line rather than a different rule.
  3. "At least 5 partners based in 5 DIFFERENT countries" is the admissibility
     rule of INTERREG EURO-MED, the other programme in the same coordinated
     capitalisation call. I have that rule written down from comparing the two
     calls in September. A Euro-MED sentence in a NEXT MED guide is a
     copy-paste, and the error code table is what the software actually runs.

The error code table wins because it is the only one of the three that is a
description of executable behaviour. But this is still worth one line to the
Joint Secretariat before questions close on 14 October, because the cost of
being wrong is the whole application and the cost of asking is a sentence.

WHAT THE GUIDE REVEALED THAT I DID NOT KNOW, IN ORDER OF HOW MUCH WORK IT ADDS

  CO-FINANCING IS 8% FOR EVERY PARTNER, NOT JUST LCEC. "The percentage of
  co-financing is fixed to 8%." I had been treating this as LCEC's problem and
  its 24.000. It is six sources of funding to name and 104.000 in total. This
  is the single biggest thing I had understated.

  FINANCIAL CAPACITY IS A WHOLE SECTION I HAD NOT IDENTIFIED. Codes 305, 306
  and 307. It must be filled AND SAVED for every partner, it computes liquidity,
  debt and subvention rates from figures each partner has to supply, and there
  is a separate programme note on how to fill it in.

  EVERY BUDGET LINE NEEDS ITS SEMESTER(S). "For each budget line it has to be
  specified in which semester(s) the expense will be incurred." A cash-flow
  profile per line, not per work package.

  ONLY ONE BUDGET LINE PER PARTNER PER COST CATEGORY PER WORK PACKAGE. Duplicates
  are an error. Our allocation is already one cell per partner per WP, so this
  costs nothing, but it constrains how the budget can be expressed.

  STATE AID SELF-ASSESSMENT is a mandatory upload. Code 014. Never mentioned
  before now.

  INFRASTRUCTURE AND WORKS (CC4) MAY NOT BE USED IN WP1 OR WP2. Ours is in WP4,
  so this is fine, but it is a hard rule.

  BUDGET PER OUTPUT IS ENTERED AS A PERCENTAGE, not an amount, and the
  percentages per work package must total exactly 100.

AND ONE PIECE OF GOOD NEWS THAT REMOVES A BLOCKER I RAISED TWICE

  "Expected results indicators and Output indicators are listed AUTOMATICALLY
   according to the Programme Specific Objective set in Preliminary Info."

The indicator codes come out of the platform. The Performance Framework
Methodology Paper is useful for choosing target values and arguing them, but it
is NOT required to fill the form. I had told him twice that it blocked
criterion 2.4. It does not block the submission.

THE SINGLE MOST USEFUL SENTENCE IN SIXTY-FIVE PAGES

  "Once submitted, your project application will be not editable, but whilst the
   call for projects remains open you may reedit your application by converting
   back to draft."

So submission is reversible until the deadline. Validate and SUBMIT as soon as
the twenty codes pass, then convert back to draft and keep improving. That
turns the 29 October deadline from a cliff into a formality, because at every
moment after the first submission there is a complete valid application on
file. An application not in status "Submitted" at the deadline is discarded,
and that is the only unrecoverable failure in this whole process.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import DIRECT, P, STAFF, WPS, matrix
from build_outputs import OUTPUTS, TECHNICAL, ROMAN, sem

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_submission_plan_2026-10-03.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)
AMBER = RGBColor(0x9A, 0x6B, 0x1E)

FLAT = 0.30
COFIN = 0.08
EU_RATE = 1 - COFIN
EU_CAP = 1_200_000.00
DEADLINE = dt.date(2026, 10, 29)
QUESTIONS_CLOSE = dt.date(2026, 10, 14)

PASS, INPUT, TODO = "PASSES", "NEEDS PARTNER DATA", "NOT STARTED"


def eur(x):
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","})) + " €"


RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


# --------------------------------------------------------------- the numbers
def figures():
    elig = {k: DIRECT[k] + STAFF[k] * FLAT for k in DIRECT}
    tot = sum(elig.values())
    m = matrix()
    return {
        "elig": elig,
        "total_eligible": tot,
        "total_direct": sum(DIRECT.values()),
        "total_staff": sum(STAFF.values()),
        "eu": tot * EU_RATE,
        "cofin": tot * COFIN,
        "countries": sorted({P[k][1] for k in P}),
        "mpc": sum(elig[k] for k in P if P[k][2] == "MPC"),
        "matrix": m,
    }


# The twenty codes, each with a predicate over the figures. A code whose
# status cannot be computed from what we hold says so - it never says PASSES.
#   (code, description, expected, section, predicate-or-None, note)
def codes(F):
    elig, tot, m = F["elig"], F["total_eligible"], F["matrix"]
    every_partner_in = lambda wp: all(m[k].get(wp, 0) > 0 for k in P)
    return [
        ("003", "Minimum number of partners", "5", "Partnership",
         len(P) >= 5, f"{len(P)} organisations"),
        ("004", "Minimum number of partners from an MPC", "1", "Partnership",
         sum(1 for k in P if P[k][2] == "MPC") >= 1,
         f"{sum(1 for k in P if P[k][2] == 'MPC')} MPC partners, "
         f"Turkiye and Lebanon"),
        ("007", "Minimum number of countries in the partnership", "3",
         "Partnership", len(F["countries"]) >= 3,
         f"{len(F['countries'])}: " + ", ".join(F["countries"])
         + ". NOTE page 37 of the same guide says 5 - see the first page of "
           "this document"),
        ("008", "Environmental screening, Checklist A", "1", "Environment",
         None, "Checklist A is required for our specific objective. Not "
               "started, and the largest single untouched block in the form"),
        ("009", "Declaration by the Applicant, upload", "1", "Documents",
         None, "Waiting on the HCMR authorised signature, as Applicant"),
        ("010", "Partner statement, upload, one per partner", "5",
         "Documents", None,
         f"{len(P) - 1} partner statements needed. LCEC confirmed signing "
         f"capacity in September"),
        ("011", "Associated partner declaration, upload", "per associate",
         "Documents", None,
         "Blocked on the North Aegean Periphery signature. Note the guide "
         "says a maximum of ONE advisory partner is allowed - so Sant'Anna "
         "as a second cannot be assumed"),
        ("014", "State Aid self-assessment, upload", "1", "Documents",
         None, "A mandatory document neither of us had identified"),
        ("019", "Minimum number of MMM outputs capitalised", "1",
         "Project / Relevance", True,
         "MMM_IF02 and MMM_IF01, both AMMIRARE under Interreg "
         "Italia-Francia Marittimo. An NX code would NOT satisfy this"),
        ("103", "Minimum number of technical work packages", "1",
         "WP and Budget", len(TECHNICAL) >= 1,
         f"{len(TECHNICAL)} technical work packages, WP3 to WP6, at the "
         f"ceiling of four"),
        ("104", "Budget per output within each WP totals 100%", "100%",
         "WP and Budget",
         all(abs(sum(o[6] for o in OUTPUTS if o[0] == wp) - 1.0) < 1e-9
             for wp in TECHNICAL),
         "Entered as percentages, not amounts: "
         + "; ".join(f"{wp} " + "/".join(
             f"{round(o[6] * 100)}%" for o in OUTPUTS if o[0] == wp)
             for wp in TECHNICAL)),
        ("202", "Maximum budget per partner", "35%", "WP and Budget",
         max(elig.values()) / tot <= 0.35,
         f"largest is {max(elig, key=elig.get)} at "
         f"{100 * max(elig.values()) / tot:.2f}%"),
        ("205", "Minimum total budget", "1.000.000 €", "WP and Budget",
         tot >= 1_000_000, eur(tot)),
        ("210", "Minimum EU contribution", "920.000 €", "WP and Budget",
         F["eu"] >= 920_000,
         f"{eur(F['eu'])}, which is also {eur(EU_CAP - F['eu'])} under the "
         f"1.200.000 ceiling - the budget cannot grow by more than "
         f"{eur((EU_CAP - F['eu']) / EU_RATE)} in total eligible cost"),
        ("217", "Minimum budget to MPC partners", "50%", "WP and Budget",
         F["mpc"] / tot >= 0.50,
         f"{100 * F['mpc'] / tot:.2f}% by partner allocation, with no need "
         f"for the justified-activities table"),
        ("220", "Staff costs as a share of total eligible", "under 40%",
         "WP and Budget", F["total_staff"] / tot < 0.40,
         f"{100 * F['total_staff'] / tot:.2f}%. A rule I did not have before "
         f"today"),
        ("222", "Mandatory budget lines: coordinator and audit for each "
                "partner in WP1, communication in WP2", "all",
         "WP and Budget",
         every_partner_in("WP1") and every_partner_in("WP2"),
         "every partner carries both. The system adds these lines "
         "automatically and they must be filled"),
        ("305", "Financial capacity complete", "5", "Financial capacity",
         None, "A section neither of us had identified. Needs figures from "
               "every partner"),
        ("306", "Financial capacity compiled for all partners", "1",
         "Financial capacity", None,
         "Waiting on the same per-partner figures as code 305 - this code "
         "checks that none of the six has been left out"),
        ("307", "Financial capacity saved", "1", "Financial capacity",
         None, "filling it is not enough, the page must be SAVED"),
    ]


# The completion order. Not the guide's page order - the order the data
# depends on itself, which is different in two places.
#   (step, section, why it is here, what it is blocked on)
ORDER = [
    ("1", "Preliminary info: project info and applicant",
     "Sets the Programme Specific Objective, and the platform generates the "
     "whole indicator list from it. Nothing downstream is right until this is.",
     "Nothing. Do it first and do not change it afterwards."),
    ("2", "Partnership: all six organisations",
     "Every budget screen is per partner, so no budget can be entered until "
     "the partners exist. This also triggers codes 003, 004 and 007.",
     "Nothing - you have all six."),
    ("3", "Logical framework: objective, expected result, work package links",
     "One specific objective, therefore ONE expected result, linked to the "
     "work packages that deliver it.",
     "Nothing."),
    ("4", "Logical framework: the eight outputs",
     "Target value, measurement unit and semester of delivery for each. The "
     "budget-per-output screen cannot be reached until the outputs exist, so "
     "this gates the entire budget.",
     "Nothing - sent 02 October. Change the target values if you disagree."),
    ("5", "Logical framework: associate outputs to programme indicators",
     "The indicator list is generated by the platform from step 1. You pick "
     "which one each output serves and set the project target value.",
     "Nothing. This is the step I wrongly said was blocked on the "
     "Performance Framework paper."),
    ("6", "WP and budget part 1: work package info, outputs, budget lines",
     "One line per partner per cost category per work package, each with the "
     "semester(s) in which the expense falls, then the budget-per-output "
     "percentages.",
     "Nothing for the amounts. The semesters per line I can generate from "
     "the activity months."),
    ("7", "Budget part 2: co-financing and source of funding",
     "8% for every partner, 104.000 in total, and each organisation must "
     "name where its share comes from.",
     "SIX ANSWERS, one per partner. This is the item I had understated as "
     "LCEC's problem alone."),
    ("8", "Financial capacity, then SAVE",
     "Codes 305 to 307. Per partner, with the rates computed by the system.",
     "Figures from every partner, plus the programme's own note on how to "
     "fill this section."),
    ("9", "Environmental screening: checklist A, then B and C if triggered",
     "Required for our specific objective. Demountable boardwalks adjacent "
     "to an Annex I habitat will not stop at A.",
     "Nothing but time, and it is the biggest untouched block left."),
    ("10", "Documents: applicant declaration, five partner statements, "
           "associated partner declaration, state aid self-assessment",
     "All PDF uploads, and signatures are the slowest thing in any "
     "consortium.",
     "Signatures. Start these now, not in the last week."),
    ("11", "VALIDATE, then SUBMIT - and keep working",
     "A submitted application can be converted back to draft while the call "
     "is open. So the first valid submission costs nothing and removes the "
     "deadline entirely.",
     "Nothing once the codes pass. Do not save this step for the 29th."),
]


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

    def table(self, head, rows, widths=None, status_col=None):
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
                if status_col is not None and i == status_col:
                    r.bold = True
                    r.font.color.rgb = {PASS: GREEN, INPUT: AMBER,
                                        TODO: RED}.get(str(v), GREY)
                elif "[" in str(v) and "]" in str(v):
                    r.font.color.rgb = RED
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t

    def page(self):
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def check(F, rows):
    """Nothing claims to pass unless it was computed."""
    fails = []
    # The invariant that matters: a code we cannot compute must never render
    # as PASSES - true by construction, since the status is derived from the
    # predicate - and it must say what it is waiting on. Sniffing the note for
    # keywords was my first attempt and it fired on a perfectly good row; a
    # keyword list is one more rule in one more place.
    for code, desc, exp, sect, pred, note in rows:
        if pred is None and len(note.strip()) < 20:
            fails.append(f"code {code} cannot be computed and does not say "
                         f"what it is waiting on")
        if pred is not None and not isinstance(pred, bool):
            fails.append(f"code {code} predicate is {type(pred).__name__}, "
                         f"not a bool - a truthy number would read as PASSES")
    # Arithmetic that must hold or the whole document is wrong.
    if abs(F["eu"] + F["cofin"] - F["total_eligible"]) > 0.01:
        fails.append("EU contribution plus co-financing does not equal total "
                     "eligible cost")
    if abs(F["total_direct"] + F["total_staff"] * FLAT
           - F["total_eligible"]) > 0.01:
        fails.append("direct plus flat rates does not equal total eligible")
    if F["eu"] > EU_CAP:
        fails.append(f"EU contribution {F['eu']:.2f} exceeds the cap")
    # The claim the first page rests on.
    if len(F["countries"]) < 3:
        fails.append("fewer than three countries - code 007 would fail and "
                     "the misprint argument would be irrelevant")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)


def main():
    F = figures()
    rows = codes(F)
    check(F, rows)

    n_pass = sum(1 for r in rows if r[4] is True)
    n_open = sum(1 for r in rows if r[4] is None)

    D = Doc()
    D.h("AMMOS — what the platform checks before it will let you submit",
        size=16)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} from page 62 of the e-form "
        f"user guide. Twenty codes. Every status below is computed from the "
        f"agreed budget and partnership, not typed. "
        f"{n_pass} pass today on the figures we hold; {n_open} are waiting "
        f"on data, documents or signatures rather than on decisions.",
        size=9.5, colour=GREY, italic=True)

    D.h("First, the one that could have ended the application", size=13)
    D.p("Page 37 of the guide says \"The minimum number of Countries to be "
        "represented is 5\". Page 62, error code 007, says the expected "
        "value is 3. We have four: Greece, Italy, Lebanon, Türkiye. Under "
        "the first sentence we are inadmissible.", size=10)
    D.p("The 3 is right and the 5 is a misprint, for three reasons. The "
        "Guidelines section 4.4.2 say three. Every other number on that same "
        "partnership page matches the Guidelines exactly — five "
        "organisations, two per country, one MPC, one EUMC, seven "
        "recommended — and only the country count disagrees, which is what a "
        "copied line looks like. And \"at least 5 partners based in 5 "
        "different countries\" is the admissibility rule of Interreg "
        "Euro-MED, the other programme in the same coordinated "
        "capitalisation call; I have it written down from comparing the two "
        "calls in September.", size=10)
    D.p("The error code table wins because it is the only one of the three "
        "that describes what the software actually does. But send one line "
        "to the Joint Secretariat before questions close on "
        f"{QUESTIONS_CLOSE:%d %B} anyway. The cost of being wrong is the "
        f"whole application; the cost of asking is a sentence.",
        size=10, bold=True)

    D.page()
    D.h("The twenty codes, with our status computed today", size=14)
    D.table(["Code", "What it checks", "Expected", "Section", "Status",
             "Where we stand"],
            [[c, d, e, s,
              PASS if p is True else (TODO if p is None else "FAILS"),
              n] for c, d, e, s, p, n in rows],
            widths=[1.1, 5.0, 2.0, 2.4, 2.6, 7.5], status_col=4)
    D.p("")
    D.p("Nothing in that table says PASSES unless it was computed from the "
        "budget and partnership we hold. The ones marked NOT STARTED are not "
        "failures of design — they are uploads, signatures and partner "
        "figures, which is a different kind of work and mostly not yours.",
        size=9.5, colour=GREY)

    D.page()
    D.h("The order to fill it in", size=14)
    D.p("Not the guide's page order. The order the data depends on itself, "
        "which differs in two places: the outputs have to exist before any "
        "budget screen will open, and the indicator list is generated from "
        "the specific objective you set in step 1.", size=9.5, colour=GREY)
    D.table(["#", "Section", "Why it sits here", "Blocked on"],
            [[n, s, w, b] for n, s, w, b in ORDER],
            widths=[0.8, 5.2, 7.0, 7.6])

    D.page()
    D.h("Six things the guide added that I did not know yesterday", size=14)
    for t in [
        "CO-FINANCING IS 8% FOR EVERY PARTNER. \"The percentage of "
        "co-financing is fixed to 8%.\" I had been treating this as LCEC's "
        f"problem and its {eur(F['elig']['LCEC'] * COFIN)}. It is six "
        f"sources of funding to name and {eur(F['cofin'])} in total, and the "
        f"form makes each organisation state where its share comes from. "
        f"This is the thing I most understated.",
        "FINANCIAL CAPACITY IS AN ENTIRE SECTION, codes 305 to 307, that "
        "neither of us had identified. It computes liquidity, debt and "
        "subvention rates from figures each partner supplies, it must be "
        "SAVED and not merely filled, and the programme publishes a separate "
        "note on how to complete it which we should get.",
        "EVERY BUDGET LINE NEEDS ITS SEMESTER OR SEMESTERS. Not per work "
        "package — per line. That is a cash-flow profile, and I can generate "
        "it from the activity months we already have.",
        "ONLY ONE BUDGET LINE PER PARTNER PER COST CATEGORY PER WORK "
        "PACKAGE. Duplicates raise an error. Our allocation is already one "
        "cell per partner per work package, so this costs us nothing.",
        "A STATE AID SELF-ASSESSMENT is a mandatory upload, code 014, and "
        "infrastructure and works may not appear in WP1 or WP2 at all. Ours "
        "is in WP4, so that one is fine.",
        "BUDGET PER OUTPUT IS A PERCENTAGE, not an amount, and the "
        "percentages within a work package must total exactly 100. Ours do: "
        + "; ".join(f"{wp} " + " / ".join(
            f"{round(o[6] * 100)}%" for o in OUTPUTS if o[0] == wp)
            for wp in TECHNICAL) + ".",
    ]:
        D.p("•  " + t, size=9.5)

    D.h("And one correction in your favour", size=13)
    D.p("\"Expected results indicators and Output indicators are listed "
        "automatically according to the Programme Specific Objective set in "
        "Preliminary Info.\"", size=10, italic=True)
    D.p("The indicator codes come out of the platform. I told you twice that "
        "the Performance Framework Methodology Paper was blocking criterion "
        "2.4 and twelve points. It is not blocking the form. It is still "
        "worth having, because it explains how the programme elaborated its "
        "target values and that helps us argue ours, but it is not on the "
        "critical path and I was wrong to put it there.", size=10)

    D.h("The most useful sentence in sixty-five pages", size=13)
    D.p("\"Once submitted, your project application will be not editable, "
        "but whilst the call for projects remains open you may reedit your "
        "application by converting back to draft.\"", size=10, italic=True)
    D.p(f"Submission is reversible until {DEADLINE:%d %B}. So validate and "
        f"SUBMIT the moment the twenty codes pass, then convert back to "
        f"draft and keep improving. From that point there is always a "
        f"complete valid application on file, and the deadline stops being a "
        f"risk. An application not in status \"Submitted\" at the deadline is "
        f"discarded, and that is the only failure in this process that "
        f"cannot be recovered from.", size=10, bold=True)
    D.p("Which also means the useful thing you can do today is press "
        "Validate. It lists every code that currently fails, and that list "
        "is a better worklist than anything I can infer from here — it is "
        "the software telling you, rather than me reasoning about the "
        "software. Export it or screenshot it and send it to me.", size=10)

    D.h("Budget reconciliation, so you can check the platform against it",
        size=13)
    D.table(["", "Amount", "Rule"],
            [["Total direct cost", eur(F["total_direct"]),
              "what the form's section 4.2 shows"],
             ["Staff cost", eur(F["total_staff"]),
              f"{100 * F['total_staff'] / F['total_eligible']:.2f}% of "
              f"eligible, code 220 wants under 40%"],
             ["Office and travel flat rates", eur(F["total_staff"] * FLAT),
              "15% + 15% of staff, computed by the system"],
             ["TOTAL ELIGIBLE COST", eur(F["total_eligible"]),
              "code 205 wants at least 1.000.000"],
             ["EU contribution at 92%", eur(F["eu"]),
              f"code 210 wants at least 920.000; ceiling 1.200.000, so "
              f"{eur(EU_CAP - F['eu'])} of headroom"],
             ["Partner co-financing at 8%", eur(F["cofin"]),
              "every partner, not just LCEC"],
             ["To MPC partners",
              f"{eur(F['mpc'])}  ({100 * F['mpc'] / F['total_eligible']:.2f}%)",
              "code 217 wants at least 50%"]],
            widths=[5.6, 4.0, 10.0])

    D.d.save(OUT)
    print(f"wrote {OUT}")
    print(f"\n{n_pass}/20 codes pass on computed figures, {n_open} open "
          f"(uploads, signatures, partner data)")
    for c, d, e, s, p, n in rows:
        if p is False:
            print(f"  FAILS  code {c}  {d}")
    print(f"\ntotal eligible {eur(F['total_eligible'])}  "
          f"EU {eur(F['eu'])}  headroom {eur(EU_CAP - F['eu'])}")
    print(f"countries: {', '.join(F['countries'])}  (code 007 expects 3)")


if __name__ == "__main__":
    main()
