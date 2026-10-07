#!/usr/bin/env python3
"""Environmental screening, the audit costs, and the WP1 corrections.

    python3 build_screening_and_wp1_fix.py

Five things, and the first two matter most.

1. CHECKLIST C DOES NOT APPLY TO US, AND THAT IS ABOUT FORTY-FIVE BOXES SAVED.

   Checklist A says it, Checklist C repeats it: the environmental report in C
   is required only for "projects including an infrastructure of at least 1M
   Euro or/and the projects including investments in infrastructure with a
   lifespan of 5 years or more".

   AMMOS has ZERO in Infrastructures and works - the whole budget model v6
   carries nothing in that category - and the installations are demountable
   and modular by design, specified to be removable without lasting trace.
   Neither threshold is met. Not the million, and not the five years.

   That is also the concrete payoff of the budget decision we are about to
   write into 3.6.2: works bought in as fabrication and installation services,
   on reversible structures, keep us out of the heavier assessment entirely.

2. HIS CHECKLIST A TOGGLES ARE CURRENTLY THE WORST POSSIBLE COMBINATION.

   The screenshot shows question 1 - does the project foresee the realization
   of infrastructures - switched ON, and questions 2, 3 and 4 switched OFF.
   Read literally that says: we are building infrastructure, we have NOT taken
   account of environmental protection and sustainable development, we have
   NOT taken account of environmental law, and the environmental context did
   NOT inform our strategy.

   For a project installing structures beside an Annex I habitat that is close
   to the worst sentence in the application. Those three are YES, obviously
   and demonstrably so - the project is ecosystem-based adaptation, it cites
   the Habitats Directive and the EEA litter protocol, and the whole design
   follows from site characterisation.

   Question 5, potential impact on the environment, is also YES. Saying no
   while installing boardwalks on a dune invites exactly the scrutiny the
   honest answer avoids: yes, here is the impact, here is why it is reversible,
   here is what we measure.

3. THE AUDIT COSTS HE ASKED ABOUT - AND MY FIGURE WAS WRONG.

   My proportional allocation spread each partner's whole external expertise
   budget across work packages by work package share, which dropped a large
   slice into WP1 and labelled it "Auditor costs". It produced 15.112 for HCMR
   and 42.412 in total. That is not what first-level control costs; it is a
   share of a partner's consultancy budget wearing the wrong label.

   First-level control is a FIXED cost, not a proportional one: four reporting
   periods, one certificate each, per partner. 23.000 across six partners,
   with the Applicant higher because it also consolidates. The remainder of
   each partner's external expertise moves to the work packages that actually
   buy services.

   Every partner total and every cost-category total is unchanged. Only the
   work package distribution moves, and WP1 falls from 90.472 to 71.060 -
   6,1% of direct cost, which is a defensible management share.

4. WP1 OUTPUT SEMESTERS: why he could not pick 24 months.

   The activity month selector is bounded by its output's semester of
   delivery. Output1.1 was semester I, so an activity running M1-M24 cannot be
   added under it. He fixed it by changing the output to semester IV. That
   works, but it loses something: an output genuinely delivered in the first
   semester is good news in a Gantt. The cleaner fix is to move the continuous
   activity to Output1.2, which is semester IV anyway.

5. WP1 HAS NO TARGET GROUP FIELD. He checked; only the others do. So the two
   WP1 target groups I wrote are not needed - the remaining ten are.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import DIRECT, P, PP, STAFF, WPS, matrix

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Environmental-screening_and_WP1-corrections_2026-10-09.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)

FLAT = 0.30
CATS = ["ST", "EC", "ES", "IW"]
CAT_NAME = {"ST": "Staff cost", "EC": "Equipment costs",
            "ES": "External expertise and services costs",
            "IW": "Infrastructures and works costs"}
CC = {"HCMR": [173500., 30000., 94450., 0.], "UAEG": [60000., 8000., 119000., 0.],
      "LCEC": [80000., 20000., 176000., 0.], "PIYA": [50000., 10000., 155000., 0.],
      "ACIP": [30000., 12000., 84000., 0.], "CNR": [50000., 0., 15000., 0.]}
# First-level control: a fixed cost per partner, four reporting periods.
AUDIT = {"HCMR": 6000., "UAEG": 3500., "PIYA": 3500., "ACIP": 3000.,
         "CNR": 3000., "LCEC": 4000.}

LIM_A, LIM_B1, LIM_B2 = 500, 1000, 500

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


def allocate():
    """ES in WP1 is the audit and nothing else; the rest goes to WP2-WP6."""
    m = matrix()
    wp_ids = [w[0] for w in WPS]
    out = {}
    for k in P:
        share = {wp: m[k].get(wp, 0.) / DIRECT[k] for wp in wp_ids
                 if m[k].get(wp, 0.) > 0}
        st, ec, es, iw = CC[k]
        rest = es - AUDIT[k]
        non1 = sum(s for wp, s in share.items() if wp != "WP1")
        out[k] = {}
        for wp, s in share.items():
            out[k][wp] = {"ST": st * s, "EC": ec * s, "IW": iw * s,
                          "ES": AUDIT[k] if wp == "WP1" else rest * s / non1}
    return out, m, wp_ids


# ------------------------------------------------------------- CHECKLIST A
A_TOGGLES = [
    (1, "Does the project foresee the realization of infrastructures?", "YES",
     "Yes, and answering yes is both honest and safe: the installations are "
     "physical, and the threshold that triggers Checklist C is not the "
     "existence of works but their scale and lifespan."),
    (2, "Does the project take into account the principles of environmental "
        "protection and sustainable development?", "YES",
     "Currently switched OFF in the form, and it must be YES. The project's "
     "whole purpose is reducing erosion of sediment shores under visitor "
     "pressure through an ecosystem-based approach."),
    (3, "Does the project take into account the relevant International, "
        "National and Regional directives, laws, agreements and strategies?",
     "YES",
     "Currently OFF, and must be YES. The Greek site is an EU Habitats "
     "Directive Annex I habitat type; litter monitoring follows the European "
     "Environment Agency aligned protocol; the proposal is written against "
     "the MSSD, the 2030 Greener Med Agenda and the European Ocean Pact."),
    (4, "Has the environmental context been taken into account when deciding "
        "on the strategies and activities of the project proposal?", "YES",
     "Currently OFF, and must be YES. Nothing is installed until each site is "
     "characterised, and the design follows that baseline rather than a "
     "template - that is literally the environmental context deciding the "
     "activities."),
    (5, "Does the project have any potential impact on the environment?",
     "YES",
     "Say yes. Claiming no impact while installing boardwalks on a dune "
     "system invites exactly the scrutiny the honest answer avoids. Yes, "
     "here is the impact, here is why it is reversible, here is what we "
     "measure before and after."),
]

A_Q1 = (
    "Reversible, demountable visitor-channelling installations at three "
    "sites: boundary marking, interpretation signage, modular boardwalks over "
    "the vulnerable sediment surface, and demarcated zones where trampling "
    "degrades vegetation or banks. Every component is removable without "
    "lasting trace. No dredging, no bottom or underwater cleaning, no "
    "sediment importation, no non-native planting. The budget carries nothing "
    "in Infrastructures and works; fabrication and installation are procured "
    "as services."
)

A_EXTRA = (
    "AMMOS is an environmental project: it transfers a published method for "
    "managing sediment shores under tourism pressure, to reduce erosion and "
    "trampling. Each site is characterised before intervention - topography, "
    "sediment, vegetation, ecological status - and monitored through a full "
    "visitor season after, so the project evidences its own effect rather "
    "than asserting it. Each site passes national consent and screening "
    "first. If monitoring shows harm, the installation is removed."
)

# ------------------------------------------------------------- CHECKLIST B
B_AREA = (
    "Three demonstration sites in three countries, each a sediment shore "
    "under visitor pressure and each a different shore type. GOMATI, Lemnos, "
    "North Aegean, Greece: a marine dune and beach system with an Annex I "
    "habitat type, lying outside Natura 2000 site GR4110006 and outside the "
    "adjacent Faraklou Geological Park, so under pressure and covered by "
    "neither designation. UCARI, Acipayam, Denizli province, Turkiye: an "
    "artificial recreation pond fed by groundwater, with heavy seasonal "
    "visitor concentration on unconsolidated sediment and no outlet to the "
    "sea. A COASTAL SITE IN LEBANON, where the national coastal management "
    "policy framework exists but the applicable field method does not. The "
    "three were chosen because they are dissimilar: the set spans marine "
    "dune, marine coast and inland freshwater, which is what establishes "
    "where the transferred method works and where it stops."
)

B_IDENT = (
    "Gomati, Lemnos, Region of North Aegean, Greece - marine dune and beach "
    "system, Annex I habitat type, outside Natura 2000 GR4110006. Ucari pond, "
    "Acipayam, Denizli province, Turkiye - artificial groundwater-fed "
    "recreation pond, no marine connection. Lebanese coastal site - to be "
    "confirmed with LCEC. Exact coordinates and site boundaries are in the "
    "map attached to this checklist."
)

B_INTERVENTION = (
    "At each site, visitor channelling and shore protection designed from "
    "that site's own baseline rather than from a template. Components: "
    "boundary marking poles delimiting the protected surface; interpretation "
    "signage stating what is protected and why; demountable modular "
    "boardwalks carrying foot traffic over the vulnerable sediment; and "
    "demarcated or lightly fenced zones where trampling is actively degrading "
    "vegetation or banks. Every component is modular and removable without "
    "lasting trace. Scale is small and local - footpath-width boardwalk runs "
    "and signage, not engineered coastal defence. There is no excavation "
    "beyond post footings, no dredging, no bottom or underwater cleaning, no "
    "sediment importation and no planting of non-native species. The design "
    "at each site is signed off by the authority that will own it, and each "
    "site passes its own national consent and environmental "
    "screening before works. Installation completes at month 20, leaving a "
    "full monitored season inside the project."
)

C_ARGUMENT = [
    "Checklist C is required only for proposals \"including an "
    "infrastructure of at least 1M Euro or/and the projects including "
    "investments in infrastructure with a lifespan of 5 years or more\". "
    "AMMOS meets neither threshold, and the reasons are on the record rather "
    "than asserted.",
    "THE MILLION. The budget carries nothing at all in the Infrastructures "
    "and works cost category - zero for every one of the six partners in "
    "model v6. The whole project is 1.300.000 € of which the physical "
    "installation component is a fraction, procured as fabrication and "
    "installation services across three sites.",
    "THE FIVE YEARS. The installations are demountable and modular by "
    "design, specified to be removable without lasting trace. Reversibility "
    "is stated in the work package, in section 3.5.1 environmental level and "
    "in the do-no-significant-harm box of section 3.7. A structure designed "
    "to be removed is not an investment in infrastructure with a lifespan of "
    "five years or more.",
    "So A and B are completed and C is left empty. If the Joint Secretariat "
    "asks why, the answer is one sentence with two figures in it. That is "
    "about forty-five 500-character boxes of environmental reporting we do "
    "not have to write, and it is the concrete payoff of budgeting the works "
    "as services on reversible structures.",
]

REWORD_361 = (
    "INFRASTRUCTURE AND WORKS IS NOT USED. The demonstration installations "
    "are procured as external expertise and services - fabrication and "
    "installation of demountable modular components - rather than as "
    "capitalised infrastructure, because every component is specified to be "
    "removable without lasting trace. That is deliberate on three counts: it "
    "keeps the intervention proportionate beside an Annex I habitat, it keeps "
    "the environmental screening proportionate to a reversible works "
    "programme, and it allows a section to be repaired or altered rather than "
    "a structure replaced."
)


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(1.6)
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

    def table(self, head, rows, widths=None, small=8.5, flag_col=None):
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
                if flag_col is not None and i == flag_col:
                    r.bold = True
                    r.font.color.rgb = GREEN if "YES" in str(v) else RED
                if str(row[0]).startswith("WP1 TOTAL"):
                    r.bold = True
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def main():
    alloc, m, wp_ids = allocate()
    new_wp = {wp: sum(sum(alloc[k][wp].values()) for k in P if wp in alloc[k])
              for wp in wp_ids}
    old_wp = {wp: sum(m[k].get(wp, 0.) for k in P) for wp in wp_ids}

    fails = []
    for k in P:
        for i, c in enumerate(CATS):
            got = sum(alloc[k][wp][c] for wp in alloc[k])
            if abs(got - CC[k][i]) > 0.01:
                fails.append(f"{k} {c} {got:.2f} vs {CC[k][i]:.2f}")
        if abs(sum(sum(alloc[k][wp].values()) for wp in alloc[k])
               - DIRECT[k]) > 0.01:
            fails.append(f"{k} total moved")
        if alloc[k]["WP1"]["ES"] != AUDIT[k]:
            fails.append(f"{k} WP1 external expertise is not exactly the "
                         f"audit")
    if abs(sum(new_wp.values()) - sum(DIRECT.values())) > 0.05:
        fails.append("grand total moved")
    if any(alloc[k][wp]["IW"] for k in P for wp in alloc[k]):
        fails.append("infrastructure is non-zero somewhere - the Checklist C "
                     "exemption argument rests on it being zero")
    for num, q, ans, why in A_TOGGLES:
        if ans != "YES":
            fails.append(f"Checklist A q{num} recommended {ans} - check the "
                         f"reasoning before shipping a NO")
    for name, text, lim in (("A q1", A_Q1, LIM_A), ("A extra", A_EXTRA, LIM_A),
                            ("B area", B_AREA, LIM_B1),
                            ("B identification", B_IDENT, LIM_B2),
                            ("B intervention", B_INTERVENTION, LIM_B1),
                            ("3.6.1 reword", REWORD_361, 2000)):
        n = len(clean(text))
        if n > lim:
            fails.append(f"{name} is {n} chars, limit {lim} - cut {n - lim}")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — environmental screening, the audit costs, and three WP1 "
        "corrections", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}.", size=9.5, colour=GREY,
        italic=True)

    # ---------------------------------------------------------- Checklist C
    D.h("First: CHECKLIST C DOES NOT APPLY TO US", size=14)
    for t in C_ARGUMENT:
        D.p(t)
    D.p("Do not fill it. Fill A and B.", size=11, bold=True)

    # ---------------------------------------------------------- Checklist A
    D.h("Checklist A — and three of your toggles are the wrong way round",
        size=14)
    D.p("Your screenshot shows question 1 ON and questions 2, 3 and 4 OFF. "
        "Read literally that says: we are building infrastructure, we have "
        "NOT taken account of environmental protection, we have NOT taken "
        "account of environmental law, and the environmental context did NOT "
        "inform our strategy. For a project installing structures beside an "
        "Annex I habitat that is close to the worst sentence in the "
        "application - and it is also untrue of us.", size=10.5, bold=True)
    D.table(["#", "Question", "Answer", "Why"],
            [[str(n), q, a, w] for n, q, a, w in A_TOGGLES],
            widths=[0.8, 5.6, 1.6, 8.4], flag_col=2)
    D.p("")
    D.p(f"Question 1 \"Specify what\" — {len(clean(A_Q1))}/{LIM_A}:",
        size=9, colour=GREY, italic=True)
    D.p(A_Q1)
    D.p(f"Additional information — {len(clean(A_EXTRA))}/{LIM_A}:", size=9,
        colour=GREY, italic=True)
    D.p(A_EXTRA)

    # ---------------------------------------------------------- Checklist B
    D.h("Checklist B, Section 1", size=14)
    D.p(f"Description of the indicative area — "
        f"{len(clean(B_AREA))}/{LIM_B1}:", size=9, colour=GREY, italic=True)
    D.p(B_AREA)
    D.p(f"Area identification — {len(clean(B_IDENT))}/{LIM_B2}:", size=9,
        colour=GREY, italic=True)
    D.p(B_IDENT)
    D.p(f"Description of the intervention/infrastructure — "
        f"{len(clean(B_INTERVENTION))}/{LIM_B1}:", size=9, colour=GREY,
        italic=True)
    D.p(B_INTERVENTION)
    D.p("MAP OF THE INDICATIVE AREA — this is an upload and I cannot make it "
        "for you. One A4 PDF with three panels, one per site, each showing "
        "the site boundary and, for Gomati, the Natura 2000 GR4110006 "
        "boundary alongside it so the gap between designations is visible. "
        "Ask the Aegean for the Gomati panel and Acipayam for Ucari; they "
        "will have GIS already. Send me the three and I will assemble the "
        "PDF.", size=10.5, bold=True)
    D.p("Sections 2 and 3 of Checklist B I have not seen - send me those "
        "screenshots and I will write them the same way.", size=10.5)

    # ---------------------------------------------------------- audit costs
    D.h("The audit costs you asked about — and my figure was wrong", size=14)
    D.p("My proportional allocation spread each partner's whole external "
        "expertise budget across work packages by work package share. That "
        "dropped a large slice into WP1 and labelled it Auditor costs: "
        "15.112 for HCMR, 42.412 in total. That is not what first-level "
        "control costs. It was a share of a consultancy budget wearing the "
        "wrong label, and you were right to ask.", size=10.5, bold=True)
    D.p("First-level control is a FIXED cost, not a proportional one: four "
        "reporting periods, one certificate each, per partner. The Applicant "
        "is higher because it also consolidates six partners' claims.",
        size=10.5)
    D.table(["Partner", "Staff cost", "Auditor costs", "Equipment",
             "WP1 total"],
            [[f"{PP[k]} {P[k][0]}", eur(alloc[k]['WP1']['ST']),
              eur(alloc[k]['WP1']['ES']), eur(alloc[k]['WP1']['EC']),
              eur(sum(alloc[k]['WP1'].values()))]
             for k in sorted(P, key=lambda x: PP[x])]
            + [["WP1 TOTAL", eur(sum(alloc[k]['WP1']['ST'] for k in P)),
                eur(sum(AUDIT.values())),
                eur(sum(alloc[k]['WP1']['EC'] for k in P)),
                eur(new_wp['WP1'])]],
            widths=[5.4, 3.0, 3.2, 2.8, 3.0])
    D.p("")
    D.p("Every partner total and every cost-category total is UNCHANGED. "
        "Only the work package distribution moves, because the external "
        "expertise that was sitting in WP1 goes to the work packages that "
        "actually buy services:", size=10.5)
    D.table(["WP", "Before", "After", "Change"],
            [[wp, eur(old_wp[wp]), eur(new_wp[wp]),
              f"{new_wp[wp] - old_wp[wp]:+,.2f}".replace(",", ".")]
             for wp in wp_ids]
            + [["TOTAL", eur(sum(old_wp.values())), eur(sum(new_wp.values())),
                "0,00"]],
            widths=[1.6, 3.4, 3.4, 3.4])
    D.p("")
    D.p(f"WP1 falls from {eur(old_wp['WP1'])} to {eur(new_wp['WP1'])}, which "
        f"is {100 * new_wp['WP1'] / sum(DIRECT.values()):.1f}% of direct "
        f"cost - a defensible management share. You said the budget is "
        f"corrected except WP1, so these are the rows to enter.", size=10.5)

    # --------------------------------------------------------- WP1 outputs
    D.h("Why you could not choose 24 months, and the cleaner fix", size=13)
    D.p("The activity month selector is bounded by its output's semester of "
        "delivery. Output1.1 was semester I, so an activity running M1-M24 "
        "cannot be added under it. Changing the output to semester IV works, "
        "but it costs you something: an output genuinely delivered in the "
        "first semester looks good in a Gantt, and moving it to IV makes the "
        "project look back-loaded for no reason.", size=10.5)
    D.p("The cleaner fix is to move the continuous activity instead. Keep "
        "Output1.1 at semester I with only the two activities that finish "
        "early, and put the financial circuit - the one that genuinely runs "
        "M1-M24 - under Output1.2, which is semester IV anyway:", size=10.5,
        bold=True)
    D.table(["Output", "Semester", "Activities"],
            [["Output1.1", "I",
              "1.1.1 Coordination structure, roles and RACI table (M1-M6); "
              "1.1.2 Partnership agreement signed and decision rules adopted "
              "(M1-M6)"],
             ["Output1.2", "IV",
              "1.2.1 Financial circuit and first-level control (M1-M24); "
              "1.2.2 Quality assurance and two-stage deliverable check "
              "(M3-M24); 1.2.3 Reporting, risk review and indicator "
              "monitoring (M6-M24)"]],
            widths=[2.0, 1.8, 11.4])
    D.p("")
    D.p("That is a renumbering of what I sent yesterday, not new content: "
        "the old 1.1.2 becomes 1.2.1 and the old 1.2.1 and 1.2.2 become "
        "1.2.2 and 1.2.3. If you have already entered them under semester "
        "IV and would rather not redo it, leave it - nothing is wrong, it is "
        "only less flattering in the overview.", size=10.5)

    # ----------------------------------------------------------- 3.6 reword
    D.h("3.6.1 and 3.6.2 — the infrastructure sentence, as approved", size=13)
    D.p("Replace the sentence beginning \"INFRASTRUCTURE AND WORKS covers "
        "only the demonstration installations\" with this:", size=10.5)
    D.p(REWORD_361, italic=True)
    D.p("It is honest, it matches the budget, and the third clause is the "
        "one that pays: it is the sentence that keeps us out of Checklist C.",
        size=10.5)

    D.h("And two small confirmations", size=12)
    D.p("WP1 has no Target group field - you checked, and you are right. So "
        "the two WP1 target groups I sent are not needed; the other ten are. "
        "WP2 and the four technical work packages all have them.", size=10.5)
    D.p("WP2's mandatory budget line is the single Communication Manager "
        "staff cost, which is the equivalent of WP1's coordinator line. Every "
        "partner still needs a WP2 line, which they have.", size=10.5)

    D.d.save(OUT)
    print(f"wrote {Path(OUT).name}")
    print(f"  audit total {eur(sum(AUDIT.values()))} "
          f"(proportional split had given 42.412,00 €)")
    print(f"  WP1 {eur(old_wp['WP1'])} -> {eur(new_wp['WP1'])}  "
          f"({100 * new_wp['WP1'] / sum(DIRECT.values()):.1f}% of direct)")
    for wp in wp_ids:
        print(f"    {wp}  {eur(new_wp[wp])}")
    print(f"  grand {eur(sum(new_wp.values()))} - unchanged")
    print(f"  Checklist A q1 {len(clean(A_Q1))}/{LIM_A}, "
          f"extra {len(clean(A_EXTRA))}/{LIM_A}")
    print(f"  Checklist B {len(clean(B_AREA))}/{LIM_B1}, "
          f"{len(clean(B_IDENT))}/{LIM_B2}, "
          f"{len(clean(B_INTERVENTION))}/{LIM_B1}")
    print(f"  Checklist C: DOES NOT APPLY")


if __name__ == "__main__":
    main()
