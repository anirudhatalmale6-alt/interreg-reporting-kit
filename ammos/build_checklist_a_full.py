#!/usr/bin/env python3
"""Checklist A in full - ten questions and eleven boxes, not five.

    python3 build_checklist_a_full.py

"CHEKCLIST A-WHEN PRESSING 5 MORE MENU CAME UP" - the toggles reveal the rest
as you answer them. The real shape:

    1  realization of infrastructures                    500
    2  principles of environmental protection            500
    3  international, national, regional directives      500
    4  environmental context informed the strategy       500
    5  potential impact on the environment               500
    6  HOW the plan reduces negative and strengthens
       positive impacts                                 1000   <- the big one
    7  advocacy / awareness-raising message              500
    8  partners' role in reducing impact or building
       stakeholder capacity                              500
    9  environmental monitoring system                   500
    10 costs provided in the budget plan                 500
       Additional information                            500

Q6 is 1000 and everything else is 500, which is the form telling us where the
substance belongs - the same signal the Effectiveness section gave when the
communication box came out at 3000.

THREE OF THESE ARE GIFTS FOR THIS PARTICULAR PROJECT

Q9, environmental monitoring system: we are building one. A shared web-GIS
platform running a common protocol across three shore types, with field data
entry in three languages and a cross-typology comparability layer. Most
applicants answer this question with a promise; we answer it with an output.

Q10, costs provided in the budget plan: we can answer with figures rather than
a yes, because the budget is itemised per work package and the monitoring and
reversible-installation costs are identifiable lines in it.

Q7, advocacy and awareness-raising: nine school custodian groups, on-site
interpretation at three sites in three languages, and a decalogue adapted from
an existing output. Again an output rather than an intention.

And Q5 stays YES. Saying no while installing boardwalks on a dune invites
exactly the scrutiny the honest answer avoids.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import DIRECT, P, STAFF, WPS, matrix

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Checklist-A_all-ten-questions_2026-10-09.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)

AUDIT = {"HCMR": 6000., "UAEG": 3500., "PIYA": 3500., "ACIP": 3000.,
         "CNR": 3000., "LCEC": 4000.}
CC = {"HCMR": [173500., 30000., 94450., 0.], "UAEG": [60000., 8000., 119000., 0.],
      "LCEC": [80000., 20000., 176000., 0.], "PIYA": [50000., 10000., 155000., 0.],
      "ACIP": [30000., 12000., 84000., 0.], "CNR": [50000., 0., 15000., 0.]}
FLAT = 0.30

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


def wp_totals():
    m = matrix()
    wp_ids = [w[0] for w in WPS]
    out = {}
    for k in P:
        share = {wp: m[k].get(wp, 0.) / DIRECT[k] for wp in wp_ids
                 if m[k].get(wp, 0.) > 0}
        rest = CC[k][2] - AUDIT[k]
        non1 = sum(s for wp, s in share.items() if wp != "WP1")
        out[k] = {wp: {"ST": CC[k][0] * s, "EC": CC[k][1] * s,
                       "IW": CC[k][3] * s,
                       "ES": AUDIT[k] if wp == "WP1" else rest * s / non1}
                  for wp, s in share.items()}
    return {wp: sum(sum(out[k][wp].values()) for k in P if wp in out[k])
            for wp in wp_ids}


# (number, question, answer, limit, text)
Q = [
    (1, "Does the project foresee the realization of infrastructures?", "YES",
     500,
     "Reversible, demountable visitor-channelling installations at three "
     "sites: boundary marking, interpretation signage, modular boardwalks "
     "over the vulnerable sediment surface, and demarcated zones where "
     "trampling degrades vegetation or banks. Every component is removable "
     "without lasting trace. No dredging, no bottom or underwater cleaning, "
     "no sediment importation, no non-native planting. The budget carries "
     "nothing in Infrastructures and works; fabrication and installation are "
     "procured as services."),

    (2, "Does the project take into account the principles of environmental "
        "protection and sustainable development?", "YES", 500,
     "They are the subject matter, not a safeguard bolted on. AMMOS exists "
     "to let tourism continue on Mediterranean sediment shores without "
     "consuming the surface it depends on, through an ecosystem-based "
     "approach under specific objective 2.2. All three dimensions are "
     "addressed without trading one against another: reduced erosion and "
     "trampling, measured before and after; the shore valued as the asset "
     "the local economy rests on; and local staff, schools and standing "
     "committees left able to maintain it."),

    (3, "Does the project take into account the relevant International, "
        "National and Regional directives, laws, agreements and strategies?",
     "YES", 500,
     "The Greek site is an EU Habitats Directive Annex I habitat type, and "
     "the method produces the condition data any future designation would "
     "require. Litter monitoring is adopted from the published European "
     "Environment Agency aligned protocol, so the series stays comparable "
     "with the European dataset. The proposal is written against the "
     "Mediterranean Strategy for Sustainable Development, the 2030 Greener "
     "Med Agenda and the European Ocean Pact, and each site passes its "
     "national consent regime first."),

    (4, "Has the environmental context been taken into account when deciding "
        "on the strategies and activities of the project proposal?", "YES",
     500,
     "The environmental context decides the activities rather than being "
     "consulted about them. Nothing is procured, installed or trained until "
     "each site is characterised - topography, sediment, vegetation, "
     "ecological status - and each installation is designed from that site's "
     "own baseline rather than a template. The choice of three dissimilar "
     "shore types, including an artificial freshwater pond, follows from the "
     "environmental question the project asks: where does this method stop "
     "working, and why?"),

    (5, "Does the project have any potential impact on the environment?",
     "YES", 500,
     "Yes, and stating it is the point. The project places structures on "
     "sediment shores, one an Annex I habitat type under recorded visitor "
     "pressure: localised disturbance during installation, post footings in "
     "the sediment, and concentrated foot traffic along the channelled "
     "route. The intended net effect is less trampling and erosion across "
     "the wider surface, and the project measures that rather than asserting "
     "it. Effects are local, reversible and monitored through a full visitor "
     "season."),

    (6, "How will the project plan reduce the negative impacts and strengthen "
        "the positive impacts?", "", 1000,
     "REDUCING THE NEGATIVE. Reversibility: every component is demountable "
     "and removable without lasting trace, so a design that performs badly "
     "in its first season can be altered rather than defended. Measure "
     "before building: nothing is installed until the site "
     "is characterised and the design follows that baseline. Nothing "
     "extractive: no dredging, no bottom cleaning, no sediment importation, "
     "no non-native planting. National consent and screening at each site "
     "before works. And a stopping rule: if monitoring shows harm, the "
     "installation is removed.\n\n"
     "STRENGTHENING THE POSITIVE. Three things make the reduction in "
     "trampling and erosion outlast the funding. The protocol stays with "
     "trained municipal staff and nine school custodian groups, so "
     "observation continues. Each site ends with a costed management plan "
     "formally adopted by its authority and overseen by a standing "
     "stewardship committee. And the adapted method returns to the MMM "
     "database under open licence, so a fourth territory can use it."),

    (7, "Is any message on advocacy or awareness-raising related to "
        "environmental issues foreseen?", "YES", 500,
     "Three strands, all outputs rather than intentions. On-site "
     "interpretation at all three sites explaining what is protected and "
     "why, in Greek, Turkish and Arabic, because a boundary nobody "
     "understands is a boundary people walk around. Nine school custodian "
     "groups across three countries, each running a monitoring round and an "
     "awareness action at its own shore. And the BEach CLEAN decalogue "
     "adapted per site for visitors and operators, extending an existing "
     "Arabic version."),

    (8, "Does any of the partners or associates involved have a role in "
        "reducing the negative impact or strengthening stakeholders' capacity "
        "to cope with it?", "YES", 500,
     "All of them, by design. The two research partners carry the "
     "characterisation and ecological status assessment that bound the "
     "intervention. The youth partner runs the custodian groups. The "
     "municipality leads site-manager training, as the authority that will "
     "run a site afterwards. The Lebanese national agency leads "
     "institutional uptake so measures enter policy. The associated regional "
     "authority carries the result into its own planning. Sixty staff are "
     "trained to run the monitoring themselves."),

    (9, "Does the project foresee an environmental monitoring system?", "YES",
     500,
     "It builds one, as a project output. A shared web-GIS platform carrying "
     "one common protocol across all three territories, with field data "
     "entry by site staff in their own language, under open licence with "
     "documented export formats. Litter monitoring follows the EEA-aligned "
     "protocol so the data is comparable with the European dataset. The "
     "cross-typology comparability layer lets a freshwater pond and a marine "
     "dune be read side by side - the precondition for detecting trend."),

    (10, "Have the costs for the above-mentioned measures been adequately "
         "provided in the budget plan?", "YES", 500,
     "Yes, and as identifiable lines rather than a contingency. The "
     "demonstration work package [WP4] covers the reversible installations "
     "and the monitoring platform with its comparability layer. "
     "Site characterisation before intervention sits in [WP3]. Training the "
     "people who will run the monitoring sits in [WP5]. "
     "Consent and screening support is budgeted under external expertise. "
     "Equipment is specified after characterisation, so nothing is bought "
     "against an assumption."),

    (0, "Additional information", "", 500,
     "AMMOS is an environmental project: it transfers a published method for "
     "managing sediment shores under tourism pressure, to reduce erosion and "
     "trampling. Each site is characterised before intervention and "
     "monitored through a full visitor season after, so the project "
     "evidences its own effect rather than asserting it. Checklist C is not "
     "provided because neither threshold is met: nothing in Infrastructures "
     "and works, and the installations are demountable, not five-year "
     "investments."),
]


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(1.7)
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

    def table(self, head, rows, widths=None, small=8.5, flag=None):
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
                if flag is not None and i == flag and str(v).strip():
                    r.bold = True
                    r.font.color.rgb = GREEN if "YES" in str(v) else RED
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def main():
    wp = wp_totals()
    subs = {"[WP3]": eur(wp["WP3"]), "[WP4]": eur(wp["WP4"]),
            "[WP5]": eur(wp["WP5"])}

    fails, report = [], []
    for num, q, ans, lim, text in Q:
        t = clean(text)
        for k, v in subs.items():
            t = t.replace(k, v)
        n = len(t)
        if n > lim:
            fails.append(f"Q{num} is {n} chars, limit {lim} - cut {n - lim}")
        if re.search(r"\[[A-Z][A-Z0-9_ ]{1,}\]", t):
            fails.append(f"Q{num} still has an unsubstituted placeholder")
        report.append((num, q, ans, lim, n, t))
    # No answered toggle may be NO without a reason being argued.
    for num, q, ans, lim, text in Q:
        if ans not in ("YES", ""):
            fails.append(f"Q{num} answer is {ans!r} - every question here is "
                         f"answerable YES for this project; check before "
                         f"shipping a NO")
    blob = " ".join(x[5] for x in report)
    for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson", r"xxxx"):
        if re.search(forbidden, blob, re.I):
            fails.append(f"text contains {forbidden!r}")
    # Q6 is the 1000 box; it must actually use the space.
    q6 = [r for r in report if r[0] == 6][0]
    if q6[4] < 850:
        fails.append(f"Q6 is only {q6[4]}/1000 - the form gave it double the "
                     f"space for a reason")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — Checklist A, all ten questions and eleven boxes", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}. The toggles reveal the rest as "
        f"you answer them, which is why five became ten. Every box checked "
        f"against its own counter.", size=9.5, colour=GREY, italic=True)

    D.p("Notice the sizes: question 6 is 1000 characters and every other box "
        "is 500. That is the form telling us where the substance belongs - "
        "the same signal the Effectiveness section gave when the "
        "communication box came out at 3000 while its neighbours were 2000. "
        "So question 6 gets the most work of the eleven.", size=10.5)

    D.h("Answers at a glance", size=13)
    D.table(["Q", "Question", "Toggle", "Chars"],
            [[(str(n) if n else "extra"), q, a, f"{c}/{l}"]
             for n, q, a, l, c, _ in report],
            widths=[1.0, 9.4, 1.8, 2.0], flag=2)
    D.p("")
    D.p("All ten are YES, and none of that is a stretch. Three of them are "
        "gifts for this particular project: question 9 asks whether we "
        "foresee an environmental monitoring system and we are building one "
        "as a project output; question 7 asks about awareness-raising and we "
        "have nine custodian groups and on-site interpretation in three "
        "languages; question 10 asks whether the costs are in the budget and "
        "we can answer with work package figures rather than a yes. Most "
        "applicants answer those three with promises.", size=10.5)

    for num, q, ans, lim, n, text in report:
        D.h(f"{'Additional information' if not num else f'Question {num}'}"
            f"{'' if not num else f' — {ans}'}", size=12, space=10)
        if num:
            D.p(q, size=10, italic=True, colour=GREY)
        D.p(f"{n}/{lim} characters:", size=9, colour=GREY, italic=True)
        D.p(text)

    D.h("One thing to watch when you paste question 6", size=12)
    D.p("It has a blank line in the middle, separating the negative-impact "
        "paragraph from the positive one. If the box strips the line break "
        "the two paragraphs will run together and still read correctly - but "
        "check it after saving, because 1000 characters in one block is hard "
        "on a reader and the capital headings are doing the work of the "
        "break.", size=10.5)

    D.d.save(OUT)
    print(f"wrote {Path(OUT).name}")
    for n, q, a, l, c, _ in report:
        label = "extra" if not n else f"Q{n}"
        print(f"   {label:6s} {a or '-':4s} {c:>4}/{l}")
    print(f"\n   WP figures used: WP3 {eur(wp['WP3'])}, WP4 {eur(wp['WP4'])}, "
          f"WP5 {eur(wp['WP5'])}")


if __name__ == "__main__":
    main()
