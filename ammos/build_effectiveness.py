#!/usr/bin/env python3
"""Section 3.4 Effectiveness - all four boxes at their exact limits.

    python3 build_effectiveness.py

His screenshot gave the limits, so this is written to them rather than to a
guess:

    3.4.1 Methodology            3000
    3.4.2 Work plan              2000
    3.4.3 Capitalisation         2000
    3.4.4 Communication strategy 3000

Criterion 4 is 20 points, the largest on the grid. And the field sizes
corroborate something I worked out from the evaluation grid in September:
3.4.4 gets 3000 characters, the same as the methodology, because criterion 4.4
COUNTS DOUBLE - 8 of the 88 points at Step 1, the biggest single item in the
whole assessment. The form is telling us where to spend effort and it agrees
with the grid.

THE TRAP IN 3.4.1, WHICH IS WHY IT IS WRITTEN THE WAY IT IS

We have already filled 3.2.1 "Transfer and adaptation methodology" with 2000
characters about how the transfer works. If 3.4.1 "Methodology" repeats that,
we burn 3000 characters of a 20-point criterion saying what we already said in
a 4-point one, and an evaluator reading them back to back sees padding.

So they are split deliberately:
  3.2.1 is about the OUTPUT - how the method gets transferred and adapted.
  3.4.1 is about the PROJECT - how it is run so that the transfer succeeds:
        the sequence and why it is that sequence, how expertise is matched to
        task, how quality is assured, and what happens when something slips.

NO PLACEHOLDERS IN A SCORED BOX

The communication strategy wants audience sizes we still do not have from
partners. Rather than ship [FROM PARTNER] into a double-weighted field, 3.4.4
is quantified on the figures we DO own - 3 sites, 4 languages, 9 custodian
groups, 60 trained staff, 6 local workshops, 2 press cycles, 1 conference, and
the actual WP2 budget. Audience reach is described as a measured outcome of
the baseline survey rather than guessed at. That is honest and it is also
stronger: a number we can defend beats a bigger number we cannot.
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
OUT = f"{BASE}/AMMOS_3.4_effectiveness_2026-10-06.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

# Measured off his screenshot.
LIMITS = {"3.4.1": 3000, "3.4.2": 2000, "3.4.3": 2000, "3.4.4": 3000}
THIN = 0.80
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


# ------------------------------------------------- 3.4.1 Methodology, 3000
S341 = [
    "AMMOS is run as a decide-test-demonstrate-embed sequence, and the order "
    "is the method rather than a convenience.",

    "DECIDE FIRST, BECAUSE EVERYTHING DOWNSTREAM DEPENDS ON IT. WP3 opens "
    "with the 21-cell transferability assessment: seven method sheets against "
    "three shore types, each cell recorded as applied, substituted or not "
    "applicable with its reason. Nothing is procured, installed or trained "
    "until that matrix exists, because it determines what each site can "
    "monitor. Installing first and assessing afterwards would commit budget "
    "against assumptions.",

    "TEST BY USING, NOT BY REVIEWING. The adapted method is validated by "
    "producing the three territorial baselines with it. If a substituted "
    "procedure does not work in the field, it fails in WP3 at the cost of a "
    "protocol revision, not in WP4 at the cost of an installation. The "
    "baselines are both an output and the method's own test.",

    "DEMONSTRATE REVERSIBLY. WP4 installations are demountable and modular. "
    "That keeps intervention proportionate beside an Annex I habitat, keeps "
    "screening proportionate, and means a design that performs "
    "badly in its first season can be altered rather than defended.",

    "EMBED THROUGH PEOPLE WHO STAY. WP5 trains the municipal and authority "
    "staff who hold the mandate over each site before WP6 asks those same "
    "authorities to adopt a plan. Plans written by people who have used the "
    "method are adopted; plans written for them are filed.",

    "EXPERTISE IS MATCHED TO TASK, NOT TO INSTITUTION. CNR leads the "
    "topographic and forcings sheets, where its researchers authored the "
    "source; HCMR carries the biological and ecological sheets and the "
    "scientific coordination; the University of the Aegean leads demonstration "
    "design and the tourism baseline; Pi Youth Association "
    "leads youth engagement because it is their field rather than a side "
    "activity; Acipayam Municipality leads site-manager training as the "
    "authority that will run a site afterwards; LCEC leads institutional "
    "uptake and policy embedding, which is what a national public agency does "
    "well. Each partner leads something it would defend in its own name.",

    "QUALITY AND RISK. Deliverables pass a two-stage check: the activity lead "
    "signs technical content, the scientific coordinator signs consistency "
    "across sites, because a protocol correct at one site and incompatible "
    "with another defeats the comparability the project exists to create. The "
    "main risks are seasonal: a campaign missed through weather or a delayed "
    "consent cannot be rerun for twelve months, so each site holds a second "
    "window, and the "
    "installations complete at month 20 rather than at the end, leaving four "
    "months of margin before project close.",

    "ADAPTIVE BY DESIGN. Each semester the management committee reviews the "
    "transferability matrix against field experience and may reclassify a "
    "cell. A reclassification is a finding, recorded with its evidence, not a "
    "deviation to be explained away - the boundary of a method's "
    "applicability is precisely what the next territory needs to know.",
]

# ---------------------------------------------------- 3.4.2 Work plan, 2000
S342 = [
    "Six work packages over 24 months: WP1 Management and WP2 Communication "
    "run throughout under the Applicant; four technical work packages, the "
    "maximum the call allows, deliver eight outputs.",

    "WP3 CHARACTERISING AND ADAPTING, M1-M12, HCMR. The transfer workshop and "
    "the 21-cell decision (M1-M8), then three territorial baselines including "
    "the tourism and valuation layer (M5-M12). Gates everything after it.",

    "WP4 DEMONSTRATING, M8-M22, University of the Aegean. Site designs and "
    "consents (M8-M13) overlap the late baselines deliberately, so design "
    "starts from measured conditions. Installation and one full visitor "
    "season (M12-M20). The shared monitoring platform and comparability layer "
    "(M10-M18).",

    "WP5 BUILDING CAPACITY, M6-M22, Pi Youth Association. Custodian material "
    "adapted (M6-M12) and nine school groups running (M10-M18); "
    "site-manager training (M6-M16) and visitor awareness (M12-M20). Starts "
    "before the demonstration so trained staff are present when it happens.",

    "WP6 MAKING IT STICK, M12-M24, LCEC. Costed plans and consultation "
    "(M12-M19), adoption decisions and stewardship committees (M16-M22), the "
    "transferability guide (M18-M23) and the policy recommendations, "
    "cooperation agreement and final conference (M20-M24).",

    "DELIVERY BY SEMESTER: two outputs in semester II, two in III, four in "
    "IV. We state the back-loading rather than disguise it, because it is "
    "unavoidable - a plan cannot be adopted before the demonstration it rests "
    "on - and because the mitigation is real: every semester-IV output has a "
    "defined earlier reportable state (plans consulted before adopted, guide "
    "drafted before lodged), and the pilot installations complete at M20, "
    "four months before close, because the pilot indicator counts only what "
    "is finished.",
]

# ------------------------------------------------- 3.4.3 Capitalisation, 2000
S343 = [
    "AMMOS is built on an identified output and returns more than it takes.",

    "WHAT WE CAPITALISE. MMM_IF02, Guidelines for the Sustainable Management "
    "of Beaches, the first output of AMMIRARE under Interreg Italia-Francia "
    "Marittimo 2021-2027 - an MMM programme other than Interreg NEXT MED and "
    "within the period the Terms of Reference prioritise. MMM_IF01, Beach "
    "Custodians, from the same project, as the engagement instrument. "
    "Alongside them, outputs from COMMON, MEDUSA and CROSSDEV under ENI CBC "
    "MED: the beach litter protocol adopted as published so our data stays "
    "comparable with the European Environment Agency dataset, the "
    "capacity-building design behind our training, the BEach CLEAN decalogue "
    "already available in Arabic, and the Lebanese national contribution our "
    "Lebanese partner extends rather than restarts.",

    "HOW, CONCRETELY. Not by citation. The source method is assessed cell by "
    "cell against each receiving shore type, translated into Greek, Turkish "
    "and Arabic, and tested by producing three baselines with it. The "
    "organisation that authored it is a partner and leads the adaptation of "
    "its physical sheets, so the transfer runs between people rather than "
    "between documents.",

    "WHAT RETURNS TO THE MECHANISM. Three solutions go back into the MMM "
    "common database under open licence: the adapted method with its 21 "
    "recorded decisions, the cross-typology comparability layer, and the "
    "transferability guide in English and Arabic. Each carries an explicit "
    "section stating the actions an adopting organisation must take, so the "
    "next territory inherits a usable instrument rather than a report. The "
    "boundary of applicability - where this method stops working, and why - "
    "is itself the most transferable thing we produce.",
]

# ------------------------------------- 3.4.4 Communication strategy, 3000
S344 = [
    "OBJECTIVE. Communication in AMMOS has one job: to move a method from the "
    "people who wrote it to the people who must apply it, and then to people "
    "we will never meet. Everything below is measured against that.",

    "AUDIENCES AND WHAT EACH NEEDS. (1) Site managers and municipal staff at "
    "the three territories - they need the method in their own language and "
    "in a form they can act on, not a journal article. (2) The authorities "
    "that will adopt the management plans - they need costs, obligations and "
    "evidence. (3) Tourism operators and visitors at each site - they need to "
    "know why a boardwalk and a marked boundary exist, or they walk around "
    "them. (4) Schools and young people - the custodians who maintain the "
    "sites after us. (5) The MMM programmes and other Mediterranean "
    "territories - the replication audience. (6) The scientific community.",

    "LANGUAGE IS THE STRATEGY, NOT A DETAIL. The source method exists only in "
    "Italian; our users work in Greek, Turkish and Arabic. Every instrument "
    "that reaches a site manager, a school or a visitor is produced in the "
    "local language, with English for the programme and scientific audiences. "
    "Four languages, costed in WP2, and the Arabic versions extend material "
    "that already exists rather than starting again.",

    "CHANNELS AND ACTIVITIES, WITH NUMBERS. A project website with a public "
    "area and a password-protected partner area, embedded in the municipal, "
    "prefecture and regional sites of the three territories so it is found "
    "where people already look. Social media in four languages, run by the "
    "youth partner with the custodian groups generating content from their "
    "own monitoring rounds - nine groups across three countries, which is "
    "the only sustainable way to keep a project channel alive. Two "
    "newsletter cycles to local sectoral press in every partner country, at "
    "month 1 on objectives and month 24 on results. On-site interpretation at "
    "all three sites and a visitor booklet. Six local workshops, one "
    "stakeholder and one results workshop per territory. A final "
    "transferability conference. One peer-reviewed paper. The transferability "
    "guide in English and Arabic lodged in the MMM database.",

    "PARTNER ROLES. The Applicant coordinates WP2 and cannot delegate it. Pi "
    "Youth Association runs youth channels and social media; each site "
    "partner owns its local press, workshops and on-site material in its own "
    "language, because national press lists and local credibility cannot be "
    "run centrally; LCEC leads policy-facing communication to national "
    "authorities; HCMR leads scientific dissemination.",

    "BUDGET AND MEASUREMENT. WP2 carries a dedicated budget of "
    "[WP2_BUDGET] of direct cost. Reach is not guessed: the WP3 visitor "
    "survey measures the actual audience at each site in its first season, "
    "and targets for the second are set from it. Progress is reported against "
    "the output indicator for participations in joint actions, which the "
    "custodian groups, the training and the workshops all feed.",
]


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(1.9)
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
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def main():
    m = matrix()
    wp2 = sum(m[k].get("WP2", 0.0) for k in P)
    total_direct = sum(DIRECT.values())
    total_elig = total_direct + sum(STAFF.values()) * FLAT

    sections = {}
    report = []
    fails = []
    for num, parts in (("3.4.1", S341), ("3.4.2", S342), ("3.4.3", S343),
                       ("3.4.4", S344)):
        text = clean("\n\n".join(parts)).replace(
            "[WP2_BUDGET]", f"{eur(wp2)} ({100 * wp2 / total_direct:.1f}% of "
                            f"direct cost)")
        n, lim = len(text), LIMITS[num]
        if n > lim:
            fails.append(f"{num} is {n} chars, limit {lim} - cut {n - lim}")
        sections[num] = text
        report.append((num, n, lim))

    # A scored box must not ship a placeholder.
    for num, text in sections.items():
        for ph in re.findall(r"\[[A-Z][A-Z0-9_ ]{2,}\]", text):
            fails.append(f"{num} still contains the placeholder {ph}")

    # 3.4.1 must not simply restate 3.2.1. Cheap structural check: it has to
    # talk about running the project, not only about transferring the method.
    for word in ("risk", "quality", "margin", "leads"):
        if word not in sections["3.4.1"].lower():
            fails.append(f"3.4.1 never mentions {word!r} - it is the project "
                         f"methodology box, not a second transfer box")

    # 3.4.4 is the double-weighted one; it has to name audiences, channels,
    # a budget and how reach is measured, or it scores like everyone else's.
    for word, why in [("audience", "criterion 4.4 asks for target audiences"),
                      ("social media", "the form names social media outreach"),
                      ("budget", "the form asks for planned budget"),
                      ("language", "our users do not read the source"),
                      ("measur", "a strategy with no measurement is a wish")]:
        if word not in sections["3.4.4"].lower():
            fails.append(f"3.4.4 never says {word!r} - {why}")

    for num, text in sections.items():
        for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson",
                          r"xxxx", r"887", r"ELKE"):
            if re.search(forbidden, text, re.I):
                fails.append(f"{num} contains {forbidden!r}")

    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — section 3.4 Effectiveness, all four boxes", size=16)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} to the limits in your "
        f"screenshot: 3.4.1 3000, 3.4.2 2000, 3.4.3 2000, 3.4.4 3000. Each "
        f"checked against its own counter before this document was produced.",
        size=9.5, colour=GREY, italic=True)

    D.p("Criterion 4 is 20 points, the largest single criterion on the grid. "
        "And notice what the field sizes tell you: 3.4.4 Communication gets "
        "3000 characters, the same as the methodology box. That matches what "
        "I worked out from the evaluation grid in September — criterion 4.4 "
        "COUNTS DOUBLE, 8 of the 88 points at Step 1, the biggest single item "
        "in the assessment. The form is pointing at where the effort belongs "
        "and it agrees with the grid.", size=10.5)

    D.table(["Box", "Characters", "Limit", "Fill"],
            [[n, str(c), str(l), f"{100 * c / l:.0f}%"]
             for n, c, l in report], widths=[3.0, 3.0, 2.4, 2.4])

    D.p("")
    D.p("One deliberate choice worth flagging before you read it. We have "
        "already filled 3.2.1 Transfer and adaptation methodology. If 3.4.1 "
        "repeats it we spend 3000 characters of a 20-point criterion saying "
        "what we already said in a 4-point one, and an evaluator reading them "
        "back to back sees padding. So 3.2.1 is about the OUTPUT — how the "
        "method gets transferred — and 3.4.1 is about the PROJECT: the "
        "sequence and why it is that sequence, how expertise is matched to "
        "task, quality assurance, and what happens when something slips.",
        size=10.5)

    for num, title in (("3.4.1", "Methodology"), ("3.4.2", "Work plan"),
                       ("3.4.3", "Capitalisation"),
                       ("3.4.4", "Communication strategy")):
        n = len(sections[num])
        D.h(f"{num} {title}", size=14)
        D.p(f"{n} of {LIMITS[num]} characters.", size=9, colour=GREY,
            italic=True)
        D.p(sections[num])

    D.h("What I did NOT put in 3.4.4, and why", size=13)
    D.p("The communication strategy wants audience sizes, and we still do not "
        "have visitor numbers from Acipayam or reach figures from the other "
        "partners. I have not written [FROM PARTNER] into a double-weighted "
        "field. Instead it is quantified on the figures we actually own — "
        "three sites, four languages, nine custodian groups, sixty trained "
        "staff, six local workshops, two press cycles, one conference, and "
        f"the real WP2 budget of {eur(wp2)} — and audience reach is described "
        f"as something the WP3 visitor survey MEASURES in the first season, "
        f"with second-season targets set from it. That is honest, and a "
        f"number we can defend beats a bigger one we cannot.", size=10.5)
    D.p("If the partner numbers arrive before the 29th, send them and I will "
        "fold them in — there is room in the box.", size=10.5)

    D.h("Still waiting on three tabs", size=13)
    D.p("Sustainability (12 points), Cost-effectiveness (12) and Horizontal "
        "principles (4) are still at zero and I have not seen their fields. "
        "One screenshot each, as you did for this one, and I will write all "
        "three the same way — to the real limits, first time. This "
        "Effectiveness section took one pass because you sent the counters.",
        size=10.5, bold=True)

    D.d.save(OUT)
    print(f"wrote {OUT}")
    for n, c, l in report:
        flag = "" if c >= l * THIN else "   <- thin"
        print(f"  {n}  {c:>5} / {l}   {100 * c / l:3.0f}%{flag}")
    print(f"\nWP2 budget {eur(wp2)} = {100 * wp2 / total_direct:.1f}% of "
          f"direct, {100 * wp2 / total_elig:.1f}% of total eligible")


if __name__ == "__main__":
    main()
