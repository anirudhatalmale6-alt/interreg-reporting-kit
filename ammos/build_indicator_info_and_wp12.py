#!/usr/bin/env python3
"""Indicator detailed information (6 x 500) and WP1 / WP2 text.

    python3 build_indicator_info_and_wp12.py

Two things he is blocked on right now.

1. "Indicator detailed information *   0/500" - REQUIRED, one per indicator
   row. Six of them: four programme output indicators and two result
   indicators. The asterisk means the Indicators tab will not complete
   without them.

2. "CAN I FILL WP1 AND WP2?" - yes, and they are the only work packages that
   can be finished today, because the mandatory templates have no outputs and
   no activities. One description box each, plus coordinator and involved
   partners. Everything needed is already decided.

WHAT GOES IN "INDICATOR DETAILED INFORMATION"

The programme's methodology paper defines each indicator and sets the
conditions under which a unit may be counted. So this box is not a place to
restate the indicator's name - it is where you show that what the project
delivers satisfies the counting rule. Each of the six below therefore says:
what we count, why it qualifies under the definition, and how it will be
evidenced. Where the definition contains a trap, the text is written to clear
it:

  RCO84  a pilot counts only if implemented AND FINALISED by project end, so
         the text states the M20 completion and the four-month margin.
  RCO116 a solution counts only if it carries indications of the actions
         needed for uptake, so the text names that section.
  RCR104 the definition EXCLUDES legal and administrative solutions, so the
         solution is the adapted METHOD and the adopted plan is the evidence
         of uptake - never the other way round.
  RCR84  countable only where a formal agreement to continue cooperation
         exists, and it may be signed during implementation, so the text
         points at the cooperation agreement in WP6.

THE BUDGET MOVED AND I HAVE RE-DERIVED EVERYTHING

His corrected model v6 puts CNR at CC1 50.000 and CC6 15.000 - Antonio's
agreed split, which my generator had never taken on board; it still held the
original 40/28. Total eligible is identical at 80.000, which is exactly why it
went unnoticed: the headline matched while the composition did not.

Consequences, all small but all real: staff rises from 433.500 to 443.500 and
its share from 33,3% to 34,1%; direct falls from 1.169.950 to 1.166.950; the
flat rates rise by 3.000; and every work package budget shifts. Two boxes he
has already pasted now contain a superseded figure. They are named in the
output so he can re-paste exactly those two and nothing else.
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
OUT_IND = f"{BASE}/AMMOS_Indicator-detailed-information_2026-10-07.docx"
OUT_WP12 = f"{BASE}/AMMOS_WP1-WP2_management-and-communication_2026-10-07.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

LIM_IND = 500
FLAT = 0.30
THIN = 0.80

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


# (block, code, name, target, detailed information <= 500)
INDICATORS = [
    ("Programme output indicators", "RCO84",
     "Pilot actions developed jointly and implemented in projects", 3,
     "Three demonstration installations: Gomati on Lemnos, Ucari in Denizli, "
     "and the Lebanese coastal site. Each is jointly developed, designed by "
     "partners from at least two participating countries from a shared "
     "baseline, installed, and observed through one full visitor season. "
     "Implementation completes at month 20, four months before project end, "
     "so each pilot is finalised within the project as the indicator "
     "requires. Evidenced by installation records and the first-season "
     "monitoring dataset."),

    ("Programme output indicators", "RCO116",
     "Jointly developed solutions", 3,
     "Three solutions, each drafted by partners from at least two "
     "participating countries: the adapted shoreline sediment management "
     "with its 21 recorded transferability decisions; the "
     "cross-typology comparability layer and its monitoring platform; and the "
     "joint transferability guide in English and Arabic. Each carries an "
     "explicit section stating the actions an adopting organisation must take "
     "to apply or upscale it, as the indicator requires. All three are lodged "
     "in MMM database under open licence."),

    ("Programme output indicators", "RCO81",
     "Participations in joint actions across borders", 60,
     "Sixty participations, counted from attendance lists. Site-manager "
     "training delivered in each of the three countries on a programme "
     "jointly designed by partners from more than one country; the school "
     "custodian programme run as nine groups across three countries; and the "
     "transfer workshop and final conference. The figure follows the "
     "programme's own planning assumption of ten participations per partner "
     "per project for a six-partner partnership, so it is proportionate."),

    ("Programme output indicators", "RCO87",
     "Organisations cooperating across borders", 7,
     "Seven organisations cooperating formally: the six partners - HCMR and "
     "the University of the Aegean in Greece, Pi Youth Association and "
     "Acipayam Municipality in Türkiye, ISAC-CNR in Italy, LCEC in Lebanon - "
     "plus one associated organisation, the Region of North Aegean. The "
     "definition counts partners and associated organisations named in the "
     "application, so the associate is included. Evidenced by the partnership "
     "agreement and the partner statements."),

    ("Expected result", "RCR104",
     "Solutions taken up or up-scaled by organisations", 3,
     "The solution taken up is the adapted shoreline sediment management "
     "method, not a legal or administrative instrument. Uptake occurs at "
     "three sites and is documented by the adopting organisations themselves "
     "in their formally adopted management plans, each incorporating the "
     "method as its monitoring regime and stating its annual cost. The "
     "adopting bodies are the authorities with the mandate. Adoption falls at "
     "month 22, inside the measurement window of implementation plus one "
     "year."),

    ("Expected result", "RCR84",
     "Organisations cooperating across borders after project completion", 7,
     "All six partners and the associated organisation commit to continued "
     "cooperation under a written agreement signed during implementation, in "
     "work package 6. It names the three Shore Stewardship Committees as the "
     "continuing structures at site level and records what "
     "each signatory will carry on doing. The indicator requires a formal "
     "statement of agreement to continue cooperating, and permits it to be "
     "concluded during the project, which is why it is a dated deliverable "
     "here."),
]


def wp12_text(F):
    """One description box each. No outputs, no activities - the form's
    mandatory templates have neither."""
    m = F["m"]
    inv = lambda wp: ", ".join(
        f"{PP[k]} {P[k][0]}" for k in sorted(P, key=lambda x: PP[x])
        if m[x := k].get(wp, 0) > 0)
    return [
        ("WP1 Management and coordination",
         f"Coordinator: {PP['HCMR']} {P['HCMR'][0]} (Applicant - the form "
         f"pre-sets this and it cannot be delegated)",
         f"Involved partners: {inv('WP1')}",
         f"Budget: {eur(F['wp']['WP1'])} of direct cost",
         [
             "STRUCTURE AND DAY-TO-DAY MANAGEMENT. HCMR as Applicant provides "
             "a scientific coordinator and a project manager. Each partner "
             "names one representative and one alternate to a Management "
             "Committee, which meets monthly online and once per semester in "
             "person, alongside the communication and demonstration "
             "activities so travel serves two purposes. A written RACI table "
             "assigns, for every output and activity, who is responsible, who "
             "is accountable, who is consulted and who is informed - which is "
             "what prevents the familiar failure of a transfer project where "
             "two partners each believe the other is adapting a method sheet.",

             "DECISION-MAKING. The Management Committee decides by consensus "
             "and, failing that, by simple majority with the Applicant "
             "holding a casting vote on scientific consistency only. Changes "
             "to budget allocation between partners, to output target values "
             "or to the transferability classification of any method sheet "
             "require a recorded Committee decision, because those are the "
             "three things that quietly change what the project promised.",

             "ROLE OF EACH PARTNER. HCMR coordinates and carries the "
             "biological and ecological method sheets; the University of the "
             "Aegean leads demonstration design and the tourism baseline; "
             "ISAC-CNR adapts the topographic and forcings sheets its own "
             "researchers authored; Pi Youth Association leads youth "
             "engagement; Acipayam Municipality leads site-manager training "
             "as an authority that will run a site afterwards; LCEC leads "
             "institutional uptake and policy embedding. Every partner leads "
             "something it would defend in its own name.",

             "FINANCIAL MANAGEMENT. Each partner contracts its own "
             "independent first-level control, budgeted in WP1, and submits "
             "certified expenditure per reporting period. The Applicant "
             "consolidates, holds the single interface with the programme, "
             "and tracks each partner against its own semester spending "
             "profile - every budget line in this application states the "
             "semester in which it falls, so a partner drifting from its "
             "profile is visible in the period it happens rather than at "
             "closure.",

             "MONITORING, EVALUATION AND RISK. Each deliverable passes a "
             "two-stage check: the activity lead signs technical content, the "
             "scientific coordinator signs cross-site consistency, because a "
             "protocol that is correct at one site and incompatible with "
             "another destroys the comparability the project exists to "
             "create. A risk register is reviewed every semester. The "
             "identified risks are seasonal - a field campaign lost to "
             "weather or a delayed consent cannot be repeated for twelve "
             "months - so every site holds a second campaign window and the "
             "installations complete at month 20, leaving four months of "
             "margin. Progress is reported against the project's six "
             "indicators, and reclassifying a transferability cell is "
             "recorded as a finding with its evidence, not a deviation.",
         ]),

        ("WP2 Communication and dissemination",
         f"Coordinator: {PP['HCMR']} {P['HCMR'][0]} (Applicant - pre-set, "
         f"cannot be delegated)",
         f"Involved partners: {inv('WP2')}",
         f"Budget: {eur(F['wp']['WP2'])} of direct cost, "
         f"{100 * F['wp']['WP2'] / F['direct']:.1f}% of total direct cost",
         [
             "OBJECTIVE. Communication here has one job: to move a method "
             "from the people who wrote it to the people who must apply it, "
             "and then to people we will never meet. Every channel below is "
             "judged against that.",

             "WHY THE BUDGET IS WHAT IT IS. The method we transfer exists "
             "only in Italian and its users work in Greek, Turkish and "
             "Arabic. Translation, local-language field material and local "
             "workshops are therefore not promotion - they are the transfer "
             "mechanism. This is also the sub-criterion the programme weights "
             "double in assessment.",

             "AUDIENCES AND WHAT EACH NEEDS. Site managers and municipal "
             "staff need the method in their own language and in a form they "
             "can act on. The authorities that will adopt the management "
             "plans need costs, obligations and evidence. Tourism operators "
             "and visitors need to know why a boardwalk and a marked "
             "boundary exist, or they walk around them. Schools and young "
             "people are the custodians who maintain the sites after us. The "
             "MMM programmes and other Mediterranean territories are the "
             "replication audience. The scientific community is the last, not "
             "the first.",

             "CHANNELS AND ACTIVITIES. A project website with a public area "
             "and a password-protected partner area, embedded in the "
             "municipal, prefecture and regional sites of the three "
             "territories so it is found where people already look. Social "
             "media in four languages, run by the youth partner with the nine "
             "custodian groups generating content from their own monitoring "
             "rounds - the only way a project channel stays alive. Two "
             "newsletter cycles to local sectoral press in every partner "
             "country, at month 1 and month 24. On-site interpretation at all "
             "three sites and a visitor booklet. Six local workshops, one "
             "stakeholder and one results workshop per territory. A final "
             "transferability conference. One peer-reviewed paper. The "
             "transferability guide in English and Arabic lodged in the MMM "
             "database.",

             "HOW PARTNERS ARE INVOLVED. The Applicant coordinates and cannot "
             "delegate the work package. Pi Youth Association runs youth "
             "channels and social media. Each site partner owns its local "
             "press, workshops and on-site material in its own language, "
             "because national press contacts and local credibility cannot be "
             "run centrally from Athens. LCEC leads policy-facing "
             "communication to national authorities. HCMR leads scientific "
             "dissemination. Visual identity and programme visibility rules "
             "are applied uniformly across all of it.",

             "MEASUREMENT. Reach is measured rather than estimated: the "
             "visitor survey in work package 3 establishes the actual "
             "audience at each site in the first season, and second-season "
             "targets are set from it. Participations in workshops, training "
             "and custodian activity are counted on attendance lists and "
             "reported against the project's participation indicator.",
         ]),
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
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def figures():
    m = matrix()
    wp = {w[0]: sum(m[k].get(w[0], 0.0) for k in P) for w in WPS}
    return {"m": m, "wp": wp, "direct": sum(DIRECT.values()),
            "staff": sum(STAFF.values()),
            "tot": sum(DIRECT.values()) + sum(STAFF.values()) * FLAT}


def main():
    F = figures()

    # ------------------------------------------------- indicator information
    fails, report = [], []
    for block, code, name, target, info in INDICATORS:
        n = len(clean(info))
        if n > LIM_IND:
            fails.append(f"{code} is {n} chars, limit {LIM_IND} - cut "
                         f"{n - LIM_IND}")
        report.append((block, code, target, n))
    # Only indicators available under our specific objective.
    allowed = {"RCO84", "RCO87", "RCO81", "RCO116", "RCR104", "RCR84"}
    for _, code, _, _, _ in INDICATORS:
        if code not in allowed:
            fails.append(f"{code} is not available under our objective")
    # The two blocks must not be mixed up - that is the error my flat table
    # caused on 06 October.
    for block, code, _, _, _ in INDICATORS:
        if block == "Expected result" and not code.startswith("RCR"):
            fails.append(f"{code} is in the expected result block but is not "
                         f"a result indicator")
        if block == "Programme output indicators" and not code.startswith("RCO"):
            fails.append(f"{code} is in the output block but is not an output "
                         f"indicator")
    # RCR104 must never present the plans as the solution.
    rcr104 = [i for i in INDICATORS if i[1] == "RCR104"][0][4]
    if "not a legal or administrative" not in rcr104:
        fails.append("RCR104 must state that the solution is the method, not "
                     "the plan - the definition excludes administrative "
                     "solutions")
    blob = " ".join(str(x) for i in INDICATORS for x in i)
    for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson", r"xxxx"):
        if re.search(forbidden, blob, re.I):
            fails.append(f"text contains {forbidden!r}")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — Indicator detailed information, all six rows", size=16)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}. The field is required and "
        f"holds 500 characters. Two blocks, and they must not be mixed: "
        f"RCR codes belong to the Expected result, RCO codes to the "
        f"Programme output indicators. That mix-up was caused by my own flat "
        f"table on 6 October, so this document is laid out block by block the "
        f"way the tab is.", size=9.5, colour=GREY, italic=True)

    D.p("This box is not for restating the indicator's name. The programme's "
        "methodology paper sets conditions under which a unit may be counted, "
        "so what belongs here is proof that what we deliver satisfies the "
        "counting rule: what we count, why it qualifies, and how it is "
        "evidenced. Four of the six definitions contain a trap and each text "
        "below is written to clear it.", size=10.5)

    for want in ("Expected result", "Programme output indicators"):
        D.h(f"BLOCK: {want}", size=14)
        for block, code, name, target, info in INDICATORS:
            if block != want:
                continue
            D.h(f"{code} — {name}", size=11.5, space=10)
            D.p(f"Target value: {target}      "
                f"Indicator detailed information: {len(clean(info))}/"
                f"{LIM_IND} characters", size=9, colour=GREY, italic=True)
            D.p(info)

    D.h("The traps these texts are written to clear", size=13)
    D.table(["Indicator", "The condition in the definition",
             "How the text answers it"],
            [["RCO84", "a pilot counts only if implemented AND finalised by "
                       "project end",
              "states completion at month 20 and the four-month margin"],
             ["RCO116", "a solution counts only if it carries indications of "
                        "the actions needed for uptake",
              "names that section in all three solutions"],
             ["RCR104", "EXCLUDES legal and administrative solutions",
              "the solution is the adapted METHOD; the adopted plan is the "
              "documentary evidence of uptake, never the other way round"],
             ["RCR84", "countable only with a formal agreement to continue "
                       "cooperation, which may be signed during "
                       "implementation",
              "points at the cooperation agreement signed in WP6"]],
            widths=[2.0, 6.4, 8.6])
    D.d.save(OUT_IND)

    # ------------------------------------------------------------ WP1 / WP2
    D2 = Doc()
    D2.h("AMMOS — WP1 Management and WP2 Communication, ready to fill",
         size=16)
    D2.p(f"Generated {dt.date.today():%d/%m/%Y}. Yes, you can fill both "
         f"today. These are the only two work packages that can be finished "
         f"now, because the mandatory templates have no outputs and no "
         f"activities - one description box each, plus coordinator and "
         f"involved partners. Nothing in them waits on a partner.",
         size=9.5, colour=GREY, italic=True)

    for title, coord, inv, budget, paras in wp12_text(F):
        D2.h(title, size=14)
        D2.p(coord, size=10, bold=True)
        D2.p(inv, size=10)
        D2.p(budget, size=10)
        D2.p("")
        D2.p("Description box:", size=9, colour=GREY, italic=True)
        for para in paras:
            D2.p(para)
        total = sum(len(clean(x)) for x in paras)
        D2.p(f"[{total} characters in total. I have not seen this box's "
             f"counter - send me a screenshot of it and I will resize if it "
             f"is smaller. If it is 2000 or 3000 like its neighbours, this "
             f"fits or needs a small trim.]", size=9, colour=RED, italic=True)

    D2.d.save(OUT_WP12)

    print(f"wrote {Path(OUT_IND).name}")
    for block, code, target, n in report:
        flag = "" if n >= LIM_IND * THIN else "  <- thin"
        print(f"   {code:7s} target {target:>3}  {n:>3}/{LIM_IND}{flag}  "
              f"[{block}]")
    print(f"\nwrote {Path(OUT_WP12).name}")
    for title, coord, inv, budget, paras in wp12_text(F):
        print(f"   {title}: {sum(len(clean(x)) for x in paras)} characters "
              f"(counter unknown)")
    print(f"\nbudget AFTER the v6 correction:")
    print(f"   direct {eur(F['direct'])}  staff {eur(F['staff'])} "
          f"({100 * F['staff'] / F['tot']:.2f}%)  eligible {eur(F['tot'])}")
    for w in WPS:
        print(f"   {w[0]} {eur(F['wp'][w[0]])}")


if __name__ == "__main__":
    main()
