#!/usr/bin/env python3
"""Indicators from the Methodology Paper, and our standing against the ToR.

    python3 build_indicators_and_tor.py

Both missing documents arrived on 03 October. Between them they unlock the
twelve points I have been flagging since 30 September - criterion 1.1 against
the Terms of Reference, 4 points at Step 1 and 4 more at Step 2, and criterion
2.4 indicators, 4 points.

They also contain five things that change the proposal, four of them gaps in
what we have built. I would rather find them now than have an evaluator find
them on 29 October.

WHAT THE METHODOLOGY PAPER SETTLES

Our specific objective has exactly FOUR output indicators and TWO result
indicators. Not a menu to choose freely from - these four and no others:

    RCO84   Pilot actions developed jointly and implemented in projects   60%
    RCO87   Organisations cooperating across borders                     100%
    RCO81   Participations in joint actions across borders                70%
    RCO116  Jointly developed solutions                                   70%
    RCR84   Organisations cooperating across borders after completion
    RCR104  Solutions taken up or up-scaled by organisations

(The percentages are the programme's own assumption of how many projects under
this objective use each indicator. RCO87 at 100% means every project reports
it.)

THERE IS NO TRAINING INDICATOR UNDER OUR OBJECTIVE. RCO85, participations in
joint training schemes, exists in the programme and is assumed by 80% of
projects - under SO 1.3, SO 4.2 and SO 4.5. It is NOT in our list. So our 60
trained site managers cannot be reported as a training indicator at all; they
are reported as RCO81 participations in joint actions, which the programme's
own definition confirms includes training schemes.

THREE DEFINITIONS THAT CHANGE WHAT WE PROMISE

1.  RCO84 pilot actions: "the pilot action needs not only to be developed, but
    also implemented within the project" and "the implementation of the pilot
    action should be FINALISED BY THE END OF THE PROJECT."

    This is binary. A demonstration installation that is 80% complete at month
    24 counts as ZERO, not as 0.8. I warned yesterday that four of our eight
    outputs land in semester IV and that a slip would cost completeness. That
    was too mild. For the pilot indicator a slip costs the whole unit, and the
    partial-state mitigation I proposed does not help here.

2.  RCO116 jointly developed solutions: "an identified solution should include
    indications of the ACTIONS NEEDED FOR IT TO BE TAKEN UP or to be upscaled."

    That is a design requirement on the output itself, not a reporting note. A
    method document that does not say what an adopting organisation must do is
    not countable. Ours must each carry that section explicitly.

3.  RCR104 solutions taken up: "the number of solutions, OTHER THAN LEGAL OR
    ADMINISTRATIVE SOLUTIONS, that are developed by supported projects and are
    taken up or upscaled... The uptake should be DOCUMENTED BY THE ADOPTING
    ORGANISATIONS in, for instance, strategies, action plans etc."

    This one is a trap and it changes our wording. If we present the three
    formally adopted management plans AS the solution, an evaluator can call
    them administrative and the indicator reads zero. The correct framing, and
    it is also the true one: the SOLUTION is the adapted method; the adopted
    management plan is the DOCUMENTARY EVIDENCE that an organisation took it
    up. That is precisely what the definition asks for.

AND ONE PIECE OF FREE CALIBRATION

The programme computed its own targets on stated assumptions: average 6
partners per project, "in average 10 participants per partner per project
where joint actions are addressed", average project size 1.890.000 € under our
priority. We have 6 partners and a 1.300.000 € budget - about 69% of the
average size. So targets should be proportionate rather than impressive. Ours
are, and now I can say so with the programme's own arithmetic rather than by
assertion.

WHAT THE ToR REQUIRES THAT WE DO NOT YET HAVE

Four cross-cutting elements are mandatory. We are strong on one, partial on
two and MISSING one outright:

  3. GOVERNANCE & PARTICIPATION: "establish a permanent or long-lived steering
     committee (public, private, academia, civil society). Include mechanisms
     for intergenerational dialogue/advisory groups, and youth participation."

     We have a project management committee of one representative per partner.
     That is not what this asks for. It asks for a multi-stakeholder body that
     OUTLIVES the project and includes private sector and civil society. This
     is a mandatory element of the call and we do not have it.

     Youth participation we have in abundance - a youth association as a
     partner and nine school custodian groups. Intergenerational dialogue is
     not named anywhere and should be, because it is sitting there unclaimed
     in a programme that pairs schoolchildren with municipal staff.

  1. CAPITALISATION & TRANSFER PROTOCOL also requires "dedicated budget
     allocations for scaling and sustaining project results, including
     FOLLOW-UP FUNDING and visibility mechanisms". We have the Limnos geopark
     maintenance route, which is a good answer, but not a budget line and not
     a follow-up funding mechanism.

  4. DATA, MONITORING AND OPEN TOOLS asks us to "use/strengthen existing
     SUSTAINABLE TOURISM OBSERVATORIES". We reference none. Our platform
     answers the rest of the element well.

  2. SYNERGIES & EMBEDDING names the frameworks it wants to see, and our
     synergies section misses most of them. This is the cheapest points on the
     table: criterion 1.4 is 4 points and the ToR has told us the answer.

ONE THING WE GET RIGHT THAT WE HAVE NEVER CLAIMED

"Priority will be given to outputs from the 2021-2027 programming period
within MMM Programmes, whenever possible."

AMMIRARE is Interreg Italia-Francia Marittimo 2021-2027. Our anchor output is
in the priority period. COMMON, MEDUSA and CROSSDEV are 2014-2020. So the
choice we made for other reasons is also the preferred one, and nothing in our
text says so.

AND ONE SENTENCE THAT COSTS NOTHING AND PAYS TWICE

The ToR lists "Memoranda of Understanding (MOUs) or cooperation agreements"
among the expected outputs. RCR84 requires, to be counted at all, "a statement
that the entities have a formal agreement to continue cooperation after the end
of the supported project", and allows it to be signed DURING implementation.

The same document satisfies a ToR expected output and makes a result indicator
reportable instead of hopeful. It needs no budget. It should be in WP6.
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

from build_wp_blended import DIRECT, P, STAFF, WPS
from build_outputs import OUTPUTS, TECHNICAL, ROMAN, sem

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_indicators_and_ToR_2026-10-03.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)
AMBER = RGBColor(0x9A, 0x6B, 0x1E)

FLAT = 0.30
PRIORITY_AVG_PROJECT = 1_890_000.00   # Methodology paper, Table 7, PO2
PROG_PARTNERS_AVG = 6                 # Methodology paper, assumption 6
PROG_PARTICIPANTS_PER_PARTNER = 10    # Methodology paper, assumption 7
N_ASSOCIATES_PLANNED = 1              # max 1 advisory partner per the e-form

# The only output indicators available under our specific objective, with the
# programme's own assumption of how many projects use each. Anything not in
# this list cannot be reported, however well it fits the project.
SO24_OUTPUT = {
    "RCO84": ("Pilot actions developed jointly and implemented in projects",
              "nr of pilot actions", 60),
    "RCO87": ("Organisations cooperating across borders",
              "number of organisations", 100),
    "RCO81": ("Participations in joint actions across borders",
              "number of participations", 70),
    "RCO116": ("Jointly developed solutions", "number of solutions", 70),
}
SO24_RESULT = {
    "RCR84": ("Organisations cooperating across borders after project "
              "completion", "number of organisations"),
    "RCR104": ("Solutions taken up or up-scaled by organisations",
               "number of solutions"),
}

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


# ------------------------------------------------- our outputs -> indicators
# (indicator, target, unit, which of our outputs carries it, justification)
MAPPING = [
    ("RCO84", 3, "nr of pilot actions", "Output4.1",
     "The three demonstration installations, one per site, each developed "
     "jointly and implemented. The definition requires implementation to be "
     "FINALISED by project end, so this is three or nothing per site - "
     "Output4.1 delivers by M20, which leaves four months of margin and that "
     "margin is the reason the target is credible."),
    ("RCO87", len(P) + N_ASSOCIATES_PLANNED, "number of organisations",
     "project level, not a single output",
     f"{len(P)} project partners plus {N_ASSOCIATES_PLANNED} associated "
     f"organisation. The definition explicitly counts \"project partners AND "
     f"associated organizations, as mentioned in the financing agreement\", "
     f"so the North Aegean Periphery statement adds a unit here as well as "
     f"four points on criterion 3.2. The programme assumes 100% of projects "
     f"report this indicator and an average of {PROG_PARTNERS_AVG} partners, "
     f"so we are at the average plus the associate."),
    ("RCO81", 60, "number of participations",
     "Output5.1 and Output5.2, plus the WP3 transfer workshop and the final "
     "conference",
     f"The programme's own assumption is {PROG_PARTICIPANTS_PER_PARTNER} "
     f"participants per partner per project, which for {len(P)} partners is "
     f"{PROG_PARTNERS_AVG * PROG_PARTICIPANTS_PER_PARTNER}. We claim exactly "
     f"that and no more. Counted on attendance lists, as the definition "
     f"requires, and a joint action needs organisations from at least two "
     f"participating countries - which the custodian programme, the training "
     f"and the conference all satisfy. THIS is where the 60 trained site "
     f"managers are reported, because there is no training indicator under "
     f"our objective."),
    ("RCO116", 3, "number of solutions",
     "Output3.1, Output4.2, Output6.2",
     "The adapted method with its 21 transferability decisions; the "
     "cross-typology comparability layer and its platform; the joint "
     "transferability guide. Each must carry, by the definition, "
     "\"indications of the actions needed for it to be taken up or to be "
     "upscaled\" - so each of the three documents needs an explicit uptake "
     "section. That is a change to the outputs, not a reporting note."),
]

RESULT_MAPPING = [
    ("RCR104", 3, "number of solutions",
     "The adapted method taken up by the authority responsible for each of "
     "the three sites, evidenced by the formally adopted management plan at "
     "each one.",
     "WORD THIS CAREFULLY. The definition excludes \"legal or administrative "
     "solutions\". If the adopted management plan is presented as the "
     "solution, it can be called administrative and counted as zero. The "
     "solution is the ADAPTED METHOD; the adopted plan is the documentary "
     "evidence of uptake, which is exactly what the definition asks for when "
     "it says the uptake \"should be documented by the adopting organisations "
     "in, for instance, strategies, action plans etc.\" Measurement runs "
     "during implementation or up to one year after completion, so Output6.1 "
     "at M22 sits inside the window."),
    ("RCR84", len(P) + N_ASSOCIATES_PLANNED, "number of organisations",
     "All partners and the associated organisation continuing to cooperate "
     "after the project, under a signed cooperation agreement.",
     "This indicator is only countable with \"a statement that the entities "
     "have a formal agreement to continue cooperation after the end of the "
     "supported project\", and the agreement MAY BE SIGNED DURING "
     "IMPLEMENTATION. The ToR separately lists MOUs and cooperation "
     "agreements among its expected outputs. One document, signed in WP6, "
     "satisfies a ToR expected output and converts this indicator from a hope "
     "into a deliverable. It costs nothing."),
]

# -------------------------------------------- the four ToR cross-cutting elements
#   (element, what the ToR requires, our standing, verdict, what to do)
STRONG, PARTIAL, MISSING = "STRONG", "PARTIAL", "MISSING"

TOR = [
    ("1. Capitalisation & transfer protocol",
     "A clear plan showing which existing results will be transferred, who "
     "will adopt them (identify transferability partners), AND dedicated "
     "budget allocations for scaling and sustaining project results, "
     "including follow-up funding and visibility mechanisms.",
     "The transfer plan is our strongest asset - 21 named decisions across "
     "seven method sheets and three sites, with the adopting authority named "
     "at each site. But there is no budget line labelled for scaling and "
     "sustaining, and no follow-up funding mechanism.",
     PARTIAL,
     "The Limnos geopark route is already a real sustaining mechanism and "
     "should be named as one rather than buried in an activity. Add a named "
     "allocation inside Output6.1 for the adoption and maintenance work, and "
     "one sentence identifying the follow-up funding each site will pursue."),

    ("2. Synergies & embedding",
     "Demonstrate alignment with national and regional tourism strategies, "
     "S3, climate adaptation plans, macro-regional strategies and the "
     "Mediterranean Strategy for Sustainable Development, where relevant. The "
     "ToR then names the frameworks it wants considered.",
     "Our synergies section has the Habitats Directive, the European "
     "Environment Agency litter methodology and LIFE IP 4 NATURA. It is "
     "missing almost everything the ToR actually names.",
     PARTIAL,
     "This is the cheapest scoring work left in the proposal. The ToR has "
     "told us the answer: MSSD, the Union for the Mediterranean Ministerial "
     "Declaration on Environment and Climate Change and the 2030 Greener Med "
     "Agenda, the Pact for the Mediterranean and its Action Plan, the "
     "European Ocean Pact, WestMED, BLUEMED, EUSAIR, the EU Strategy for "
     "Sustainable Tourism, and the smart specialisation strategies of the "
     "three receiving regions. Criterion 1.4 is 4 points at Step 2."),

    ("3. Governance & participation",
     "Establish a PERMANENT OR LONG-LIVED STEERING COMMITTEE spanning public "
     "sector, private sector, academia and civil society. Include mechanisms "
     "for intergenerational dialogue or advisory groups, and youth "
     "participation.",
     "We have a project management committee of one representative and one "
     "alternate per partner. That is a project governance body that ends with "
     "the project, and it contains no private sector and no civil society. "
     "Youth participation, by contrast, is a real strength: a youth "
     "association is a full partner and nine school custodian groups run in "
     "three countries.",
     MISSING,
     "This is a mandatory element of the call and the one place we are "
     "genuinely short. The fix is not cosmetic but it is cheap: constitute a "
     "Shore Stewardship Committee at each site - the responsible authority, "
     "a tourism operator, the research partner and a civil society or youth "
     "representative - established in WP6 alongside the management plan it "
     "will oversee, with its continuation written into the cooperation "
     "agreement. That single structure answers this element, strengthens "
     "RCR84, and gives the management plans an owner after we leave. "
     "Intergenerational dialogue should also be named explicitly: a programme "
     "that puts schoolchildren and municipal site staff on the same shore "
     "with the same protocol already IS that, and we have never said so."),

    ("4. Data, monitoring and open tools",
     "Use or strengthen existing Sustainable Tourism Observatories, shared "
     "digital platforms for visitor and environmental monitoring, and open "
     "datasets to support transfer and policy uptake.",
     "The shared web-GIS platform with one protocol across three shore types "
     "and open licensing is a direct answer to two thirds of this element. We "
     "reference no existing observatory.",
     STRONG,
     "Identify whether any of the three regions sits inside an existing "
     "Sustainable Tourism Observatory, and if one does, connect to it rather "
     "than beside it. If none does, say that plainly and position the "
     "platform as the first observatory-grade dataset for these shore types - "
     "an honest answer scores better than a vague claim of alignment."),
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

    def table(self, head, rows, widths=None, verdict_col=None):
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
                if verdict_col is not None and i == verdict_col:
                    r.bold = True
                    r.font.color.rgb = {STRONG: GREEN, PARTIAL: AMBER,
                                        MISSING: RED}.get(str(v), GREY)
                elif "[" in str(v) and "]" in str(v):
                    r.font.color.rgb = RED
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t

    def page(self):
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def check():
    """Refuse to ship a document that invents an indicator or a target."""
    fails = []
    out_codes = {m[0] for m in MAPPING}
    res_codes = {m[0] for m in RESULT_MAPPING}

    # 1. Every indicator we use must exist in our specific objective's list.
    for c in out_codes:
        if c not in SO24_OUTPUT:
            fails.append(f"{c} is not an output indicator under our specific "
                         f"objective - using it is not possible")
    for c in res_codes:
        if c not in SO24_RESULT:
            fails.append(f"{c} is not a result indicator under our objective")

    # 2. No indicator from another objective may appear anywhere in the text.
    #    RCO85 is the dangerous one: it fits our training perfectly and is not
    #    available to us.
    blob = " ".join(
        [str(x) for m in MAPPING + RESULT_MAPPING for x in m]
        + [str(x) for t in TOR for x in t])
    for forbidden in ("RCO85", "RCO01", "RCO02", "RCO90", "RCO82", "RCR01",
                      "RCR03", "RCR81", "RCR85"):
        if re.search(rf"\b{forbidden}\b", blob):
            fails.append(f"{forbidden} appears in the mapping but is not "
                         f"available under our specific objective")

    # 3. RCO84's target cannot exceed the number of pilot sites, and the
    #    output carrying it must finish before the project does, because a
    #    pilot not finalised by the end counts as zero.
    pilots = [m for m in MAPPING if m[0] == "RCO84"][0]
    o41 = [o for o in OUTPUTS if o[0] == "WP4" and o[1] == 1]
    if not o41:
        fails.append("Output4.1 no longer exists but RCO84 is mapped to it")
    else:
        fin = o41[0][5]
        if fin >= 24:
            fails.append(f"Output4.1 finishes M{fin} and RCO84 requires the "
                         f"pilot to be FINALISED before project end - no "
                         f"margin at all")
        if pilots[1] != 3:
            fails.append(f"RCO84 target is {pilots[1]} but there are three "
                         f"demonstration sites")

    # 4. RCO87 and RCR84 count organisations; they cannot exceed partners plus
    #    associates, and must not silently assume an unsigned associate.
    cap = len(P) + N_ASSOCIATES_PLANNED
    for code, target in [("RCO87", [m for m in MAPPING
                                    if m[0] == "RCO87"][0][1]),
                         ("RCR84", [m for m in RESULT_MAPPING
                                    if m[0] == "RCR84"][0][1])]:
        if target > cap:
            fails.append(f"{code} target {target} exceeds {len(P)} partners "
                         f"plus {N_ASSOCIATES_PLANNED} associate")

    # 5. RCO81 must not exceed the programme's own per-partner assumption, or
    #    we are claiming more reach than the methodology it will be read
    #    against.
    rco81 = [m for m in MAPPING if m[0] == "RCO81"][0][1]
    ceiling = len(P) * PROG_PARTICIPANTS_PER_PARTNER
    if rco81 > ceiling:
        fails.append(f"RCO81 target {rco81} exceeds the programme's own "
                     f"assumption of {ceiling} for {len(P)} partners")

    # 6. RCO116 must not claim more solutions than we have outputs to carry.
    rco116 = [m for m in MAPPING if m[0] == "RCO116"][0]
    named = re.findall(r"Output(\d)\.(\d)", rco116[3])
    if len(named) != rco116[1]:
        fails.append(f"RCO116 target is {rco116[1]} but {len(named)} outputs "
                     f"are named as carrying it")

    # 7. A ToR element marked MISSING must come with something to do about it.
    for name, req, standing, verdict, action in TOR:
        if verdict == MISSING and len(action) < 100:
            fails.append(f"{name} is MISSING and the remedy is one line - "
                         f"that is a complaint, not a plan")

    # 8. The house rules.
    for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson"):
        if re.search(forbidden, blob):
            fails.append(f"the text contains {forbidden!r}")

    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)


def main():
    check()
    tot_elig = sum(DIRECT[k] + STAFF[k] * FLAT for k in DIRECT)

    D = Doc()
    D.h("AMMOS — indicators, and where we stand against the Terms of "
        "Reference", size=16)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} from the Performance Framework "
        f"Methodology Paper and the ToR. Between them these unlock the twelve "
        f"points I have been flagging since 30 September. They also contain "
        f"four gaps in what we have built, and one of them is a mandatory "
        f"element of the call.", size=9.5, colour=GREY, italic=True)

    # ------------------------------------------------------------ indicators
    D.h("Our specific objective has exactly four output indicators", size=13)
    D.p("Not a menu. These four and no others, with the programme's own "
        "assumption of how many projects under this objective report each "
        "one:", size=10)
    D.table(["Code", "Name", "Unit", "Used by"],
            [[c, v[0], v[1], f"{v[2]}% of projects"]
             for c, v in SO24_OUTPUT.items()],
            widths=[1.8, 8.0, 4.2, 3.4])
    D.p("")
    D.p("THERE IS NO TRAINING INDICATOR UNDER OUR OBJECTIVE. RCO85, "
        "participations in joint training schemes, is assumed by 80% of "
        "projects — under three other specific objectives, not ours. So the "
        "60 trained site managers cannot be reported as training. They are "
        "reported as RCO81 participations in joint actions, which the "
        "programme's definition confirms includes training schemes. Worth "
        "knowing before you reach that dropdown and look for the obvious "
        "option that is not there.", size=10, bold=True)

    D.h("How our eight outputs map onto the four", size=13)
    D.table(["Indicator", "Target", "Unit", "Carried by", "Why this number"],
            [[c, str(t), u, w, j] for c, t, u, w, j in MAPPING],
            widths=[1.6, 1.3, 3.0, 3.6, 8.0])
    D.p("")
    D.p("Output3.2, the three site baselines, and Output6.1, the three "
        "adopted management plans, carry no output indicator. That is "
        "correct, not an omission — Output6.1 is where the RESULT indicator "
        "is evidenced, which is a stronger place to sit.",
        size=9.5, colour=GREY)

    D.h("The two result indicators, and one that needs careful wording",
        size=13)
    D.table(["Indicator", "Target", "Unit", "What it measures for us",
             "The catch"],
            [[c, str(t), u, w, j] for c, t, u, w, j in RESULT_MAPPING],
            widths=[1.6, 1.3, 3.2, 5.4, 6.0])

    D.page()
    D.h("Three definitions that change what we promise", size=14)
    for n, (title, body) in enumerate([
        ("RCO84 is binary, and that makes the semester-IV problem worse "
         "than I said yesterday",
         "\"The pilot action needs not only to be developed, but also "
         "implemented within the project, and the implementation of the pilot "
         "action should be FINALISED BY THE END OF THE PROJECT.\" An "
         "installation that is 80% complete at month 24 counts as zero, not "
         "as 0.8. Yesterday I told you a slip would cost completeness and "
         "that a reportable partial state would mitigate it. For this "
         "indicator that is wrong — a partial pilot is no pilot. Output4.1 "
         "delivering at M20 is what makes the target of 3 credible, and that "
         "four-month margin should now be treated as untouchable rather than "
         "as slack."),
        ("RCO116 imposes a requirement on the documents themselves",
         "\"An identified solution should include indications of the actions "
         "needed for it to be taken up or to be upscaled.\" So the adapted "
         "method, the comparability layer and the transferability guide each "
         "need an explicit section saying what an adopting organisation must "
         "do. Without it they are not countable. This is a change to the "
         "outputs, not a note for the reporting phase."),
        ("RCR104 excludes administrative solutions, and our instinct was to "
         "present an administrative one",
         "\"The number of solutions, OTHER THAN LEGAL OR ADMINISTRATIVE "
         "SOLUTIONS, that are developed by supported projects and are taken "
         "up or upscaled.\" Our three formally adopted management plans are "
         "the most tangible thing the project produces, and presenting them "
         "as the solution invites an evaluator to call them administrative "
         "and score zero. The solution is the ADAPTED METHOD. The adopted "
         "plan is the documentary evidence that an organisation took it up — "
         "which is exactly what the definition asks for when it says uptake "
         "\"should be documented by the adopting organisations in, for "
         "instance, strategies, action plans etc.\" Same facts, different "
         "sentence, and the sentence is worth four points."),
    ], 1):
        D.h(f"{n}.  {title}", size=11, space=10)
        D.p(body, size=9.5)

    D.h("Free calibration from the programme's own arithmetic", size=12)
    D.p(f"The methodology paper states the assumptions behind the programme's "
        f"own targets: an average of {PROG_PARTNERS_AVG} partners per "
        f"project, \"in average {PROG_PARTICIPANTS_PER_PARTNER} participants "
        f"per partner per project where joint actions are addressed\", and an "
        f"average project size of {eur(PRIORITY_AVG_PROJECT)} under our "
        f"priority. We have {len(P)} partners and "
        f"{eur(tot_elig)} — about "
        f"{100 * tot_elig / PRIORITY_AVG_PROJECT:.0f}% of the average size. "
        f"That is an argument for proportionate targets rather than "
        f"impressive ones, and it means our RCO81 of 60 is not a guess: it is "
        f"the programme's own figure for a partnership of our size.", size=10)

    D.page()
    D.h("The ToR's four mandatory cross-cutting elements", size=14)
    D.p("Criterion 1.1 is coherence with the Terms of Reference — 4 points at "
        "Step 1 and 4 more from the Assessment Board at Step 2, so 8 of the "
        "100. These four elements are where that is decided.",
        size=9.5, colour=GREY)
    D.table(["Element", "What the ToR requires", "Where we actually stand",
             "Verdict", "What to do"],
            [[n, r, s, v, a] for n, r, s, v, a in TOR],
            widths=[3.0, 4.6, 4.6, 1.8, 4.6], verdict_col=3)

    D.page()
    D.h("The one gap that is mandatory, and the fix that solves three things "
        "at once", size=14)
    D.p("The ToR requires \"a permanent or long-lived steering committee "
        "(public, private, academia, civil society)\" with \"mechanisms for "
        "intergenerational dialogue/advisory groups, and youth "
        "participation\".", size=10, italic=True)
    D.p("We have a project management committee — one representative and one "
        "alternate per partner — which ends when the project ends and "
        "contains no private sector and no civil society. That is not what "
        "this asks for, and it is a mandatory element.", size=10)
    D.p("The fix: a SHORE STEWARDSHIP COMMITTEE at each of the three sites — "
        "the responsible authority, a tourism operator, the research partner, "
        "and a civil society or youth representative. Constituted in WP6 "
        "alongside the management plan it exists to oversee, with its "
        "continuation written into the cooperation agreement.", size=10,
        bold=True)
    D.p("That one structure does four jobs. It answers the mandatory "
        "governance element. It gives the three management plans an owner "
        "after we leave, which is the sustainability criterion. It is itself "
        "one of the ToR's listed expected outputs, \"setting-up of governance "
        "structures to capitalise on strategies\". And a signed continuation "
        "agreement is the precise evidence RCR84 requires, turning a result "
        "indicator from a hope into a deliverable.", size=10)
    D.p("Youth participation needs no work — a youth association is a full "
        "partner and nine school custodian groups run across three "
        "countries. Intergenerational dialogue needs only to be NAMED: a "
        "programme that puts schoolchildren and municipal site staff on the "
        "same shore with the same protocol already is intergenerational "
        "dialogue, and nothing in our text says so.", size=10)

    D.h("One thing we get right and have never claimed", size=12)
    D.p("\"Priority will be given to outputs from the 2021-2027 programming "
        "period within MMM Programmes, whenever possible.\"", size=10,
        italic=True)
    D.p("AMMIRARE is Interreg Italia-Francia Marittimo 2021-2027. Our anchor "
        "output is in the priority period; COMMON, MEDUSA and CROSSDEV are "
        "2014-2020. The choice we made in September for entirely different "
        "reasons is also the preferred one, and our text has never said so. "
        "One sentence in section 3.1.4.", size=10)

    D.h("And the cheapest points left on the table", size=12)
    D.p("The ToR names the policy frameworks it wants to see considered, and "
        "our synergies section misses most of them. This is criterion 1.4, 4 "
        "points, and the answer has been handed to us:", size=10)
    for f in ["Mediterranean Strategy for Sustainable Development (MSSD)",
              "Union for the Mediterranean — Ministerial Declaration on "
              "Environment and Climate Change, and the 2030 Greener Med "
              "Agenda",
              "Union for the Mediterranean — Ministerial Declaration on Blue "
              "Economy",
              "Joint Communication on the Pact for the Mediterranean and its "
              "Action Plan",
              "Communication on the European Ocean Pact",
              "Communication on a Sustainable Blue Economy (DG MARE)",
              "WestMED sea-basin initiative; BLUEMED initiative",
              "EUSAIR and EUSALP macro-regional strategies",
              "EU Strategy for Sustainable Tourism",
              "The smart specialisation strategy (S3) of each of the three "
              "receiving regions"]:
        D.p("•  " + f, size=9.5)
    D.p("Not all of these are genuinely relevant to a shore management "
        "method, and listing frameworks we do not touch is exactly what the "
        "ToR warns against elsewhere: \"the simple listing of projects or "
        "outputs is not sufficient\". I would take MSSD, 2030 Greener Med, "
        "the Pact for the Mediterranean, the Ocean Pact, the EU Strategy for "
        "Sustainable Tourism and the three regional S3s, and say something "
        "specific about each. Six real connections beat ten decorative ones.",
        size=10)

    D.d.save(OUT)
    print(f"wrote {OUT}")
    print(f"\nindicators mapped: {len(MAPPING)} output, "
          f"{len(RESULT_MAPPING)} result, all inside our objective's list")
    for c, t, u, w, _ in MAPPING + [(m[0], m[1], m[2], m[3], None)
                                    for m in RESULT_MAPPING]:
        print(f"  {c:7s} target {t:>3}  {u:26s} <- {w[:46]}")
    print("\nToR cross-cutting elements:")
    for n, r, s, v, a in TOR:
        print(f"  {v:8s} {n}")


if __name__ == "__main__":
    main()
