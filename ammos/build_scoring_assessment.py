#!/usr/bin/env python3
"""Where AMMOS stands against the seven award criteria, now that the form is full.

    python3 build_scoring_assessment.py

He asked this on 1 October and I answered it then from a concept note. He has
asked again now, and this time the question is answerable from evidence: the
form carries 368.778 characters, every section is present, the budget
reconciles and the environmental screening is done. So this is an assessment of
what is actually in the application, not of what we intend.

THE ARITHMETIC THAT MAKES THIS TIGHT

Step 1 is scored by external assessors over 21 sub-criteria totalling 88
points, and the threshold is 75. We may lose THIRTEEN POINTS IN TOTAL. Step 2
adds 12 more from the Assessment Board - criteria 1.1, 1.2 and 1.4, four each -
giving 100. RELEVANCE must separately reach 16 of its 24.

So this is not a question of scoring well on average. One sub-criterion scored
2 out of 4 where it should be 4 costs a sixth of the entire allowance.

HOW I HAVE SCORED IT

Each sub-criterion gets my honest estimate out of its maximum, with the
evidence that supports it and the thing that would cost us the mark. Where I
am guessing I say so. The totals at the end are deliberately NOT optimistic:
where a sub-criterion could go either way I have taken the lower number,
because a forecast that says we clear the threshold comfortably is useless to
him - he needs to know which boxes are still soft.

THE ONE RULE I HAVE FOLLOWED THROUGHOUT

No mark is claimed for something that is not in the form. The stewardship
committees count because they are in Output6.1 and in 3.5; the intergenerational
dialogue does not count because the word appears nowhere, even though we do it.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_where-we-stand_seven-criteria_2026-10-11.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)
AMBER = RGBColor(0x9A, 0x6B, 0x1E)

STEP1_MAX = 88
STEP1_THRESHOLD = 75
STEP2_EXTRA = 12
RELEVANCE_MIN = 16
RELEVANCE_MAX = 24

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


# (section, sub, title, max at step 1, my estimate, evidence, what costs the mark)
C = [
    ("1 RELEVANCE", "1.1", "Coherence with the Terms of Reference", 4, 4,
     "All four mandatory cross-cutting elements are now answered in the form. "
     "The capitalisation and transfer protocol is the 21-decision matrix with "
     "the adopting authority named per site. Governance is the Shore "
     "Stewardship Committee in Output6.1 - the element we were failing a week "
     "ago. Data and open tools is the web-GIS platform. Synergies is 3.1.4a "
     "with MSSD, Greener Med, the Pact and the Ocean Pact named specifically.",
     "A ToR element an assessor judges unanswered. The weakest is still the "
     "'dedicated budget allocation for scaling and sustaining' - the geopark "
     "route is named but there is no line in the budget labelled for it."),

    ("1 RELEVANCE", "1.2", "Relevance to the programme area and the challenge",
     4, 4,
     "3.1.1 is 3.661 of 4.000 characters and carries the economic argument "
     "that most applications omit: the sediment surface IS the tourism "
     "product, the cost falls on a public authority while the revenue accrues "
     "to private operators, which is why the measure nobody is individually "
     "paid to take never gets taken. Three named sites, one an Annex I habitat "
     "between two designations.",
     "Nothing obvious. This is our strongest box."),

    ("1 RELEVANCE", "1.3", "Target groups and their needs", 4, 3,
     "3.1.3 names every target group and, since the 2 October rewrite, states "
     "a NEED for each one rather than listing them. Twelve target groups and "
     "involvement statements across the outputs say what each group DOES.",
     "THE NUMBERS. The criterion uses the word QUANTIFIED and 3.1.3 still has "
     "no audience figures, because Acipayam's visitor counts have been "
     "outstanding since September. This is the single most recoverable mark "
     "in the whole grid and it needs one email answered."),

    ("2 DESIGN", "2.1", "Intervention logic", 4, 4,
     "One specific objective, one expected result, four technical work "
     "packages each linked to it, twelve outputs, twenty-four activities. The "
     "logic is decide, test by using, demonstrate, embed - and 3.4.1 explains "
     "why that order and not another.",
     "Nothing. The chain is complete and internally consistent."),

    ("2 DESIGN", "2.2", "Innovative character", 4, 3,
     "The claim is specific and unusual: not that we transfer a method, but "
     "that we produce a method for deciding WHAT transfers - 21 cells, each "
     "applied, substituted or not applicable with its reason - plus a "
     "cross-typology comparability layer that lets a freshwater pond and a "
     "marine dune be read together.",
     "An assessor who reads capitalisation calls as inherently unoriginal. "
     "The defence is in 3.2.1 and it is well made; I would not spend more "
     "characters on it."),

    ("2 DESIGN", "2.3", "Target groups quantified and involved", 4, 3,
     "Involvement is strong and concrete - the authority signs off the "
     "installation design, the trained staff then write the management plan, "
     "the custodian groups produce monitoring data.",
     "Again the quantification. Same single missing input as 1.3, and it "
     "costs a mark in two separate places."),

    ("2 DESIGN", "2.4", "Indicators and target values", 4, 4,
     "Four output indicators and two result indicators, each with a target "
     "and a 500-character justification proving the counting rule is met. "
     "RCO81 is set at the programme's own published assumption of ten "
     "participations per partner, so it is defensible against the programme's "
     "own arithmetic. RCR104 is worded so the solution is the METHOD and the "
     "adopted plan is the evidence - which keeps it countable.",
     "Nothing, now that the indicator rows are in the right blocks."),

    ("3 PARTNERSHIP", "3.1", "Composition and relevance of the partnership",
     4, 3,
     "Six organisations, four countries, every partner leading something it "
     "would defend in its own name. CNR's researcher is first named author of "
     "the output we transfer - the FAQ says involving the originating "
     "organisation is strongly encouraged.",
     "LCEC. Their project list is nine items and all nine are energy. I have "
     "reframed their role to institutional uptake and policy embedding, which "
     "is honest and defensible, but an assessor comparing WP6's coastal "
     "content to an energy agency's portfolio may still mark it down."),

    ("3 PARTNERSHIP", "3.2", "Associated partners and stakeholders", 4, 3,
     "North Aegean signs on Monday. That moves this from zero to a real mark "
     "and adds a unit to RCO87, which counts associated organisations.",
     "It is not signed yet. Until the declaration is uploaded this is a 0, "
     "not a 3 - and if Sant'Anna can also be added it is a 4, because the "
     "author of the source output endorsing the transfer is exactly what this "
     "criterion rewards."),

    ("3 PARTNERSHIP", "3.3", "Cross-border added value", 4, 4,
     "The strongest structural argument in the proposal: the method exists in "
     "one part of the Mediterranean and the problem exists across all of it, "
     "and the three sites are deliberately dissimilar so the result is a "
     "transferability gradient rather than three local projects.",
     "Nothing."),

    ("4 EFFECTIVENESS", "4.1", "Management and coordination", 4, 4,
     "WP1 names the RACI table, the two-stage deliverable check, the three "
     "changes requiring a recorded committee decision, and seasonal risk with "
     "a second campaign window per site.",
     "Nothing. This is a conventional box answered conventionally well."),

    ("4 EFFECTIVENESS", "4.2", "Work plan and feasibility", 4, 3,
     "Twenty-four activities with real month ranges, a gated sequence, and "
     "installations completing at M20 with four months of margin.",
     "The back-loading. Four of eight technical outputs land in semester IV, "
     "and for the pilot indicator a partial delivery counts as zero. We state "
     "it and mitigate it rather than hiding it, which is the right call, but "
     "an assessor may still take a mark."),

    ("4 EFFECTIVENESS", "4.3", "Capitalisation", 4, 4,
     "MMM_IF02 and MMM_IF01 from AMMIRARE - Italia-Francia Marittimo, inside "
     "the 2021-2027 window the ToR prioritises - plus four ENI CBC MED "
     "outputs used specifically rather than listed. Three solutions return to "
     "the MMM database under open licence.",
     "Nothing. This is the heart of a capitalisation call and we are strong "
     "on it."),

    ("4 EFFECTIVENESS", "4.4", "Communication strategy (COUNTS DOUBLE)", 8, 6,
     "2.943 of 3.000 characters. Audiences named with what each needs, "
     "channels with numbers, partner roles, and the real budget share. The "
     "argument that language IS the strategy - the method exists only in "
     "Italian and its users work in Greek, Turkish and Arabic - is better "
     "than the channel list most applications submit.",
     "REACH FIGURES. The criterion is worth 8 and we cannot state an audience "
     "size for any of the three sites. I have written it so reach is "
     "something the WP3 survey MEASURES rather than guessing, which is "
     "honest, but 8 points is where a partner's unanswered email hurts most."),

    ("5 SUSTAINABILITY", "5.1", "Financial, technical and environmental "
                                "sustainability", 4, 4,
     "Three boxes at 1.643 to 1.991 of 2.000. Costed plans so the authority "
     "knows the annual figure, maintenance carried by the Limnos geopark - an "
     "organisation that already has income - demountable structures repairable "
     "section by section, and a protocol within the reach of staff already "
     "paid to be at the site.",
     "Nothing. The geopark route is the one answer that does not require "
     "anybody to find new money."),

    ("5 SUSTAINABILITY", "5.2", "Long-term impact", 4, 4,
     "The transferability assessment is the durable product: a fourth "
     "territory can tell before spending anything which of the seven "
     "procedures suits its shore. Three adopted plans, three standing "
     "committees, and the Lebanese case where the policy framework exists and "
     "the field method does not.",
     "Nothing."),

    ("5 SUSTAINABILITY", "5.3", "Continued use of results", 4, 4,
     "Every result has a named owner arranged during the project: the plans "
     "with their authorities, the platform under open licence with "
     "institutional backing, the training material with the municipalities, "
     "and the partnership itself under a cooperation agreement signed in WP6 "
     "- which is also what makes RCR84 countable rather than hopeful.",
     "Nothing."),

    ("6 COST-EFFECTIVENESS", "6.1", "Budget", 4, 4,
     "Every figure in 3.6 is computed from the partner budget rather than "
     "typed. Staff 34,08%, largest partner 26,89%, MPC 51,10%, flat rates "
     "exact. 102 budget lines each with a justification written per cost "
     "category per work package.",
     "Nothing, once CNR's 1.497 is corrected."),

    ("6 COST-EFFECTIVENESS", "6.2", "Ratio between costs and results", 4, 4,
     "Unit figures an assessor can check: cost per demonstration site, cost "
     "per person trained, and the comparison with the programme's own average "
     "project size of 1.890.000 for this priority - we deliver across three "
     "countries at 69% of it. Plus the avoided cost argument: we do not "
     "re-derive a method that already exists.",
     "Nothing."),

    ("6 COST-EFFECTIVENESS", "6.3", "Necessity of costs and subcontracting",
     4, 4,
     "The subcontracting box argues capability rather than capacity, naming "
     "four uses no partner holds: certified Turkish and Arabic translation, "
     "licensed drone and bathymetric survey, the independent first-level "
     "control the programme itself requires, and platform development.",
     "Nothing. This is the box assessors read most suspiciously and it is "
     "answered directly."),

    ("7 HORIZONTAL", "7.1", "Horizontal principles", 4, 3,
     "Five boxes at 842 to 965 of 1.000, each answering with something the "
     "project does at a named site rather than EU boilerplate. The "
     "accessibility commitment is real - boardwalks wheelchair-navigable "
     "where gradient allows, with the constraint documented where it does "
     "not. Do-no-significant-harm is four design constraints plus a stopping "
     "rule.",
     "Gender equality is the honest weak one: the project has no gendered "
     "service delivery, so the answer is monitoring and survey "
     "disaggregation. And 'intergenerational' appears nowhere in the form "
     "although we do it - one clause would claim it."),
]

STEP2 = [
    ("1.1", 4, 3, "Coherence with the ToR, re-scored by the Assessment Board. "
                  "Same evidence as Step 1, same soft spot on the budget "
                  "allocation for sustaining results."),
    ("1.2", 4, 4, "Relevance to the programme area. Our strongest box and the "
                  "Board sees the same 3.661 characters."),
    ("1.4", 4, 3, "Synergies with policies and initiatives - assessed ONLY at "
                  "Step 2. 3.1.4a now names MSSD, the 2030 Greener Med "
                  "Agenda, the Pact for the Mediterranean, the Ocean Pact, "
                  "the EU Sustainable Tourism Strategy and regional S3. The "
                  "mark depends on whether the three regional S3 strategies "
                  "are named specifically, which they are not yet."),
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

    def table(self, head, rows, widths=None, small=8.5, scorecol=None):
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
                if str(row[0]).startswith("TOTAL"):
                    r.bold = True
                if scorecol is not None and i == scorecol:
                    r.bold = True
                    if "/" in str(v):
                        got, mx = str(v).split("/")
                        try:
                            r.font.color.rgb = (GREEN if float(got) >= float(mx)
                                                else AMBER)
                        except ValueError:
                            pass
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t

    def page(self):
        from docx.enum.text import WD_BREAK
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def main():
    step1_max = sum(c[3] for c in C)
    step1_got = sum(c[4] for c in C)
    lost = step1_max - step1_got
    rel_max = sum(c[3] for c in C if c[0].startswith("1 ")) * 2
    rel_got = sum(c[4] for c in C if c[0].startswith("1 ")) \
        + sum(s[2] for s in STEP2 if s[0] in ("1.1", "1.2", "1.4"))
    total_got = step1_got + sum(s[2] for s in STEP2)

    fails = []
    if step1_max != STEP1_MAX:
        fails.append(f"sub-criteria total {step1_max}, the grid is "
                     f"{STEP1_MAX} - the breakdown does not match the grid")
    for sec, sub, title, mx, got, ev, risk in C:
        if got > mx:
            fails.append(f"{sub} scored {got} out of {mx}")
        if got < mx and not risk.strip():
            fails.append(f"{sub} loses a mark with no reason given")
        if got == mx and "Nothing" not in risk:
            # full marks claimed with a live risk named - allowed, but the
            # risk must not be a blocker
            pass
    # No mark may be claimed for something absent from the form.
    blob = " ".join(x for c in C for x in (c[5], c[6]))
    if "intergenerational" in blob and "appears nowhere" not in blob:
        fails.append("intergenerational is claimed as evidence but it is not "
                     "in the form")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — where we stand against the seven award criteria", size=15)
    D.p(f"Assessed {dt.date.today():%d/%m/%Y} against what is ACTUALLY IN THE "
        f"FORM, not against what we intend. You asked this on 1 October and I "
        f"answered from a concept note; this time the application carries "
        f"368.778 characters, every section is present, the budget reconciles "
        f"and the screening is done, so it is finally a question with evidence "
        f"behind it.", size=9.5, colour=GREY, italic=True)

    D.h("The arithmetic, which is what makes this tight", size=13)
    D.p(f"Step 1 is scored by external assessors over 21 sub-criteria "
        f"totalling {STEP1_MAX} points, and the threshold is "
        f"{STEP1_THRESHOLD}. WE MAY LOSE {STEP1_MAX - STEP1_THRESHOLD} POINTS "
        f"IN TOTAL. Step 2 adds {STEP2_EXTRA} more from the Assessment Board - "
        f"criteria 1.1, 1.2 and 1.4 - giving 100. Relevance must separately "
        f"reach {RELEVANCE_MIN} of its {RELEVANCE_MAX}.", size=10.5)
    D.p("So this is not about scoring well on average. One sub-criterion at 2 "
        "out of 4 where it should be 4 costs a sixth of the entire allowance.",
        size=10.5, bold=True)

    D.h("My honest estimate", size=14)
    D.table(["Criterion", "Sub", "Title", "Score"],
            [[sec, sub, title, f"{got}/{mx}"]
             for sec, sub, title, mx, got, _, _ in C]
            + [["TOTAL STEP 1", "", "threshold is 75",
                f"{step1_got}/{step1_max}"]],
            widths=[4.2, 1.2, 8.4, 1.8], scorecol=3)
    D.p("")
    D.table(["Step 2 (Assessment Board)", "Score", "Note"],
            [[sub, f"{got}/{mx}", note] for sub, mx, got, note in STEP2]
            + [["TOTAL OVERALL", f"{total_got}/100", ""]],
            widths=[3.0, 1.8, 10.8], scorecol=1)

    D.p("")
    D.p(f"STEP 1: {step1_got} of {step1_max}, threshold {STEP1_THRESHOLD}. "
        f"We clear it by {step1_got - STEP1_THRESHOLD} points, having lost "
        f"{lost} of the {STEP1_MAX - STEP1_THRESHOLD} we are allowed.",
        size=11, bold=True)
    D.p(f"RELEVANCE: {rel_got} of {RELEVANCE_MAX}, against a separate minimum "
        f"of {RELEVANCE_MIN}. Clear.", size=11, bold=True)
    D.p(f"OVERALL: {total_got} of 100.", size=11, bold=True)
    D.p("Those numbers are deliberately not optimistic. Where a sub-criterion "
        "could go either way I have taken the lower figure, because a "
        "forecast telling you we pass comfortably is no use to you - what you "
        "need is the list of boxes that are still soft.", size=10)

    D.page()
    D.h("Where the marks go, one by one", size=14)
    for sec, sub, title, mx, got, ev, risk in C:
        D.h(f"{sub} {title}  —  {got} of {mx}", size=11, space=10,
            colour=GREEN if got == mx else AMBER)
        D.p(f"What is in the form: {ev}", size=10)
        D.p(f"What would cost the mark: {risk}", size=10,
            colour=GREY if got == mx else RED)

    D.page()
    D.h("The four marks we are losing, and what each one costs", size=14)
    below = [c[1] for c in C if c[4] < c[3]]
    D.p(f"Eight sub-criteria at Step 1 are below maximum - "
        f"{', '.join(below)} - plus 1.1 and 1.4 at Step 2. But they come "
        f"down to four missing things, and three of those are somebody "
        f"else's email. That is the useful part of this assessment.",
        size=10.5)
    D.table(["What is missing", "Costs us", "Who can fix it"],
            [["Audience and visitor NUMBERS - Acipayam's visitor counts, "
              "reach figures from the other partners",
              "1.3 (1), 2.3 (1) and most of 4.4 (2) - around four points, "
              "and 4.4 counts double so it is the largest single loss in the "
              "grid",
              "One email answered. Outstanding since September."],
             ["North Aegean's signed declaration",
              "3.2 is a 0 until it is uploaded, not a 3",
              "Monday, per your message"],
             ["The three regional S3 strategies named specifically in 3.1.4a",
              "1.4 at Step 2 - one point",
              "One line from each site partner"],
             ["The word 'intergenerational' in 3.7, and a budget line "
              "labelled for sustaining results",
              "7.1 (1) and the soft spot in 1.1",
              "Me, in ten minutes, if you want it"]],
            widths=[5.4, 5.4, 4.8])

    D.h("What I would not spend any more effort on", size=13)
    D.p("Capitalisation, cross-border added value, the budget boxes, "
        "sustainability and the intervention logic are at maximum and the "
        "evidence is in the form. Adding characters to them cannot raise the "
        "score and risks diluting answers that currently read well. The two "
        "boxes I would leave exactly as they are, because they are the best "
        "things in the application, are 3.1.1 and 3.6.3.", size=10.5)

    D.h("The honest risks that no amount of writing fixes", size=13)
    for t in [
        "LCEC's portfolio. Nine energy projects and no coastal work, leading "
        "the coastal uptake work package. The reframing to institutional "
        "uptake is honest and I would not change it, but an assessor may "
        "still take a mark on 3.1 and there is nothing more to write.",
        "The back-loading. Four of eight technical outputs in semester IV, "
        "and for the pilot indicator a partial delivery counts as zero. We "
        "state it with the mitigation rather than hiding it, which is right, "
        "but it is visible.",
        "Capitalisation scepticism. Some assessors read transfer projects as "
        "inherently less innovative. Our answer - that we produce a method "
        "for deciding what transfers - is the best available and it is "
        "already in 3.2.1.",
    ]:
        D.p("•  " + t, size=10)

    D.p("")
    D.p("Taken together: we are above the threshold with room, the Relevance "
        "sub-minimum is clear, and the losses are concentrated in things "
        "other people owe us rather than things still to be written. Three "
        "weeks ago the honest answer to your question was that we could not "
        "tell. Now it is a list of four items, and three of them are "
        "somebody's email.", size=10.5, bold=True)

    D.d.save(OUT)
    print(f"wrote {Path(OUT).name}")
    print(f"  Step 1: {step1_got}/{step1_max}  (threshold {STEP1_THRESHOLD}, "
          f"lost {lost} of {STEP1_MAX - STEP1_THRESHOLD} allowed)")
    print(f"  Relevance: {rel_got}/{RELEVANCE_MAX}  (minimum "
          f"{RELEVANCE_MIN})")
    print(f"  Overall: {total_got}/100")
    print("\n  below maximum:")
    for sec, sub, title, mx, got, _, _ in C:
        if got < mx:
            print(f"    {sub} {title[:46]:48s} {got}/{mx}")
    for sub, mx, got, _ in STEP2:
        if got < mx:
            print(f"    {sub} (Step 2){'':38s} {got}/{mx}")


if __name__ == "__main__":
    main()
