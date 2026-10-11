#!/usr/bin/env python3
"""Checklist B section 3: every one of the nineteen toggles says YES.

    python3 build_checklist_b3_toggle_fix.py

HE ASKED whether the Environment section was fixed, and said he thought only
the maps, the associated partner and co-financing were still missing.

The Environment section is filled. That is the good news and it answers his
question. But read the nineteen answers of Checklist B section 3 as one
sentence, which is how an assessor reads a checkbox grid:

    the project has out-of-scale features, its visual effect extends over a
    large area, it risks breaching environmental standards, its effects are
    cumulative, permanent, continuous, frequent, IRREVERSIBLE, and difficult
    to avoid, reduce, repair or compensate.

Next to an Annex I habitat type. For a project whose entire environmental
argument is reversibility, that is the most damaging thing in the application,
and it is worse than the three items he listed because those are known gaps
while this one looks complete.

NINE OF THE NINETEEN CONTRADICT THEIR OWN EXPLANATION TEXT. Question 15 reads
"Yes" beside the words "Temporary by design". Question 18 reads "Yes" beside
the word "No." Question 2 reads "Yes" beside "No. The built elements are
footpath-width boardwalk runs". The toggle and the paragraph underneath it
disagree in plain sight.

WHY THIS HAPPENED, AND IT IS MY HANDOFF

I gave him nineteen blocks of text and no toggles. Not one of my blocks said
whether the answer was yes or no - I wrote each paragraph to argue the position
and assumed the position was obvious from the prose. It is obvious to someone
who has read the paragraph; it is not obvious to someone pasting nineteen boxes
in a row, and the default on each control is apparently Yes.

This is the second time on this application. In October I found Checklist A's
toggles inverted in a way that read as "we build infrastructure and considered
none of it". I fixed the toggles that time and did not change how I hand over
checklists, so the same defect arrived again in a different section.

So this file leads with the TOGGLE COLUMN, and the text is secondary. The
generator refuses to emit a question that has text but no explicit toggle.

AND TWO ANSWERS WERE PASTED TWICE

Question 10 in the live form carries question 9's text, word for word.
Question 18 carries question 17's text, word for word. The two paragraphs I
wrote for 10 and 17 are not in the application at all. Both are included below.
This is what a nineteen-box copy-paste looks like when it goes wrong, and it is
another argument for handing over a table rather than a sequence of prose
blocks.

WHAT I AM NOT CHANGING

Section 2 of Checklist B - the location grid - is correct and stays. Densely
populated No, cultural heritage No, natural protected area No, wetlands No,
coastal Yes, marine Yes, estuaries No, mountain N/A, BUFFER ZONE YES, special
biodiversity area No, vulnerable landscape Yes. That grid is honest and
internally consistent with section 3.1.1, and the buffer-zone Yes is the one
that would have contradicted our own narrative had it been No.

Checklist A is also correct as it now stands: infrastructure Yes, environmental
protection Yes, context taken into account Yes, potential impact Yes,
monitoring Yes, costs in the budget Yes. Those are Yes because they should be.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_checklist_b_2_3 import S3 as SECTION3

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Checklist-B-section3_TOGGLES_2026-10-11.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)

LIM = 500

# What the live form says today, from his export of 11-10. All nineteen Yes.
LIVE = {n: "Yes" for n in range(1, 20)}

# The toggle each answer actually argues for, with the reason stated so that
# nobody - him, me, or a partner - has to re-derive it from the paragraph.
#   n: (toggle, why)
TOGGLE = {
    1:  ("Yes", "The intervention does change environmental conditions, and "
                "the change is the improvement the project exists to produce. "
                "Denying it would deny the project."),
    2:  ("No",  "Footpath-width boardwalk, marker poles and signage. Nothing "
                "is engineered coastal defence and nothing rises above eye "
                "level. The answer text already begins with the word No."),
    3:  ("Yes", "There is a visual effect and it is honest to say so. The "
                "text bounds it: modest, local, and removable."),
    4:  ("No",  "The effect is confined to the immediate shore surface at "
                "each site. The text already begins with No."),
    5:  ("Yes", "This one stays Yes deliberately. There IS a transnational "
                "effect - a monitoring protocol and transferability "
                "assessment usable basin-wide - and it is positive. The text "
                "says the physical effect is three footpath corridors and the "
                "methodological effect is basin-wide, which reads correctly "
                "under Yes and contradicts itself under No."),
    6:  ("Yes", "A quantity question forced into a yes/no control. The answer "
                "quantifies, so Yes is the only coherent toggle."),
    7:  ("Yes", "Same - a quantity question. The text enumerates flora, "
                "fauna, businesses and infrastructure."),
    8:  ("Yes", "Valuable features ARE affected, positively, and that is the "
                "point of the project. Declaring it is stronger than denying "
                "something an assessor can see from the site description."),
    9:  ("No",  "No emissions, no discharge, no hazardous materials, no "
                "standard approached. The text says so in its first sentence."),
    10: ("Yes", "Keep Yes. Gomati adjoins Natura 2000 GR4110006, and section "
                "2 of this same checklist already declares the buffer-zone "
                "situation as Yes. A No here would contradict our own grid "
                "two rows up, and an internal contradiction costs more than "
                "an uncomfortable Yes with a control attached."),
    11: ("Yes", "A likelihood question. The answer gives the likelihood."),
    12: ("Yes", "The positive effect is intended to continue indefinitely, "
                "which is what sustainability means here."),
    13: ("No",  "Three independent sites, far apart, no other development "
                "planned. No cumulative NEGATIVE effect. The text says so."),
    14: ("Yes", "Direct effects exist - the footprint and its disturbance - "
                "so Yes is correct. The text then explains that the larger "
                "effect is indirect and positive."),
    15: ("No",  "The answer text is the words 'Temporary by design'. Yes "
                "here says the opposite of every other environmental "
                "statement in the application."),
    16: ("No",  "The answer text is the word 'Intermittent'. Installation is "
                "short and defined; pressure is seasonal."),
    17: ("No",  "The project's own interventions are rare events - one "
                "installation period per site, then observation."),
    18: ("No",  "This is the single most important toggle in the checklist. "
                "Reversibility is argued in the work package, in 3.5.1, in "
                "the do-no-significant-harm answer and in the decision to "
                "budget zero infrastructure. A Yes here contradicts all four "
                "and reads as an admission beside a protected habitat."),
    19: ("No",  "Modularity means a section can be repaired or realigned and "
                "the whole installation removed. The text begins with No."),
}

# The two paragraphs that are not in the application, because another answer's
# text was pasted over them. Taken from the shipped generator, not retyped.
LOST = (10, 17)

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


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

    def table(self, head, rows, widths=None, small=8.5, bold_col=None):
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
                if bold_col is not None and i == bold_col:
                    r.bold = True
                    r.font.color.rgb = (RED if clean(v).strip() == "No"
                                        else GREEN)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def main():
    # Entry 0 is the free-text "Other comments/information" box - no toggle,
    # and it is already correct in the form.
    byn = {n: (q, a) for n, q, a in SECTION3 if n}
    comments = next(a for n, _, a in SECTION3 if not n)

    fails = []
    if set(TOGGLE) != set(byn):
        fails.append("a question has text but no explicit toggle, which is "
                     "exactly the defect this file exists to fix: "
                     f"{sorted(set(byn) ^ set(TOGGLE))}")
    for n, (tog, why) in TOGGLE.items():
        if tog not in ("Yes", "No", "N/A"):
            fails.append(f"Q{n} toggle {tog!r} is not a form value")
        if not why.strip():
            fails.append(f"Q{n} has no reason - a toggle without a reason "
                         f"gets re-decided by the next person in a hurry")
        if len(clean(byn[n][1])) > LIM:
            fails.append(f"Q{n} text {len(clean(byn[n][1]))} > {LIM}")
    # The claim the document leads with has to be true of his export.
    changes = [n for n in TOGGLE if TOGGLE[n][0] != LIVE[n]]
    if not changes:
        fails.append("no toggle differs from the live form - if he has "
                     "already fixed them, this document is wrong")
    # Q18 and Q15 are the two that contradict the whole environmental case.
    for n in (15, 18):
        if TOGGLE[n][0] != "No":
            fails.append(f"Q{n} must be No - reversibility is argued in four "
                         f"other sections and they all depend on it")
    # Internal consistency with section 2's buffer-zone YES.
    if TOGGLE[10][0] != "Yes":
        fails.append("Q10 must stay Yes to match the buffer-zone Yes in "
                     "section 2 of the same checklist")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS - Checklist B, section 3: the nineteen toggles", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} from your export of 11-10. "
        f"The Environment section IS filled - that answers your question. But "
        f"all nineteen answers in section 3 are set to Yes, and {len(changes)} "
        f"of them should be No.", colour=GREY, italic=True)

    D.h("Read as one sentence, which is how it is assessed", size=12,
        colour=RED)
    D.p("\"The project has out-of-scale features, its visual effect extends "
        "over a large area, it risks breaching environmental standards, its "
        "effects are cumulative, permanent, continuous, frequent, "
        "IRREVERSIBLE, and difficult to avoid, reduce, repair or "
        "compensate.\"")
    D.p("Beside an Annex I habitat type. For a project whose whole "
        "environmental case is reversibility, that is the most damaging thing "
        "in the application - and more urgent than the maps, the associated "
        "partner or co-financing, because those are known gaps and this one "
        "looks finished.")
    D.p("Nine of the nineteen contradict the paragraph printed next to them. "
        "Question 15 says Yes beside the words \"Temporary by design\". "
        "Question 18 says Yes beside the word \"No.\" Question 2 says Yes "
        "beside \"No. The built elements are footpath-width boardwalk "
        "runs...\"")

    D.h("This is my fault and here is the mechanism", size=12)
    D.p("I sent you nineteen blocks of text and not one toggle. I wrote each "
        "paragraph to argue its position and assumed the position was obvious "
        "from the prose. It is obvious if you have read the paragraph; it is "
        "not obvious while pasting nineteen boxes in a row, and the control "
        "appears to default to Yes.")
    D.p("I found the same defect in Checklist A in October and fixed the "
        "toggles without changing how I hand checklists over, so it arrived "
        "again in a different section. From here every checklist I send you "
        "leads with the answer column and the text comes second.", colour=GREY)

    D.h(f"The {len(changes)} to change", size=12, colour=RED)
    D.table(["Q", "Question", "Now", "Set to", "Why"],
            [[str(n), byn[n][0], LIVE[n], TOGGLE[n][0], TOGGLE[n][1]]
             for n in changes],
            widths=[0.9, 5.4, 1.0, 1.2, 8.5], bold_col=3)

    D.h("The ten that are already right - do not touch these", size=12,
        colour=GREEN)
    D.table(["Q", "Question", "Answer", "Why it stays"],
            [[str(n), byn[n][0], TOGGLE[n][0], TOGGLE[n][1]]
             for n in sorted(set(TOGGLE) - set(changes))],
            widths=[0.9, 5.4, 1.2, 9.5])

    D.h("And two answers were pasted twice", size=12, colour=RED)
    D.p("Question 10 in your form carries question 9's text word for word. "
        "Question 18 carries question 17's text word for word. The two "
        "paragraphs I wrote for 10 and 17 are not in the application at all. "
        "Here they are - replace the duplicated text in those two boxes.")
    for n in LOST:
        D.h(f"Question {n}: {byn[n][0]}", size=11, space=10)
        D.p(f"Toggle: {TOGGLE[n][0]}", bold=True)
        D.p(clean(byn[n][1]))
        D.p(f"{len(clean(byn[n][1]))}/{LIM} characters.", size=8.5,
            colour=GREY, italic=True)

    D.h("What is correct and stays as it is", size=12, colour=GREEN)
    D.p("Checklist B section 2, the location grid: densely populated No, "
        "cultural heritage No, natural protected area No, wetlands No, "
        "coastal Yes, marine Yes, estuaries No, mountain N/A, buffer zone "
        "YES, special biodiversity area No, vulnerable landscape Yes. That "
        "grid is honest and it agrees with section 3.1.1 of the application - "
        "the buffer-zone Yes is the one that would have contradicted our own "
        "narrative had it been No.")
    D.p("Checklist A: infrastructure Yes, environmental protection Yes, "
        "context taken into account Yes, potential impact Yes, monitoring "
        "Yes, costs in the budget Yes. Those are Yes because they should be.")
    D.p("Checklist C remains not required - nothing in Infrastructures and "
        "works, and demountable installations are not five-year investments. "
        "Both thresholds fail and the statement is already in the form.")

    D.d.save(OUT)
    print(f"wrote {OUT}")
    print(f"  toggles to change: {changes}")
    print(f"  texts to replace:  {list(LOST)}")
    print(f"  staying as-is:     {sorted(set(TOGGLE) - set(changes))}")


if __name__ == "__main__":
    main()
