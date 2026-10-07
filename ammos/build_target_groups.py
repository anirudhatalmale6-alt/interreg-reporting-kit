#!/usr/bin/env python3
"""Target group (100) and Target group involvement (200) for every output,
plus the WP1 and WP2 outputs I wrongly said did not exist.

    python3 build_target_groups.py

HIS QUESTION: "Dear friend they also have Target group" - yes, and it is per
OUTPUT, not per work package. Two fields, and they are the tightest in the
whole application:

    Target group               100 characters
    Target group involvement   200 characters

One hundred characters is a phrase, not a sentence. So each one names the
group and nothing else, and the involvement box says what that group actually
does rather than that it will be "engaged".

AND THE CORRECTION I OWE HIM

On 2 October I told him the mandatory work packages have no outputs, because
the exported Word form showed no output rows under WP1 and WP2. He then found
that WP Budget per output will not work until WP1 and WP2 have outputs - and
the programme's own Courtesy Budget confirms it, with Output 1.1 through 1.4
under WP1.

The export was not the platform. I treated a Word rendering of the form as the
form, and it was missing rows the live system requires. Four outputs below for
WP1 and WP2, with target values, units and semesters.

If he has already created his own - he said "WP BUDGET PER OUTPUTS DONE" -
then his stay and only the target group text is needed.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_outputs import OUTPUTS

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Target-groups_and_WP1-WP2-outputs_2026-10-08.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

LIM_TG, LIM_INV = 100, 200
LIM_O_TITLE, LIM_O_DESC, LIM_O_UNIT = 100, 500, 50
ROMAN = {1: "I", 2: "II", 3: "III", 4: "IV"}

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


# code -> (target group <=100, involvement <=200)
TG = {
    "Output1.1": (
        "Project partners and their financial and administrative staff",
        "They sit on the Management Committee with a named representative and "
        "alternate, apply the RACI table to their own activities, and submit "
        "certified expenditure each reporting period."),
    "Output1.2": (
        "The Managing Authority, the Joint Secretariat and the partnership",
        "They receive the semester risk review and the indicator progress "
        "report; partners supply the data and the Committee records any "
        "reclassified transferability cell as a finding with its evidence."),
    "Output2.1": (
        "Site managers, municipal staff and visitors at the three sites",
        "They receive the method and the on-site material in Greek, Turkish "
        "or Arabic rather than in the source language, and the partner in "
        "their own country produces and distributes it."),
    "Output2.2": (
        "Mediterranean territories, MMM programmes and the scientific "
        "community",
        "They attend six local workshops and the final transferability "
        "conference, and take the guide from the MMM database; the "
        "replication audience is reached through the mechanism, not only "
        "through us."),

    "Output3.1": (
        "Site managers and technical staff of the three site authorities",
        "They state at the transfer workshop what their site can and cannot "
        "support - that is the input the 21 decisions are built from - and "
        "they receive the adapted method in their own language."),
    "Output3.2": (
        "The three site authorities and their planning departments",
        "They host the field campaigns, supply local data and visitor "
        "records, and receive a baseline for their own shore that they can "
        "use in planning decisions independently of this project."),
    "Output4.1": (
        "Visitors, tourism operators and the three site authorities",
        "Operators and residents are consulted on the installation design; "
        "the authority that will own it signs the design off; visitors use "
        "the channelled access through one full season under monitoring."),
    "Output4.2": (
        "Site monitoring staff and future adopting territories",
        "Site staff enter field observations in their own language and own "
        "the resulting series; a new territory can join the protocol using "
        "the uptake section, without renegotiating it with us."),
    "Output5.1": (
        "Schools and pupils near the three sites, and their teachers",
        "Nine groups run at least one monitoring round and one awareness "
        "action at their own shore, with their observations entering the "
        "shared platform - they produce data, not just attend."),
    "Output5.2": (
        "Municipal and authority staff with a mandate over each site",
        "Sixty staff are trained in their own language with a practical "
        "module at the site itself, and the same people then write the "
        "management plan in WP6 rather than receiving one."),
    "Output6.1": (
        "The authority with the mandate over each site, and local "
        "stakeholders",
        "Each authority takes a costed plan through consultation to a formal "
        "adoption decision, and sits on the Shore Stewardship Committee that "
        "oversees it with an operator and a civil society or youth voice."),
    "Output6.2": (
        "National and regional authorities, and territories beyond the "
        "partnership",
        "They receive policy recommendations in English and Arabic and the "
        "guide lodged in the MMM database, which states the actions needed to "
        "adopt or upscale the method elsewhere."),
}

# WP1 and WP2 outputs. (code, wp, title<=100, target, unit<=50, semester,
#                       share, description<=500)
MANDATORY_OUTPUTS = [
    ("Output1.1", "WP1",
     "Project management and coordination system in operation",
     1, "management and coordination system", 1, 0.50,
     "The instruments that keep a six-partner transfer project consistent: "
     "the Management Committee with a named representative and alternate per "
     "partner, the RACI table assigning every output and activity, the "
     "decision rules naming what requires a recorded decision, and the "
     "financial circuit in which each partner contracts its own independent "
     "first-level control and the Applicant consolidates. Established in the "
     "first semester because everything else depends on it."),

    ("Output1.2", "WP1",
     "Monitoring, evaluation and risk management cycle completed",
     4, "reporting periods completed", 4, 0.50,
     "Four reporting periods, each with certified expenditure, indicator "
     "progress against the project's six indicators, and a semester risk "
     "review. Deliverables pass two checks: the activity lead signs technical "
     "content and the scientific coordinator signs cross-site consistency. "
     "Reclassifying a transferability cell is recorded as a finding with its "
     "evidence rather than as a deviation, because the boundary of a method's "
     "applicability is itself a result."),

    ("Output2.1", "WP2",
     "Communication and visibility package in four languages",
     4, "languages (EN, EL, TR, AR)", 2, 0.50,
     "The website with its public and partner areas embedded in the "
     "municipal, prefecture and regional sites of the three territories; "
     "social media in four languages; on-site interpretation at all three "
     "sites and a visitor booklet; and the programme visibility rules applied "
     "uniformly. Produced in Greek, Turkish and Arabic as well as English, "
     "because the method's users do not read the language it was written in."),

    ("Output2.2", "WP2",
     "Dissemination and transferability outreach delivered",
     7, "workshops, press cycles and the final conference", 4, 0.50,
     "Six local workshops - one stakeholder and one results workshop per "
     "territory - and the final transferability conference, plus two "
     "newsletter cycles to local sectoral press in every partner country at "
     "month 1 and month 24, and one peer-reviewed paper. Reach is measured "
     "from the WP3 visitor survey rather than estimated, and participations "
     "are counted on attendance lists against the project's participation "
     "indicator."),
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
    codes = [f"Output{o[0][2]}.{o[1]}" for o in OUTPUTS]
    all_codes = [m[0] for m in MANDATORY_OUTPUTS] + codes

    fails = []
    for c in all_codes:
        if c not in TG:
            fails.append(f"{c} has no target group")
    for c, (tg, inv) in TG.items():
        if c not in all_codes:
            fails.append(f"{c} is a target group for an output that does not "
                         f"exist")
        if len(clean(tg)) > LIM_TG:
            fails.append(f"{c} target group {len(clean(tg))} > {LIM_TG} - "
                         f"cut {len(clean(tg)) - LIM_TG}")
        if len(clean(inv)) > LIM_INV:
            fails.append(f"{c} involvement {len(clean(inv))} > {LIM_INV} - "
                         f"cut {len(clean(inv)) - LIM_INV}")
    for code, wp, title, tv, unit, s, share, desc in MANDATORY_OUTPUTS:
        if len(clean(title)) > LIM_O_TITLE:
            fails.append(f"{code} title {len(clean(title))} > {LIM_O_TITLE}")
        if len(clean(desc)) > LIM_O_DESC:
            fails.append(f"{code} description {len(clean(desc))} > "
                         f"{LIM_O_DESC} - cut {len(clean(desc)) - LIM_O_DESC}")
        if len(clean(unit)) > LIM_O_UNIT:
            fails.append(f"{code} unit {len(clean(unit))} > {LIM_O_UNIT}")
    for wp in ("WP1", "WP2"):
        tot = sum(m[6] for m in MANDATORY_OUTPUTS if m[1] == wp)
        if abs(tot - 1.0) > 1e-9:
            fails.append(f"{wp} output shares total {tot}, not 1.0")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — Target group and Target group involvement, every output",
        size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}. Target group 100 characters, "
        f"Target group involvement 200. The two tightest fields in the "
        f"application, so each names the group and nothing else, and the "
        f"involvement box says what the group does rather than that it will "
        f"be engaged.", size=9.5, colour=GREY, italic=True)

    D.h("A correction I owe you first", size=13)
    D.p("On 2 October I told you the mandatory work packages have no outputs. "
        "You found that WP Budget per output will not work until WP1 and WP2 "
        "have them, and the programme's own Courtesy Budget confirms you are "
        "right - its budget-per-output sheet shows Output 1.1 to 1.4 under "
        "WP1.", size=10)
    D.p("My mistake was treating the exported Word form as the form. The "
        "export showed no output rows under WP1 and WP2 and I concluded they "
        "had none; the live platform requires them. Four outputs are below. "
        "If you have already created your own - you said budget per outputs "
        "was done - keep yours and take only the target group text.",
        size=10, bold=True)

    D.h("WP1 and WP2 outputs", size=14)
    D.table(["Code", "Title", "Target", "Unit", "Sem.", "Budget share"],
            [[c, t, str(tv), u, ROMAN[s], f"{round(sh * 100)}%"]
             for c, wp, t, tv, u, s, sh, d in MANDATORY_OUTPUTS],
            widths=[2.0, 5.6, 1.3, 4.4, 1.2, 2.2])
    for code, wp, title, tv, unit, s, share, desc in MANDATORY_OUTPUTS:
        D.h(f"{code} — {wp}", size=11, space=10)
        D.p(f"Title ({len(clean(title))}/{LIM_O_TITLE}): {title}", size=10)
        D.p(f"Target value {tv}  ·  Unit ({len(clean(unit))}/{LIM_O_UNIT}): "
            f"{unit}  ·  Semester {ROMAN[s]}  ·  Budget share "
            f"{round(share * 100)}%", size=9, colour=GREY)
        D.p(f"Description ({len(clean(desc))}/{LIM_O_DESC}): {desc}", size=10)

    D.h("Target group and involvement — all twelve outputs", size=14)
    D.p("Paste into the Target tab of each output. The counts are shown so "
        "you can see nothing is near the edge.", size=9.5, colour=GREY)
    D.table(["Output", f"Target group (max {LIM_TG})",
             f"Target group involvement (max {LIM_INV})", "Counts"],
            [[c, TG[c][0], TG[c][1],
              f"{len(clean(TG[c][0]))} / {len(clean(TG[c][1]))}"]
             for c in all_codes],
            widths=[1.9, 5.4, 8.8, 1.6], small=8)

    D.h("And one inconsistency in the budget that you should decide on",
        size=13)
    D.p("Your budget model v6 has ZERO in Infrastructures and works for every "
        "partner. The demonstration installations - boardwalks, boundary "
        "marking, signage - are therefore budgeted as External expertise and "
        "services, which is defensible: fabrication and installation bought "
        "in as a service.", size=10)
    D.p("But section 3.6.2, which you have already pasted, says "
        "\"Infrastructure and works covers only the demonstration "
        "installations\". That describes a cost category with nothing in it, "
        "and an assessor reading 3.6 against the budget would notice.",
        size=10, bold=True)
    D.p("Two ways to fix it. Either I reword that sentence to say the "
        "installations are procured as external services and the "
        "infrastructure category is not used - quick, honest, and it also "
        "keeps the environmental screening lighter. Or we reclassify part of "
        "WP4 into Infrastructures and works, which is arguably more accurate "
        "for fixed works but means changing the partners' agreed figures and "
        "may invite heavier scrutiny of the works. I would reword, but it is "
        "your call and the partners' money.", size=10)

    D.d.save(OUT)
    print(f"wrote {Path(OUT).name}")
    print(f"  {len(MANDATORY_OUTPUTS)} mandatory-WP outputs, "
          f"{len(all_codes)} target groups")
    for c in all_codes:
        print(f"    {c:11s} tg {len(clean(TG[c][0])):>3}/{LIM_TG}   "
              f"inv {len(clean(TG[c][1])):>3}/{LIM_INV}")


if __name__ == "__main__":
    main()
