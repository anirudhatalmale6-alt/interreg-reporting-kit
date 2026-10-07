#!/usr/bin/env python3
"""Activities for the WP1 and WP2 outputs.

    python3 build_wp12_activities.py

"wp1 and wp2 have also tab activities where can i find this?"

Nowhere yet - because I never wrote them. On 2 October I concluded the
mandatory work packages have no outputs and therefore no activities, so when I
drafted WP1 and WP2 I put everything into the single description box and told
him the numbered activities "stop being numbered". They do not. His screenshot
shows the two WP1 outputs he has entered, each with an "Add Activities (0)"
link beside it.

So the activities come back, numbered the way everything else is:

    Activity {WP}.{Output}.{Activity}

Eight of them, two per output, same field sizes as the technical work packages
since it is the same form: Activity title 100, Description 500, Implementing
period as a month multi-select.

WHERE THE CONTENT COMES FROM

Not invented. The original work package document had three management
activities and three communication activities, drafted in September and
carried in the WP1 and WP2 description boxes ever since. They are re-cut into
two per output, because that is what the outputs now require, and nothing in
them contradicts the description boxes he has already pasted - the description
is the summary, the activities are the same work itemised.

THE ONE JUDGEMENT CALL

WP1 and WP2 run M1-M24, so every activity could claim all 24 months and none
of it would be wrong. That would also be useless: an implementing period that
spans the whole project tells a monitoring officer nothing. So each activity
carries the months in which its work actually falls - the coordination
structure is set up in the first semester and then maintained, the reporting
cycle runs to the end, the website is built early and the final conference is
at the end. Only the two that genuinely run continuously claim M1-M24.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import P, PP

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_WP1-WP2_activities_2026-10-08.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

LIM_TITLE, LIM_DESC = 100, 500
MONTHS = 24

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


# (code, output, title<=100, m_from, m_to, lead, description<=500)
ACTS = [
    ("1.1.1", "Output1.1",
     "Coordination structure, roles and the RACI table established",
     1, 6, "HCMR",
     "The Applicant appoints a scientific coordinator and project manager. "
     "Each partner names a representative and an alternate to the "
     "Management Committee, which meets monthly online and once per semester "
     "in person, scheduled alongside demonstration and communication work so "
     "one trip serves two purposes. A written RACI table assigns, for every "
     "output and activity, who is responsible, accountable, consulted and "
     "informed. The decision rules name the three changes requiring a "
     "recorded Committee vote."),

    ("1.1.2", "Output1.1",
     "Financial circuit and first-level control arrangements in place",
     1, 24, "HCMR",
     "Each partner contracts its own independent first-level control and "
     "submits certified expenditure per reporting period. The Applicant "
     "consolidates, holds the single interface with the Programme, and tracks "
     "each partner against its own semester spending profile - every budget "
     "line in this application states the semester it falls in, so a partner "
     "drifting from profile is visible in the period it happens rather than "
     "at closure. Partnership agreement signed in the first semester."),

    ("1.2.1", "Output1.2",
     "Quality assurance and two-stage deliverable check applied",
     3, 24, "HCMR",
     "Every deliverable passes two checks before release: the activity lead "
     "signs the technical content, and the scientific coordinator signs "
     "consistency across the three sites. The second check exists because a "
     "protocol that is correct at one site and incompatible with another "
     "destroys the comparability the project is built on. Reclassifying a "
     "transferability cell is recorded as a finding with its evidence rather "
     "than as a deviation."),

    ("1.2.2", "Output1.2",
     "Reporting, risk review and indicator monitoring across four periods",
     6, 24, "HCMR",
     "Four reporting periods, each carrying certified expenditure, progress "
     "against the project's six programme indicators, and a semester risk "
     "review. The principal risks are seasonal: a field campaign lost to "
     "weather or a delayed consent cannot be repeated for twelve months, so "
     "each site holds a second campaign window and the installations complete "
     "at month 20, four months before closure. The register is reviewed by "
     "the Committee each semester, not only when something goes wrong."),

    ("2.1.1", "Output2.1",
     "Project website, partner area and multilingual channels launched",
     1, 8, "HCMR",
     "A website with a public area and a password-protected partner area, "
     "embedded in the municipal, prefecture and regional sites of the three "
     "territories so that it is found where people already look rather than "
     "only where we put it. Social media channels in English, Greek, Turkish "
     "and Arabic, run by the youth partner, with the nine school custodian "
     "groups generating content from their own monitoring rounds. Programme "
     "visibility rules applied uniformly from the first publication."),

    ("2.1.2", "Output2.1",
     "On-site interpretation and visitor material produced in local languages",
     8, 20, "PIYA",
     "Interpretation panels and a visitor booklet at all three sites, each "
     "produced in the local language and designed to be understood without "
     "specialist knowledge, using pictograms alongside text. Each site "
     "partner produces and distributes its own, because local credibility and "
     "local distribution cannot be run centrally. The BEach CLEAN decalogue "
     "is adapted rather than rewritten, and its existing Arabic version is "
     "extended."),

    ("2.2.1", "Output2.2",
     "Local workshops in the three territories and press cycles",
     1, 22, "LCEC",
     "Six local workshops, one stakeholder workshop and one results workshop "
     "per territory, run by the partner in that country in its own language. "
     "Two newsletter cycles to local sectoral press in every partner country: "
     "month 1 on objectives and month 24 on results. LCEC leads the "
     "policy-facing communication to national authorities, where a national "
     "public agency is heard differently from a research institute."),

    ("2.2.2", "Output2.2",
     "Final transferability conference, scientific paper and MMM submission",
     20, 24, "HCMR",
     "The final transferability conference presents the 21 decisions, the "
     "comparability layer and the adopted plans to the replication audience. "
     "One peer-reviewed paper is submitted. The transferability guide in "
     "English and Arabic is lodged in the MMM common database under open "
     "licence, with its section stating what an adopting organisation must do "
     "- which is what makes it a countable solution rather than a "
     "publication. Reach is reported from the WP3 visitor survey, measured "
     "rather than estimated."),
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
    fails = []
    for code, out, title, m0, m1, lead, desc in ACTS:
        if len(clean(title)) > LIM_TITLE:
            fails.append(f"{code} title {len(clean(title))} > {LIM_TITLE} - "
                         f"cut {len(clean(title)) - LIM_TITLE}")
        if len(clean(desc)) > LIM_DESC:
            fails.append(f"{code} description {len(clean(desc))} > "
                         f"{LIM_DESC} - cut {len(clean(desc)) - LIM_DESC}")
        if not (1 <= m0 <= m1 <= MONTHS):
            fails.append(f"{code} months {m0}-{m1}")
        if lead not in P:
            fails.append(f"{code} lead {lead!r} is not a partner")
        wp, oi, ai = code.split(".")
        if out != f"Output{wp}.{oi}":
            fails.append(f"{code} is filed under {out}")
    # Two activities per output, contiguous numbering.
    for out in ("Output1.1", "Output1.2", "Output2.1", "Output2.2"):
        idx = [int(a[0].split(".")[2]) for a in ACTS if a[1] == out]
        if idx != list(range(1, len(idx) + 1)):
            fails.append(f"{out} activity numbers are {idx}, must be 1..n")
        if not idx:
            fails.append(f"{out} has no activities")
    # An implementing period covering the whole project tells nobody anything;
    # allow it only where the work really is continuous.
    whole = [a[0] for a in ACTS if a[3] == 1 and a[4] == MONTHS]
    if len(whole) > 2:
        fails.append(f"{len(whole)} activities claim M1-M24 ({whole}) - that "
                     f"is a schedule with no information in it")
    blob = " ".join(str(x) for a in ACTS for x in a)
    for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson", r"xxxx"):
        if re.search(forbidden, blob, re.I):
            fails.append(f"text contains {forbidden!r}")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS — WP1 and WP2 activities, all eight", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}. Activity title 100, "
        f"Description 500, Implementing period as a month multi-select - the "
        f"same form as the technical work packages. Every field checked "
        f"against its counter.", size=9.5, colour=GREY, italic=True)

    D.h("Why you could not find these", size=13)
    D.p("Because I never wrote them. On 2 October I decided the mandatory "
        "work packages have no outputs and therefore no activities, so I put "
        "all the management and communication detail into the single "
        "description box and told you the numbered activities stop being "
        "numbered. They do not - your screenshot shows Add Activities beside "
        "each of your two WP1 outputs, both at zero.", size=10.5)
    D.p("The content is not new. These are the three management and three "
        "communication activities drafted in September, re-cut into two per "
        "output because that is what the outputs now need. Nothing in them "
        "contradicts the description boxes you have already pasted: the "
        "description is the summary, the activities are the same work "
        "itemised.", size=10.5)

    D.h("One judgement call worth knowing about", size=12)
    D.p("WP1 and WP2 both run M1 to M24, so every activity could claim all "
        "twenty-four months and none of it would be false. It would also be "
        "useless - an implementing period that spans the whole project tells "
        "a monitoring officer nothing. So each activity carries the months "
        "its work really falls in, and only the two that are genuinely "
        "continuous claim M1-M24. The build refuses to produce this document "
        "if more than two do.", size=10.5)

    D.table(["Code", "Under output", "Title", "Months", "Leads"],
            [[c, o, t, f"M{m0}-M{m1}", f"{PP[l]} {P[l][0]}"]
             for c, o, t, m0, m1, l, d in ACTS],
            widths=[1.4, 2.0, 6.6, 1.8, 4.2])

    for code, out, title, m0, m1, lead, desc in ACTS:
        D.h(f"Activity {code}  —  add under {out}", size=11.5, space=10)
        D.p(f"Title ({len(clean(title))}/{LIM_TITLE}): {title}", size=10)
        D.p(f"Implementing period: months "
            f"{', '.join(str(x) for x in range(m0, m1 + 1))}", size=10)
        D.p(f"Leads: {PP[lead]} {P[lead][0]}", size=9, colour=GREY)
        D.p(f"Description ({len(clean(desc))}/{LIM_DESC}): {desc}", size=10)

    D.h("That completes the work package section", size=13)
    D.p("With these eight, all six work packages have outputs, activities, "
        "target groups, involved partners, final beneficiaries and budget "
        "rows. Twelve outputs and twenty-four activities in total.",
        size=10.5, bold=True)
    D.p("What is left in the application: the six co-financing sources, "
        "financial capacity for six partners, the environmental screening "
        "checklists, and the document uploads. The screening is the largest "
        "and is still untouched - send me a screenshot of Checklist A when "
        "you reach it.", size=10.5)

    D.d.save(OUT)
    print(f"wrote {Path(OUT).name}")
    for c, o, t, m0, m1, l, d in ACTS:
        print(f"   {c}  {o:11s} title {len(clean(t)):>3}/{LIM_TITLE}  "
              f"desc {len(clean(d)):>3}/{LIM_DESC}  M{m0}-M{m1}  {l}")
    print(f"\n   {len(ACTS)} activities; "
          f"{len([a for a in ACTS if a[3] == 1 and a[4] == MONTHS])} claim "
          f"the full M1-M24")


if __name__ == "__main__":
    main()
