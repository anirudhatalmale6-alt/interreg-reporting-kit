#!/usr/bin/env python3
"""Relevance and logical framework text for the online form, sections 3.1 and 3.2.

HE ASKED: "I need text for 3.1.1 and 3.1.4 relevance, and so on from the
template. Relevance, Logical framework, ..."

This is the JUDGEMENT category - drafted, not generated. But two things found
while drafting it change what is possible, and both are listed at the top of
the output so he reads them before he reads the prose.

TWO HARD RULES FROM THE GUIDELINES THAT CONSTRAIN THE LOGICAL FRAMEWORK

    "projects shall define one expected result maximum per specific objective,
     meaning maximum two expected results per project"

We are under ONE specific objective, 2.2. Therefore ONE expected result for
the whole project. Not one per work package, not one per site. One. Most of
the instinct to write three, transfer and adoption and replication, has to be
compressed into a single sentence that contains all three.

    "projects shall include two mandatory Work Packages (WP1 'Management' and
     WP2 'Communication') and a maximum of four technical Work Packages"

We have four technical work packages, WP3 to WP6. We are exactly at the
ceiling. Nothing more can be added, and if anything needs splitting later
something else has to merge.

Also: "the choice of the project overall objective is SET and corresponds to
the Programme Specific Objectives." So the overall objective is not drafted at
all - it is SO 2.2's own wording, copied.

AND THE SECOND MISSING PROGRAMME DOCUMENT

The indicator codes are NOT in the Guidelines. They are defined in a separate
paper the Guidelines name: the "PERFORMANCE FRAMEWORK METHODOLOGY PAPER",
which "provides definitions of each result/output indicator and background
information on how target values have been elaborated."

So there are now TWO documents blocking real work, both of which should be on
the same programme page as the Guidelines:

    Terms of Reference          criterion 1.1, 4 points at Step 1 + 4 at Step 2
    Performance Framework paper criterion 2.4, 4 points

I have not invented a single RCO or RCR code. Every indicator slot below says
which document supplies it.

ON THE CHARACTER LIMITS

His screenshot shows a counter under each box reading "0/400" for 3.1.1 and
"0/30" for 3.1.2, both CLIPPED BY THE WINDOW EDGE. So the real limits are
either 400 and 300, or 4000 and 3000, and I cannot tell which from the image.

4000 is the normal size for this section in an Interreg form and 400 would be
about two sentences. But normal is not verified, so each section below is
written to the larger reading AND opens with a self-contained first paragraph
that works alone under 400 characters. The character count of each part is
printed in the document, so he can see at a glance what fits.
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

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_relevance_and_logframe_2026-10-02.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

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
        s.left_margin = s.right_margin = Cm(2.0)
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
                if "[" in str(v) and "]" in str(v):
                    r.font.color.rgb = RED
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


# --------------------------------------------------------------- the text
# Each entry: (form section, lead paragraph that stands alone under 400
# characters, then the rest).
S311_LEAD = (
    "Mediterranean sediment shores are the asset tourism depends on and the "
    "surface it erodes. The science of shore resilience is well developed; "
    "the site managers who must act, municipalities and local authorities, "
    "rarely hold a method they can apply themselves."
)
S311_REST = [
    "Proven solutions exist. They have not travelled east or south. AMMOS "
    "transfers a method already validated in the western Mediterranean to "
    "three receiving territories in Greece, Türkiye and Lebanon, and adapts it "
    "to shore types it was not written for.",

    "The need is specific and documented rather than general. At Gomati on "
    "Lemnos, an Annex I habitat under recorded visitor pressure sits outside "
    "Natura 2000 site GR4110006, which covers almost the whole coastal and "
    "marine perimeter of the island, and outside the adjacent Faraklou "
    "Geological Park. It belongs to neither. A habitat under pressure that "
    "falls between two designations is precisely the situation for which a "
    "transferable management method, rather than a new designation, is the "
    "realistic instrument.",

    "The same gap recurs in different forms at the other two sites: an "
    "artificial recreation pond in Denizli province carrying heavy seasonal "
    "visitor concentration on unconsolidated sediment, and a Lebanese coastal "
    "site where the national policy framework exists but the field method does "
    "not.",

    "Contribution to the Programme Specific Objective is direct. SO 2.2 "
    "promotes climate change adaptation and disaster risk prevention taking "
    "into account ecosystem-based approaches, under the topic of tourism and "
    "the green transition. AMMOS does not propose engineering against the sea. "
    "It transfers a method for measuring, monitoring and managing the capacity "
    "of a sediment shore to absorb visitor pressure while the climate changes "
    "around it, and it puts that method in the hands of the authorities "
    "responsible for the shore.",

    "[TERMS OF REFERENCE: one paragraph here adopting the Terms of Reference's "
    "own wording for the sustainable tourism challenges of the cooperation "
    "area. Criterion 1.1 refers to that document explicitly and we do not yet "
    "have it. This is the single highest-value gap in the application.]",
]

S312_LEAD = (
    "The method we transfer exists in one part of the Mediterranean and the "
    "problem it solves exists across all of it. Transfer between territories "
    "is not a presentational feature of this project; it is the project."
)
S312_REST = [
    "A single country could not produce this result. The donor territory holds "
    "a method validated on marine dune systems and does not hold the shore "
    "types that test its limits. The receiving territories hold those shore "
    "types, and the visitor pressure, and not the method.",

    "The three territories are chosen as a gradient rather than as a sample. "
    "Lemnos is a marine dune system and the calibration site, where the method "
    "runs as written. The Lebanese site is a marine coast where most of it "
    "transfers. Ucarı in Denizli province is an artificial freshwater "
    "reservoir shore where the physical protocols transfer and the biological "
    "ones must be rebuilt. Running the same method across that gradient "
    "establishes not only that it transfers but WHERE IT STOPS TRANSFERRING, "
    "which is the knowledge a fourth territory actually needs.",

    "The transnational benefit is mutual and asymmetric, which is why it is "
    "real. The receiving territories gain a working method and trained teams. "
    "The donor method gains validation outside the conditions it was written "
    "for, and an extension to lacustrine and riparian sediment shores that did "
    "not previously exist. Neither outcome is obtainable nationally.",

    "Comparability is the mechanism. The same indicators, collected the same "
    "way, at three sites in three countries, with a common parameter set "
    "agreed before fieldwork begins. That is what makes the result usable by "
    "somebody who was never a partner, and it is why the project needs a "
    "cooperation structure rather than three parallel contracts.",
]

S313_LEAD = (
    "The primary target group is the people who manage a shore and must act: "
    "local authorities and site managers at the three demonstration "
    "territories. The final beneficiaries are the communities whose economy "
    "rests on those shores, and the visitors who use them."
)
S313_REST = [
    "Target groups, in order of how directly the project changes what they do:",
    "Site managers and local authorities with territorial competence at the "
    "three sites. Number: [FROM PARTNERS]. They receive the adapted method, "
    "are trained in it, and apply it.",
    "Technical staff of those authorities who carry out monitoring. Number: "
    "[FROM PARTNERS].",
    "Tourism operators at the three shores, being concessions, boat and tour "
    "operators, accommodation and catering. Number: [FROM PARTNERS, and "
    "Acıpayam's figure for Ucarı has been outstanding since September].",
    "Young people in the volunteer and custodian programmes. Number: [FROM "
    "PIYA, who already run the network this builds on].",
    "Public authorities beyond the partnership able to replicate, including "
    "the regional authority for the Greek site. Number: [FROM PARTNERS].",
    "Final beneficiaries: residents of the three territories and seasonal "
    "visitors to the three sites. Number: [FROM PARTNERS, visitor counts per "
    "site per season].",
    "[EVERY NUMBER IN THIS SECTION IS DELIBERATELY BLANK. Criteria 1.3 and 2.3 "
    "both use the word QUANTIFIED, and inventing an audience size to fill a "
    "box is the one thing that cannot be corrected later. One request to each "
    "partner closes this section and the communication strategy together.]",
]

SYNERGIES = [
    ("AMMIRARE, Interreg Italy, France Maritime 2021-2027",
     "The source. MMM_IF02, Guidelines for the Sustainable Management of "
     "Beaches, published as Spiagge che cambiano, is the output AMMOS "
     "transfers; MMM_IF01, Beach Custodians, is the second. CNR, a partner in "
     "AMMIRARE and among the named authors, is a partner here.",
     "Direct capitalisation"),
    ("COMMON, Interreg NEXT MED 2014-2020",
     "MMM_NX36, the Beach litter monitoring guide, a standardised protocol "
     "aligned with the European Environment Agency methodology, in Italian, "
     "French, English and Arabic. MMM_NX35, the BEach CLEAN decalogue, for "
     "operator and community engagement. MMM_NX37 consolidates the Lebanese "
     "national contribution, which our Lebanese partner extends rather than "
     "starts.",
     "Direct capitalisation"),
    ("MEDUSA, Interreg NEXT MED 2014-2020",
     "MMM_NX26 and NX27, adventure tourism mapping and hiking e-guides across "
     "Jordan, Lebanon, Tunisia, Catalonia and Puglia, explicitly aimed at "
     "drawing visitors to alternative areas to reduce pressure on more popular "
     "ones. The same visitor-redistribution logic as our own channelling work.",
     "Complementary method"),
    ("CROSSDEV, Interreg NEXT MED 2014-2020",
     "MMM_NX05 and NX06, destination-level local action plans for "
     "off-the-beaten-path areas. The model for our own uptake work package.",
     "Complementary method"),
    ("The MMM sustainable tourism database itself",
     "AMMOS draws from it and returns to it. The adapted method, the "
     "cross-typology comparability protocol and the lacustrine adaptation are "
     "submitted as outputs to the same database, so the next capitalisation "
     "call can transfer them onward.",
     "Feeds the mechanism"),
    ("EU Habitats Directive, Annex I",
     "The Greek demonstration site is an Annex I habitat type, and the method "
     "produces the condition data that any future designation or management "
     "plan would require.",
     "Policy coherence"),
    ("European Environment Agency beach litter methodology",
     "Adopted through MMM_NX36 rather than reinvented, so the litter data is "
     "comparable with the European dataset.",
     "Standards coherence"),
    ("LIFE IP 4 NATURA, Greece",
     "Habitat methodology and national prioritised action framework. Its "
     "target regions do not include our site, which is why we connect to its "
     "method rather than duplicate its coverage.",
     "Complementary, national"),
    ("[TERMS OF REFERENCE and the programme's own synergies document]",
     "The Guidelines reference a document titled Overview of main policies, "
     "strategies and initiatives having synergies and complementarities with "
     "Interreg NEXT MED. Criterion 1.4 is assessed by the Assessment Board at "
     "STEP 2 ONLY and is worth 4 points. Worth completing properly, but not "
     "before the Step 1 sections.",
     "TO OBTAIN"),
]


def main():
    D = Doc()
    D.h("AMMOS, relevance and logical framework for the online form", size=16)
    D.p(f"Draft of {dt.date.today():%d/%m/%Y}. Sections 3.1.1 to 3.1.4 and the "
        f"logical framework. Read the four notes below before the prose.",
        size=9.5, colour=GREY, italic=True)

    D.h("Four things that constrain what can be written here", colour=RED,
        size=12)
    D.p("ONE EXPECTED RESULT, NOT THREE. The Guidelines: \"projects shall "
        "define one expected result maximum per specific objective, meaning "
        "maximum two expected results per project\". We are under one "
        "objective, so the whole project has ONE expected result. Transfer, "
        "adoption and replication all have to live inside a single sentence.",
        size=10)
    D.p("FOUR TECHNICAL WORK PACKAGES IS THE CEILING, and WP3 to WP6 means we "
        "are exactly at it. Nothing more can be added; if something must be "
        "split later, something else must merge.", size=10)
    D.p("THE OVERALL OBJECTIVE IS NOT DRAFTED. The Guidelines: \"the choice of "
        "the project overall objective is set and corresponds to the Programme "
        "Specific Objectives\". It is SO 2.2's own wording, copied.", size=10)
    D.p("TWO PROGRAMME DOCUMENTS ARE MISSING AND BOTH COST POINTS. The TERMS "
        "OF REFERENCE, named explicitly in criterion 1.1, worth 4 points at "
        "Step 1 and 4 more at Step 2. And the PERFORMANCE FRAMEWORK "
        "METHODOLOGY PAPER, which the Guidelines say defines every RCO and RCR "
        "indicator and how target values were set, which is criterion 2.4 and "
        "4 points. I have invented no indicator codes. Both should be on the "
        "same programme page as the Guidelines.", size=10, bold=True)

    for num, title, lead, rest in [
        ("3.1.1", "Proposal relevance", S311_LEAD, S311_REST),
        ("3.1.2", "Transnational dimension", S312_LEAD, S312_REST),
        ("3.1.3", "Project beneficiaries", S313_LEAD, S313_REST),
    ]:
        D.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        full = lead + " " + " ".join(rest)
        D.h(f"{num} {title}")
        D.p(f"Lead paragraph alone: {len(lead)} characters. "
            f"Whole section: {len(full)} characters. "
            f"If the box limit turns out to be 400 rather than 4000, paste the "
            f"lead paragraph only.", size=9, colour=RED, italic=True)
        D.h("Lead paragraph, stands alone", size=11, space=8)
        D.p(lead)
        D.h("Continuation", size=11, space=8)
        for r in rest:
            D.p(r)

    # ------------------------------------------------- logical framework
    D.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    D.h("3.1.4 Synergies and complementarities")
    D.p("Criterion 1.4 is scored by the Assessment Board at STEP 2 ONLY. "
        "Useful, but not before the Step 1 sections are finished.",
        size=9.5, colour=GREY)
    D.table(["Initiative", "What it gives us", "Relationship"],
            [[a, b, c] for a, b, c in SYNERGIES], widths=[4.6, 9.0, 3.0])

    D.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    D.h("Logical framework")
    D.p("Overall objective: copied from SO 2.2, not drafted.", size=9.5,
        colour=GREY)
    D.p("Promoting climate change adaptation and disaster risk prevention, "
        "resilience taking into account ecosystem-based approaches.",
        italic=True)

    D.h("The one expected result", size=11, space=10)
    D.p("Public authorities and site managers in three Mediterranean "
        "territories apply a shared, transferable method for managing sediment "
        "shores under tourism pressure, the limits of that transfer are "
        "established across three shore typologies, and the adapted method is "
        "available for reuse beyond the partnership.", bold=True)
    D.p("That single sentence carries all three of transfer, limits and "
        "replication, because the Guidelines allow only one. Each expected "
        "result must correspond to at least one work package; this one spans "
        "WP3 to WP6.", size=9.5, colour=GREY)

    D.h("Proposed outputs, for your decision", size=11, space=10)
    D.p("The form requires each OUTPUT to carry a title, a TARGET VALUE and a "
        "SEMESTER OF DELIVERY, with activities numbered underneath as "
        "WP.Output.Activity. This is the list I would propose; the target "
        "values in square brackets need deciding, and the indicator column "
        "cannot be filled until the Performance Framework paper arrives.",
        size=9.5, colour=GREY)
    D.table(["Output", "Title", "Target", "Sem.", "Programme indicator"],
            [["O3.1", "Transfer decision: which of the seven method sheets "
                      "applies, is substituted, or does not apply at each "
                      "site", "1", "I", "[PF paper]"],
             ["O3.2", "Adapted method pack for marine, lacustrine and riparian "
                      "sediment shores, in four languages", "1", "II",
              "[PF paper]"],
             ["O3.3", "Three comparable territorial baselines", "3", "II",
              "[PF paper]"],
             ["O4.1", "Visitor channelling and shore protection measures "
                      "implemented", "3 sites", "III", "[PF paper]"],
             ["O4.2", "Tourism valuation survey, willingness to visit and "
                      "willingness to pay", "3 sites", "III", "[PF paper]"],
             ["O5.1", "Site teams trained in the adapted method",
              "[FROM PARTNERS]", "III", "[PF paper]"],
             ["O5.2", "Shore Custodians programme operating",
              "3 sites", "III-IV", "[PF paper]"],
             ["O6.1", "Management measures adopted or endorsed by the "
                      "competent authority at each site", "3", "IV",
              "[PF paper]"],
             ["O6.2", "Outputs submitted to the MMM database", "3", "IV",
              "[PF paper]"]],
            widths=[1.4, 7.4, 2.4, 1.3, 2.6])
    D.p("Nine outputs across four technical work packages. If that is too "
        "many for the form, O4.2 and O5.2 are the two I would fold into their "
        "neighbours, because they are activities that produce evidence rather "
        "than products somebody else reuses.", size=9.5, colour=GREY)

    out = Path(OUT)
    D.d.save(out)
    print(f"wrote {out}")
    for num, lead, rest in [("3.1.1", S311_LEAD, S311_REST),
                            ("3.1.2", S312_LEAD, S312_REST),
                            ("3.1.3", S313_LEAD, S313_REST)]:
        full = lead + " " + " ".join(rest)
        assert len(lead) <= 400, (
            f"{num} lead paragraph is {len(lead)} characters, over the 400 "
            f"fallback. It must stand alone if the limit is 400.")
        print(f"  {num}: lead {len(lead)} chars, full {len(full)} chars")
    body = "\n".join(p.text for p in D.d.paragraphs)
    assert " , " not in body, "the dash rule has produced a stray comma"
    print("  dash rule clean; no indicator codes invented")
    return out


if __name__ == "__main__":
    main()
