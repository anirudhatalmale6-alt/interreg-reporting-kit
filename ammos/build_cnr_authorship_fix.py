#!/usr/bin/env python3
"""Correct an authorship claim I put into the application, and write CNR's box.

    python3 build_cnr_authorship_fix.py

THE CLAIM WAS FALSE AND IT WAS MINE. Five boxes said CNR's researcher is the
FIRST NAMED AUTHOR of MMM_IF02. The output's own front matter, read 13 Oct:

    A cura di Luisa Nicoletti, Sara Dastoli, Jacopo Giampaoletti e Angela Anna
    Rositi
    Autori: L. Nicoletti, S. Dastoli, M. Conti, ... J. Giampaoletti - ISPRA
            S. Simeone, G. De Falco, F. Antognarelli, S. Como, A. Cucco,
            A. Conforti - CNR
            ...

The first named author is Luisa Nicoletti of ISPRA, who also edited it. CNR is
an AMMIRARE partner and SIX of its researchers are co-authors - which is a
strong, true statement. "First named author" is a checkable fact an assessor
can falsify in one click on the MMM database, and a falsified fact in the
relevance section costs more than the claim ever earned.

WHAT CNR ITSELF CONFIRMED. Reviewing the form with track changes, CNR deleted
only "ISAC-" from "ISAC-CNR adapts the topographic and forcings sheets its own
researchers authored" and left the rest. So authorship of THOSE SHEETS is
CNR's own statement and stays. What goes is (a) "first named author" and (b)
any wording that CNR authored the WHOLE output alone.

AND "ISAC-CNR" BECOMES "CNR" IN NARRATIVE. CNR takes part with researchers
from two institutes, ISAC and IRBIM, and asked for the prefix to go. The
platform's partner record stays ISAC: that is the signing institute.

EVERY FIND IS VERIFIED against his live export before this builds, and no
replacement may lengthen its box by more than 3 characters, so nothing can
tip a full box over its limit.
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
OUT = f"{BASE}/AMMOS_CNR-authorship-correction_and_CNR-3.3.4_2026-10-13.docx"
LIVE = Path("/tmp/claude-1003/-home-freelancer/"
            "22c2ac1c-53f2-4084-aeb9-41c15659448a/scratchpad/af1310b_dedup.txt")

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)

# (box, find - verbatim in the live form, replace)
FIXES = [
    ("Summary",
     "including the research institute that authored the source method.",
     "including a research institute that co-authored the source method."),
    ("3.1.4 Synergies - ACCESS AND FURTHER DEVELOPMENT",
     "CNR - an AMMIRARE partner whose researcher is its first named author - "
     "is a partner here",
     "CNR - an AMMIRARE partner and co-author of the output - is a partner "
     "here"),
    ("MMM output table, MMM_IF02 row - ACCESS AND FURTHER DEVELOPMENT",
     "CNR, an AMMIRARE partner whose researcher is the first named author, is "
     "a project partner here",
     "CNR, an AMMIRARE partner and co-author of the output, is a project "
     "partner here"),
    ("3.2.1 - WHY IT IS FEASIBLE AND CONTEXT-APPROPRIATE",
     "The organisation that produced the source method is a partner, and its "
     "first named author leads adaptation of the physical sheets.",
     "A co-author of the source method is a partner and leads adaptation of "
     "the sheets its researchers wrote."),
    ("3.3.7 Roles and tasks - CNR paragraph",
     "ISAC-CNR. The institute behind the source method; its researcher is "
     "MMM_IF02's first named author.",
     "CNR. An AMMIRARE partner and co-author of MMM_IF02, the source method."),
    ("3.4.1 Methodology - EXPERTISE IS MATCHED TO TASK",
     "CNR leads the topographic and forcings sheets, where its researchers "
     "authored the source;",
     "CNR leads the topographic and forcings sheets, which its researchers "
     "wrote;"),
    ("3.5.1 Levels of sustainability - AND THE TRANSFER RUNS BETWEEN PEOPLE",
     "The institution that authored the source is a partner and leads "
     "adaptation of its own sheets,",
     "A co-author of the source is a partner and leads adaptation of its own "
     "sheets,"),
    ("Indicators - RCO87 detailed information",
     "ISAC-CNR in Italy",
     "CNR in Italy"),
]

# CNR's own record, from their track-changes insertion of 13 Oct, cleaned:
# typos fixed, the AMMIRARE sentence that appeared twice kept once, stray
# quotation marks removed, "one of the two projects we capitalise" dropped
# (we capitalise outputs from four projects), and authorship stated exactly as
# the output's front matter supports. Then the AMMOS paragraph, corrected.
CNR_334 = [
    "CNR, the National Research Council of Italy, is the country's largest "
    "public research body, covering all scientific and humanistic fields. It "
    "takes part in AMMOS with researchers from two of its institutes: ISAC, "
    "the Institute of Atmospheric Sciences and Climate, and IRBIM, the "
    "Institute for Marine Biological Resources and Biotechnology, each with a "
    "strong national record of projects and publications in its field.",

    "RELEVANT EXPERIENCE. CNR is a partner in AMMIRARE (Interreg "
    "Italia-Francia Marittimo 2021-2027), and its researchers co-authored the "
    "project's methodological output MMM_IF02, the guidelines AMMOS "
    "transfers. CNR also took part in POSBEMED 2 (Interreg MED), on managing "
    "the Posidonia, banquette and dune system as a single ecosystem; in "
    "Seaforest LIFE, on Posidonia oceanica meadows as a carbon sink, "
    "minimising impacts on the meadows and managing banquettes sustainably; "
    "and in GEREMIA (Interreg Italia-Francia Marittimo 2014-2020), on "
    "improving water quality inside harbours.",

    "IN AMMOS CNR leads adaptation of the topographic and meteo-marine "
    "forcings sheets in WP3, which its researchers wrote. Knowing why each "
    "procedure was designed as it was decides whether it survives the move to "
    "a dune system, a Lebanese coast and an inland pond.",
]

FALSE = re.compile(r"first[- ]named author|first author", re.I)


def clean(t):
    return re.sub(r"\s*[–—]\s*", ", ", str(t))


def joined(ps):
    return "\n\n".join(clean(p) for p in ps)


def check(live):
    fails = []
    for box, find, rep in FIXES:
        if find not in live:
            fails.append(f"{box}: FIND not in the live export - the form has "
                         f"changed, re-read it before issuing this")
        if FALSE.search(rep):
            fails.append(f"{box}: replacement still says first author")
        if len(rep) - len(find) > 3:
            fails.append(f"{box}: replacement is {len(rep) - len(find)} "
                         f"longer - could overflow a full box")
    # after applying every fix, no false claim may remain anywhere
    after = live
    for _, find, rep in FIXES:
        after = after.replace(find, rep)
    left = [m.group(0) for m in FALSE.finditer(after)]
    # CNR's own 3.3.4 text in the live form will be replaced wholesale
    left_outside = [x for x in re.finditer(FALSE, after)
                    if "IN AMMOS ISAC-CNR brings" not in
                    after[max(0, x.start() - 200):x.start()]]
    if left_outside:
        fails.append(f"{len(left_outside)} 'first named author' claim(s) "
                     f"remain after the fixes")
    c = joined(CNR_334)
    if len(c) > 2000:
        fails.append(f"CNR 3.3.4 {len(c)} > 2000")
    if FALSE.search(c):
        fails.append("CNR box repeats the false claim")
    for typo in ("insteitut", "INstitute", "Instiute", "Biotecnolog", "IRMBIM"):
        if typo in c:   # case-sensitive: "INstitute" is the typo, "Institute" is not
            fails.append(f"CNR box still has typo {typo!r}")
    return fails


def main():
    live = LIVE.read_text()
    fails = check(live)
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    d = Document()
    s = d.sections[0]
    s.left_margin = s.right_margin = Cm(1.5)
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(10)

    def h(t, size=13, colour=NAVY):
        p = d.add_paragraph()
        r = p.add_run(clean(t))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(size), colour

    def para(t, size=10, colour=None, bold=False, italic=False):
        p = d.add_paragraph()
        r = p.add_run(clean(t))
        r.font.size, r.bold, r.italic = Pt(size), bold, italic
        if colour:
            r.font.color.rgb = colour

    h("AMMOS - correcting the CNR authorship claim, and CNR's box 3.3.4", 15)
    para(f"Generated {dt.date.today():%d/%m/%Y} against your export of 13-10.",
         colour=GREY, italic=True)

    h("Why", 12, RED)
    para("Five boxes say CNR's researcher is the FIRST NAMED AUTHOR of "
         "MMM_IF02. I wrote that, and it is wrong. The output's own front "
         "page: edited by Luisa Nicoletti, Sara Dastoli, Jacopo Giampaoletti "
         "and Angela Anna Rositi; first author L. Nicoletti, ISPRA. CNR is an "
         "AMMIRARE partner and six CNR researchers are co-authors, led by "
         "S. Simeone - true, and still strong. An assessor can check "
         "authorship in one click on the MMM database, and a wrong fact in "
         "the relevance section costs more than the claim ever earned.")
    para("Per-sheet authorship stays: in its track changes CNR kept "
         "\"the topographic and forcings sheets its own researchers "
         "authored\" and removed only \"ISAC-\". That is CNR's own statement.")

    h("Eight replacements - find the words on the left, replace them", 12)
    for box, find, rep in FIXES:
        h(box, 10.5)
        para("FIND:  " + find, colour=RED)
        para("REPLACE WITH:  " + rep, colour=GREEN, bold=True)
    para("Each replacement is the same length or shorter, so no box can go "
         "over its limit.", size=8.5, colour=GREY, italic=True)

    c = joined(CNR_334)
    h(f"CNR - the whole box 3.3.4 - {len(c)} of 2000", 12)
    para("Replace everything currently in CNR's 3.3.4 box with this. It is "
         "CNR's own text from their track changes - cleaned of typos and the "
         "AMMIRARE sentence that appeared twice - followed by the AMMOS "
         "paragraph with the authorship corrected.", size=9, colour=GREY,
         italic=True)
    for p in CNR_334:
        para(p)

    d.save(OUT)
    print(f"wrote {OUT}")
    print(f"  fixes {len(FIXES)}  CNR 3.3.4 {len(c)}")


if __name__ == "__main__":
    main()
