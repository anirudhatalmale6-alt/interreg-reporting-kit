#!/usr/bin/env python3
"""Pre-fill the programme's associate declaration for North Aegean.

    python3 build_associate_declaration.py

The Directorate signs tomorrow and, in the client's words, has no idea what it
will be doing. This fills the programme's own template - not a lookalike - so
the wording the Joint Secretariat expects is untouched and only the blanks
change: the official name in Greek and English, associate number 1, the legal
status, the reasons, and the past experience. Signature, date, name and stamp
are left for the Directorate, and the text still has to be moved onto their
headed paper, as the template's first line requires.

LEGAL STATUS. The template offers "Ministry or other national public
administration" and "Regional or local public administration". The
Decentralized Administration of the Aegean is state administration - the org
chart the client sent places the Water Directorate under its General
Directorate of Spatial, Environmental and Agricultural Policy - not the Region
of North Aegean, which is regional self-government. So the first box. This is
the same distinction as the RCO87 correction, and the declaration is where it
becomes a tick rather than a sentence.

EVERY FACT is from Presidential Decree 143/2010 (the Water Directorate of North
Aegean article) or the org chart: the Lemnos and Samos Water Offices are both
on it.
"""
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

BASE = Path(str(here(__file__)))
SRC = BASE.parent / "5.-Interreg-NEXT-MED_Associate-declaration_01.07.2026.docx"
OUT = BASE / "AMMOS_Associate-declaration_North-Aegean-Water-Directorate_PREFILLED.docx"

NAME = ("Αποκεντρωμένη Διοίκηση Αιγαίου – Διεύθυνση Υδάτων Βορείου Αιγαίου / "
        "Decentralized Administration of the Aegean – North Aegean Water "
        "Directorate")
ASSOC_N = "1"

REASONS = (
    "The Directorate is the statutory water authority for the North Aegean, "
    "including Lemnos, where the AMMOS demonstration site at Gomati lies, and "
    "its participation links the project to the plans it prepares and "
    "implements. As associate it will (1) share existing data for the Gomati "
    "baseline - bathing water quality records, the river basin plan's "
    "characterisation of the island's coastal waters and the protected-area "
    "entries; (2) take part, through its Lemnos Water Office, in the Gomati "
    "Shore Stewardship Committee; and (3) receive the Gomati monitoring "
    "results and the adopted management plan as input to the next review of "
    "the river basin and flood risk management plans and their public "
    "consultation. It manages no project funds and takes on no obligation "
    "beyond its statutory work.")

EXPERIENCE = (
    "Under Presidential Decree 143/2010 the Directorate, seated in Mytilene "
    "with Water Offices on Lemnos and Samos, prepares and implements the River "
    "Basin Management Plans and the Flood Risk Management Plans for the North "
    "Aegean, including flood hazard and flood risk maps; manages bathing water "
    "quality; keeps the register of protected areas; assesses the impact of "
    "human activities on water status; and runs public participation in the "
    "preparation and review of both plans.")

STATUS_LINE = "Ministry or other national public administration"
DOTS = "…"


def set_text(par, text, bold=None):
    """Put text in the first run, empty the rest - keeps the run's font."""
    runs = par.runs
    runs[0].text = text
    if bold is not None:
        runs[0].bold = bold
    for r in runs[1:]:
        r.text = ""


def main():
    assert SRC.exists(), f"template missing: {SRC}"
    d = Document(SRC)
    P = d.paragraphs

    # 2 - the opening sentence: name and associate number, nothing else moves
    p2 = P[2]
    texts = [r.text for r in p2.runs]
    i_name = texts.index("Official name of the organisation in Original "
                         "Language + Translation into English")
    i_num = texts.index("...")
    p2.runs[i_name].text = NAME
    p2.runs[i_num].text = ASSOC_N

    # legal status - tick exactly one
    hits = [p for p in P if p.text.strip() == STATUS_LINE]
    assert len(hits) == 1, f"legal status line found {len(hits)} times"
    set_text(hits[0], "[X]  " + STATUS_LINE + " - state administration "
             "(Decentralized Administration of the Aegean)", bold=True)

    # reasons: the first dotted line after "expresses its will"
    k = next(i for i, p in enumerate(P)
             if p.text.startswith("expresses its will"))
    set_text(P[k + 1], REASONS)
    # the next paragraph is dots + the fixed "reports the following
    # experience" wording: drop only the dots
    p_exp = P[k + 2]
    for r in p_exp.runs:
        if set(r.text) <= {DOTS, "."} and r.text:
            r.text = ""
    assert p_exp.text.strip().startswith("reports the following experience")
    set_text(P[k + 3], EXPERIENCE)
    set_text(P[k + 4], "")

    # nothing that reads as a blank may survive in the body we filled
    # Only the paragraphs I filled. The unselected options keep their own
    # "(please specify)……" dots - they are the template, not my blanks.
    filled = [P[2]] + P[k + 1:k + 5]
    body = "\n".join(p.text for p in P[:k + 5])
    fails = []
    if any(DOTS * 3 in p.text for p in filled):
        fails.append("dotted placeholder still in a paragraph I filled")
    for must in (NAME, ASSOC_N, REASONS, EXPERIENCE, "[X]"):
        if must not in body:
            fails.append(f"missing: {must[:40]}")
    if body.count("[X]") != 1:
        fails.append("more or fewer than one legal status ticked")
    assert not fails, "NOT SAVING:\n  - " + "\n  - ".join(fails)

    d.save(OUT)
    print(f"wrote {OUT}")
    print(f"  reasons {len(REASONS)}  experience {len(EXPERIENCE)}")


if __name__ == "__main__":
    main()
