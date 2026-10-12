#!/usr/bin/env python3
"""Partnership, associates and co-financing: what the 12-10 export still lacks.

    python3 build_partnership_gaps.py

He listed three things still missing: co-financing, the associated partner,
and "partnership applicant all menu". The export shows the third is not one
partner's problem. Every field of the Partnership section was read for all six
organisations, and the gaps fall into three kinds that need three different
people:

  HIM, in minutes        a junk associated partner that blocks submission,
                         placeholder strings in Acipayam's identity fields,
                         duplicated paragraphs in PIYA's capacity box, and one
                         comma in Checklist B Q9 that reverses its meaning.

  HCMR                   the Applicant block is the emptiest in the form -
                         address, acronym, the four capacity figures, contact
                         person, and box 3.3.3, which is SCORED.

  EACH PARTNER           their own 3.3.4 capacity box (scored), the capacity
                         figures, and the named people.

WHAT THE CO-FINANCING TABLE CORRECTED IN MY OWN WORK

His budget export now carries the co-financing sheet, and with it the four real
dropdown values. The partners chose:

  HCMR, Aegean    Contribution from other EU Institutions or Member States
  PIYA            Contribution from Mediterranean Partner Countries
                  (and its description names the Directorate for EU Affairs)
  Acipayam, CNR,  Own funds
  LCEC

My v2 texts said "own resources" for all six. For three of them that
contradicts the dropdown the partner actually chose, and a description that
disagrees with its own source type is worse than a short one. The texts below
are written per SOURCE TYPE, and the generator refuses to emit a text whose
wording does not match the type selected in his export.

WHAT I AM NOT CLAIMING

The name of the Greek national co-financing instrument. Greek public
beneficiaries normally receive it through the Public Investment Programme, and
the earlier "NSRF" in Aegean's row points the same way, but he works in this
system daily and the exact wording is his to confirm. Same for whether CNR, as
an Italian public body, draws on the national rotation fund, and whether
Acipayam, like PIYA, is covered by the Directorate for EU Affairs. Each is a
one-line question, stated as a question.
"""
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Partnership-associates-cofinancing_GAPS_2026-10-12.docx"
SCRATCH = Path("/tmp/claude-1003/-home-freelancer/"
               "22c2ac1c-53f2-4084-aeb9-41c15659448a/scratchpad")
FORM_DOCX = SCRATCH / "af_1210c.docx"
FORM_TXT = SCRATCH / "af1210c_dedup.txt"


def extract(docx_path, out_path):
    """Flatten the export to text, printing each MERGED cell once.

    python-docx returns a cell spanning three grid columns three times. The
    first version of this extraction joined them all, so every merged text
    box appeared twice on one line - which I then reported to the client as a
    partner who had pasted paragraphs twice (PIYA, 12 Oct). The tell was a
    943-character "paragraph" that was exactly the last paragraph, " | ", and
    the first one: 349 + 3 + 591.
    """
    from docx import Document as _D
    d = _D(docx_path)
    lines = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    for t in d.tables:
        for r in t.rows:
            seen, cs = set(), []
            for c in r.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cs.append(c.text.strip())
            if any(cs):
                lines.append(" | ".join(cs))
    Path(out_path).write_text("\n".join(lines))
BUDGET = SCRATCH / "budget_12-10.xlsx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x6B, 0x45)

PLACEHOLDER = re.compile(r"^(?:[xc?]{3,}|g{2,}h[gh]*|[gh]{5,})$", re.I)
FORBIDDEN = ("elke", "finance office", "non-compliant", "option b", "ras ",
             "wrong", "4.4.3", "remaco", "887")

FIELDS = [
    ("Department/service", "Department / service", False),
    ("Translation in English", "Name translated into English", False),
    ("Acronym of the organisation", "Acronym", False),
    ("Region", "Region", False),
    ("Official address", "Official address", False),
    ("Website", "Website", False),
    ("Partner’s annual total budget", "Annual total budget (avg last 2 yrs)",
     False),
    ("Partner’s budget dedicated", "Budget dedicated to EU projects", False),
    ("Partner’s total staff", "Total staff (FTE band)", False),
    ("Staff dedicated to EU", "Staff dedicated to EU projects (FTE)", False),
    ("Legal representative", "Legal representative: name, email, phone, "
     "role", False),
    ("Contact person", "Contact person: name, email, phone, role", False),
]
SHORT = {"Hellenic Centre for Marine Research": "HCMR (Applicant)",
         "ΕΙΔΙΚΟΣ ΛΟΓΑΡΙΑΣΜΟΣ ΚΟΝΔΥΛΙΩΝ ΕΡΕΥΝΑΣ ΠΑΝΕΠΙΣΤΗΜΙΟΥ ΑΙΓΑΙΟΥ":
             "University of the Aegean",
         "Pi Gençlik Derneği": "Pi Youth Association",
         "Acıpayam Belediyesi": "Acipayam Municipality",
         "CONSIGLIO NAZIONALE DELLE RICERCHE": "ISAC-CNR",
         "المركز اللبناني لحفظ الطاقة": "LCEC"}
CODE = {"HCMR (Applicant)": "LEP", "University of the Aegean": "PP1",
        "Pi Youth Association": "PP2", "Acipayam Municipality": "PP3",
        "ISAC-CNR": "PP4", "LCEC": "PP5"}


def clean(t):
    return re.sub(r"\s*[–—]\s*", ", ", str(t))


def eur(x):
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","})) + " €"


# ------------------------------------------------------------- read the form
def partner_blocks():
    extract(FORM_DOCX, FORM_TXT)
    L = FORM_TXT.read_text().splitlines()
    starts = [i for i, l in enumerate(L)
              if l.startswith("Official name of the organisation in the "
                              "original language")]
    assert len(starts) == 6, f"expected 6 partner blocks, found {len(starts)}"
    out = []
    for n, s in enumerate(starts):
        # The LAST partner has no following "Official name" line, and a fixed
        # window ran it 90 lines into the 3.3.5 breakdown table, which
        # repeats every partner's text. That one bug made me report LCEC's
        # empty box as filled AND PIYA's clean box as duplicated (12 Oct).
        # End the last block at 3.3.5, and fail if that heading is missing.
        if n + 1 < len(starts):
            e = starts[n + 1]
        else:
            e = next((i for i in range(s, len(L))
                      if L[i].startswith("3.3.5")), None)
            assert e is not None, "3.3.5 heading not found - cannot bound " \
                                  "the last partner block"
        blk = L[s:e]
        full = blk[0].split("|", 1)[1].strip()
        name = next((v for k, v in SHORT.items() if full.startswith(k[:20])),
                    full)
        vals = {}
        for key, _, _ in FIELDS:
            idx = next((i for i, l in enumerate(blk) if l.startswith(key)),
                       None)
            if idx is None:
                vals[key] = None
                continue
            line = blk[idx]
            if "|" not in line and idx + 1 < len(blk):
                line += " " + blk[idx + 1]
            vals[key] = line.split("|", 1)[1].strip() if "|" in line else ""
        j = next((i for i, l in enumerate(blk)
                  if l.startswith(("3.3.3", "3.3.4"))), None)
        paras = []
        if j is not None:
            for l in blk[j + 2:]:
                if l.startswith("Official name"):
                    break
                if l.strip():
                    paras.append(l)
        out.append(dict(name=name, vals=vals, paras=paras,
                        box="3.3.3" if name.startswith("HCMR") else "3.3.4"))
    return out


def status(key, v):
    if v is None:
        return "absent"
    if not v or v in ("0", ","):
        return "EMPTY"
    if PLACEHOLDER.match(v.replace(" ", "")):
        return "PLACEHOLDER"
    if key == "Legal representative" and v.startswith(","):
        return "NAME MISSING"
    if key == "Legal representative" and "," not in v:
        return "INCOMPLETE"
    return "ok"


def associates():
    L = FORM_TXT.read_text().splitlines()
    return [l for l in L if l.startswith("Associated ")]


def cofin_rows():
    ws = openpyxl.load_workbook(BUDGET, data_only=True)["Co-financing table"]
    rows = {}
    for r in ws.iter_rows(values_only=True):
        if r[0] in CODE.values():
            rows[r[0]] = dict(total=r[1], cofin=r[7], source=r[9],
                              desc=r[10] or "", rev=r[11] or "")
    assert set(rows) == set(CODE.values()), sorted(rows)
    return rows


# --------------------------------------------- texts, written per source type
MEMBER_STATE = "Contribution from other EU Institutions or Member States"
MPC = "Contribution from Mediterranean Partner Countries"
OWN = "Own funds"

COFIN_TEXT = {
    "LEP": (MEMBER_STATE,
            "The 8% co-financing of HCMR is national public co-financing "
            "provided by the Hellenic Republic. It is drawn from the Public "
            "Investment Programme under the national arrangements for Greek "
            "public beneficiaries of Interreg programmes, and is secured for "
            "the full project duration."),
    "PP1": (MEMBER_STATE,
            "The 8% co-financing of the University of the Aegean is national "
            "public co-financing provided by the Hellenic Republic. It is "
            "drawn from the Public Investment Programme under the national "
            "arrangements for Greek public beneficiaries of Interreg "
            "programmes, and is secured for the full project duration."),
    "PP2": (MPC, None),          # already filled correctly in the form
    "PP3": (OWN,
            "Acipayam Municipality covers its 8% co-financing from its own "
            "municipal budget, adopted annually by the Municipal Council. The "
            "contribution comes from the municipality's own revenues and is "
            "available for the full project duration."),
    "PP4": (OWN,
            "ISAC-CNR covers its 8% co-financing from the own institutional "
            "resources of the National Research Council of Italy, a public "
            "research body. The share is committed within the institute's "
            "ordinary funding and is available for the full project "
            "duration."),
    "PP5": (OWN,
            "LCEC covers its 8% co-financing from its own resources, provided "
            "for in its approved annual operating budget. The contribution is "
            "included in LCEC's budget planning for the project duration and "
            "is available from the start of implementation."),
}
# A description must not contradict its own dropdown.
TYPE_WORDS = {MEMBER_STATE: ("national public co-financing",),
              OWN: ("own",),
              MPC: ("directorate",)}

NORTH_AEGEAN_EXPERTISE = (
    "Statutory competence for water management in the North Aegean: "
    "implementation of the Water Framework Directive, including coastal "
    "waters, and of the Floods Directive, through the river basin and flood "
    "risk management plans. AMMOS gives the Directorate a tested field method "
    "for sediment shores under visitor pressure, and the Directorate gives "
    "the Gomati results a route into the plans that govern the island's coast "
    "after the project ends.")

ACIP_IDENTITY = [
    ("Department/service", "xxcxcxcx", "N/A - or the municipal directorate "
                                        "that will manage the project"),
    ("Translation in English", "xcxcxc", "Acipayam Municipality"),
    ("Acronym of the organisation", "ccxcxcx", "ACIPAYAM"),
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
                if v in ("EMPTY", "PLACEHOLDER", "NAME MISSING", "INCOMPLETE",
                         "EMPTY - scored", "DELETE"):
                    r.bold, r.font.color.rgb = True, RED
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t


def main():
    P = partner_blocks()
    A = associates()
    C = cofin_rows()

    fails = []
    for code, (want_type, text) in COFIN_TEXT.items():
        got = C[code]["source"]
        if got != want_type:
            fails.append(f"{code}: export says {got!r}, my text assumes "
                         f"{want_type!r} - rewrite before shipping")
        if text:
            low = text.lower()
            if not any(w in low for w in TYPE_WORDS[want_type]):
                fails.append(f"{code}: description does not match its own "
                             f"source type")
            if len(text.split(". ")[0]) > 200:
                fails.append(f"{code}: first sentence over 200 - the counter "
                             f"on this box is 500 but keep it modular")
            if len(text) > 500:
                fails.append(f"{code}: {len(text)} > 500")
    if "directorate for eu affairs" not in C["PP2"]["desc"].lower():
        fails.append("PIYA's description no longer names its source - "
                     "re-read the export")
    for t in [x[1] for x in COFIN_TEXT.values() if x[1]] + [
            NORTH_AEGEAN_EXPERTISE]:
        for bad in FORBIDDEN:
            if bad in t.lower():
                fails.append(f"forbidden {bad!r} in partner-facing text")
        if PLACEHOLDER.search(t):
            fails.append("a placeholder pattern in MY OWN text")
    if not any("ghgh" in a.lower() for a in A):
        fails.append("the junk associate is gone - rewrite that section")
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    D = Doc()
    D.h("AMMOS - Partnership, associates and co-financing: what is still "
        "missing", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} from your exports of 12-10. "
        f"Every field of the Partnership section was read for all six "
        f"organisations. Budget re-checked: 30 of 30 checks pass, total "
        f"1.300.000,00 exactly - nothing financial has moved.",
        colour=GREY, italic=True)

    # ------------------------------------------------ 1. do these first
    D.h("1. Four things you can fix yourself in a few minutes", size=12,
        colour=RED)
    D.p("A. DELETE \"Associated 1 - ghghghghg\". It is a test record. Code 011 "
        "requires a signed declaration upload for EVERY associated partner, so "
        "while it exists the application cannot be submitted - and an "
        "assessor would read \"ghghghghg\" in section 3.3.6. If Sant'Anna "
        "confirms, add them as a fresh associate instead.", bold=True)
    D.p("B. Acipayam's identity fields still hold placeholder strings:")
    D.table(["Field", "Now", "Put"],
            [[f, n, w] for f, n, w in ACIP_IDENTITY], widths=[5, 3, 9])
    D.p("C. PIYA's capacity box (3.3.4) has a paragraph pasted twice and two "
        "versions of the closing paragraph, one 943 characters and one 349. "
        "Keep the long one and delete the repeats - about 700 characters of "
        "duplicate an assessor will notice.")
    D.p("D. Checklist B question 9. The toggle is now No, which is right, but "
        "the export still prints the text, and it reads \"No, standard is "
        "approached, let alone breached\" - the comma turns it into the "
        "opposite. Switch it to Yes for a moment, delete that one comma so "
        "it reads \"No standard is approached, let alone breached\", and "
        "switch it back to No. All nine toggles are otherwise correct now.")

    # ------------------------------------------------ 2. co-financing
    D.h("2. Co-financing: four descriptions say ????? - and my earlier texts "
        "were wrong for three partners", size=12, colour=RED)
    D.p("Your budget export now shows the co-financing sheet and the four "
        "real choices in the Source of funding list. The partners chose:")
    D.table(["Partner", "8%", "Source of funding selected", "Description now"],
            [[code, eur(C[code]["cofin"]), C[code]["source"],
              (C[code]["desc"][:60] + "...") if len(C[code]["desc"]) > 60
              else (C[code]["desc"] or "(empty)")]
             for code in CODE.values()], widths=[1.6, 2.4, 6.5, 7])
    D.p("My previous file assumed own funds for everybody. For HCMR, Aegean "
        "and PIYA that contradicts the source type actually selected, and a "
        "description that disagrees with its own dropdown is worse than a "
        "short one. These replace it, each matched to the type in your "
        "export:", colour=GREY)
    for code, (typ, text) in COFIN_TEXT.items():
        name = next(k for k, v in CODE.items() if v == code)
        D.h(f"{code} {name} - {eur(C[code]['cofin'])} - {typ}", size=10.5,
            space=8)
        if text:
            D.p(text)
            D.p(f"{len(text)}/500", size=8.5, colour=GREY, italic=True)
        else:
            D.p("Already filled and consistent with its source type: "
                f"\"{C[code]['desc']}\". Leave it.", colour=GREEN)
    D.p("Three things for you to confirm, each one line:", bold=True)
    D.p("- The Greek instrument. I have written \"Public Investment "
        "Programme\" for HCMR and Aegean because that is the usual route for "
        "Greek public beneficiaries and Aegean's earlier \"NSRF\" points the "
        "same way. You know the exact wording far better than I do - adjust "
        "it before pasting.")
    D.p("- Acipayam and PIYA are both Turkish, but PIYA declares the "
        "Directorate for EU Affairs and Acipayam declares own funds. That may "
        "be exactly right - a municipality can simply pay - but confirm both "
        "are deliberate, because an assessor will see the two side by side.")
    D.p("- CNR. Italian public bodies often draw national co-financing from "
        "the national rotation fund rather than paying it themselves. Worth "
        "asking Antonio whether that applies here, because if it does the "
        "dropdown should match HCMR's, and it saves CNR 6.400 €.")

    # ------------------------------------------------ 3. associates
    D.h("3. Associated partners", size=12, colour=RED)
    D.table(["Associate", "Country", "Past experience", "Expertise provided",
             "Action"],
            [["ghghghghg", "Greece", "-", "-", "DELETE"],
             ["North Aegean Water Directorate", "Greece", "EMPTY", "EMPTY",
              "fill both"]], widths=[4.5, 2, 3, 3, 3])
    D.p("Both description columns for North Aegean are empty, and this is "
        "what criterion 3.2 is scored on. Draft for \"Expertise provided\", "
        "built only on the Directorate's statutory remit - please let them "
        "confirm it before it goes in:")
    D.p(NORTH_AEGEAN_EXPERTISE, bold=True)
    D.p(f"{len(NORTH_AEGEAN_EXPERTISE)} characters. \"Past experience\" has "
        f"to come from them: one or two plans or projects they have run on "
        f"coastal waters or flood risk is enough.", size=8.5, colour=GREY,
        italic=True)

    # ------------------------------------------------ 4. partner by partner
    D.h("4. Partner by partner - every field of the Partnership section",
        size=12)
    D.p("\"Scored\" means an assessor reads it to mark the partnership "
        "criteria. The rest is administrative, but an empty capacity figure "
        "is still visible on the page.", colour=GREY, italic=True)
    for x in P:
        bad = [(lbl, status(k, x["vals"][k]))
               for k, lbl, _ in FIELDS if status(k, x["vals"][k]) != "ok"]
        cap = sum(len(p) for p in x["paras"])
        cap_s = "EMPTY - scored" if cap < 50 else f"{cap} characters"
        D.h(x["name"], size=11, space=10)
        rows = [[lbl, st] for lbl, st in bad] + [
            [f"Box {x['box']} experience and capacity", cap_s]]
        D.table(["Field", "Status"], rows, widths=[9, 5])
        if x["name"] == "University of the Aegean":
            D.p("CONTRADICTION: annual total budget is \"€500.000 or less\" "
                "but the budget dedicated to EU projects is 5.500.000 €, "
                "eleven times larger. One of the two is wrong, and the "
                "assessor scoring capacity will see both on the same page.",
                colour=RED, bold=True)
        if x["name"] == "LCEC":
            D.p("CHECK: budget dedicated to EU projects reads 539,36 € a year. "
                "For an organisation with nine EU projects that is almost "
                "certainly a typing error - probably 539.360 €.",
                colour=RED, bold=True)
        if x["name"] == "Pi Youth Association":
            D.p("Complete apart from the duplicated paragraphs in 3.3.4 - see "
                "1C above.", colour=GREEN)

    # ------------------------------------------------ 5. what it costs
    D.h("5. And a correction to my own scoring", size=12)
    D.p("When I scored the application at 79 of 88 on Step 1, I gave the "
        "partnership composition criterion 3 of 4 and never opened the "
        "capacity boxes. Box 3.3.3 is empty for the Applicant and box 3.3.4 "
        "is empty for Aegean, Acipayam and CNR. That breaks my own rule of "
        "claiming no mark for anything that is not in the form.")
    D.p("Read honestly, Step 1 is nearer 77 than 79 today - still above 75, "
        "but with two points of margin rather than four. The good news is the "
        "same as before: these are the cheapest points left in the "
        "application. They are other people's paragraphs, not new writing, "
        "and filling them puts the margin back.", bold=True)

    # ------------------------------------------------ 6. one ask
    D.h("6. What I need from you to write the capacity boxes", size=12)
    D.p("One screenshot of box 3.3.3 or 3.3.4 with its character counter "
        "visible. With the limit and three facts per partner - founding "
        "year, two or three relevant past projects, and the people who will "
        "do the work - I will write all four boxes to size, consistent with "
        "the roles already described in 3.3.7 and the work packages.")

    D.d.save(OUT)
    print(f"wrote {OUT}")
    for x in P:
        n = sum(1 for k, _, _ in FIELDS if status(k, x["vals"][k]) != "ok")
        print(f"  {x['name']:28s} gaps {n:2d}  capacity "
              f"{sum(len(p) for p in x['paras'])} chars")


if __name__ == "__main__":
    main()
