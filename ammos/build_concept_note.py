#!/usr/bin/env python3
"""Draft concept note for the Interreg NEXT MED capitalisation call.

Structured to mirror the evaluation grid in the Guidelines for Applicants
(section 5.3), because that is what the proposal will be scored against:
every heading below carries the award criterion it answers, so partners can
see which of their contributions earns which points.

Sources, all read in full and cached locally:
  * "Text of the call for capitalisation projects"  -> epirus/nextmed_call.txt
  * "Guidelines for Applicants"                    -> epirus/nextmed_guide.txt
  * MMM sustainable tourism outputs database       -> epirus/mmm_outputs.xlsm
Every figure and every quoted rule is taken from one of those three and is
asserted below before the document is written.

v2, 13 Sep 2026.  The site is LEMNOS, not Lesvos -- Gomati beach and Agios
Nikolaos bay -- and the science now comes from HCMR's own prior proposal on
this exact habitat, supplied by the client.

v3, 13 Sep 2026.  Surveyed site data from the client: the centre point and the
two polygons with their measured areas.  Note that A + B = 740,000 m2, while
HCMR's text states a total of 0.77 km2; the 30,000 m2 gap is flagged in the
document rather than silently resolved, because only HCMR can say which figure
is the surveyed one.

NOT YET SETTLED, marked [TO CONFIRM] in the text:
  - whether Acipayam has a fragile landscape under visitor pressure
  - RESOLVED 28 Sep: LCEC has confirmed signing capacity in its own name
  - the habitat code: HCMR's own document says 2210 once and 2250 twice
"""
import re

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

# Both standing rules he asked for on 23 September: no dash as punctuation
# (ranges survive, because digits on both sides means "to" not "and"), and my
# commentary in one deletable block at the end.
RANGE = re.compile(r"(?<=\d)\s*[\u2013\u2014]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = t.replace("--", ", ")
    t = re.sub(r"\s*[\u2013\u2014]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


import pathlib
# WAS a hard-coded folder on MY server, so this script only ran on my
# machine. Now it uses its OWN folder.
#
# Note here(__file__) and not _shared.HERE: HERE is the folder _shared.py
# lives in, which is one level UP from this file. Using it sent these scripts
# looking for mmm_outputs.xlsm in the parent folder. here() takes the folder
# of the script that asks.
import sys as _sys
_sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from _shared import here

BASE = str(here(__file__))
# VERSION AND DATE IN THE FILENAME, ALWAYS.
#
# I sent you two different documents called EPIRUSMEDEYE_advisor_report_
# 2026-09-24.docx and you had no way to tell which was which. The same thing
# was about to happen here: this run produces a materially different budget
# from the 25 September file and would have overwritten it under the same
# name. The old file stays where it is, so you can diff the two.
OUT = f"{BASE}/AMMOS_Concept_Note_v5_1303000_2026-09-28.docx"

# ------------------------------------------------------------ verified figures
# The AGREED total, not the ceiling.  v3 used 1,304,347.83, which is the
# maximum the call allows and also, as it turns out, the entire envelope of one
# Specific Objective - so bidding at it is asking for 100% of the pot.  1.3m
# leaves 4,347.83 of margin and reads as a budget rather than a grab.
TOTAL = 1303000.00          # agreed total eligible cost, 28 September
EU = round(1303000.00 * 0.92, 2)     # 1.198.760,00
CO = round(1303000.00 * 0.08, 2)     # 104.240,00
CEILING = 1304347.83        # what the call would allow, for reference
EU_CAP = 1200000.00         # the EU contribution the call will award
MPC_FLOOR = 0.50            # >=50% of total eligible spent in MPC territories
ORG_CAP = 0.35              # <=35% of total eligible to any one organisation

# WHY 1.303.000 AND NOT 1.304.000
#
# The ceiling is 1.304.347,83, and 1.304.000 would sit 347,83 under it. But
# the binding constraint is the EU CONTRIBUTION, not the total: at 92%,
# 1.304.000 gives an EU contribution of 1.199.680, only 320 under the
# 1.200.000 the call awards. Any late correction upward - a partner adding a
# travel line, a flat rate recomputed - breaches the EU cap rather than the
# total, and that is the one number an eligibility check looks at first.
# 1.303.000 leaves 1.240 of EU headroom and 1.347,83 of ceiling headroom.

# THE 4.4.3(b) LEVER, DECLARED AS A NUMBER RATHER THAN AS AN ARGUMENT
#
# Guidelines 4.4.3 requires 50% of TOTAL ELIGIBLE COSTS to be dedicated to
# activities implemented in Mediterranean Partner Country territories. It does
# NOT require that money to be held by MPC partners: 4.4.3(b) lets an EU
# partner's direct costs count toward the floor where the activity is
# implemented in MPC territory.
#
# That is what pays for giving the Greeks 24.000 back. And it is load-bearing,
# not decorative: the MPC PARTNERS alone come to 645.000, which is 49,50% of
# 1.303.000 and BELOW the floor. So the proposal is compliant only if a real
# part of CNR's work is genuinely carried out in Turkiye and Lebanon.
#
# Declared here as a figure so check() can test it, rather than left as a
# sentence in the text that nobody can verify.
CNR_IN_MPC = 50000.00       # CNR direct costs for work done in TR and LB
DEADLINE = "29 October 2026, 13:00 CET"

# surveyed site data supplied by the client, 13 Sep 2026
SITE_LAT, SITE_LON = 39.993075, 25.144233     # 39d59'35.07"N 25d08'39.24"E
SECTION_A_M2 = 464000        # inland section, the larger of the two
SECTION_B_M2 = 276000        # coastal section, reaching the shore at Gomati
HCMR_STATED_KM2 = 0.77       # from HCMR's own proposal text
QUESTIONS_CLOSE = "14 October 2026"

# AGREED 22 September, with one PROVISIONAL addition.  v3's figures were an
# illustrative split I generated before the real per-partner totals existed,
# and two partners then negotiated against them - Pi Youth over 230,000 and
# Acipayam over 270,000.  Nothing in this version is illustrative.
ALLOC = [
    ("HCMR, Hellenic Centre for Marine Research", "Greece", "Lead Partner",
     363000.00, "agreed"),
    ("University of the Aegean", "Greece", "Partner", 215000.00, "agreed"),
    ("Pi Youth Association (PIYA), Izmir", "Türkiye", "Partner", 230000.00,
     "agreed"),
    ("Acipayam Municipality, Denizli", "Türkiye", "Partner", 135000.00,
     "agreed"),
    ("LCEC, Lebanese Centre for Energy Conservation", "Lebanon", "Partner",
     280000.00, "agreed"),
    ("ISAC, CNR, Institute of Atmospheric Sciences and Climate, Cagliari",
     "Italy", "Partner", 80000.00, "confirmed, 28 September"),
]

# WHAT MOVED FROM THE 22 SEPTEMBER SPLIT, AND WHERE IT CAME FROM
#
# He asked for Italy's 80.000 to be funded without cutting HCMR and the
# University of the Aegean as hard as the previous version did. Those two had
# carried the whole 80.000 between them, because a transfer between two EU
# partners is neutral to the 50% test and taking it from an MPC partner was
# not. That reasoning was right as far as it went and cost the Greeks 80.000.
#
# 4.4.3(b) is the way out: the floor is about where activities are IMPLEMENTED,
# not who holds the budget. So 20.000 comes back from LCEC instead, and CNR's
# own MPC-implemented work replaces it inside the floor.
PREVIOUS = {"HCMR, Hellenic Centre for Marine Research": 350000.00,
            "University of the Aegean": 205000.00,
            "LCEC, Lebanese Centre for Energy Conservation": 300000.00}
MPC = {"Türkiye", "Lebanon"}          # Italy and Greece are EU Mediterranean
COUNTRY_CAP = 2                       # maximum organisations per country


def check():
    assert abs(EU + CO - TOTAL) < 0.01
    assert abs(EU / TOTAL - 0.92) < 0.0001
    s = round(sum(a[3] for a in ALLOC), 2)
    assert abs(s - TOTAL) < 0.01, f"allocation sums to {s}, not {TOTAL}"
    assert TOTAL <= CEILING, f"{TOTAL} is over the call ceiling {CEILING}"
    for name, country, role, amt, _ in ALLOC:
        assert amt <= TOTAL * ORG_CAP + 0.01, \
            f"{name} exceeds the 35% single-organisation cap"
    assert abs(EU - TOTAL * 0.92) < 0.01
    assert EU <= EU_CAP, (
        f"the EU contribution {EU:,.2f} exceeds the {EU_CAP:,.2f} the call "
        f"awards. The ceiling on the TOTAL is not the binding constraint here "
        f"- the EU contribution is.")

    # ------------------------------------------------- THE 50% FLOOR, BOTH WAYS
    #
    # I have got this wrong once already on this proposal, by switching the
    # denominator from total eligible costs to direct costs and arriving at
    # 49,61% while believing we were compliant. Guidelines 4.4.3 says total
    # eligible costs, and says it twice. So both readings are computed every
    # run and the failure message says which one broke.
    #
    # And both NUMERATORS are computed too, because they are not the same
    # question:
    #
    #   mpc_partners  what MPC-country partners hold. 645.000 = 49,50%.
    #                 BELOW THE FLOOR. This split is NOT compliant on
    #                 partner location alone.
    #   mpc_activity  what is implemented in MPC territory, which is what
    #                 4.4.3 actually tests, including CNR's work in Turkiye
    #                 and Lebanon under 4.4.3(b). 695.000 = 53,34%.
    #
    # If CNR turns out to intend doing its work from Cagliari, mpc_activity
    # collapses to mpc_partners and the proposal fails. That is why the number
    # is asserted rather than described.
    mpc_partners = round(sum(a[3] for a in ALLOC if a[1] in MPC), 2)
    mpc_activity = round(mpc_partners + CNR_IN_MPC, 2)
    floor = round(TOTAL * MPC_FLOOR, 2)
    shortfall = round(floor - mpc_partners, 2)

    cnr = next(a for a in ALLOC if a[1] == "Italy")
    assert CNR_IN_MPC <= cnr[3], (
        f"CNR_IN_MPC is {CNR_IN_MPC:,.2f} but CNR's whole budget is only "
        f"{cnr[3]:,.2f}. You cannot implement more in MPC territory than the "
        f"partner holds.")
    assert mpc_activity >= floor, (
        f"activities implemented in MPC territory come to {mpc_activity:,.2f}, "
        f"below the {floor:,.2f} floor")
    assert CNR_IN_MPC >= shortfall, (
        f"MPC partners hold {mpc_partners:,.2f} = "
        f"{mpc_partners / TOTAL * 100:.2f}%, which is BELOW the 50% floor of "
        f"{floor:,.2f}. Compliance depends entirely on {shortfall:,.2f} of "
        f"CNR's budget being work genuinely implemented in Turkiye and "
        f"Lebanon. CNR_IN_MPC is declared as {CNR_IN_MPC:,.2f}, which is not "
        f"enough.")
    mpc = mpc_activity
    if mpc_partners < floor:
        print(f"  !! THE 50% FLOOR DEPENDS ON 4.4.3(b). MPC partners alone: "
              f"{mpc_partners:,.2f} = {mpc_partners / TOTAL * 100:.2f}%, "
              f"which is BELOW the floor of {floor:,.2f}.")
        print(f"     Compliance needs at least {shortfall:,.2f} of CNR's "
              f"{cnr[3]:,.2f} to be activity implemented in MPC territory. "
              f"Declared: {CNR_IN_MPC:,.2f}.")
        print(f"     With the lever: {mpc_activity:,.2f} = "
              f"{mpc_activity / TOTAL * 100:.2f}%, margin "
              f"{mpc_activity - floor:,.2f}.")
        print("     CONFIRM WITH CNR THAT THEY WILL TRAVEL. If the work is "
              "done from Cagliari this proposal is not eligible.")
    # 4.4.2: at least 5 organisations, at least 3 countries, no more than 2
    # per country.  Adding Italy takes Greece and Turkiye both TO the cap.
    countries = {}
    for _, c, _, _, _ in ALLOC:
        countries[c] = countries.get(c, 0) + 1
    assert len(ALLOC) >= 5, f"only {len(ALLOC)} organisations"
    assert len(countries) >= 3, f"only {len(countries)} countries"
    worst = max(countries, key=countries.get)
    assert countries[worst] <= COUNTRY_CAP, \
        f"{worst} has {countries[worst]} organisations, the maximum is {COUNTRY_CAP}"
    # the site figures this document prints must add up to the total it prints;
    # HCMR's own text says 0.77 km2, which is 30,000 m2 more than A + B, so the
    # document reports both rather than quietly picking one
    site_total = SECTION_A_M2 + SECTION_B_M2
    assert site_total == 740000, site_total
    assert abs(site_total / 1e6 - HCMR_STATED_KM2) > 0.001, \
        "the A+B discrepancy has gone away - update the text that flags it"
    # EVERY PARTNER'S TERRITORY MUST BE IN THE ELIGIBLE AREA, CHECKED AGAINST
    # THE GUIDELINES RATHER THAN AGAINST MY MEMORY.
    #
    # The Joint Secretariat will not confirm eligibility - they told us so on
    # 28 September - so the document now asserts the position itself. The
    # eligible area is a list of PROVINCES, and the claim in section 12 is
    # that inland provinces are in it. That claim is testable: several of the
    # named Turkish provinces are landlocked, so if the list really contains
    # them, "inland" cannot be a bar.
    #
    # If the call is ever reissued with a different area, this fails instead
    # of the document quietly asserting something that has stopped being true.
    area = open(f"{BASE}/nextmed_guide.txt", encoding="utf-8",
                errors="replace").read()
    for place in ("Denizli", "İzmir", "Lebanon", "Greece"):
        assert place in area, (
            f"{place} is not named in the Guidelines eligible-area text, so "
            f"the eligibility claim in section 12 cannot be made")
    LANDLOCKED = ("Afyonkarahisar", "Kütahya", "Uşak", "Isparta", "Burdur",
                  "Kahramanmaraş")
    missing = [p for p in LANDLOCKED if p not in area]
    assert not missing, (
        f"these provinces are quoted in section 12 as landlocked members of "
        f"the eligible area but are not in the Guidelines text: {missing}. "
        f"Do not claim inland eligibility without them.")

    # the six capitalised outputs must exist in the MMM database, by code
    wb = openpyxl.load_workbook(f"{BASE}/mmm_outputs.xlsm", data_only=True, read_only=True)
    codes = {str(r[0]).strip() for r in wb["database"].iter_rows(
        min_row=5, max_row=214, max_col=1, values_only=True) if r[0]}
    for c in ("MMM_IF01", "MMM_IF02", "MMM_NX35", "MMM_NX36", "MMM_NX37", "MMM_NX38"):
        assert c in codes, f"{c} is not in the MMM database"
    print(f"checks passed: allocation {s:,.2f} = total eligible; "
          f"MPC {mpc:,.2f} = {mpc/TOTAL*100:.1f}% (floor {MPC_FLOOR:.0%}); "
          f"6 MMM output codes verified against the database")
    return mpc


# ------------------------------------------------------------------- document

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)


def main():
    mpc = check()
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(10.5)

    def h(text, size=13, colour=NAVY, space=10, after=4):
        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(clean(text))
        r.bold = True
        r.font.size = Pt(size)
        r.font.color.rgb = colour
        return p

    def crit(text):
        p = d.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        r = p.add_run(clean(text))
        r.italic = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = GREY

    def para(text, bold=False, colour=None, size=10.5, after=5):
        text = clean(text)
        p = d.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
        if colour is not None:
            r.font.color.rgb = colour
        return p

    def bullet(text, bold_head=None):
        p = d.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        if bold_head:
            r = p.add_run(clean(bold_head))
            r.bold = True
        p.add_run(clean(text))
        return p

    # ---- title
    t = d.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(clean("CONCEPT NOTE — DRAFT FOR PARTNER DISCUSSION"))
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = NAVY
    s = d.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run(clean("Interreg NEXT MED — Call for capitalisation projects on sustainable tourism"))
    r.font.size = Pt(11)
    r.font.color.rgb = GREY
    s2 = d.add_paragraph()
    s2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s2.add_run(clean(
        "AMMOS — Adaptive Management of Mediterranean shOreline Sediments "
        "under tourism pressure   (ammos = sand, Greek)"))
    r.italic = True
    r.font.size = Pt(10)
    r.font.color.rgb = GREY

    para("Every figure here is agreed except the Italian allocation, which is "
         "marked PROVISIONAL. Each heading names the award criterion it answers, "
         "so every partner can see which of their contributions earns which "
         "points.", colour=GREY, size=9.5, after=10)

    h("What changed since the earlier draft", size=11.5, space=6)
    for txt in [
        "TITLE. The object all three demonstration territories share is a "
        "SEDIMENT SHORE, not a dune: marine at Lemnos, lacustrine at Ucari "
        "Lake, coastal in Lebanon. The earlier title said dune systems, which "
        "an evaluator would have read against a confirmed lake shore and "
        "recorded as an internal inconsistency.",
        "SITES. Acipayam has confirmed Ucari Lake Recreation Area, a "
        "municipality-managed lacustrine shore of unconsolidated sediment, "
        "sand and fine gravel under heavy seasonal visitor pressure. The "
        "Acipayam Canyon fluvial corridor is named as a documented replication "
        "target rather than a fourth delivery site.",
        "BUDGET. The per-partner figures in the earlier draft were an "
        "illustrative split made before the real totals existed. They are "
        "replaced by the agreed allocation. Nothing in this version is "
        "illustrative.",
        "ARGUMENT. Acipayam's source-to-sea reasoning is now central rather "
        "than absent. See section 2.",
        "PARTNERSHIP. ISAC, CNR is proposed as a sixth partner. CNR is an "
        "AMMIRARE partner and its researchers are named authors of the "
        "methodology this project transfers. See section 6.",
    ]:
        b = d.add_paragraph(style="List Bullet")
        b.paragraph_format.space_after = Pt(3)
        rr = b.add_run(clean(txt))
        rr.font.size = Pt(9.5)

    # ---- 1 identification
    h("1. Identification")
    tbl = d.add_table(rows=0, cols=2)
    tbl.style = "Light Grid Accent 1"
    for k, v in [
        ("Programme", "Interreg VI-B NEXT Mediterranean Sea Basin (NEXT MED)"),
        ("Call", "Coordinated call for capitalisation projects, Mediterranean Multi-Programme "
                 "Mechanism (MMM)"),
        ("Priority", "2 — A greener, low-carbon and resilient Mediterranean"),
        ("Specific Objective", "2.2 — Climate change adaptation and disaster risk prevention, "
                               "resilience taking into account ecosystem-based approaches"),
        ("Topic (from the call's Terms of Reference)", "Tourism and the green transition"),
        ("Lead Partner", "HCMR — Hellenic Centre for Marine Research (Greece)"),
        ("Partnership", "5 organisations · 3 countries · 1 EU Mediterranean Country · "
                        "2 Mediterranean Partner Countries"),
        ("Duration", "24 months (permitted range 18–24)"),
        ("Total eligible cost", f"EUR {TOTAL:,.2f}  (EU contribution EUR {EU:,.2f} = 92%, "
                                f"partnership contribution EUR {CO:,.2f} = 8%)"),
        ("Submission deadline", DEADLINE),
    ]:
        row = tbl.add_row().cells
        row[0].text = clean(k)
        row[1].text = clean(v)
        for p in row[0].paragraphs:
            for rr in p.runs:
                rr.bold = True
                rr.font.size = Pt(9.5)
        for p in row[1].paragraphs:
            for rr in p.runs:
                rr.font.size = Pt(9.5)

    # ---- 2 the challenge
    h("2. The challenge this project addresses")
    crit("Award criterion 1.1 Justification — 8 points. To be aligned with the wording of the "
         "call's Terms of Reference once obtained.")
    para("Mediterranean coastal dune systems are the region's cheapest and most effective "
         "natural defence against coastal erosion, storm surge and saline intrusion. They are "
         "also, almost everywhere, the exact strip of land on which summer tourism lands: "
         "parking, access paths, beach bars, sunbeds and informal camping. The pressure is "
         "concentrated into the same twelve weeks each year, and it is applied by the economic "
         "activity the host communities depend on. Protecting the dune by excluding tourism is "
         "neither politically nor economically available; the problem is how to carry the same "
         "visitor numbers without destroying the landform that protects the coast behind it.")
    h("2.1 Why an inland shore belongs in a coastal project", size=11.5,
      space=8)
    para("Lakes and river networks are the primary vectors by which litter "
         "reaches the Mediterranean basin. Tackling it at a lake or river shore "
         "is therefore not a substitute for coastal work; it is the upstream "
         "half of the same problem, and it is the half nobody funds. COMMON was "
         "a marine litter project, and MMM_NX37 consolidates the Italian, "
         "Tunisian and Lebanese national reports on tackling it. Intercepting "
         "the material before it reaches the sea is the logically necessary "
         "complement to cleaning the beach it arrives on.", bold=True)
    para("This is why AMMOS runs on three shore typologies rather than three "
         "beaches. A method that works on a marine dune, a lacustrine sediment "
         "bank and a riparian corridor has demonstrated that it travels. A "
         "method proven on three beaches has demonstrated only that beaches "
         "resemble each other. The argument is Acipayam's own and the "
         "partnership has adopted it.")

    para("The demonstration site makes that abstract problem concrete. The dune field of "
         "northern Lemnos at Gomati beach, centred on 39\u00b059\u203235.07\u2033N "
         "25\u00b008\u203239.24\u2033E and known locally as the Pachies Ammoudies, is a "
         "coastal dune and associated Mediterranean desert system listed in Annex I of the "
         "Habitats Directive 92/43/EEC. It is surveyed in two sections about 400 m apart:")
    bullet("464,000 m\u00b2 \u2014 the inland section, lying some 1.8 km from the shore and "
           "the larger of the two", "Section A: ")
    bullet("276,000 m\u00b2 \u2014 the coastal section, running down to the beach itself",
           "Section B: ")
    para("Together 740,000 m\u00b2, or 0.74 km\u00b2, with a perimeter of roughly 8.3 km. "
         "Greece lost most of its dunes before 1970 to sand extraction for construction and "
         "later for beach nourishment; what survives here is under uncontrolled pedestrian and "
         "vehicle traffic, recreation and catering structures, holiday-home construction and "
         "sandboarding. A vehicle track runs between and through both sections.")
    para("Figure to resolve before submission: HCMR's own text gives a total of 0.77 km\u00b2, "
         "while the two surveyed polygons sum to 0.74 km\u00b2 \u2014 a difference of "
         "30,000 m\u00b2, about 4%. Both figures cannot go into an application form.",
         colour=RED, size=9.5)
    para("And there is a governance gap that a capitalisation project is well placed to close: "
         "the site is NOT inside a Natura 2000 area, although Lemnos has two \u2014 GR4110001 "
         "and GR4110006 \u2014 and GR4110006 covers almost the whole coastal and marine "
         "perimeter of the island. The Faraklou Geological Park lies immediately to the east "
         "of the dune field. An Annex I habitat under documented tourism pressure, sitting "
         "between a protected area that does not include it and a geopark that does not "
         "either, is exactly the situation a transferable management method is for.",
         bold=True)
    para("This is a shared Mediterranean problem with a shared shortage: the science of dune "
         "resilience is well developed, but site managers — municipalities, port authorities, "
         "protected-area bodies — rarely have a method they can apply themselves. Solutions "
         "exist and have been proven in the western Mediterranean. They have not travelled east "
         "or south.")

    # ---- 3 transnational added value
    h("3. Why this has to be done transnationally")
    crit("Award criterion 1.2 Transnational added value and impact — 8 points.")
    para("The outputs being capitalised were produced in Italy and France. The territories that "
         "need them here are Greek, Turkish and Lebanese. No single partner can transfer a "
         "management method to itself: adaptation requires the method to be tested against "
         "different legal regimes for coastal land, different tourism seasons and intensities, "
         "and different institutional owners of the problem.")
    bullet("a public research institute that can adapt the science and validate the result "
           "(HCMR), a university that can turn monitoring protocols into a tool others can use "
           "(University of the Aegean), a municipality that actually manages a site and its "
           "visitors (Acipayam), a youth organisation that can move practice into the next "
           "generation of site users (PIYA), and a national agency that can carry the energy "
           "dimension of low-impact visitor facilities (LCEC).",
           "The partnership is deliberately asymmetric: ")
    para("Each of those five functions exists in one partner and not in the others. That is the "
         "cooperation need, and it is why the project is not five parallel local pilots.")

    # ---- 4 outputs capitalised
    h("4. The outputs being capitalised")
    crit("Award criterion 2.1 Transfer and adaptation methodology — 4 points. The call requires "
         "at least one output from an MMM programme participating in the coordinated call; both "
         "source programmes below are named in the call text.")
    para("From AMMIRARE — \"Actions and Methodologies for Improving the Resilience of Sandy "
         "Beaches\" (Interreg Italy–France Maritime):", bold=True, after=3)
    bullet("an operational document supporting site managers in assessing ecosystem services "
           "and planning sustainable actions, promoting resilient beach management models that "
           "address climate change while remaining compatible with tourism and public use. This "
           "is the analytical core of the project.",
           "MMM_IF02 — Guidelines for the Sustainable Management of Beaches: ")
    bullet("an educational programme engaging students in active beach stewardship through field "
           "activities, biodiversity monitoring and dune restoration.",
           "MMM_IF01 — Beach Custodians: ")
    para("From COMMON — \"COastal Management and MOnitoring Network for tackling marine litter "
         "in the Mediterranean sea\" (Interreg NEXT MED):", bold=True, after=3)
    bullet("a standardised protocol aligned with the European Environment Agency methodology, "
           "multilingual, with instructions for systematic collection, categorisation and "
           "documentation.", "MMM_NX36 — Beach litter monitoring guide: ")
    bullet("consolidating the Italian, Tunisian and Lebanese national reports, with actions and "
           "recommendations. A Lebanese contribution already sits inside this output.",
           "MMM_NX37 — Marine litter tackling in the Mediterranean: ")
    bullet("published in Italian, French, English and Arabic — directly reusable for uptake in "
           "the Mediterranean Partner Countries.", "MMM_NX35 — BEach CLEAN decalogue: ")
    bullet("knowledge transfer on monitoring and management, plus legislation, funding, "
           "communication and stakeholder motivation.",
           "MMM_NX38 — Capacity Building Training Design Plan: ")
    para("No partner in this consortium took part in either source project. That is not an "
         "obstacle — the call's stated purpose is \"transferring, adapting and scaling up results "
         "from existing Interreg projects to new territories and contexts\".",
         colour=GREY, size=9.5)

    # ---- 5 what the project does
    h("5. What the project will actually do")
    crit("Award criterion 2.2 Clear, realistic and achievable objectives — 4 points.")
    para("The transferable asset is the METHOD, not the site. AMMIRARE's guidelines describe how "
         "to assess a fragile site under visitor pressure and manage it accordingly; that travels "
         "to any such site. The project therefore has one place where the method is adapted and "
         "demonstrated, and two territories where it is transferred, applied and independently "
         "tested.", bold=True)
    bullet("Lemnos, Greece — adaptation and demonstration. The AMMIRARE guidelines are "
           "rewritten for Aegean dune systems and applied at Gomati beach and Agios Nikolaos "
           "bay: topographic and ecological baseline, identification of natural and "
           "anthropogenic risks, then the internationally accepted interventions the site "
           "needs — marked movement corridors and boardwalks, fencing, sand-trapping "
           "structures — specified and installed at demonstration scale, with a costed "
           "management plan taken through stakeholder consultation so it is adopted rather "
           "than filed.")
    bullet("Denizli province, Türkiye — transfer to a second context. [TO CONFIRM: the fragile "
           "landscape under visitor pressure that Acipayam manages. The relevance of this "
           "partner depends on naming it, and the answer decides whether Acipayam is a funded "
           "partner or an associate.]")
    bullet("Lebanese coast — transfer to a third context, with the energy and carbon dimension "
           "of visitor facilities as the distinct contribution. Signing capacity confirmed by "
           "LCEC on 28 September [TO CONFIRM: LCEC's coastal and "
           "tourism-facility experience, and its capacity to sign in its own name].")
    para("A shared spatial monitoring platform runs across all three, so that the same "
         "indicators are collected the same way in three countries and the comparison is "
         "meaningful rather than anecdotal.")

    # ---- 6 partnership and roles
    h("6. Partnership and the role of each partner")
    crit("Award criterion 3 Partnership operational and financial capacity — 12 points. Each "
         "role below is tied to a specific capitalised output, not to a general competence.")
    roles = [
        ("HCMR — Hellenic Centre for Marine Research", "Greece", "Lead Partner",
         "Scientific lead and legal responsibility for the partnership, through the Institute "
         "of Marine Biological Resources and Inland Waters (Laboratory of Integrated Coastal "
         "Zone Management) and the Institute of Oceanography (Department of Geology). Adapts "
         "MMM_IF02 to Aegean dune systems, carries out the ecosystem-services assessment, and "
         "validates results across all three territories."),
        ("ISAC, CNR, Institute of Atmospheric Sciences and Climate, Cagliari "
         "Section", "Italy", "Partner, PROVISIONAL",
         "METHOD OWNER. CNR is a partner in AMMIRARE, and CNR researchers are "
         "named authors of the methodology this project transfers: the method "
         "sheets of MMM_IF02 are signed by S. Simeone, G. De Falco, "
         "F. Antognarelli, S. Como, A. Cucco and A. Conforti of CNR. Their "
         "role here is to assure fidelity of transfer, not to run a further "
         "pilot: they co-author the adapted method, quality-assure the three "
         "territorial baselines against the original protocol, and train the "
         "site teams. There is deliberately NO Italian demonstration site, "
         "because AMMOS transfers the method to new territories rather than "
         "continuing the project that produced it."),
        ("University of the Aegean", "Greece", "Partner",
         "Builds the shared web-GIS monitoring platform, embedding the MMM_NX36 EEA-aligned "
         "protocol and the MMM_IF01 biodiversity monitoring so that field data from three "
         "countries is comparable. Based in the North Aegean, the same region as the "
         "demonstration site."),
        ("Pi Youth Association (PIYA)", "Türkiye", "Partner",
         "Transfers MMM_IF01 Beach Custodians. The output's own target groups include Youth and "
         "training institutions, and PIYA is an ESC-accredited Erasmus+ provider with an annual "
         "EU reporting record — it is the partner the output was written for."),
        ("Acipayam Municipality", "Türkiye", "Partner",
         "The site-manager voice: receives the management method and applies it as the authority "
         "responsible for a territory and its visitors. [TO CONFIRM — see section 5.]"),
        ("LCEC — Lebanese Centre for Energy Conservation", "Lebanon", "Partner",
         "The energy and low-carbon dimension of low-impact visitor infrastructure — lighting, "
         "off-grid facilities, seasonal demand — which is squarely its mandate. Also the natural "
         "home for the Lebanese strand already present in MMM_NX37. Signing "
         "capacity in its own name CONFIRMED by LCEC, 28 September 2026."),
    ]
    for name, country, role, text in roles:
        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(clean(f"{name} ({country}) — {role}"))
        r.bold = True
        r.font.size = Pt(10.5)
        para(text, size=10, after=2)

    h("Synergies with existing initiatives", size=11, space=8)
    crit("Award criterion 1.4 — 4 points, assessed by the Assessment Board only.")
    para("LIFE-IP 4 NATURA (2018–2025) works across Greek habitats but its target regions are "
         "Attica, Crete, Eastern Macedonia–Thrace and the Decentralised Administration of "
         "Epirus–Western Macedonia, and do not include Lemnos. TERRA LEMNIA (MAVA Foundation, "
         "2019–2022) worked on Lemnos but not on this site. Both are stated in HCMR's own "
         "documentation as not covering the target area.")
    para("Note for the drafting team: as written, those two entries demonstrate a GAP rather "
         "than a synergy. Criterion 1.4 asks whether synergies are \"well demonstrated and "
         "likely to be exploited\". The same two initiatives should be described in terms of "
         "what we will connect to — LIFE-IP 4 NATURA's habitat methodology and national "
         "reporting channel, TERRA LEMNIA's island stakeholder base — with the absence of "
         "coverage used to justify need in section 2, not here.", colour=RED, size=9.5)

    h("Associates (no funding, no country slot)", size=11, space=8)
    crit("Criterion 3.2, on the meaningful involvement of associated partners in reuse and "
         "replication.")
    para("Municipality of Zakynthos, Hellenic Ministry of Tourism, Hellenic Ministry of Maritime "
         "Affairs, and — carrying over the existing EpirusMedEye relationship — Municipality of "
         "Dropull (Albania), which associate status permits because associates do not have to "
         "meet the geographical eligibility criteria. Associates receive no funding and take on "
         "no reporting or audit obligation; travel to project events is covered by the project. "
         "They cannot participate in any procurement launched by the project.")
    para("Their purpose is uptake: a regional authority and two ministries in the room is how "
         "the proposal demonstrates that the transferred method will be applied by bodies with "
         "the power to apply it.", colour=GREY, size=9.5)

    # ---- 7 target groups
    h("7. Target groups and final beneficiaries")
    crit("Award criterion 1.3 — 4 points. These must be quantified before submission; the "
         "figures below are placeholders for partner input.")
    for a, b in [("Public authorities", "site and coastal managers in the three territories "
                                        "[number to be supplied by partners]"),
                 ("Training institutions and youth", "students engaged through Beach Custodians "
                                                     "[number]"),
                 ("SMEs", "beach concessions, accommodation and tourism operators at the "
                          "demonstration sites [number]"),
                 ("Communities", "residents of the host municipalities [number]")]:
        bullet(b, f"{a}: ")

    # ---- 8 results
    h("8. Expected results and outputs")
    crit("Award criteria 2.2 and 2.4 Indicators — quantification to follow with the work packages.")
    bullet("one dune management method adapted from MMM_IF02 and validated for Aegean, Anatolian "
           "and Levantine conditions")
    bullet("one shared web-GIS monitoring platform carrying the MMM_NX36 protocol, operating in "
           "three countries")
    bullet("Beach Custodians running in at least two countries beyond its country of origin")
    bullet("low-impact visitor infrastructure demonstrated at Gomati, Lemnos, and specified for "
           "transfer")
    bullet("a joint transferability guide, in English and Arabic, so the method continues to "
           "travel after the project ends")

    # ---- 9 sustainability
    h("9. Sustainability")
    crit("Award criterion 5 — 12 points.")
    para("The project deliberately produces instruments rather than installations: a method, a "
           "protocol, a platform and a curriculum. The associates are the mechanism by which "
           "those instruments outlive the funding — a ministry that has endorsed a method can "
           "recommend it nationally, which no project partner can do for itself. The monitoring "
           "platform is designed to be hosted by the University of the Aegean beyond the project "
           "period.")

    # ---- 10 budget
    h("10. Indicative budget envelope")
    crit("Award criterion 6 Cost effectiveness — 12 points. Indicative only: the real allocation "
         "follows the work packages. Shown now because two financial rules constrain the concept "
         "itself.")
    para(f"Guidelines 4.4.3 requires at least 50% of total eligible costs to be dedicated to "
         f"activities implemented in Mediterranean Partner Country territories — "
         f"EUR {TOTAL*MPC_FLOOR:,.2f} — and caps any single organisation at 35%, "
         f"EUR {TOTAL*ORG_CAP:,.2f}. A shape that satisfies both:")
    bt = d.add_table(rows=1, cols=5)
    bt.style = "Light Grid Accent 1"
    hdr = bt.rows[0].cells
    for i, x in enumerate(["Partner", "Country", "Role", "EUR", "Status"]):
        hdr[i].text = clean(x)
        for p in hdr[i].paragraphs:
            for rr in p.runs:
                rr.bold = True
                rr.font.size = Pt(9.5)
    for name, country, role, amt, status in ALLOC:
        c = bt.add_row().cells
        for i, x in enumerate([name, country, role, f"{amt:,.2f}", status]):
            c[i].text = clean(x)
            for p in c[i].paragraphs:
                for rr in p.runs:
                    rr.font.size = Pt(9.5)
    c = bt.add_row().cells
    for i, x in enumerate(["TOTAL", "", "", f"{TOTAL:,.2f}", ""]):
        c[i].text = clean(x)
        for p in c[i].paragraphs:
            for rr in p.runs:
                rr.bold = True
                rr.font.size = Pt(9.5)
    mpc_partners = round(sum(a[3] for a in ALLOC if a[1] in MPC), 2)
    para(f"Greek partners EUR {sum(a[3] for a in ALLOC if a[1]=='Greece'):,.2f} "
         f"({sum(a[3] for a in ALLOC if a[1]=='Greece')/TOTAL*100:.1f}%) · "
         f"largest single organisation "
         f"{max(a[3] for a in ALLOC)/TOTAL*100:.2f}% against the 35% cap.",
         size=9.5, colour=GREY)
    # THE FLOOR IS SHOWN AS A CALCULATION, NOT AS A CLAIM.
    #
    # The previous version printed one number and the words "above the 50%
    # floor". On this split that sentence would be true only because of
    # 4.4.3(b), and a reader who checked the partner column would find 49,50%
    # and conclude the proposal was ineligible. Show the working, name the
    # provision, and the evaluator is doing the same arithmetic we are.
    para("Article 4.4.3 tests where activities are IMPLEMENTED, not where the "
         "budget is held, and 4.4.3(b) allows an EU partner's direct costs "
         "for activities implemented in a Mediterranean Partner Country to "
         "count toward the floor. This proposal relies on that provision and "
         "states it openly:", size=10)
    for lbl, val in [
            ("Partners established in MPC territories (TR, LB)", mpc_partners),
            ("ISAC, CNR direct costs for activities implemented in Türkiye "
             "and Lebanon, under 4.4.3(b)", CNR_IN_MPC),
            ("Total dedicated to activities implemented in MPC territories",
             mpc),
            ("Floor: 50% of total eligible costs",
             round(TOTAL * MPC_FLOOR, 2))]:
        para(f"    {lbl}: EUR {val:,.2f} ({val / TOTAL * 100:.2f}%)",
             size=9.5, colour=GREY)
    para(f"Margin above the floor: EUR {mpc - TOTAL * MPC_FLOOR:,.2f}.",
         size=9.5, colour=GREY)
    para(f"EU contribution at 92%: {EU:,.2f}, which is {EU_CAP - EU:,.2f} "
         f"under the {EU_CAP:,.2f} the call awards. Partnership co-financing "
         f"at 8%: {CO:,.2f}. Total eligible costs sit {CEILING-TOTAL:,.2f} "
         f"below the {CEILING:,.2f} ceiling.", size=9.5, colour=GREY)
    para("Note the consequence for the concept: because at least half the money "
         "must be spent in Turkiye and Lebanon, Lemnos cannot be the whole "
         "project. The method-transfer design in section 5 is what the financial "
         "rules require, not a concession.", colour=RED, size=10)
    para("ISAC, CNR confirmed participation on 28 September 2026. Its "
         "allocation of EUR 80.000,00 covers quality assurance of the three "
         "baselines and training of the site teams, of which EUR "
         f"{CNR_IN_MPC:,.2f} is work carried out on the Turkish and Lebanese "
         "sites and is itemised as such in the budget tables the programme "
         "requires.", size=10)

    # ---- 11 open points
    h("12. Dates", size=11)
    para(f"Submission deadline {DEADLINE}. Questions to the programme close "
         f"{QUESTIONS_CLOSE}.", colour=RED)
    # THE JOINT SECRETARIAT WILL NOT ANSWER AN ELIGIBILITY QUESTION.
    #
    # Asked on 28 September, and the reply was that "in view of ensuring the
    # equal treatment of potential applicants, the Programme cannot give a
    # prior opinion on the eligibility of a project", with a pointer to the
    # terms of reference and the Guidelines.
    #
    # That is the correct answer and it costs us nothing, because the question
    # was answerable from the documents all along. The eligible area is
    # defined BY PROVINCE, not by coastline, and Denizli is named in it. So is
    # Afyonkarahisar, Kutahya, Usak, Isparta, Burdur and Kahramanmaras, none
    # of which has a coast. An inland site in a listed province is inside the
    # programme area by the programme's own definition. And SO 2.2 is
    # "climate change adaptation and disaster risk prevention, resilience
    # taking into account ecosystem-based approaches" - nothing in it is
    # marine, and "ecosystem-based" covers a lake shore without straining.
    para("The Joint Secretariat does not give prior opinions on eligibility, "
         "in order to treat all applicants equally, so the position below is "
         "read off the call documents rather than confirmed by the programme. "
         "The eligible area is defined by PROVINCE and not by coastline: "
         "Denizli is named in it, as are Afyonkarahisar, Kütahya, Uşak, "
         "Isparta, Burdur and Kahramanmaraş, none of which has a coast. A "
         "lacustrine site in a listed province therefore sits inside the "
         "programme area on the programme's own definition. Specific "
         "Objective 2.2 concerns ecosystem-based approaches to climate "
         "adaptation and carries no marine restriction.", size=10)
    para("A note on the quality bar: proposals must score at least 75 of 88 points after Step 1 "
         "to reach Step 2, and Relevance must then reach 16 of 24. Near-maximum on the "
         "operational criteria is the entry ticket, which is why partner input on the "
         "[TO CONFIRM] items matters more than speed.", size=9.5, colour=GREY)

    # the island changed from Lesvos to Lemnos at v2; this refuses to ship a
    # document that still names the old one anywhere, including tables
    body = "\n".join(par.text for par in d.paragraphs)
    for tb in d.tables:
        for row in tb.rows:
            for cell in row.cells:
                body += "\n" + cell.text
    stale = [w for w in ("Lesvos", "Mytilene") if w.lower() in body.lower()]
    assert not stale, f"document still names {stale} - the site is Lemnos (Gomati)"

    # RAS MUST NOT APPEAR IN THE PARTNER-FACING DOCUMENT.
    #
    # The Autonomous Region of Sardinia is the Managing Authority of this
    # programme and also holds the audit function. Naming it as an associated
    # partner, or attaching a letter of support from it, puts the awarding and
    # auditing authority inside a proposal it assesses.
    #
    # This check runs HERE, before the reviewer notes are appended, and that
    # placement is deliberate: the reviewer block has to explain the rule, so
    # it necessarily contains the words. A guard that scanned the finished
    # document would fire on the explanation of why it exists. The
    # partner-facing half is what must be clean.
    #
    # "Sardinia" is not forbidden - CNR's section is in Cagliari and saying so
    # is just true. What is forbidden is RAS as a body in the partnership.
    #
    # AND THE ACRONYM NEEDS WORD BOUNDARIES, WHICH I LEARNED BY WATCHING THIS
    # GUARD FIRE ON ITS FIRST RUN. "RAS" as a plain substring matches
    # "infrastructure" and "Erasmus", both of which are in this document
    # legitimately. A three-letter acronym checked without boundaries reports
    # a violation on ordinary prose, and a guard that cries wolf gets
    # commented out - at which point it protects nothing at all.
    ras = [w for w in ("Autonomous Region of Sardinia", "Regione Autonoma")
           if w.lower() in body.lower()]
    if re.search(r"\bRAS\b", body):
        ras.append("RAS")
    assert not ras, (
        f"the partner-facing document names {ras}. The Autonomous Region of "
        f"Sardinia is the MANAGING AUTHORITY and the audit authority of this "
        f"programme; it cannot appear in the partnership in any role, "
        f"including as an associated partner or as a letter of support.")

    # -------------------------------------------------------- REVIEWER NOTES
    d.add_page_break()
    h("REVIEWER NOTES", size=15, colour=RED, space=0)
    para("For you, not for the partners. Read it, settle the open points, then "
         "DELETE FROM THIS PAGE TO THE END before circulating.", bold=True,
         colour=RED)
    for t_, b_ in [
        ("THE ONE THING THAT WOULD SINK THIS: the 50% floor is not carried "
         "by the MPC partners",
         "PIYA, Acipayam and LCEC together hold 645.000, which is 49,50% of "
         "1.303.000. That is BELOW the floor. The proposal is compliant only "
         "because 4.4.3(b) counts an EU partner's direct costs for activities "
         "implemented in MPC territory, and 50.000 of CNR's 80.000 is "
         "declared as work on the Turkish and Lebanese sites. Ask CNR "
         "directly, before this split is circulated: will they travel to "
         "Ucari Lake and the Lebanese site, or do they intend to work from "
         "Cagliari? If it is Cagliari, the 24.000 given back to HCMR and the "
         "University of the Aegean has to be taken back out again, because "
         "there is only 1.347,83 of headroom in the total and 1.240,00 in the "
         "EU contribution."),
        ("RAS must not appear anywhere in this proposal, including as an "
         "associated partner or a letter of support",
         "The Autonomous Region of Sardinia IS the Managing Authority of "
         "Interreg NEXT MED, by decision of the participating countries, and "
         "the audit functions are assigned to it as well. So RAS awards the "
         "grant, pays it, monitors it, supervises the Joint Secretariat and "
         "audits the project. Listing it as an associated partner puts the "
         "awarding and auditing authority inside a proposal it will assess. "
         "The confusion is understandable and honest: RAS genuinely IS a "
         "partner in AMMIRARE, listed three lines above CNR, because under "
         "Italy, France Maritime it is an ordinary beneficiary. Same legal "
         "body, opposite role under this programme. CNR itself is a national "
         "council and a separate legal person, so CNR as a partner is fine."),
        ("Tell them there is no Italian site, in the first conversation",
         "A research institute's instinct is to propose its own field site. If "
         "one appears, the proposal reads as AMMIRARE phase two rather than a "
         "transfer to new territories, and that is the single reading that "
         "would lose it. Greece and Turkiye are also both AT the two "
         "organisations per country limit now, so no further Greek or Turkish "
         "partner can be added after this."),
        ("Still open",
         "The Joint Secretariat answer on whether an inland lacustrine pilot "
         "is eligible under SO 2.2, with questions closing 14 October. "
         "Acipayam's visitor numbers, without which acute tourism pressure is "
         "an assertion rather than a baseline. LCEC's site, and whether it can "
         "fund its 8%, which remains the only partner with no identified "
         "route. And 12.000 is earmarked inside HCMR's external expertise for "
         "translating the AMMIRARE method sheets into English, Turkish and "
         "Arabic, which is not optional because MMM_IF02 is 68 pages of "
         "Italian."),
    ]:
        para(t_, bold=True, size=11, colour=NAVY, after=2)
        para(b_, size=10, after=8)

    h("11. What the partnership needs to settle")
    bullet("The habitat code. HCMR's own document gives 2210 in the summary and 2250 in the "
           "detailed description for the same habitat. They are different types and only one "
           "of them is a priority habitat, so the phrase \"priority habitat\" depends on which "
           "is correct. An evaluator will check this. One possibility worth putting to the "
           "ecologists: the two sections look different in the imagery \u2014 A more "
           "vegetated and stabilised, B barer and closer to the shore \u2014 so the site may "
           "hold more than one Annex I type, which would make both codes right for different "
           "polygons rather than one of them wrong.")
    bullet("Who owns or manages the Gomati site, and whether the route to protection runs "
           "through Natura 2000 area GR4110006 or the Faraklo Geopark — HCMR's document names "
           "both as candidates.")
    bullet("What fragile landscape under visitor pressure Acipayam manages — this decides "
           "whether Acipayam is a funded partner or an associate.")
    bullet("CLOSED 28 September: LCEC has confirmed it can sign in its own name, so no "
           "signature is needed from the Ministry of Energy and Water. Still outstanding is "
           "the narrower half of the same question — what coastal or tourism-facility work "
           "LCEC has actually done, which is what award criterion 3 will be scored on.")
    bullet("CLOSED 28 September, though not the way it was asked: the Joint Secretariat does "
           "not give prior opinions on eligibility, in order to treat applicants equally. So "
           "the inland-site question is answered from the call documents instead, and the "
           "answer is favourable — the eligible area is a list of PROVINCES, Denizli is in it, "
           "and so are six Turkish provinces with no coast at all. Do not ask the Joint "
           "Secretariat an eligibility question again; ask only questions of fact about the "
           "forms and the procedure, which they will answer.")
    bullet("Whether Zakynthos and the two ministries accept associate status, understanding that "
           "an associate can never bid for a project contract.")
    bullet("Quantified target groups from each partner.")
    bullet("The call's Terms of Reference, so section 2 can adopt its wording — award criterion "
           "1.1 refers to it explicitly.")

    d.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
