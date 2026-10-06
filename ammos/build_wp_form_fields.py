#!/usr/bin/env python3
"""Every work package field, at the sizes the form actually uses.

    python3 build_wp_form_fields.py

His screenshots settled the whole pattern. Each work package is five sub-tabs:

    WP Info | WP Outputs | WP Budget | WP Budget per outputs
            | WP Budget per cost category

  WP Info          Title 100 · Description 2000 · Coordinator (dropdown)
                   Involved partners (multi-select, REQUIRED)
                   Final beneficiaries 500 · expected result read-only
  Activities       Title 100 · Description 500 · Implementing period
                   (month multi-select)
  WP Budget        a row grid: Code | Cost category | Partner | Brief
                   description | Total | Justification, plus "Semester
                   incurred" where you ctrl+click several months
  per outputs      the % split, which must total 100
  per cost cat.    read-only Partner x Cost category matrix

AND IT SETTLED THE WP1/WP2 QUESTION: DESCRIPTION IS 2000, NOT 3000.
What I sent yesterday was 2805 and 2741. Both are cut here, and both now also
carry the Final beneficiaries box I did not know existed.

THREE THINGS IN HIS SCREENSHOTS THAT NEED CORRECTING BEFORE HE GOES FURTHER

1. WP3's budget holds ONE row: Applicant, staff cost, 173.500,00. That is
   HCMR's ENTIRE project staff budget sitting in a single work package. The
   per-cost-category tab shows it: Applicant 173.500, every other partner
   nil. WP3's real total is 245.522,00 spread over six partners.

2. Output3.1 has three activities under it - 3.1.1 transfer workshop, 3.1.2
   territorial baseline, 3.1.3 tourism and valuation. The last two belong to
   Output3.2. Activity numbering is {WP}.{Output}.{Activity}, so an activity
   filed under the wrong output is numbered wrong and reads as though the
   baseline work has no output of its own.

3. Activities 3.1.2 and 3.1.3 have implementing period "1" - month one only.
   They run M5 to M12.

WHY THE BUDGET TABLE BELOW EXISTS

The form wants one row per partner per cost category per work package, and
nothing we hold was three-dimensional: the model gives each partner's split by
cost category, and the work package matrix gives each partner's split by work
package. So each partner's cost-category total is distributed across that
partner's own work packages in the same proportion as their work package
shares. That reconciles in both directions exactly - every partner's
cost-category total is preserved, and every work package total is preserved -
and the script refuses to emit the table unless both margins balance to the
cent.
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

from build_wp_blended import DIRECT, P, PP, STAFF, WPS, matrix
from build_outputs import OUTPUTS, TECHNICAL

BASE = str(here(__file__))
OUT_INFO = f"{BASE}/AMMOS_WP-Info-and-Activities_all-work-packages_2026-10-07.docx"
OUT_BUD = f"{BASE}/AMMOS_WP-Budget-rows_per-partner-per-category_2026-10-07.docx"

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

LIM_TITLE, LIM_DESC, LIM_BEN = 100, 2000, 500
LIM_ACT_TITLE, LIM_ACT_DESC = 100, 500
FLAT = 0.30

# Cost categories as the form names them. Travel and admin are flat rates the
# system computes, so they are not rows.
CATS = ["Staff cost", "Equipment costs",
        "External expertise and services costs",
        "Infrastructures and works costs"]

# Per-partner split by cost category, from the client's budget model v6.
CC = {
    "HCMR": {"Staff cost": 173_500.0, "Equipment costs": 30_000.0,
             "External expertise and services costs": 94_450.0,
             "Infrastructures and works costs": 0.0},
    "UAEG": {"Staff cost": 60_000.0, "Equipment costs": 8_000.0,
             "External expertise and services costs": 119_000.0,
             "Infrastructures and works costs": 0.0},
    "LCEC": {"Staff cost": 80_000.0, "Equipment costs": 20_000.0,
             "External expertise and services costs": 176_000.0,
             "Infrastructures and works costs": 0.0},
    "PIYA": {"Staff cost": 50_000.0, "Equipment costs": 10_000.0,
             "External expertise and services costs": 155_000.0,
             "Infrastructures and works costs": 0.0},
    "ACIP": {"Staff cost": 30_000.0, "Equipment costs": 12_000.0,
             "External expertise and services costs": 84_000.0,
             "Infrastructures and works costs": 0.0},
    "CNR": {"Staff cost": 50_000.0, "Equipment costs": 0.0,
            "External expertise and services costs": 15_000.0,
            "Infrastructures and works costs": 0.0},
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


def sems(m0, m1):
    """Semesters a month range touches, for 'Semester incurred'."""
    return sorted({(m - 1) // 6 + 1 for m in range(m0, m1 + 1)})


# ------------------------------------------------------- WP Info, all six
WP_INFO = {
    "WP1": ("Management and coordination", "Applicant",
            ["The project's coordination layer. HCMR as Applicant provides a "
             "scientific coordinator and a project manager. Each partner "
             "names one representative and one alternate to a Management "
             "Committee, meeting monthly online and once a semester in "
             "person, scheduled alongside demonstration and communication "
             "activity so travel serves two purposes. A written RACI table "
             "assigns, for every output and activity, who is responsible, "
             "accountable, consulted and informed - which prevents the "
             "familiar failure of a transfer project where two partners each "
             "assume the other is adapting a method sheet.",

             "DECISION-MAKING. The Committee decides by consensus and failing "
             "that by simple majority, with the Applicant holding a casting "
             "vote on scientific consistency only. Three things require a "
             "recorded Committee decision: reallocating budget between "
             "partners, changing an output target value, and reclassifying "
             "the transferability of any method sheet. Those are the three "
             "that quietly change what the project promised.",

             "FINANCIAL MANAGEMENT. Each partner contracts its own "
             "independent first-level control, budgeted here, and submits "
             "certified expenditure per reporting period. The Applicant "
             "consolidates, holds the single interface with the programme, "
             "and tracks each partner against its own semester spending "
             "profile, so drift is visible in the period it happens rather "
             "than at closure.",

             "MONITORING AND RISK. Every deliverable passes two checks: the "
             "activity lead signs technical content, the scientific "
             "coordinator signs cross-site consistency, because a protocol "
             "correct at one site and incompatible with another destroys the "
             "comparability the project exists to create. The risk register "
             "is reviewed each semester. The principal risks are seasonal - "
             "a field campaign lost to weather or a delayed consent cannot be "
             "repeated for twelve months - so each site holds a second "
             "campaign window and the installations complete at month 20, "
             "four months before closure."],
            "The partnership itself and the programme: coordination quality "
            "determines whether the transferred method arrives intact at "
            "three sites. Indirectly the three site authorities, who receive "
            "deliverables that are internally consistent and on time."),

    "WP2": ("Communication and dissemination", "Applicant",
            ["Communication here has one job: to move a method from the "
             "people who wrote it to the people who must apply it, and then "
             "to people we will never meet.",

             "WHY THE BUDGET IS WHAT IT IS. The method we transfer exists "
             "only in Italian and its users work in Greek, Turkish and "
             "Arabic. Translation, local-language field material and local "
             "workshops are not promotion - they are the transfer mechanism.",

             "AUDIENCES. Site managers and municipal staff need the method "
             "in their own language and in a form they can act on. The "
             "authorities that adopt the management plans need costs, "
             "obligations and evidence. Operators and visitors need to know "
             "why a boardwalk and a marked boundary exist, or they walk "
             "around them. Schools and young people are the custodians who "
             "maintain the sites after us. The MMM programmes and other "
             "Mediterranean territories are the replication audience.",

             "CHANNELS. A website with a public area and a password-protected "
             "partner area, embedded in the municipal, prefecture and "
             "regional sites of the three territories so it is found where "
             "people already look. Social media in four languages run by the "
             "youth partner with the nine custodian groups generating content "
             "from their own monitoring rounds. Two newsletter cycles to "
             "local sectoral press in every partner country, at month 1 and "
             "month 24. On-site interpretation at all three sites and a "
             "visitor booklet. Six local workshops, one stakeholder and one "
             "results workshop per territory. A final transferability "
             "conference. One peer-reviewed paper.",

             "PARTNER ROLES AND MEASUREMENT. The Applicant coordinates. The "
             "youth partner runs youth channels. Each site partner owns its "
             "local press, workshops and on-site material in its own "
             "language, because national press contacts cannot be run "
             "centrally. LCEC leads policy-facing communication; HCMR leads "
             "scientific dissemination. Reach is measured, not estimated: the "
             "WP3 visitor survey establishes the real audience at each site "
             "in season one and season two targets are set from it."],
            "Site managers and municipal staff at the three territories, who "
            "receive the method in their own language; visitors and operators "
            "at each site; schools and young people; and other Mediterranean "
            "territories reached through the MMM database and the final "
            "conference."),

    "WP3": ("Characterising and adapting the method", "Applicant",
            ["Everything the project does afterwards depends on this work "
             "package, so it comes first and it gates the rest.",

             "THE TRANSFERABILITY DECISION. The source method is seven field "
             "method sheets; the receiving sites are a marine dune system at "
             "Gomati, a marine coastal site in Lebanon and an artificial "
             "freshwater pond at Ucari. Transfer is therefore 21 separate "
             "decisions, and each is recorded as APPLIED, SUBSTITUTED or NOT "
             "APPLICABLE with its reason. Where a sheet does not transfer, a "
             "substitute procedure is drafted: a freshwater littoral "
             "vegetation protocol in place of dune vegetation, a lake-regime "
             "forcings protocol in place of meteo-marine forcings, an "
             "ecological status indicator recomposed for a system with "
             "neither dune nor banquette. Nothing is procured, installed or "
             "trained until that matrix exists.",

             "TRANSLATION IS PART OF THE METHOD, NOT AN EXTRA. The source is "
             "68 pages of Italian and no site manager will use it in that "
             "form, so the adapted method is produced in Greek, Turkish and "
             "Arabic.",

             "THEN TESTED BY BEING USED. Three comparable territorial "
             "baselines are produced with the adapted method: topography and "
             "shore state, utilities and access, sea level rise scenarios "
             "from 0.2 m to 1 m, biodiversity and ecological status, and the "
             "tourism layer - visitor numbers and seasonality, a visitor "
             "survey on willingness to visit and to pay, and the shore's "
             "economic contribution. If a substituted procedure fails in the "
             "field it fails here, at the cost of a protocol revision, rather "
             "than in WP4 at the cost of an installation.",

             "ROLES. HCMR coordinates and carries the biological and "
             "ecological sheets. ISAC-CNR adapts the topographic and forcings "
             "sheets its own researchers authored. The University of the "
             "Aegean leads the tourism and valuation baseline. The three site "
             "partners state what their site can and cannot support, which is "
             "the input the decision matrix is built from."],
            "The authorities and site managers at the three demonstration "
            "sites, who gain a method they can apply and a baseline for their "
            "own shore. Beyond them, any Mediterranean territory facing the "
            "same problem, because the 21 recorded decisions state where this "
            "method works and where it stops."),

    "WP4": ("Demonstrating it on three shores", "PP01",
            ["The method applied rather than described, at all three sites, "
             "and observed under real visitor load.",

             "WHAT IS INSTALLED. Visitor channelling and shore protection "
             "designed from each site's own baseline: boundary marking, "
             "interpretation signs showing what is protected and why, "
             "demountable modular boardwalks over the vulnerable surface, and "
             "demarcated zones where trampling is degrading vegetation or "
             "banks. Designs are signed off by the authority that will own "
             "the installation afterwards, not only by the partner building "
             "it. No underwater or bottom cleaning at any site.",

             "WHY DEMOUNTABLE. Three reasons, and they are not only "
             "environmental. It keeps intervention proportionate beside an "
             "Annex I habitat; it keeps environmental screening "
             "proportionate; and it means a design that performs badly in its "
             "first season can be altered rather than defended. Modularity is "
             "also a cost control - a section can be repaired instead of a "
             "structure replaced.",

             "AND THE PART THAT MAKES THE SITES COMPARABLE. One shared "
             "web-GIS monitoring platform carries a common protocol across "
             "all three territories, with field data entry in each local "
             "language. The COMMON beach litter protocol is adopted as "
             "published so the data stays comparable with the European "
             "Environment Agency dataset. The cross-typology comparability "
             "layer is the new element: it is what allows a freshwater pond "
             "and a marine dune to be read side by side, and it carries an "
             "explicit statement of what a new territory must do to join.",

             "TIMING IS DELIBERATE. Installation completes at month 20, four "
             "months before closure, because a pilot action counts only if it "
             "is finalised within the project."],
            "Visitors and tourism operators at the three sites, who keep "
            "access to a shore that is being protected rather than closed; "
            "the site authorities, who gain a working installation and the "
            "first season of evidence about it; and the local economies that "
            "depend on the shore remaining usable."),

    "WP5": ("Building the capacity to keep doing it", "PP02",
            ["A method nobody local can run is not transferred. This work "
             "package puts it in the hands of the people who will still be "
             "there in year three.",

             "SITE-MANAGER TRAINING. Built on the COMMON Capacity Building "
             "Training Design Plan rather than designed from scratch, "
             "delivered in each country in the local language with a "
             "practical module at the demonstration site itself. It targets "
             "municipal and authority staff holding a mandate over the site, "
             "so that the management plans in WP6 are written by people who "
             "have used the method rather than for them. Sixty staff across "
             "three countries.",

             "SHORE CUSTODIANS. The AMMIRARE Beach Custodians instrument "
             "transferred to shore types it was not written for - the same "
             "adaptation problem as the method sheets, and the same test of "
             "it. Material in Greek, Turkish and Arabic, with field exercises "
             "reworked to what each site actually has. Nine groups, three per "
             "country, run with the schools nearest each site so a group "
             "survives by being local rather than by being funded. Each runs "
             "at least one monitoring round and one awareness action, and its "
             "own observations enter the shared platform.",

             "VISITOR AWARENESS. The BEach CLEAN decalogue adapted per site "
             "and deployed as on-site material, with the Arabic version "
             "extended rather than translated afresh.",

             "WHY IT STARTS BEFORE THE DEMONSTRATION. Training opens at month "
             "6 and the custodian groups at month 10, so trained staff and "
             "active groups are in place when the installations go in rather "
             "than afterwards."],
            "Municipal and authority staff at the three sites, who acquire a "
            "transferable skill rather than attending an event; school pupils "
            "in three countries; and the sites themselves, which gain people "
            "able to monitor and maintain them once the funding stops."),

    "WP6": ("Making it stick", "PP05",
            ["The difference between a pilot and a transfer is whether "
             "anybody adopts it. This work package produces the instruments "
             "that outlive the funding and puts them in front of the bodies "
             "that can apply them.",

             "COSTED PLANS TAKEN TO ADOPTION. One management plan per site, "
             "each stating its own maintenance cost so the authority adopting "
             "it knows the annual figure it is taking on, and each taken "
             "through stakeholder consultation to a formal adoption decision "
             "by the body with the mandate. A plan that does not state its "
             "running cost is adopted and then abandoned at the first budget "
             "round.",

             "GOVERNANCE THAT OUTLASTS US. A Shore Stewardship Committee is "
             "constituted at each site alongside the plan it oversees - the "
             "responsible authority, a tourism operator, the research partner "
             "and a civil society or youth representative - so every plan has "
             "a standing multi-stakeholder owner after closure. For Lemnos "
             "the geopark route is pursued so that maintenance is carried by "
             "an organisation that already has activities and income.",

             "WHAT GOES BACK TO THE MECHANISM. A joint transferability guide "
             "in English and Arabic: the adapted method, the 21 decisions and "
             "their reasons, the comparability layer and the costs, with a "
             "closing section stating the actions a fourth territory must "
             "take to adopt or upscale it. Lodged in the MMM common database "
             "under open licence. Policy recommendations addressed to the "
             "adopting authorities and to the programme, and a cooperation "
             "agreement signed by all partners, the associated organisation "
             "and the three committees, committing them to continue "
             "cooperating after the project ends."],
            "The three site authorities, which end the project holding an "
            "adopted, costed plan and a standing committee; their residents "
            "and operators, represented on those committees; and any "
            "Mediterranean territory that later adopts the method from the "
            "guide without having been in the partnership."),
}


class Doc:
    def __init__(self, margin=1.7, size=10.5):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(margin)
        n = self.d.styles["Normal"]
        n.font.name, n.font.size = "Calibri", Pt(size)

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
                if str(row[0]).startswith("TOTAL"):
                    r.bold = True
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        return t

    def page(self):
        self.d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def allocate():
    """partner -> wp -> category -> amount, reconciling BOTH margins."""
    m = matrix()
    out = {}
    for k in P:
        share = {wp: m[k].get(wp, 0.0) / DIRECT[k] for wp in m[k]
                 if m[k].get(wp, 0.0) > 0}
        out[k] = {}
        for wp, s in share.items():
            out[k][wp] = {c: CC[k][c] * s for c in CATS}
    return out, m


def main():
    alloc, m = allocate()
    wp_tot = {w[0]: sum(m[k].get(w[0], 0.0) for k in P) for w in WPS}

    fails = []

    # ---- WP Info fits
    info_report = []
    for wp, (title, coord, paras, benef) in WP_INFO.items():
        d = clean("\n\n".join(paras))
        if len(clean(title)) > LIM_TITLE:
            fails.append(f"{wp} title {len(clean(title))} > {LIM_TITLE}")
        if len(d) > LIM_DESC:
            fails.append(f"{wp} description {len(d)} > {LIM_DESC} - cut "
                         f"{len(d) - LIM_DESC}")
        if len(clean(benef)) > LIM_BEN:
            fails.append(f"{wp} final beneficiaries {len(clean(benef))} > "
                         f"{LIM_BEN} - cut {len(clean(benef)) - LIM_BEN}")
        info_report.append((wp, len(clean(title)), len(d), len(clean(benef))))

    # ---- activities fit, and sit under the right output
    act_rows = []
    for wp, oi, _t, _tv, _u, _fin, _sh, _desc, acts in OUTPUTS:
        for ai, (atitle, m0, m1, lead, adesc) in enumerate(acts, 1):
            code = f"{wp[2]}.{oi}.{ai}"
            if len(clean(atitle)) > LIM_ACT_TITLE:
                fails.append(f"activity {code} title "
                             f"{len(clean(atitle))} > {LIM_ACT_TITLE}")
            if len(clean(adesc)) > LIM_ACT_DESC:
                fails.append(f"activity {code} description "
                             f"{len(clean(adesc))} > {LIM_ACT_DESC} - cut "
                             f"{len(clean(adesc)) - LIM_ACT_DESC}")
            act_rows.append((wp, code, atitle, adesc, m0, m1, lead))

    # ---- the budget allocation must reconcile in BOTH directions
    for k in P:
        for c in CATS:
            got = sum(alloc[k][wp].get(c, 0.0) for wp in alloc[k])
            if abs(got - CC[k][c]) > 0.01:
                fails.append(f"{k} {c}: rows sum {got:.2f}, model says "
                             f"{CC[k][c]:.2f}")
        got = sum(sum(alloc[k][wp].values()) for wp in alloc[k])
        if abs(got - DIRECT[k]) > 0.01:
            fails.append(f"{k} total {got:.2f} vs direct {DIRECT[k]:.2f}")
    for wp in wp_tot:
        got = sum(sum(alloc[k][wp].values()) for k in P if wp in alloc[k])
        if abs(got - wp_tot[wp]) > 0.01:
            fails.append(f"{wp} rows sum {got:.2f} vs matrix "
                         f"{wp_tot[wp]:.2f}")
    grand = sum(sum(sum(alloc[k][wp].values()) for wp in alloc[k]) for k in P)
    if abs(grand - sum(DIRECT.values())) > 0.02:
        fails.append(f"grand total {grand:.2f} vs {sum(DIRECT.values()):.2f}")
    # Every partner must appear in WP1 and WP2 - error code 222.
    for k in P:
        for wp in ("WP1", "WP2"):
            if sum(alloc[k].get(wp, {}).values()) <= 0:
                fails.append(f"{k} has nothing in {wp} - code 222 requires "
                             f"coordinator and audit in WP1 and "
                             f"communication in WP2 for every partner")
    # No infrastructure outside the technical work packages.
    for k in P:
        for wp in ("WP1", "WP2"):
            if alloc[k].get(wp, {}).get("Infrastructures and works costs", 0):
                fails.append(f"{k} has infrastructure in {wp}, which the "
                             f"rules forbid")

    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    # =============================================== document 1: info + acts
    D = Doc()
    D.h("AMMOS — WP Info and Activities, all six work packages", size=15)
    D.p(f"Generated {dt.date.today():%d/%m/%Y} to the sizes in your "
        f"screenshots: WP Title 100, WP Description 2000, Final "
        f"beneficiaries 500, Activity title 100, Activity description 500. "
        f"Every field checked against its counter before this file was "
        f"produced.", size=9.5, colour=GREY, italic=True)

    D.h("Three things to correct first", size=13)
    for t in [
        "WP3's budget currently holds one row: Applicant, staff cost, "
        "173.500,00. That is HCMR's ENTIRE project staff budget sitting in a "
        "single work package - the per-cost-category tab shows Applicant "
        "173.500 and every other partner at nil. WP3's real total is "
        f"{eur(wp_tot['WP3'])} across six partners. The second file has every "
        "row.",
        "Output3.1 has three activities under it. 3.1.2 territorial baseline "
        "and 3.1.3 tourism and valuation belong to Output3.2. Activities are "
        "numbered WP.Output.Activity, so one filed under the wrong output is "
        "numbered wrong and makes the baseline work look as though it has no "
        "output of its own.",
        "Activities 3.1.2 and 3.1.3 have implementing period month 1 only. "
        "They run M5 to M12.",
        "And WP1 and WP2 descriptions: you told me 2000, and what I sent "
        "yesterday was 2805 and 2741. Both are cut below, and both now also "
        "carry the Final beneficiaries box I had not seen.",
    ]:
        D.p("•  " + t, size=10)

    D.table(["WP", "Title", "Description", "Final beneficiaries"],
            [[wp, f"{a}/{LIM_TITLE}", f"{b}/{LIM_DESC}", f"{c}/{LIM_BEN}"]
             for wp, a, b, c in info_report], widths=[1.6, 2.6, 3.0, 3.6])

    for wp in ("WP1", "WP2", "WP3", "WP4", "WP5", "WP6"):
        title, coord, paras, benef = WP_INFO[wp]
        D.page()
        D.h(f"{wp} — WP Info tab", size=14)
        D.p(f"Title: {title}", size=10.5, bold=True)
        D.p(f"Coordinator: {coord}"
            + ("  (pre-set to Applicant, cannot be changed)"
               if wp in ("WP1", "WP2") else ""), size=10)
        inv = ", ".join(f"{PP[k]} {P[k][0]}"
                        for k in sorted(P, key=lambda x: PP[x])
                        if sum(alloc[k].get(wp, {}).values()) > 0)
        D.p(f"Involved partners (required - this is the field that would not "
            f"save empty): {inv}", size=10)
        D.p(f"Programme expected result: SO1ER1 (read-only)", size=10)
        D.p("")
        D.p(f"Description — {len(clean(chr(10).join(paras)))} characters:",
            size=9, colour=GREY, italic=True)
        for para in paras:
            D.p(para)
        D.p("")
        D.p(f"Final beneficiaries — {len(clean(benef))}/{LIM_BEN}:", size=9,
            colour=GREY, italic=True)
        D.p(benef)

        if wp in TECHNICAL:
            D.h(f"{wp} — Activities", size=12, space=10)
            D.p("Add these under the output shown, not all under the first "
                "one. Implementing period is a month multi-select: "
                "ctrl+click the months listed.", size=9.5, colour=GREY)
            for w, code, atitle, adesc, m0, m1, lead in act_rows:
                if w != wp:
                    continue
                D.p(f"Activity {code}  —  under Output{code[0]}.{code[2]}",
                    size=10.5, bold=True)
                D.p(f"Title ({len(clean(atitle))}/{LIM_ACT_TITLE}): "
                    f"{atitle}", size=10)
                D.p(f"Implementing period: months "
                    f"{', '.join(str(x) for x in range(m0, m1 + 1))}   "
                    f"(semester {', '.join(str(s) for s in sems(m0, m1))})",
                    size=10)
                D.p(f"Leads: {PP[lead]} {P[lead][0]}", size=9, colour=GREY)
                D.p(f"Description ({len(clean(adesc))}/{LIM_ACT_DESC}): "
                    f"{adesc}", size=10)
    D.d.save(OUT_INFO)

    # ================================================ document 2: budget rows
    D2 = Doc(margin=1.2, size=9)
    D2.h("AMMOS — WP Budget rows, every partner, every cost category",
         size=15)
    D2.p(f"Generated {dt.date.today():%d/%m/%Y} from your budget model v6. "
         f"The form wants one row per partner per cost category per work "
         f"package. Nothing we held was three-dimensional, so each partner's "
         f"cost-category total is spread across that partner's own work "
         f"packages in proportion to their work package shares. It "
         f"reconciles both ways to the cent and the script refuses to print "
         f"it otherwise.", size=9, colour=GREY, italic=True)
    D2.p("Travel and administration are NOT rows - the system computes them "
         "as 15% each of staff. Infrastructure appears only in WP4, which is "
         "required: the rules forbid it in WP1 and WP2.", size=9)

    for wp in [w[0] for w in WPS]:
        D2.h(f"{wp} — total {eur(wp_tot[wp])}", size=12, space=10)
        rows = []
        for k in sorted(P, key=lambda x: PP[x]):
            cell = alloc[k].get(wp, {})
            for c in CATS:
                v = cell.get(c, 0.0)
                if v <= 0:
                    continue
                a = next((x for x in act_rows if x[0] == wp), None)
                rows.append([
                    f"{PP[k]} {P[k][0][:26]}", c,
                    eur(round(v, 2)),
                    ", ".join(str(s) for s in sems(
                        *(next(((x[4], x[5]) for x in act_rows
                                if x[0] == wp), (1, 24))))),
                ])
        rows.append(["TOTAL", "", eur(wp_tot[wp]), ""])
        D2.table(["Partner", "Cost category", "Total", "Semester incurred"],
                 rows, widths=[5.6, 6.0, 3.0, 3.4], small=8)

    D2.page()
    D2.h("Check these against the platform's own two summary tabs", size=13)
    D2.p("WP Budget per cost category is read-only and should reproduce this "
         "exactly. WP Budget per outputs takes percentages, which must total "
         "100 within each work package:", size=9)
    D2.table(["WP", "Output percentages", "WP total"],
             [[wp, " / ".join(f"O{wp[2]}.{o[1]} {round(o[6] * 100)}%"
                              for o in OUTPUTS if o[0] == wp),
               eur(wp_tot[wp])] for wp in TECHNICAL],
             widths=[1.6, 8.0, 3.4])
    D2.d.save(OUT_BUD)

    print(f"wrote {Path(OUT_INFO).name}")
    for wp, a, b, c in info_report:
        print(f"   {wp}  title {a:>3}/{LIM_TITLE}  desc {b:>4}/{LIM_DESC}  "
              f"beneficiaries {c:>3}/{LIM_BEN}")
    print(f"\n   {len(act_rows)} activities, all within 100/500")
    print(f"\nwrote {Path(OUT_BUD).name}")
    print("   budget reconciles both ways:")
    for wp in wp_tot:
        print(f"     {wp} {eur(wp_tot[wp])}")
    print(f"     TOTAL DIRECT {eur(sum(DIRECT.values()))}")


if __name__ == "__main__":
    main()
