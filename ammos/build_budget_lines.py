#!/usr/bin/env python3
"""The budget lines in the programme's own Courtesy Budget format.

    python3 build_budget_lines.py

He asked: "IN WP1 THE SYSTEM HAS ADDED AUDITOR COSTS. Perhaps if we fill these
excel files is better?" - and separately, "cAN WE HAVE A JUSTIFICATION?"

Yes to both, and the Courtesy Budget is the right place. Its sheets mirror the
platform one for one: Budget tot per budget line, Budget per CC, Budget per WP,
Budget per output, Co-financing table, 50% rule table. So the arithmetic can be
settled once, here, and then typed into the web form from a sheet that is
already balanced - instead of discovering a reconciliation error after a
hundred rows are in.

WHAT THE TEMPLATE TAUGHT ME

Budget line codes are {WP}.{CC}.{PARTNER}.{number}, and the cost category
abbreviations are fixed:

    ST  Staff Costs
    EC  Equipment costs
    ES  External expertise and services costs
    IW  Infrastructures and works costs

The Brief Description column uses standard programme wordings - "Project
Coordinator", "Human resources involved in WP3", "Equipment necessary for
WP3", "Auditor costs", "External expertise and services necessary for WP3".
Using their phrasing rather than our own costs nothing and reads as familiar.

AND IT SETTLES SOMETHING I GOT WRONG

The template's "Budget per output" sheet shows WP1 with Output 1.1, 1.2, 1.3
and 1.4. So the mandatory work packages DO have outputs, which is exactly what
he found when the budget-per-output tab would not work until WP1 and WP2 had
them. I told him on 2 October that the mandatory templates have no outputs,
because the exported Word form showed no output rows under WP1 and WP2. The
export was not the platform. I was wrong and this file fixes it.

AUDITOR COSTS: the system adding them in WP1 is correct and expected. Error
code 222 requires a coordinator line and an audit line for every partner in
WP1, and the template carries "Auditor costs" as a standard WP1 description.
First-level control has to be independent, so it is external expertise, not
staff.
"""
import datetime as dt
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import DIRECT, P, PP, STAFF, WPS, matrix

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_budget_lines_Courtesy-Budget-format_2026-10-08.xlsx"

FLAT = 0.30
COFIN = 0.08
MONTHS = 24

CC_CODE = {"Staff Costs": "ST", "Equipment costs": "EC",
           "External expertise and services costs": "ES",
           "Infrastructures and works costs": "IW"}
CATS = list(CC_CODE)

# Partner code as the template writes it.
PCODE = {"HCMR": "LeadPartner", "UAEG": "PP1", "PIYA": "PP2",
         "ACIP": "PP3", "CNR": "PP4", "LCEC": "PP5"}
PSHORT = {"HCMR": "LEP", "UAEG": "PP1", "PIYA": "PP2",
          "ACIP": "PP3", "CNR": "PP4", "LCEC": "PP5"}

# Per-partner split by cost category, from the client's budget model v6.
CC = {
    "HCMR": (173_500.0, 30_000.0, 94_450.0, 0.0),
    "UAEG": (60_000.0, 8_000.0, 119_000.0, 0.0),
    "LCEC": (80_000.0, 20_000.0, 176_000.0, 0.0),
    "PIYA": (50_000.0, 10_000.0, 155_000.0, 0.0),
    "ACIP": (30_000.0, 12_000.0, 84_000.0, 0.0),
    "CNR": (50_000.0, 0.0, 15_000.0, 0.0),
}

# First-level control is a FIXED cost, not a proportional one: four reporting
# periods, one certificate each, per partner. Spreading each partner's whole
# external-expertise budget by work package share put 42.412 into WP1 under the
# label "Auditor costs" - a consultancy budget wearing the wrong label. 09 Oct.
AUDIT = {"HCMR": 6_000.0, "UAEG": 3_500.0, "PIYA": 3_500.0, "ACIP": 3_000.0,
         "CNR": 3_000.0, "LCEC": 4_000.0}

WP_NAME = {"WP1": "WP1 Management", "WP2": "WP2 Communication",
           "WP3": "WP3", "WP4": "WP4", "WP5": "WP5", "WP6": "WP6"}

# Justification of the estimated cost. One per (cost category, work package),
# with the partner named where it changes the answer. These are the sentences
# an assessor reads when deciding whether a cost is necessary.
JUST = {
    ("Staff Costs", "WP1"):
        "Coordination and financial management effort over 24 months: "
        "scientific coordinator and project manager at the Applicant, and one "
        "representative plus one alternate per partner on the Management "
        "Committee, which meets monthly online and once per semester in "
        "person. Estimated from each organisation's own salary scales and the "
        "committee calendar, not from a percentage of budget.",
    ("Staff Costs", "WP2"):
        "Communication effort over the full project: website and partner "
        "area, four-language social media, two newsletter cycles, six local "
        "workshops and the final conference. Each site partner carries its "
        "own share because local press contacts and local-language material "
        "cannot be produced centrally.",
    ("Staff Costs", "WP3"):
        "The heaviest intellectual load in the project: assessing 21 "
        "transferability cells, drafting substituted procedures, translating "
        "the method into Greek, Turkish and Arabic, and producing three "
        "territorial baselines. Estimated from the number of method sheets "
        "and sites rather than as a block allowance.",
    ("Staff Costs", "WP4"):
        "Design of three site installations from their own baselines, "
        "supervision of installation, one full visitor season of monitoring "
        "at each site, and development and testing of the shared platform "
        "protocol. Field seasons are the cost driver and they are fixed by "
        "the calendar, not by effort.",
    ("Staff Costs", "WP5"):
        "Adapting training and custodian material into three languages, "
        "delivering training in three countries to 60 site managers and "
        "local staff, and establishing and running nine school custodian "
        "groups. Estimated per delivery rather than per participant.",
    ("Staff Costs", "WP6"):
        "Drafting and costing three site management plans, running "
        "stakeholder consultation to a formal adoption decision at each "
        "site, constituting three stewardship committees, and writing the "
        "transferability guide and policy recommendations in English and "
        "Arabic.",
    ("Equipment costs", "WP1"):
        "IT equipment for project staff, purchased within the first year as "
        "the eligibility rules require. No equipment purchase is planned in "
        "the final year, so no Managing Authority authorisation is needed "
        "late in the project.",
    ("Equipment costs", "WP2"):
        "Equipment for producing and displaying communication material at "
        "the three sites, including on-site interpretation supports and the "
        "means to produce local-language versions.",
    ("Equipment costs", "WP3"):
        "Field survey and monitoring instruments specified from each site's "
        "own characterisation requirements - topographic and bathymetric "
        "survey, vegetation and sediment sampling. Specified after the "
        "transferability decision, so nothing is bought against an "
        "assumption.",
    ("Equipment costs", "WP4"):
        "Monitoring instruments installed at the three sites and the "
        "hardware needed for field data entry into the shared platform in "
        "each local language.",
    ("Equipment costs", "WP5"):
        "Teaching and field equipment for the custodian groups and the "
        "practical training modules held at the demonstration sites.",
    ("Equipment costs", "WP6"):
        "Equipment supporting the consultation meetings and the final "
        "transferability conference.",
    ("External expertise and services costs", "WP1"):
        "Auditor costs: the first-level control the Programme requires of "
        "every partner, which must be independent by definition and "
        "therefore cannot be staff. A FIXED cost, not a share of budget - "
        "four reporting periods, one certificate each, per partner, with the "
        "Applicant higher because it also consolidates six partners' claims. "
        "One line per partner, as error code 222 requires.",
    ("External expertise and services costs", "WP2"):
        "Website development and hosting, graphic design, certified "
        "translation into Turkish and Arabic, and local press distribution. "
        "Translation is bought in because the method's usability depends on "
        "precision in languages no partner holds certified capability in.",
    ("External expertise and services costs", "WP3"):
        "Drone and specialist bathymetric survey where a partner lacks the "
        "licensed capability, certified translation of the adapted method "
        "into three languages, and the visitor and willingness-to-pay survey "
        "fieldwork. Capability that no partner holds and none should staff up "
        "to acquire for 24 months.",
    ("External expertise and services costs", "WP4"):
        "Fabrication and installation of the demountable modular components, "
        "site consents and environmental screening support, and web-GIS "
        "platform development - where buying sixteen months of developer "
        "time would cost more than the deliverable.",
    ("External expertise and services costs", "WP5"):
        "Translation and design of training and custodian material in three "
        "languages, venue and logistics for training delivered in three "
        "countries, and specialist trainers where a partner lacks the "
        "pedagogic capability.",
    ("External expertise and services costs", "WP6"):
        "Facilitation of the stakeholder consultations, legal and "
        "administrative support for the formal adoption decisions in three "
        "national systems, editing and translation of the transferability "
        "guide into Arabic, and the final conference.",
    ("Infrastructures and works costs", "WP4"):
        "The demonstration installations themselves: boundary marking, "
        "interpretation signage, demountable modular boardwalks over the "
        "vulnerable surface and demarcated zones. Reversible by design, "
        "which is both an environmental requirement beside an Annex I "
        "habitat and a cost control - a section can be repaired rather than "
        "a structure replaced.",
}


def eur(x):
    return round(float(x), 2)


def brief(cat, wp, partner, first_staff):
    """The template's own wordings."""
    if cat == "Staff Costs":
        if wp == "WP1" and first_staff:
            return "Project Coordinator"
        if wp == "WP2" and first_staff:
            return "Communication Manager"
        return f"Human resources involved in {WP_NAME[wp]}"
    if cat == "Equipment costs":
        return f"Equipment necessary for {WP_NAME[wp]}"
    if cat == "Infrastructures and works costs":
        return f"Infrastructures and works necessary for {WP_NAME[wp]}"
    if wp == "WP1":
        return "Auditor costs"
    return f"External expertise and services necessary for {WP_NAME[wp]}"


def wp_months(wp):
    m = {w[0]: w[4] for w in WPS}[wp]
    lo, hi = (int(x) for x in re.findall(r"M(\d+)", m))
    return lo, hi


def sems(wp):
    lo, hi = wp_months(wp)
    return sorted({(x - 1) // 6 + 1 for x in range(lo, hi + 1)})


def allocate():
    """partner -> wp -> cat -> amount, preserving both margins.

    External expertise in WP1 is the audit and nothing else; the remainder of
    each partner's external expertise is spread over WP2-WP6. Everything else
    stays proportional to the partner's work package shares.
    """
    m = matrix()
    out = {}
    for k in P:
        share = {wp: v / DIRECT[k] for wp, v in m[k].items() if v > 0}
        rest = CC[k][CATS.index("External expertise and services costs")] \
            - AUDIT[k]
        assert rest > 0, f"{k}: audit {AUDIT[k]} exceeds its external expertise"
        non1 = sum(s for wp, s in share.items() if wp != "WP1")
        out[k] = {}
        for wp, s in share.items():
            row = {c: CC[k][i] * s for i, c in enumerate(CATS)}
            row["External expertise and services costs"] = (
                AUDIT[k] if wp == "WP1" else rest * s / non1)
            out[k][wp] = row
    return out, m


def main():
    alloc, m = allocate()
    wp_ids = [w[0] for w in WPS]
    wp_tot = {wp: sum(sum(alloc[k][wp].values()) for k in P if wp in alloc[k])
              for wp in wp_ids}

    # ------------------------------------------------------- the budget lines
    lines, seq = [], 100
    for wp in wp_ids:
        for cat in CATS:
            first_staff = True
            for k in sorted(P, key=lambda x: PSHORT[x]):
                v = alloc[k].get(wp, {}).get(cat, 0.0)
                if v <= 0.004:
                    continue
                seq += 1
                lines.append({
                    "code": f"{wp}.{CC_CODE[cat]}.{PSHORT[k]}.{seq}",
                    "wp": wp, "cat": cat, "partner": PCODE[k], "key": k,
                    "brief": brief(cat, wp, k, first_staff and k == "HCMR"),
                    "just": JUST.get((cat, wp), ""),
                    "total": eur(v), "sems": sems(wp),
                })
                if k == "HCMR":
                    first_staff = False

    fails = []
    # Every line must carry a justification - that is the column he asked for.
    for L in lines:
        if not L["just"]:
            fails.append(f"{L['code']} has no justification")
        if L["total"] <= 0:
            fails.append(f"{L['code']} total {L['total']}")
    # Reconcile in both directions.
    for k in P:
        for i, c in enumerate(CATS):
            got = round(sum(L["total"] for L in lines
                            if L["key"] == k and L["cat"] == c), 2)
            if abs(got - CC[k][i]) > 0.05:
                fails.append(f"{k} {c}: lines {got} vs model {CC[k][i]}")
        got = round(sum(L["total"] for L in lines if L["key"] == k), 2)
        if abs(got - DIRECT[k]) > 0.05:
            fails.append(f"{k} total {got} vs direct {DIRECT[k]}")
    for wp in wp_ids:
        got = round(sum(L["total"] for L in lines if L["wp"] == wp), 2)
        if abs(got - wp_tot[wp]) > 0.05:
            fails.append(f"{wp} lines {got} vs allocation {wp_tot[wp]}")
    for k in P:
        if abs(alloc[k]["WP1"]["External expertise and services costs"]
               - AUDIT[k]) > 0.01:
            fails.append(f"{k} WP1 external expertise is not exactly the "
                         f"audit - the whole point of the 09 Oct fix")
    grand = round(sum(L["total"] for L in lines), 2)
    if abs(grand - sum(DIRECT.values())) > 0.1:
        fails.append(f"grand {grand} vs {sum(DIRECT.values())}")
    # Rule checks the programme actually runs.
    for k in P:
        for wp in ("WP1", "WP2"):
            if not any(L["key"] == k and L["wp"] == wp for L in lines):
                fails.append(f"{k} has no line in {wp} - code 222")
        if not any(L["key"] == k and L["wp"] == "WP1"
                   and L["brief"] == "Auditor costs" for L in lines):
            fails.append(f"{k} has no auditor line in WP1 - code 222 "
                         f"requires audit for each partner")
    if any(L["cat"] == "Infrastructures and works costs"
           and L["wp"] in ("WP1", "WP2") for L in lines):
        fails.append("infrastructure appears in WP1 or WP2, which is "
                     "forbidden")
    # Duplicates: one line per partner per cost category per WP.
    seen = set()
    for L in lines:
        key = (L["wp"], L["cat"], L["key"])
        if key in seen:
            fails.append(f"duplicate line for {key}")
        seen.add(key)
    assert not fails, "NOT SHIPPING:\n  - " + "\n  - ".join(fails)

    # --------------------------------------------------------------- workbook
    wb = openpyxl.Workbook()
    HDR = Font(bold=True, color="FFFFFF", size=9)
    FILL = PatternFill("solid", fgColor="1F385E")
    WRAP = Alignment(wrap_text=True, vertical="top")
    MONEY = '#,##0.00'

    def head(ws, cols, widths):
        ws.append(cols)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
            c = ws.cell(row=1, column=i)
            c.font, c.fill, c.alignment = HDR, FILL, WRAP
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "Budget tot per budget line"
    head(ws, ["Budget-line code", "WP", "Cost category", "Partner N.",
              "Brief Description", "Justification of the estimated cost",
              "Total Costs (in EUR)", "Sem 1", "Sem 2", "Sem 3", "Sem 4"],
         [22, 6, 30, 13, 42, 78, 15, 7, 7, 7, 7])
    for L in lines:
        ws.append([L["code"], L["wp"], L["cat"], L["partner"], L["brief"],
                   L["just"], L["total"]]
                  + ["x" if s in L["sems"] else "-" for s in (1, 2, 3, 4)])
        r = ws.max_row
        ws.cell(row=r, column=6).alignment = WRAP
        ws.cell(row=r, column=5).alignment = WRAP
        ws.cell(row=r, column=7).number_format = MONEY
    ws.append(["TOTAL DIRECT", "", "", "", "", "", grand])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.cell(row=ws.max_row, column=7).font = Font(bold=True)
    ws.cell(row=ws.max_row, column=7).number_format = MONEY

    ws2 = wb.create_sheet("Budget per CC")
    head(ws2, ["Partner"] + CATS + ["Subtotal direct costs",
               "Travel 15% of staff", "Admin 15% of staff",
               "Total eligible"],
         [24, 14, 14, 26, 24, 20, 18, 18, 16])
    for k in sorted(P, key=lambda x: PSHORT[x]):
        row = [f"{PSHORT[k]} {P[k][0]}"] + [CC[k][i] for i in range(4)]
        row += [DIRECT[k], STAFF[k] * 0.15, STAFF[k] * 0.15,
                DIRECT[k] + STAFF[k] * FLAT]
        ws2.append(row)
        for col in range(2, 10):
            ws2.cell(row=ws2.max_row, column=col).number_format = MONEY
    ws2.append(["TOTAL"] + [sum(CC[k][i] for k in P) for i in range(4)]
               + [sum(DIRECT.values()), sum(STAFF.values()) * 0.15,
                  sum(STAFF.values()) * 0.15,
                  sum(DIRECT.values()) + sum(STAFF.values()) * FLAT])
    for col in range(1, 10):
        ws2.cell(row=ws2.max_row, column=col).font = Font(bold=True)
        if col > 1:
            ws2.cell(row=ws2.max_row, column=col).number_format = MONEY

    ws3 = wb.create_sheet("Budget per WP")
    head(ws3, ["Partner"] + wp_ids + ["Subtotal direct costs"],
         [24] + [14] * len(wp_ids) + [20])
    for k in sorted(P, key=lambda x: PSHORT[x]):
        ws3.append([f"{PSHORT[k]} {P[k][0]}"]
                   + [sum(alloc[k].get(wp, {}).values()) for wp in wp_ids]
                   + [DIRECT[k]])
        for col in range(2, len(wp_ids) + 3):
            ws3.cell(row=ws3.max_row, column=col).number_format = MONEY
    ws3.append(["TOTAL"] + [wp_tot[wp] for wp in wp_ids]
               + [sum(DIRECT.values())])
    for col in range(1, len(wp_ids) + 3):
        ws3.cell(row=ws3.max_row, column=col).font = Font(bold=True)
        if col > 1:
            ws3.cell(row=ws3.max_row, column=col).number_format = MONEY

    ws4 = wb.create_sheet("Co-financing table")
    head(ws4, ["Partner", "Total eligible cost", "% managed",
               "EU Contribution 92%", "Revenues",
               "Co-financing 8%", "SOURCE OF FUNDING - to be named by the "
               "partner"],
         [24, 20, 12, 20, 12, 18, 54])
    tot_e = sum(DIRECT[k] + STAFF[k] * FLAT for k in P)
    for k in sorted(P, key=lambda x: PSHORT[x]):
        e = DIRECT[k] + STAFF[k] * FLAT
        ws4.append([f"{PSHORT[k]} {P[k][0]}", e, e / tot_e,
                    e * (1 - COFIN), 0, e * COFIN, ""])
        r = ws4.max_row
        for col in (2, 4, 5, 6):
            ws4.cell(row=r, column=col).number_format = MONEY
        ws4.cell(row=r, column=3).number_format = '0.0%'
    ws4.append(["TOTAL", tot_e, 1.0, tot_e * (1 - COFIN), 0, tot_e * COFIN,
                ""])
    for col in range(1, 8):
        ws4.cell(row=ws4.max_row, column=col).font = Font(bold=True)
    ws4.cell(row=ws4.max_row, column=3).number_format = '0.0%'
    for col in (2, 4, 5, 6):
        ws4.cell(row=ws4.max_row, column=col).number_format = MONEY

    ws5 = wb.create_sheet("50% rule table")
    head(ws5, ["Note"], [120])
    mpc = sum(DIRECT[k] + STAFF[k] * FLAT for k in P if P[k][2] == "MPC")
    for t in [
        "This table justifies the shortfall when LESS than 50% of total "
        "eligible cost is allocated to partners in Mediterranean Partner "
        "Countries.",
        f"AMMOS allocates {eur(mpc)} EUR to MPC partners, which is "
        f"{100 * mpc / tot_e:.2f}% of total eligible cost of {eur(tot_e)} "
        f"EUR.",
        "That is above the 50% floor by partner allocation alone, so there is "
        "no shortfall to justify and this table is left empty - deliberately, "
        "not by omission.",
        "The margin above the floor is "
        f"{eur(mpc - 0.5 * tot_e)} EUR. If any MPC partner's budget is "
        "reduced, check this figure before agreeing the change.",
    ]:
        ws5.append([t])
        ws5.cell(row=ws5.max_row, column=1).alignment = WRAP

    wb.save(OUT)

    print(f"wrote {Path(OUT).name}")
    print(f"  {len(lines)} budget lines, every one with a justification")
    print(f"  reconciles: per partner per category, per partner, per WP, "
          f"grand total")
    for wp in wp_ids:
        n = sum(1 for L in lines if L["wp"] == wp)
        print(f"    {wp}  {n:>2} lines  {eur(wp_tot[wp]):>12,.2f}")
    print(f"  TOTAL DIRECT {grand:,.2f}   total eligible {tot_e:,.2f}")
    print(f"  MPC share {100 * mpc / tot_e:.2f}%  "
          f"co-financing {tot_e * COFIN:,.2f}")


if __name__ == "__main__":
    main()
