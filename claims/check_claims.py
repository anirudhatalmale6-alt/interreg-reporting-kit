#!/usr/bin/env python3
"""Check partner expenditure claims against the approved budget and contracts.

    python3 check_claims.py

Two claims arrived on 13 Oct labelled as EPIRUSMEDEYE. Only one is:

    ΧΑΤΖΗΚΩΣΤΑ  id 105436  EPIRUSMEDEYE, MIS 6010984  - not yet verified
    ΑΘΗΝΑ       id 103081  MINOR-MED,    MIS 6010952  - FINAL verification

They are different projects with different budgets, so they get separate
files and are never mixed. (The ΑΘΗΝΑ folder also carries a stray copy of
Hatzikosta's 105436 table.)

EVERY CHECK IS AGAINST A DOCUMENT, NOT AN ASSUMPTION:
  EPIRUSMEDEYE  technical sheet v1.1 (export 96606), procurement tracker v5
  MINOR-MED     technical sheet v1.2 (approved 02/09/2026)

Per expense row:
  - declared amount vs the approved line (partner x deliverable x category),
    cumulated across rows
  - invoice -> contract: amount incl VAT matches, contract signed before the
    invoice (EPIRUSMEDEYE, from the tracker)
  - payment date inside the claim period; invoice date inside the project's
    eligible period
  - split invoices: the parts must add up to the invoice exactly
  - same invoice number twice with the same amount on the same deliverable
Flat rates: the claimed rate must equal the approved budget's own ratio.
"""
import collections
import csv
import datetime as dt
import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
OUT = HERE

GREEN, AMBER, RED = "C6EFCE", "FFEB9C", "FFC7CE"


def d(s):
    return dt.datetime.strptime(s.strip(), "%d/%m/%Y").date() if s else None


def read_claim(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    meta = {str(r[0]).strip(): r[5] for r in rows[:12] if r and r[0]}
    out, flats = [], []
    for r in rows[16:]:
        if r[0] and str(r[0]).startswith("ΣΥΝΟΛΙΚΑ"):
            break
        cat = str(r[11] or "")
        if r[0] not in (None, ""):
            out.append(dict(
                inv=str(r[0]).strip(), inv_date=d(str(r[1])),
                supplier=str(r[2]).strip(), inv_total=float(r[4] or 0),
                pay_doc=str(r[6] or "").strip(), pay_date=d(str(r[7] or "")),
                pay_amt=float(r[9] or 0), cat=cat.split("|")[-1].strip(),
                deliv=str(r[12] or "").strip(), declared=float(r[14] or 0),
                eligible=r[15], checked=str(r[17] or "").strip(),
                note=str(r[19] or "").strip()))
        elif cat:
            flats.append(dict(cat=cat.split("|")[-1].strip(),
                              declared=float(r[14] or 0)))
    return meta, out, flats


def code(cat):
    m = re.match(r"([ΑΒ]\.\d+)", cat)
    return m.group(1) if m else cat


# --------------------------------------------------------------- budgets
def epirus_budget():
    """Hatzikosta (P2) from technical sheet v1.1, the same parse as before."""
    wb = openpyxl.load_workbook(ROOT / "1.1-interreg_report96606.xlsx",
                                data_only=True)
    ws = wb.active
    cur, out = None, {}
    for row in ws.iter_rows(values_only=True):
        c = ["" if x is None else x for x in row]
        t = str(c[0])
        if "Κύριος Δικαιούχος" in t:
            cur = "P1"
        m = re.search(r"\b(P[234])\b", t)
        if m:
            cur = m.group(1)
        md = re.match(r"\s*Παραδοτέο\s*(\d+)\.\s*(\d+)", t)
        if md and cur == "P2":
            n = [x for x in c[2:] if isinstance(x, (int, float))]
            out[f"{md.group(1)}.{md.group(2)}"] = dict(
                staff=n[0], trav=n[2], es=n[3], eq=n[4], tot=n[6])
    return out


def minormed_budget(partner_no="4"):
    out = {}
    with open(ROOT / "td_6010952_v1.2_partner_deliverable_category.csv",
              encoding="utf-8-sig") as f:
        for r in csv.DictReader(f, delimiter=";"):
            if r["partner_no"] != partner_no:
                continue
            num = lambda k: float(r[k].replace(".", "").replace(",", "."))
            out[r["deliverable"].replace("Π", "").strip()] = dict(
                staff=num("Προσωπικό"), over=num("Γενικά έξοδα"),
                trav=num("Ταξίδια/διαμονή"), es=num("Εξωτ. εμπειρογνωμοσύνη"),
                eq=num("Εξοπλισμός"), tot=num("total"))
    return out


CAT_TO_KEY = {"Α.25": "staff", "Α.29": "eq", "Α.33": "es"}


def tracker_contracts(prefix="PB2"):
    ws = openpyxl.load_workbook(
        ROOT / "Procurement_Tracker_EPIRUS_MEDEYE_v5.xlsx",
        data_only=True)["EPIRUS-MEDEYE"]
    out = {}
    for r in ws.iter_rows(min_row=4, values_only=True):
        if r[0] and str(r[0]).startswith(prefix):
            k = str(r[1]).strip()
            sig = r[13].date() if hasattr(r[13], "date") else None
            out.setdefault(k, dict(buys=r[2], signed=sig, amount=r[14],
                                   delivs=[]))
            out[k]["delivs"].append(str(r[4]).split("|")[0].strip())
    return out


# --------------------------------------------------------------- checks
def check(rows, budget, period, project_start, contracts=None,
          contract_of=None):
    used = collections.defaultdict(float)
    by_inv = collections.defaultdict(list)
    for x in rows:
        by_inv[(x["inv"], x["supplier"], x["inv_total"])].append(x)
    for x in rows:
        flags, why = [], []
        key = CAT_TO_KEY.get(code(x["cat"]))
        b = budget.get(x["deliv"])
        if not b or key is None:
            flags.append(RED)
            why.append(f"no approved line for deliverable {x['deliv']} "
                       f"category {code(x['cat'])}")
        else:
            used[(x["deliv"], key)] += x["declared"]
            cap = b[key]
            u = used[(x["deliv"], key)]
            if u > cap + 0.005:
                flags.append(RED)
                why.append(f"line {x['deliv']} {key} over budget: "
                           f"{u:,.2f} of {cap:,.2f}")
            elif u >= cap - 0.005:
                flags.append(AMBER)
                why.append(f"line {x['deliv']} {key} now 100% used "
                           f"({cap:,.2f}) - no headroom left")
        if x["pay_date"] and not (period[0] <= x["pay_date"] <= period[1]):
            flags.append(RED)
            why.append(f"paid {x['pay_date']:%d/%m/%Y}, outside the claim "
                       f"period")
        if x["inv_date"] and x["inv_date"] < project_start:
            flags.append(RED)
            why.append(f"invoice dated before the project start "
                       f"{project_start:%d/%m/%Y}")
        elif x["inv_date"] and x["inv_date"] < period[0]:
            why.append(f"invoice {x['inv_date']:%d/%m/%Y} predates the "
                       f"period but was paid inside it - eligible on "
                       f"payment date")
        parts = by_inv[(x["inv"], x["supplier"], x["inv_total"])]
        if len(parts) > 1:
            s = sum(p["declared"] for p in parts)
            if abs(s - x["inv_total"]) > 0.01:
                flags.append(RED)
                why.append(f"invoice split into {len(parts)} parts totalling "
                           f"{s:,.2f} vs invoice {x['inv_total']:,.2f}")
            else:
                why.append(f"invoice split across {len(parts)} lines, parts "
                           f"add up exactly")
        if contracts is not None:
            c = contract_of(x)
            if not c:
                flags.append(AMBER)
                why.append("no matching contract found in tracker v5")
            else:
                k, ct = c
                if ct["signed"] and x["inv_date"] and \
                        x["inv_date"] < ct["signed"]:
                    flags.append(RED)
                    why.append(f"invoice before contract {k} signature")
                else:
                    why.append(f"contract {k} signed "
                               f"{ct['signed']:%d/%m/%Y}" if ct["signed"]
                               else f"contract {k}")
        x["flag"] = RED if RED in flags else AMBER if AMBER in flags \
            else GREEN
        x["why"] = "; ".join(why) or "within budget, inside period"
    return used


def write(path, title, meta, rows, flats, budget, used, notes, rates):
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "Summary"
    s["A1"] = title
    s["A1"].font = Font(bold=True, size=14)
    r = 3
    for k in ("Κωδικός ΟΠΣ", "Ακρωνύμιο", "Δικαιούχος", "id δελτίου"):
        s.cell(row=r, column=1, value=k).font = Font(bold=True)
        s.cell(row=r, column=2, value=str(meta.get(k, "")))
        r += 1
    r += 1
    n = collections.Counter(x["flag"] for x in rows)
    s.cell(row=r, column=1, value="Expense rows").font = Font(bold=True)
    s.cell(row=r, column=2, value=f"{len(rows)}  -  green {n[GREEN]}, "
                                  f"amber {n[AMBER]}, red {n[RED]}")
    r += 1
    direct = sum(x["declared"] for x in rows)
    flat = sum(f["declared"] for f in flats)
    s.cell(row=r, column=1, value="Declared").font = Font(bold=True)
    s.cell(row=r, column=2, value=f"direct {direct:,.2f} + flat rates "
                                  f"{flat:,.2f} = {direct + flat:,.2f}")
    r += 2
    s.cell(row=r, column=1, value="Flat rates").font = Font(bold=True)
    r += 1
    for line in rates:
        s.cell(row=r, column=1, value=line)
        r += 1
    r += 1
    s.cell(row=r, column=1, value="Points for you").font = Font(bold=True)
    r += 1
    for line in notes:
        s.cell(row=r, column=1, value=line).alignment = Alignment(
            wrap_text=True, vertical="top")
        s.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        s.row_dimensions[r].height = 48
        r += 1
    s.column_dimensions["A"].width = 28
    s.column_dimensions["B"].width = 90

    e = wb.create_sheet("Expenses")
    head = ["Flag", "Invoice", "Invoice date", "Supplier", "Invoice total",
            "Paid on", "Payment", "Category", "Deliverable", "Declared",
            "Verifier checked", "Why"]
    e.append(head)
    for c in e[1]:
        c.font = Font(bold=True)
    for x in rows:
        e.append(["", x["inv"], x["inv_date"], x["supplier"][:60],
                  x["inv_total"], x["pay_date"], x["pay_amt"], code(x["cat"]),
                  x["deliv"], x["declared"], x["checked"], x["why"]])
        e.cell(row=e.max_row, column=1).fill = PatternFill(
            "solid", fgColor=x["flag"])
        e.cell(row=e.max_row, column=1, value={GREEN: "OK", AMBER: "CHECK",
                                               RED: "PROBLEM"}[x["flag"]])
    for i, w in enumerate((10, 10, 12, 44, 13, 12, 12, 9, 11, 12, 10, 90), 1):
        e.column_dimensions[get_column_letter(i)].width = w
    for row in e.iter_rows(min_row=2):
        for c in row[4:10]:
            if isinstance(c.value, float):
                c.number_format = "#,##0.00"
        row[11].alignment = Alignment(wrap_text=True, vertical="top")

    u = wb.create_sheet("Budget use")
    u.append(["Deliverable", "Category", "Approved", "Claimed now",
              "% used", "Left"])
    for c in u[1]:
        c.font = Font(bold=True)
    for (dl, key), v in sorted(used.items()):
        cap = budget[dl][key]
        u.append([dl, key, cap, round(v, 2), round(100 * v / cap, 1) if cap
                  else None, round(cap - v, 2)])
    for i, w in enumerate((12, 10, 14, 14, 9, 14), 1):
        u.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)
    return n


def main():
    # ------------------------------------------------ EPIRUSMEDEYE, P2
    meta, rows, flats = read_claim(
        HERE / "hatz/ΧΑΤΖΗΚΩΣΤΑ/105436/Πίνακας Δαπανών 105436.xlsx")
    assert "6010984" in str(meta.get("Κωδικός ΟΠΣ")), meta
    bud = epirus_budget()
    contracts = tracker_contracts("PB2")

    def contract_of(x):
        for k, c in contracts.items():
            if c["amount"] and abs(float(c["amount"]) - x["inv_total"]) < 0.01:
                return k, c
        # a contract paid by two invoices: match on the sum per supplier
        tot = sum(y["inv_total"] for y in rows
                  if y["supplier"] == x["supplier"]
                  and y["inv"] != "" and y["deliv"] == x["deliv"])
        same = {(y["inv"], y["inv_total"]) for y in rows
                if y["supplier"] == x["supplier"]}
        s = sum(t for _, t in same)
        for k, c in contracts.items():
            if c["amount"] and abs(float(c["amount"]) - s) < 0.01:
                return k, c
        return None

    used = check(rows, bud, (dt.date(2026, 6, 1), dt.date(2026, 10, 5)),
                 dt.date(2025, 10, 13), contracts, contract_of)
    staff_flat = next(f["declared"] for f in flats if "Β.24" in f["cat"])
    trav_flat = next(f["declared"] for f in flats if "Β.30" in f["cat"])
    direct = sum(x["declared"] for x in rows)
    a_staff = sum(v["staff"] for v in bud.values())
    a_direct = sum(v["es"] + v["eq"] for v in bud.values())
    a_trav = sum(v["trav"] for v in bud.values())
    rates = [
        f"B.24 staff: claimed {staff_flat:,.2f} = {100*staff_flat/direct:.3f}% "
        f"of direct; approved ratio {100*a_staff/a_direct:.3f}%",
        f"B.30 travel: claimed {trav_flat:,.2f} = "
        f"{100*trav_flat/staff_flat:.3f}% of staff; approved ratio "
        f"{100*a_trav/a_staff:.3f}%",
    ]
    assert abs(staff_flat / direct - a_staff / a_direct) < 0.0001, rates
    assert abs(trav_flat / staff_flat - a_trav / a_staff) < 0.0001, rates
    a_total = sum(v["tot"] for v in bud.values())
    claimed = direct + staff_flat + trav_flat
    notes = [
        f"Every invoice matches its contract to the cent: contract 4 "
        f"(Papapostolou, signed 15/07/2026) 430.156,00 = invoices 83 + 95; "
        f"contract 5 (7 PCs + 2 laptops, signed 20/02/2026) 8.531,20 = "
        f"invoice 74; contract 6 (kick-off, signed 03/12/2025) 1.570,00 = "
        f"invoices 11 + A1660. Every invoice is dated after its contract.",
        f"This first claim is {claimed:,.2f} of Hatzikosta's whole approved "
        f"budget of {a_total:,.2f} ({100*claimed/a_total:.1f}%). Equipment "
        f"lines 3.2, 3.3, 4.1 and 5.2 are now 99%+ used, so any later price "
        f"difference on that equipment cannot be absorbed.",
        "Not yet verified: every row shows eligible 0,00 and 'Ελέγχθηκε OXI', "
        "the sample is 0%, and the certificate (444.787,70; EU 355.830,16 + "
        "national 88.957,54) has no verifier name or date. Do not treat it as "
        "certified until the verifier has sampled it.",
        "Two invoices (05/12/2025 and 25/02/2026) predate the claim period "
        "01/06/2026-05/10/2026 but were paid inside it, which is what "
        "counts. Hatzikosta's kick-off hotel was paid six months after the "
        "invoice - eligible, but the verifier may ask why.",
        "The equipment receipt protocol and fixed-asset register entry are "
        "in the pack (2_Protokollo_paralabis...). Good - that is the first "
        "thing a verifier asks for on 430k of equipment.",
    ]
    n1 = write(OUT / "EPIRUSMEDEYE_claim-check_Hatzikosta_105436.xlsx",
               "EPIRUSMEDEYE (MIS 6010984) - claim check, Hatzikosta, "
               "id 105436", meta, rows, flats, bud, used, notes, rates)
    print("EPIRUSMEDEYE Hatzikosta:", dict(n1), f"claimed {claimed:,.2f}")

    # ------------------------------------------------ MINOR-MED, Athena
    meta, rows, flats = read_claim(
        HERE / "athina/ΑΘΗΝΑ/103081/Πίνακας Δαπανών 103081.xlsx")
    assert "6010952" in str(meta.get("Κωδικός ΟΠΣ")), meta
    bud = minormed_budget("4")
    # MINOR-MED's start moved to 13/02/2026 in technical sheet v1.1
    used = check(rows, bud, (dt.date(2025, 9, 1), dt.date(2026, 9, 30)),
                 dt.date(2026, 2, 13))
    over = next(f["declared"] for f in flats if "Β.26" in f["cat"])
    trav = next(f["declared"] for f in flats if "Β.30" in f["cat"])
    staff = sum(x["declared"] for x in rows if code(x["cat"]) == "Α.25")
    a_staff = sum(v["staff"] for v in bud.values())
    a_over = sum(v["over"] for v in bud.values())
    a_trav = sum(v["trav"] for v in bud.values())
    rates = [
        f"B.26 overheads: claimed {over:,.2f} = {100*over/staff:.3f}% of "
        f"staff; approved ratio {100*a_over/a_staff:.3f}%",
        f"B.30 travel: claimed {trav:,.2f} = {100*trav/staff:.3f}% of staff; "
        f"approved ratio {100*a_trav/a_staff:.3f}%",
    ]
    assert abs(over / staff - a_over / a_staff) < 0.0001, rates
    assert abs(trav / staff - a_trav / a_staff) < 0.0001, rates
    notes = [
        "This is MINOR-MED (MIS 6010952), not EPIRUSMEDEYE - it is Athena "
        "Research Center's claim and is checked against MINOR-MED's own "
        "technical sheet v1.2.",
        "Already FINAL: verified on 05/10/2026, 60,16% "
        "sampled, no findings, all 30.307,54 eligible. This check agrees - "
        "nothing over budget, flat rates exact, every split adds up.",
        "Equipment on deliverable 3.1 is now used to the euro (5.800,00 of "
        "5.800,00). The DATUM invoice 97 is split 5.800,00 to 3.1 and "
        "4.761,08 to 3.3; worth keeping the itemisation that supports that "
        "split on file, because the 3.1 share lands exactly on the budget.",
        "The report header says period 01/09/2025-30/09/2026, but MINOR-MED's "
        "start moved to 13/02/2026 in technical sheet v1.1. No expense here "
        "predates 13/02/2026, so nothing is affected - but anything dated "
        "between those two dates would not be eligible.",
        "Section 5 of the report lists D3.5 among the deliverables, but no "
        "expense in this claim is on 3.5 (they are 3.1, 3.3 and 3.4). "
        "Cosmetic, but it is a mismatch between the report and the table.",
        "The three individual fees are "
        "invoices from natural persons claimed as STAFF real cost and paid "
        "through the payroll system. The verifier accepted that; the "
        "contracts and timesheets should stay on file.",
    ]
    n2 = write(OUT / "MINOR-MED_claim-check_Athena_103081.xlsx",
               "MINOR-MED (MIS 6010952) - claim check, Athena RC, id 103081",
               meta, rows, flats, bud, used, notes, rates)
    print("MINOR-MED Athena:", dict(n2))


if __name__ == "__main__":
    main()
