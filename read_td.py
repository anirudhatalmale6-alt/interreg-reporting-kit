#!/usr/bin/env python3
"""Read a MIS Technical Sheet PDF (Τεχνικό Δελτίο Πράξης) into tables.

YOUR QUESTION: "If I want financial spending analytically for every cost
category of every partner? Every partner has 3 tech sheets so far uploaded to
MIS 1.0, 1.1, 1.2?"

THE ANSWER IS THAT YOU ALREADY HAVE IT AND NOBODY HAS TO TYPE IT AGAIN.

Part D of the technical sheet contains a section called ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΑΝΑ
ΔΙΚΑΙΟΥΧΟ / ΠΑΡΑΔΟΤΕΑ / ΚΑΤΗΓΟΡΙΑ ΔΑΠΑΝΗΣ.  That is exactly what you asked
for: every partner, every deliverable, every cost category, one number per
cell.  It is not a picture, it is real text inside the PDF, so it can be read.

So this script takes the PDF the MIS gives you, and gives you back:

    td_<MIS>_v<version>_partner_deliverable_category.csv
    td_<MIS>_v<version>_wp_category.csv
    td_<MIS>_v<version>_summary.txt

and the CSVs open in Excel with one row per partner/deliverable/category.
Pivot them however the report needs.

ONE CORRECTION TO THE PREMISE, AND IT MATTERS

You said "financial spending".  This is NOT spending.  The technical sheet is
the APPROVED BUDGET - what you are allowed to spend.  Spending comes from the
MIS payment/verification reports (ΔΑΠΑΝΕΣ / ΑΙΤΗΜΑΤΑ ΕΠΑΛΗΘΕΥΣΗΣ), which are a
different export.  If a report puts budget figures under a heading that says
spending, the reader concludes the money is gone and it is not.

What you can legitimately build from these three PDFs alone is:
  * the approved budget, analytically, per partner and per cost category
  * what the three approved versions CHANGED (see diff_td.py)
and nothing about absorption.  For absorption you need the expenditure export
and the procurement tracker, which is what build_charts.py already reads.

WHY THE ASSERTIONS ARE THE POINT OF THIS SCRIPT

A PDF table read by text position can silently drop a cell - a number lands a
few points to the left, two columns merge, and you get a plausible table whose
total is wrong.  You would not see it.  So every table here is checked against
every other table that ought to agree with it:

    partner deliverable rows  ->  partner total printed in the PDF
    partner totals            ->  grand total printed in the PDF
    work package budgets      ->  grand total
    work package x category   ->  work package budgets, both ways
    EU contribution + national -> total, and the EU rate

If any of those disagree the script STOPS and tells you which cell to look at.
A parser without these checks is worse than typing it by hand, because typing
by hand you at least know you might have made a mistake.
"""
import csv
import re
import subprocess
import sys

from _shared import HERE, Report, newest

# ---------------------------------------------------------------- NUMBERS
# Greek convention in, Python float out.  "210.631,37" -> 210631.37
NUM = r"-?\d{1,3}(?:\.\d{3})*,\d{2}"


def gr(text):
    return float(str(text).replace(".", "").replace(",", "."))


def money(x):
    """Back out to Greek convention for printing: 887.786,27"""
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","}))


# The six cost categories of this programme, in the order the MIS prints them.
CATS = ["Προσωπικό", "Γενικά έξοδα", "Ταξίδια/διαμονή",
        "Εξωτ. εμπειρογνωμοσύνη", "Εξοπλισμός", "Υποδομή/επενδύσεις"]


def text_of(pdf):
    """pdftotext -layout, because the column layout IS the table structure."""
    try:
        out = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                             capture_output=True, check=True)
    except FileNotFoundError:
        sys.exit("pdftotext is missing. Install poppler-utils:\n"
                 "  Windows: choco install poppler   or download poppler and "
                 "put its bin folder on PATH\n"
                 "  Ubuntu:  sudo apt install poppler-utils")
    return out.stdout.decode("utf-8", "replace")


# ---------------------------------------------------------------- HEADER
def header(t):
    def one(pat, label):
        m = re.search(pat, t)
        assert m, f"cannot find {label} in this PDF. Is it a Τεχνικό Δελτίο?"
        return m.group(1).strip()

    return {
        "mis": one(r"ΚΩΔΙΚΟΣ MIS:\s*(\d+)", "the MIS code"),
        "acronym": one(r"ΑΚΡΩΝΥΜΟ ΠΡΑΞΗΣ:\s*(.+)", "the acronym"),
        "title": one(r"ΤΙΤΛΟΣ ΠΡΑΞΗΣ:\s*(.+)", "the title"),
        "start": one(r"Έναρξη:\s*(\d{2}/\d{2}/\d{4})", "the start date"),
        "end": one(r"Λήξη:\s*(\d{2}/\d{2}/\d{4})", "the end date"),
        "months": one(r"Συνολική Διάρκεια \(σε μήνες\):\s*(" + NUM + ")",
                      "the duration"),
        # "ΕΚΔΟΣΗ: 1. 2" - the MIS prints a space inside its own version number
        "version": re.sub(r"\s+", "", one(r"ΕΚΔΟΣΗ:\s*([\d.\s]+?)\s{2,}",
                                          "the version")),
        "kind": one(r"ΕΚΔΟΣΗ:[\d.\s]+\s{2,}(\S+)", "Αρχική/Επικαιροποίηση"),
        "approved": one(r"ΗΜΕΡΟΜΗΝΙΑ ΕΓΚΡΙΣΗΣ:\s*(\d{2}/\d{2}/\d{4})",
                        "the approval date"),
        # The two-column layout interleaves these: the label ΗΜΕΡΟΜΗΝΙΑ
        # ΥΠΟΒΟΛΗΣ sits on one line, the whole ΗΜΕΡΟΜΗΝΙΑ ΕΓΚΡΙΣΗΣ line lands
        # between label and value, and the submitted date arrives two lines
        # later. Read it as it is printed rather than as it looks on screen.
        "submitted": one(r"ΗΜΕΡΟΜΗΝΙΑ ΥΠΟΒΟΛΗΣ:(?:[^\n]*\n){0,2}\s*"
                         r"(\d{2}/\d{2}/\d{4}(?: \d{2}:\d{2})?)",
                         "the submission date"),
    }


# ------------------------------------------------------- WRAPPED CELL TITLES
# The MIS vertically CENTRES a long title on its own row, so a three line
# title prints as
#
#          TECHNOLOGY DEVELOPMENT, EQUIPMENT          <- above the number
#     3                                    dates  371.753,89
#          PROCUREMENT, AND IMPLEMENTATION            <- below the number
#
# and a continuation line sitting between two numbered rows could belong to
# either of them.  My first attempt just took the line above and the line
# below every row, which gave WP1 the first line of WP2's title and gave WP5
# the last word of WP4's.  Both looked like plausible titles, which is exactly
# why that class of bug survives a read-through.
#
# The rule that is actually right: attribute each run of continuation lines to
# the row that is MISSING a title, and split the run down the middle when both
# neighbours are missing one.  A row that already printed its title on its own
# line is complete and takes nothing.
_BARRIER = re.compile(r"ΤΙΤΛΟΣ|Κωδικός|ΔΕΙΚΤΕΣ|Α/Α|Μονάδα|ΣΥΝΟΛ|Ειδικός|"
                      r"Άξονας|ΠΑΚΕΤΑ")


def wrap_titles(lines, rows):
    """rows is [(line_index, own_title)]; returns the full title per row."""
    titles = {i: ([t] if t else []) for i, t in rows}
    idx = [i for i, _ in rows]
    own = dict(rows)

    # runs of continuation lines, delimited by numbered rows and by barriers
    run, prev_row = [], None
    bounds = []                                  # (prev_row, run, next_row)
    for n, line in enumerate(lines):
        txt = line.strip()
        if n in titles:
            bounds.append((prev_row, run, n))
            run, prev_row = [], n
        elif not txt or _BARRIER.search(txt):
            bounds.append((prev_row, run, None))
            run, prev_row = [], None
        else:
            run.append(txt)
    bounds.append((prev_row, run, None))

    for a, lines_between, b in bounds:
        if not lines_between:
            continue
        if a is None and b is None:
            continue
        if a is None:                            # only a following row exists
            titles[b][:0] = lines_between
        elif b is None:                          # only a preceding row exists
            titles[a].extend(lines_between)
        elif own[a] and not own[b]:
            titles[b][:0] = lines_between
        elif own[b] and not own[a]:
            titles[a].extend(lines_between)
        elif not own[a] and not own[b]:
            half = len(lines_between) // 2 + len(lines_between) % 2
            titles[a].extend(lines_between[:half])
            titles[b][:0] = lines_between[half:]
        # both rows already carry their own title: the run belongs to neither,
        # so it is dropped rather than guessed onto one of them.

    out = {}
    for i in idx:
        s = " ".join(x for x in titles[i] if x)
        # "CAPACITY- BUILDING" was one word before the line break
        out[i] = re.sub(r"(?<=[A-Za-zΑ-Ωα-ω])-\s+(?=[A-Za-zΑ-Ωα-ω])", "-", s)
    return out


def section(t, start, *ends):
    """The slab of text between one heading and whichever end comes first."""
    i = t.find(start)
    assert i >= 0, f"heading not found: {start!r}"
    i += len(start)
    j = min([k for k in (t.find(e, i) for e in ends) if k >= 0] or [len(t)])
    return t[i:j]


# ------------------------------------------------------- BUDGET BY PARTNER
def by_partner(t):
    """Part A: EU contribution, national contribution and total, per partner.

    This table is keyed by row number and country only - the MIS does not
    repeat the partner name here - so it is used for the funding split and the
    partner NAMES come from Part D below.
    """
    s = section(t, "ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΑΝΑ ΔΙΚΑΙΟΥΧΟ", "ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΑΝΑ ΧΩΡΑ")
    rows = []
    for line in s.splitlines():
        m = re.match(r"\s*(\d+)\s+(\S[^\d]*?)\s+(" + NUM + r")\s+(" + NUM +
                     r")\s+(" + NUM + r")\s+(" + NUM + r")\s+(" + NUM +
                     r")\s+(" + NUM + r")", line)
        if m:
            rows.append({"no": int(m.group(1)), "country": m.group(2).strip(),
                         "eu": gr(m.group(3)), "eu_pc": gr(m.group(4)),
                         "nat": gr(m.group(5)), "nat_pc": gr(m.group(6)),
                         "nat_pub": gr(m.group(7)), "total": gr(m.group(8))})
    assert rows, "the per-partner funding table came back empty"
    m = re.search(r"ΣΥΝΟΛΑ\s+(" + NUM + r")\s+(" + NUM + r")\s+(" + NUM +
                  r")\s+(" + NUM + r")\s+(" + NUM + ")", s)
    assert m, "cannot find the ΣΥΝΟΛΑ line of the per-partner funding table"
    return rows, {"eu": gr(m.group(1)), "nat": gr(m.group(2)),
                  "total": gr(m.group(5))}


# ----------------------------------------------------------- WORK PACKAGES
def work_packages(t):
    """Part D: work package number, dates and budget.

    The title wraps onto the line ABOVE the number in the PDF, so the title is
    taken from the WP x CC table instead, where it is not needed at all.  What
    matters here are the dates and the budget.
    """
    s = section(t, "ΠΑΚΕΤΑ ΕΡΓΑΣΙΑΣ", "ΣΥΝΟΛΙΚΟΣ ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΑΝΑ ΠΕ")
    lines = s.splitlines()
    wps, rows = [], []
    for i, line in enumerate(lines):
        m = re.match(r"\s*(\d+)\s+(.*?)\s{2,}(\d{2}/\d{2}/\d{4})\s+"
                     r"(\d{2}/\d{2}/\d{4})\s+(" + NUM + ")", line)
        if not m:
            continue
        rows.append((i, m.group(2).strip()))
        wps.append({"wp": int(m.group(1)), "line": i,
                    "start": m.group(3), "end": m.group(4),
                    "budget": gr(m.group(5))})
    assert wps, "the work package table came back empty"
    titles = wrap_titles(lines, rows)
    for w in wps:
        w["title"] = titles[w.pop("line")]
    m = re.search(r"ΣΥΝΟΛΟ\s+(" + NUM + ")", s)
    assert m, "cannot find the ΣΥΝΟΛΟ of the work package table"
    return wps, gr(m.group(1))


def wp_by_category(t):
    """Part D: work package x cost category, six categories plus a total."""
    s = section(t, "ΣΥΝΟΛΙΚΟΣ ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΑΝΑ ΠΕ",
                "ΧΡΗΜΑΤΟΔΟΤΗΣΗ ΠΡΑΞΗΣ")
    rows, totals = [], None
    for line in s.splitlines():
        m = re.match(r"\s*ΠΕ\s*(\d+)\s+((?:" + NUM + r"\s+){6})(" + NUM + ")",
                     line)
        if m:
            rows.append({"wp": int(m.group(1)),
                         "cats": [gr(x) for x in re.findall(NUM, m.group(2))],
                         "total": gr(m.group(3))})
            continue
        m = re.match(r"\s*ΣΥΝΟΛΑ\s+((?:" + NUM + r"\s+){6})(" + NUM + ")",
                     line)
        if m:
            totals = {"cats": [gr(x) for x in re.findall(NUM, m.group(1))],
                      "total": gr(m.group(2))}
    assert rows, "the work package x cost category table came back empty"
    assert totals, "cannot find the ΣΥΝΟΛΑ line of the WP x category table"
    return rows, totals


# ------------------------------------ THE ANALYTICAL TABLE YOU ASKED FOR
def analytical(t):
    """Part D: partner x deliverable x cost category.

    Structure in the PDF:

        1. ΕΠΙΤΕΛΙΚΗ ΔΟΜΗ ΕΣΠΑ ΥΠΟΥΡΓΕΙΟΥ ΥΓΕΙΑΣ      <- partner
        ΠΕ1   PROJECT MANAGEMENT AND COORDINATION       <- work package
        Π 1.2 Administrative Management  2.845,71 ...   <- the row we want
                                            ΣΥΝΟΛΟ  210.631,37   <- partner sum

    A deliverable row is recognised by "Π n.n" followed by exactly seven
    numbers.  Seven, not "at least seven": a row that yielded six or eight has
    been mis-split by the layout and must not be silently accepted.
    """
    s = section(t, "ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΑΝΑ ΔΙΚΑΙΟΥΧΟ / ΠΑΡΑΔΟΤΕΑ",
                "ΜΕΡΟΣ Ε", "ΜΕΡΟΣ ΣΤ")
    partners, cur = [], None
    for line in s.splitlines():
        # a partner heading: "3. ΓΕΝΙΚΟ ΝΟΣΟΚΟΜΕΙΟ ΦΙΛΙΑΤΩΝ"
        m = re.match(r"\s*(\d+)\.\s+([A-ZΑ-ΩΆ-Ώ][^\d]{4,})$", line)
        if m and "ΣΥΝΟΛΟ" not in line:
            cur = {"no": int(m.group(1)), "name": m.group(2).strip(),
                   "rows": [], "printed_total": None}
            partners.append(cur)
            continue
        if cur is None:
            continue
        # a work package heading inside a partner block
        m = re.match(r"\s*ΠΕ\s*(\d+)\s*(.*)$", line)
        if m and not re.search(NUM, line):
            cur["wp"] = int(m.group(1))
            continue
        # the deliverable row
        m = re.match(r"\s*Π\s*(\d+\.\d+)\s+(.*?)\s{2,}((?:" + NUM +
                     r"\s+){6})(" + NUM + r")\s*$", line)
        if m:
            cats = [gr(x) for x in re.findall(NUM, m.group(3))]
            assert len(cats) == 6, (
                f"deliverable {m.group(1)} of partner {cur['no']} produced "
                f"{len(cats)} cost-category cells instead of 6. The PDF layout "
                f"has merged or split a column on this line:\n  {line!r}")
            cur["rows"].append({"deliv": "Π " + m.group(1),
                                "wp": cur.get("wp"),
                                "title": m.group(2).strip(),
                                "cats": cats, "total": gr(m.group(4))})
            continue
        m = re.search(r"ΣΥΝΟΛΟ\s+(" + NUM + r")\s*$", line)
        if m:
            cur["printed_total"] = gr(m.group(1))
    assert partners, ("the analytical table came back empty. Check that this "
                      "PDF is the full technical sheet and not an extract.")
    return partners


def indicators(t):
    s = section(t, "ΜΕΡΟΣ ΣΤ: ΔΕΙΚΤΕΣ", "ΜΕΡΟΣ Ζ")
    lines = s.splitlines()
    out, rows = [], []
    for i, line in enumerate(lines):
        m = re.match(r"\s*(RC[OR]\d+)\s+(.*?)\s{2,}(\S+)\s{2,}(\d+)\s*$", line)
        if not m:
            continue
        rows.append((i, m.group(2).strip()))
        out.append({"code": m.group(1), "line": i, "unit": m.group(3),
                    "target": int(m.group(4)),
                    "kind": ("output" if m.group(1).startswith("RCO")
                             else "result")})
    # Same wrapping as the work packages, and here a title can wrap onto BOTH
    # sides of its own code, so the run-splitting rule earns its keep.
    titles = wrap_titles(lines, rows)
    for d in out:
        d["title"] = titles[d.pop("line")]
    return out


# ---------------------------------------------------------------- THE CHECKS
def verify(out, hdr, fund_rows, fund_tot, wps, wp_grand, wpcc, wpcc_tot,
           partners):
    """Every table against every other table that must agree with it."""
    eps = 0.02          # two cents, for rounding in the MIS itself
    fails = []

    def chk(label, a, b):
        if abs(a - b) > eps:
            fails.append(f"{label}: computed {money(a)} vs printed {money(b)}, "
                         f"difference {money(a - b)}")
        out(f"  {'ok  ' if abs(a - b) <= eps else 'FAIL'} {label:<58s} "
            f"{money(b):>14s}")

    out()
    out("CROSS-CHECKS")
    out("-" * 78)
    chk("funding rows sum to the printed ΣΥΝΟΛΑ",
        sum(r["total"] for r in fund_rows), fund_tot["total"])
    chk("EU + national equals the total",
        fund_tot["eu"] + fund_tot["nat"], fund_tot["total"])
    chk("work package budgets sum to the grand total",
        sum(w["budget"] for w in wps), wp_grand)
    chk("work package grand total equals the funding total",
        wp_grand, fund_tot["total"])

    for r in wpcc:
        chk(f"WP{r['wp']} cost categories sum to its own total",
            sum(r["cats"]), r["total"])
    for r in wpcc:
        w = next((x for x in wps if x["wp"] == r["wp"]), None)
        if w:
            chk(f"WP{r['wp']} category total equals its work package budget",
                r["total"], w["budget"])
    for i, cat in enumerate(CATS):
        chk(f"column total, {cat}",
            sum(r["cats"][i] for r in wpcc), wpcc_tot["cats"][i])

    for p in partners:
        if p["printed_total"] is None:
            fails.append(f"partner {p['no']} {p['name']}: no ΣΥΝΟΛΟ line found")
            continue
        chk(f"partner {p['no']} deliverable rows sum to its ΣΥΝΟΛΟ",
            sum(r["total"] for r in p["rows"]), p["printed_total"])
        for r in p["rows"]:
            if abs(sum(r["cats"]) - r["total"]) > eps:
                fails.append(
                    f"partner {p['no']} {r['deliv']}: cost categories sum to "
                    f"{money(sum(r['cats']))} against a printed row total of "
                    f"{money(r['total'])}")
    chk("partner totals sum to the grand total",
        sum(p["printed_total"] or 0 for p in partners), fund_tot["total"])

    # the analytical table aggregated by work package must reproduce the WP
    # table.  This is the check that catches a whole deliverable row lost to a
    # page break, which no per-row check can see.
    agg = {}
    for p in partners:
        for r in p["rows"]:
            agg[r["wp"]] = agg.get(r["wp"], 0.0) + r["total"]
    for w in wps:
        chk(f"WP{w['wp']} from the analytical table equals the WP table",
            agg.get(w["wp"], 0.0), w["budget"])

    # the funding table is keyed by row number, the analytical table by partner
    # number.  They are the same partner in the same order, so the totals must
    # match row by row - and if the MIS ever reorders one of them, this fails
    # loudly instead of pairing the wrong partner with the wrong money.
    for f, p in zip(fund_rows, partners):
        chk(f"partner {p['no']} funding row equals its analytical ΣΥΝΟΛΟ",
            f["total"], p["printed_total"] or 0)

    # ---------------------------------------------------------------- FLAGS
    # These are LEGAL but need explaining in a report, so they are printed as
    # flags rather than failures. The one that found something real here: when
    # the implementation period was shifted by modification, five of six work
    # packages moved with it and one did not, so WP5 still carries dates from
    # the superseded calendar. That is invisible in any single version - it
    # only shows up when you ask whether each WP sits inside the project.
    def d(s):
        p = s.split("/")
        return (int(p[2]), int(p[1]), int(p[0]))

    flags = []
    ps, pe = d(hdr["start"]), d(hdr["end"])
    for w in wps:
        ws, we = d(w["start"]), d(w["end"])
        if ws < ps or we > pe:
            flags.append(f"WP{w['wp']} runs {w['start']} to {w['end']}, "
                         f"OUTSIDE the project period {hdr['start']} to "
                         f"{hdr['end']}")
        elif (ws, we) != (ps, pe):
            flags.append(f"WP{w['wp']} runs {w['start']} to {w['end']}, "
                         f"inside but not equal to the project period "
                         f"{hdr['start']} to {hdr['end']}")
    # A work package with no procurement-bearing category is not a problem,
    # but a work package with NO money at all means a row was lost.
    for w in wps:
        if w["budget"] <= 0:
            fails.append(f"WP{w['wp']} has a budget of zero, which almost "
                         f"certainly means its row was not read")
    if flags:
        out()
        out("FLAGS, legal but they need a sentence in the report")
        out("-" * 78)
        for f in flags:
            out(f"  ! {f}")
    return fails


# ---------------------------------------------------------------- OUTPUT
def main(pdf=None):
    pdf = pdf or newest("*.pdf")
    t = text_of(pdf)
    hdr = header(t)
    fund_rows, fund_tot = by_partner(t)
    wps, wp_grand = work_packages(t)
    wpcc, wpcc_tot = wp_by_category(t)
    partners = analytical(t)
    inds = indicators(t)

    tag = f"{hdr['mis']}_v{hdr['version']}"
    out = Report()
    out("=" * 78)
    out(f"{hdr['acronym']}, MIS {hdr['mis']}, technical sheet version "
        f"{hdr['version']} ({hdr['kind']})")
    out("=" * 78)
    out(f"title      {hdr['title']}")
    out(f"duration   {hdr['start']} to {hdr['end']}, {hdr['months']} months")
    out(f"submitted  {hdr['submitted']}")
    out(f"approved   {hdr['approved']}")
    out(f"source     {pdf.name}")
    out()
    out("APPROVED BUDGET")
    out("-" * 78)
    out(f"  EU contribution      {money(fund_tot['eu']):>14s}   "
        f"{fund_tot['eu'] / fund_tot['total']:.2%}")
    out(f"  national             {money(fund_tot['nat']):>14s}   "
        f"{fund_tot['nat'] / fund_tot['total']:.2%}")
    out(f"  TOTAL ELIGIBLE       {money(fund_tot['total']):>14s}")
    out()
    out("BY PARTNER")
    for p in partners:
        f = fund_rows[p["no"] - 1] if p["no"] <= len(fund_rows) else {}
        out(f"  {p['no']}. {p['name'][:52]:<52s} "
            f"{money(p['printed_total'] or 0):>14s}  "
            f"{f.get('country', '')}")
    out()
    out("BY COST CATEGORY")
    for i, cat in enumerate(CATS):
        v = wpcc_tot["cats"][i]
        out(f"  {cat:<26s} {money(v):>14s}   {v / wpcc_tot['total']:6.2%}")
    out(f"  {'TOTAL':<26s} {money(wpcc_tot['total']):>14s}")
    out()
    out("BY WORK PACKAGE")
    for w in wps:
        out(f"  WP{w['wp']} {w['title'][:44]:<44s} "
            f"{money(w['budget']):>14s}   {w['start']} to {w['end']}")
    if inds:
        out()
        out("INDICATORS")
        for d in inds:
            out(f"  {d['code']:<8s} {d['title'][:44]:<44s} "
                f"{'' if d['target'] is None else d['target']} {d['unit']}")

    fails = verify(out, hdr, fund_rows, fund_tot, wps, wp_grand, wpcc,
                   wpcc_tot, partners)

    # ------------------------------------------------------------ the CSVs
    f1 = HERE / f"td_{tag}_partner_deliverable_category.csv"
    with open(f1, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["MIS", "version", "partner_no", "partner", "country", "wp",
                    "deliverable", "deliverable_title"] + CATS + ["total"])
        for p in partners:
            country = (fund_rows[p["no"] - 1]["country"]
                       if p["no"] <= len(fund_rows) else "")
            for r in p["rows"]:
                w.writerow([hdr["mis"], hdr["version"], p["no"], p["name"],
                            country, r["wp"], r["deliv"], r["title"]]
                           + [f"{v:.2f}".replace(".", ",") for v in r["cats"]]
                           + [f"{r['total']:.2f}".replace(".", ",")])

    f2 = HERE / f"td_{tag}_wp_category.csv"
    with open(f2, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["MIS", "version", "wp", "title", "start", "end"] + CATS
                   + ["total"])
        for r in wpcc:
            wp = next((x for x in wps if x["wp"] == r["wp"]), {})
            w.writerow([hdr["mis"], hdr["version"], r["wp"],
                        wp.get("title", ""), wp.get("start", ""),
                        wp.get("end", "")]
                       + [f"{v:.2f}".replace(".", ",") for v in r["cats"]]
                       + [f"{r['total']:.2f}".replace(".", ",")])

    out()
    out(f"[written] {f1.name}")
    out(f"[written] {f2.name}")
    out.save(HERE / f"td_{tag}_summary.txt")

    if fails:
        print("\n" + "!" * 78)
        for f in fails:
            print("  " + f)
        raise AssertionError(
            f"{len(fails)} cross-check(s) failed. The CSVs were written so you "
            f"can look, but DO NOT put these numbers in a report until the "
            f"failures above are explained.")
    print(f"\n[all cross-checks passed on {pdf.name}]")
    return hdr, partners, fund_rows, fund_tot, wps, wpcc, wpcc_tot, inds


if __name__ == "__main__":
    import pathlib
    main(pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else None)
