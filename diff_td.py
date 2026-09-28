#!/usr/bin/env python3
"""Compare technical sheet versions and say what each modification changed.

    python3 diff_td.py report_26346.pdf report_118212.pdf report_140109.pdf

or with no arguments at all, in which case it takes every PDF beside it that
is a technical sheet, sorts them by version, and walks the chain.

WHY THIS IS THE SCRIPT THAT EARNS ITS KEEP, NOT THE PARSER

You have three approved versions of one technical sheet.  The question a
progress report has to answer is not "what does the budget say" - it is "what
changed, when was it approved, and what has to be reported differently as a
result".  Nobody can answer that by opening three 28 page PDFs side by side,
and the person who tries will miss something in the middle of page 20.

So this compares them field by field and prints only the differences.  The
sections it walks are the ones a modification can legally touch: duration,
budget per partner, budget per cost category, budget per work package, work
package dates, indicator targets, legal representatives, bank details, and
the list of documents attached to the sheet.

THE FINDING THAT MATTERS MOST IS USUALLY A ZERO

If a modification was approved and NOTHING in the budget moved, that is a
finding, not an absence of one.  It means the modification was administrative,
and a report that describes it as a budget revision is wrong.  The summary at
the end says so explicitly rather than leaving you to infer it from an empty
list.
"""
import pathlib
import re
import sys

from _shared import HERE, Report
from read_td import (analytical, by_partner, gr, header, indicators, money,
                     section, text_of, work_packages, wp_by_category, CATS,
                     NUM)

EPS = 0.02


# --------------------------------------------- fields a modification can move
def attachments(t):
    """The ΜΕΡΟΣ Η list: which documents are attached to THIS version."""
    try:
        s = section(t, "ΚΑΤΑΛΟΓΟΣ ΣΥΝΗΜΜΕΝΩΝ", "Τεχνικό δελτίο Πράξης")
    except AssertionError:
        return []
    out = []
    for line in s.splitlines():
        m = re.search(r"/\s*(\S.*?\.(?:pdf|xls|xlsx|doc|docx|zip))", line,
                      re.I)
        if m:
            out.append(m.group(1).strip())
    return sorted(set(out))


def reps(t):
    """Legal representatives, in the order Part C prints the partners."""
    return re.findall(r"Νόμιμος Εκπρόσωπος\s+(\S.*?)\s*$", t, re.M)


def banks(t):
    return re.findall(r"IBAN\s+([A-Z]{2}\d[\w]+)", t)


def read(pdf):
    t = text_of(pdf)
    hdr = header(t)
    fund_rows, fund_tot = by_partner(t)
    wps, wp_grand = work_packages(t)
    wpcc, wpcc_tot = wp_by_category(t)
    parts = analytical(t)
    return {
        "pdf": pdf, "hdr": hdr, "fund": fund_rows, "tot": fund_tot,
        "wps": wps, "wpcc": wpcc, "wpcc_tot": wpcc_tot, "parts": parts,
        "inds": indicators(t), "att": attachments(t), "reps": reps(t),
        "banks": banks(t),
        # partner name -> total, and partner name -> {deliverable: total}
        "by_partner": {p["name"]: p["printed_total"] for p in parts},
        "by_deliv": {(p["name"], r["deliv"]): r["total"]
                     for p in parts for r in p["rows"]},
        "by_cat": {c: v for c, v in zip(CATS, wpcc_tot["cats"])},
        "by_wp": {w["wp"]: w["budget"] for w in wps},
        "wp_dates": {w["wp"]: (w["start"], w["end"]) for w in wps},
        "wp_titles": {w["wp"]: w["title"] for w in wps},
        "ind_targets": {d["code"]: d["target"] for d in indicators(t)},
    }


# ------------------------------------------------------------------ comparing
def cmp_money(out, label, a, b, fmt=lambda k: str(k)):
    """Compare two {key: amount} maps.  Returns the number of differences."""
    keys = list(dict.fromkeys(list(a) + list(b)))
    hits = 0
    for k in keys:
        x, y = a.get(k), b.get(k)
        if x is None:
            out(f"    ADDED    {fmt(k)}: {money(y)}")
            hits += 1
        elif y is None:
            out(f"    REMOVED  {fmt(k)}: {money(x)}")
            hits += 1
        elif abs(x - y) > EPS:
            out(f"    CHANGED  {fmt(k)}: {money(x)} -> {money(y)} "
                f"({'+' if y > x else ''}{money(y - x)})")
            hits += 1
    if not hits:
        out(f"    no change in {label}")
    return hits


def cmp_text(out, label, a, b, fmt=lambda k: str(k)):
    keys = list(dict.fromkeys(list(a) + list(b)))
    hits = 0
    for k in keys:
        x, y = a.get(k), b.get(k)
        if x != y:
            out(f"    CHANGED  {fmt(k)}: {x} -> {y}")
            hits += 1
    if not hits:
        out(f"    no change in {label}")
    return hits


def cmp_list(out, label, a, b):
    add, rem = [x for x in b if x not in a], [x for x in a if x not in b]
    for x in rem:
        out(f"    REMOVED  {x}")
    for x in add:
        out(f"    ADDED    {x}")
    if not add and not rem:
        out(f"    no change in {label}")
    return len(add) + len(rem)


def compare(out, A, B):
    a, b = A["hdr"], B["hdr"]
    out()
    out("=" * 78)
    out(f"VERSION {a['version']} ({a['kind']}, approved {a['approved']})  ->  "
        f"VERSION {b['version']} ({b['kind']}, approved {b['approved']})")
    out("=" * 78)
    out(f"  {A['pdf'].name}  ->  {B['pdf'].name}")
    out()

    out("  DURATION")
    dur = cmp_text(out, "the duration",
                   {"start": a["start"], "end": a["end"],
                    "months": a["months"]},
                   {"start": b["start"], "end": b["end"],
                    "months": b["months"]})
    out()
    out("  TOTAL ELIGIBLE COST")
    tot = cmp_money(out, "the total", {"total": A["tot"]["total"],
                                       "EU": A["tot"]["eu"],
                                       "national": A["tot"]["nat"]},
                    {"total": B["tot"]["total"], "EU": B["tot"]["eu"],
                     "national": B["tot"]["nat"]})
    out()
    out("  BUDGET PER PARTNER")
    per_p = cmp_money(out, "any partner's budget", A["by_partner"],
                      B["by_partner"], lambda k: k[:46])
    out()
    out("  BUDGET PER COST CATEGORY")
    per_c = cmp_money(out, "any cost category", A["by_cat"], B["by_cat"])
    out()
    out("  BUDGET PER WORK PACKAGE")
    per_w = cmp_money(out, "any work package budget", A["by_wp"], B["by_wp"],
                      lambda k: f"WP{k}")
    out()
    out("  BUDGET PER PARTNER AND DELIVERABLE")
    per_d = cmp_money(out, "any deliverable's budget", A["by_deliv"],
                      B["by_deliv"], lambda k: f"{k[0][:34]} {k[1]}")
    out()
    out("  WORK PACKAGE DATES AND TITLES")
    wpd = cmp_text(out, "work package dates", A["wp_dates"], B["wp_dates"],
                   lambda k: f"WP{k} dates")
    wpd += cmp_text(out, "work package titles", A["wp_titles"],
                    B["wp_titles"], lambda k: f"WP{k} title")
    out()
    out("  INDICATOR TARGETS")
    ind = cmp_text(out, "indicator targets", A["ind_targets"],
                   B["ind_targets"])
    out()
    out("  LEGAL REPRESENTATIVES")
    rep = cmp_list(out, "legal representatives", A["reps"], B["reps"])
    out()
    out("  BANK ACCOUNTS (IBAN)")
    bank = cmp_list(out, "bank accounts", A["banks"], B["banks"])
    out()
    out("  ATTACHED DOCUMENTS")
    out("    Read REMOVED here with care. The MIS attaches documents to a")
    out("    technical sheet VERSION, not to the operation, so a new version")
    out("    starts with a short list and everything carried by the previous")
    out("    version reads as removed. That is the system's behaviour, not a")
    out("    withdrawal. What is worth checking is the opposite: a document")
    out("    the MA expects to find on the CURRENT version and cannot.")
    att = cmp_list(out, "the attachment list", A["att"], B["att"])

    money_moved = tot + per_p + per_c + per_w + per_d
    out()
    out("  WHAT THIS MODIFICATION ACTUALLY WAS")
    out("  " + "-" * 74)
    if money_moved == 0:
        out("  NO BUDGET FIGURE MOVED. Not the total, not a partner, not a cost")
        out("  category, not a work package, not a single deliverable. So this")
        out("  modification cannot be described in a report as a budget")
        out("  revision, and Article 3.7 or any absorption target calculated")
        out("  from the previous version still stands unchanged.")
    else:
        out(f"  {money_moved} budget figure(s) moved. Every percentage in any")
        out(f"  report written against version {a['version']} has to be")
        out(f"  recomputed against {b['version']}.")
    # THE CHECK THAT FOUND SOMETHING REAL ON THIS PROJECT
    #
    # When a modification shifts the implementation period, every work package
    # is supposed to shift with it. Here five of six did and one did not, so
    # the surviving work package still carries dates from the superseded
    # calendar. Nothing in a single version shows this, and the plain date
    # diff above shows it only by ABSENCE - the WP that did not change simply
    # is not listed, which is the easiest thing in the world to read past.
    if (a["start"], a["end"]) != (b["start"], b["end"]):
        stale = [w for w in A["wp_dates"]
                 if w in B["wp_dates"]
                 and A["wp_dates"][w] == B["wp_dates"][w]]
        if stale:
            out()
            out("  WORK PACKAGES LEFT ON THE OLD CALENDAR")
            out("  " + "-" * 74)
            out(f"  The project period moved from {a['start']}, {a['end']} to "
                f"{b['start']}, {b['end']},")
            out("  but these work packages kept the dates they had before:")
            for w in stale:
                s, e = B["wp_dates"][w]
                out(f"    WP{w} {B['wp_titles'].get(w, '')[:44]:<44s} "
                    f"{s} to {e}")
            out("  Either that is deliberate and the report should say why, or")
            out("  it is an oversight in the modification and the MA will ask.")
            out("  Decide which before the report is filed, not after.")

    nonmoney = []
    if dur:
        nonmoney.append("the implementation period")
    if wpd:
        nonmoney.append("work package dates or titles")
    if ind:
        nonmoney.append("indicator targets")
    if rep:
        nonmoney.append("a legal representative")
    if bank:
        nonmoney.append("bank account details")
    if att:
        nonmoney.append("the list of attached documents")
    if nonmoney:
        out("  It changed: " + ", ".join(nonmoney) + ".")
    elif money_moved == 0:
        out("  It changed nothing this script checks, which means the change is")
        out("  in the narrative text of Part B. Read the two PDFs side by side")
        out("  for that section before describing the modification.")
    return money_moved


def main(pdfs):
    if not pdfs:
        pdfs = sorted(HERE.glob("*.pdf"))
    read_ok = []
    for p in pdfs:
        try:
            read_ok.append(read(p))
        except (AssertionError, SystemExit) as e:
            print(f"[skipping {p.name}: {e}]")
    assert len(read_ok) >= 2, (
        "need at least two technical sheet PDFs to compare. Found "
        f"{len(read_ok)} readable one(s) in {HERE}.")

    def vkey(d):
        return [int(x) for x in d["hdr"]["version"].split(".")]
    read_ok.sort(key=vkey)

    mis = {d["hdr"]["mis"] for d in read_ok}
    assert len(mis) == 1, (
        f"these PDFs are different projects: MIS {sorted(mis)}. Comparing "
        f"versions across two different operations would produce a diff in "
        f"which every single line is a difference and none of it means "
        f"anything. Put one project's sheets in one folder.")

    out = Report()
    h = read_ok[-1]["hdr"]
    out("=" * 78)
    out(f"{h['acronym']}, MIS {h['mis']}: what changed across "
        f"{len(read_ok)} approved technical sheet versions")
    out("=" * 78)
    for d in read_ok:
        out(f"  v{d['hdr']['version']:<5s} {d['hdr']['kind']:<16s} "
            f"submitted {d['hdr']['submitted']:<17s} "
            f"approved {d['hdr']['approved']}   "
            f"total {money(d['tot']['total'])}")

    moved = 0
    for A, B in zip(read_ok, read_ok[1:]):
        moved += compare(out, A, B)

    out()
    out("=" * 78)
    out("ACROSS ALL VERSIONS")
    out("=" * 78)
    first, last = read_ok[0], read_ok[-1]
    out(f"  total eligible cost   {money(first['tot']['total'])}  ->  "
        f"{money(last['tot']['total'])}")
    out(f"  implementation period {first['hdr']['start']} to "
        f"{first['hdr']['end']}  ->  {last['hdr']['start']} to "
        f"{last['hdr']['end']}")
    out(f"  budget figures moved  {moved}")
    out()
    out("  THE CURRENT VERSION IS THE ONLY ONE TO REPORT AGAINST")
    out(f"  v{last['hdr']['version']}, approved {last['hdr']['approved']}, "
        f"total {money(last['tot']['total'])}, "
        f"{last['hdr']['start']} to {last['hdr']['end']}.")
    out("  Superseded versions are evidence of the modification history and")
    out("  nothing else. Do not compute a percentage against one of them.")

    out.save(HERE / f"td_{h['mis']}_version_history.txt")


if __name__ == "__main__":
    main([pathlib.Path(a) for a in sys.argv[1:]])
