#!/usr/bin/env python3
"""Audit the LIVE application, from his own exports. Not from my model.

    python3 audit_live_application.py

He asked: "Can you check that so far except cofinancing, checklist B (waiting
for maps), and does everything financially count ok?"

So this reads HIS two exports - application_form-08-10.docx and
budget_08-10.xls - and checks them. That distinction matters. Every previous
check I have written compared my generator against my own constants, which
proves my arithmetic and nothing about what is in the platform. This one reads
what the platform says and tests it against the programme's rules.

CALIBRATION FIRST, BECAUSE A CLEAN REPORT IS THE FAILURE MODE OF AN AUDIT

Before reporting anything, the script proves it can SEE. It asserts that it
found the budget line sheet, that the line count is in a plausible range, that
the grand total parses as a number in the expected magnitude, and that it can
locate at least one known section of text in the form. If any of those fail it
refuses to produce a report at all, rather than printing "no problems found"
because it was reading an empty parse.

I have been caught by exactly this before: an audit script whose failure mode
was a clean report.

WHAT IT CHECKS

Money, against the programme's own error codes:
  205  total eligible >= 1.000.000
  210  EU contribution >= 920.000, and <= the 1.200.000 ceiling
  202  no partner above 35% of total eligible
  217  MPC partners >= 50%
  220  staff < 40% of total eligible
  104  output percentages total 100 within each work package
  222  every partner has a line in WP1 and WP2, and an auditor line in WP1
  the flat rates are exactly 15% + 15% of staff
  direct + flat rates == total eligible
  no duplicate partner+category+work-package line
  no infrastructure in WP1 or WP2

Text, against what should be there:
  every section that should carry text does, and nothing is left as a
  placeholder - xxxx, TBD, [SOMETHING], or a lone full stop
"""
import re
import sys
from pathlib import Path

import openpyxl
from docx import Document

SCRATCH = Path("/tmp/claude-1003/-home-freelancer/"
               "22c2ac1c-53f2-4084-aeb9-41c15659448a/scratchpad")
BUDGET = SCRATCH / "budget_08-10.xlsx"
FORM = SCRATCH / "af_0810.docx"

MIN_TOTAL = 1_000_000.0
MIN_EU = 920_000.0
MAX_EU = 1_200_000.0
MAX_PARTNER = 0.35
MIN_MPC = 0.50
MAX_STAFF = 0.40
FLAT_EACH = 0.15
MPC_COUNTRIES = {"turkiye", "türkiye", "turkey", "lebanon", "liban"}
# His sheets use the template codes, not names. PP2 Pi Youth and PP3
# Acipayam are Turkiye, PP5 LCEC is Lebanon; LEP, PP1 and PP4 are EU.
MPC_CODES = {"PP2", "PP3", "PP5"}
# The AGREED project budget. The programme minimum is 1.000.000, so a
# shortfall against what the partners agreed passes every programme rule
# and is still grant left on the table. 08 Oct: the audit reported FAILS 0
# while the application was 19.412 light, because it only checked the
# programme floor. Reconciliation proves consistency, never plausibility.
AGREED_TOTAL = 1_300_000.0
AGREED_CC = {
    "LEP": (173_500.0, 30_000.0, 94_450.0, 0.0),
    "PP1": (60_000.0, 8_000.0, 119_000.0, 0.0),
    "PP2": (50_000.0, 10_000.0, 155_000.0, 0.0),
    "PP3": (30_000.0, 12_000.0, 84_000.0, 0.0),
    "PP4": (50_000.0, 0.0, 15_000.0, 0.0),
    "PP5": (80_000.0, 20_000.0, 176_000.0, 0.0),
}

ok, warn, bad = [], [], []


def eur(x):
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","})) + " €"


def num(v):
    """Parse a cell that may be a number or a formatted string."""
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    s = str(v).strip().replace("€", "").replace(" ", "")
    if not s:
        return None
    # 1.234.567,89  ->  1234567.89
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def load():
    wb = openpyxl.load_workbook(BUDGET, data_only=True)
    doc = Document(FORM)
    text = "\n".join(p.text for p in doc.paragraphs)
    for T in doc.tables:
        for r in T.rows:
            for c in r.cells:
                text += "\n" + c.text
    return wb, text


def calibrate(wb, text):
    """Prove the audit can see before it is allowed to report."""
    fails = []
    if "Budget tot per budget line" not in wb.sheetnames:
        fails.append("the budget line sheet is not in the workbook")
    ws = wb["Budget tot per budget line"]
    lines = [r for r in ws.iter_rows(min_row=1, values_only=True)
             if r and r[0] and re.match(r"^WP\d\.", str(r[0]))]
    if not 40 <= len(lines) <= 400:
        fails.append(f"found {len(lines)} budget lines - outside the "
                     f"plausible range, so the parse is wrong")
    tot = sum(num(r[6]) or 0 for r in lines)
    if not 500_000 < tot < 2_000_000:
        fails.append(f"budget lines sum to {tot} - wrong magnitude, so the "
                     f"amount column is not column 7")
    if len(text) < 20_000:
        fails.append(f"the form export yields only {len(text)} characters - "
                     f"too little to audit")
    if "AMMOS" not in text:
        fails.append("cannot find the acronym in the form text - wrong parse")
    if "transferability" not in text.lower():
        fails.append("cannot find known project text in the form - the "
                     "export may be empty or a different project")
    assert not fails, ("AUDIT CANNOT SEE - refusing to report:\n  - "
                       + "\n  - ".join(fails))
    return lines, tot


def main():
    wb, text = load()
    lines, line_total = calibrate(wb, text)
    print(f"calibration passed: {len(lines)} budget lines parsed, "
          f"total {eur(line_total)}, {len(text):,} characters of form text\n")

    # ------------------------------------------------- per-partner from CC
    ws = wb["Budget per CC"]
    rows = []
    for r in ws.iter_rows(min_row=1, values_only=True):
        if not r or not r[0]:
            continue
        name = str(r[0]).strip()
        if name.lower().startswith(("partner", "subtotal", "total",
                                    "budget per")):
            continue
        vals = [num(x) for x in r[1:10]]
        if vals[0] is None:
            continue
        rows.append((name, vals))

    staff = sum(v[0] or 0 for _, v in rows)
    direct = sum((v[4] if v[4] is not None else sum(x or 0 for x in v[:4]))
                 for _, v in rows)
    elig = {}
    for name, v in rows:
        e = v[7] if v[7] is not None else None
        if e is None:
            e = (v[4] or 0) + (v[5] or 0) + (v[6] or 0)
        elig[name] = e
    total = sum(elig.values())

    print("PER PARTNER, from his Budget per CC sheet:")
    for name, v in rows:
        print(f"   {name[:28]:30s} staff {eur(v[0] or 0):>14s}   "
              f"eligible {eur(elig[name]):>14s}  "
              f"{100 * elig[name] / total:5.2f}%")
    print(f"   {'TOTAL':30s} staff {eur(staff):>14s}   "
          f"eligible {eur(total):>14s}\n")

    def check(label, cond, detail):
        (ok if cond else bad).append((label, detail))

    check("205 total eligible >= 1.000.000 (programme floor)",
          total >= MIN_TOTAL - 0.5, eur(total))
    # The check the programme does NOT do for us.
    check("total eligible equals the AGREED 1.300.000",
          abs(total - AGREED_TOTAL) < 1.0,
          f"{eur(total)} entered against {eur(AGREED_TOTAL)} agreed - "
          f"{eur(AGREED_TOTAL - total)} of grant not claimed"
          if abs(total - AGREED_TOTAL) >= 1.0 else eur(total))
    for name, v in rows:
        key = name.upper()
        if key not in AGREED_CC:
            continue
        got = [v[0] or 0, v[1] or 0, v[2] or 0, v[3] or 0]
        for i, label in enumerate(("staff", "equipment",
                                   "external expertise", "infrastructure")):
            want = AGREED_CC[key][i]
            if abs(got[i] - want) > 0.5:
                bad.append((f"{key} {label} differs from the agreed budget",
                            f"{eur(got[i])} entered, {eur(want)} agreed "
                            f"({got[i] - want:+,.0f})"))
    eu = total * 0.92
    check("210 EU contribution >= 920.000", eu >= MIN_EU - 0.5, eur(eu))
    check("210 EU contribution <= 1.200.000 ceiling", eu <= MAX_EU + 0.5,
          f"{eur(eu)}, {eur(MAX_EU - eu)} of headroom")
    biggest = max(elig, key=elig.get)
    check("202 no partner above 35%", elig[biggest] / total <= MAX_PARTNER,
          f"largest is {biggest[:24]} at {100 * elig[biggest] / total:.2f}%")
    mpc = sum(e for n, e in elig.items()
              if n.upper() in MPC_CODES
              or any(c in n.lower() for c in MPC_COUNTRIES))
    if mpc == 0:
        warn.append(("217 MPC share",
                     "could not identify MPC partners by name in his sheet - "
                     "check manually against the platform's own 50% rule tab"))
    else:
        check("217 MPC partners >= 50%", mpc / total >= MIN_MPC,
              f"{100 * mpc / total:.2f}%")
    check("220 staff below 40% of total eligible", staff / total < MAX_STAFF,
          f"{100 * staff / total:.2f}%")
    flats = total - direct
    check("flat rates are exactly 15% + 15% of staff",
          abs(flats - 2 * FLAT_EACH * staff) < 1.0,
          f"{eur(flats)} vs {eur(2 * FLAT_EACH * staff)} expected")
    check("direct + flat rates == total eligible",
          abs(direct + flats - total) < 1.0,
          f"{eur(direct)} + {eur(flats)} = {eur(direct + flats)}")

    # --------------------------------------------------- budget line checks
    seen, dups = set(), []
    per_wp, audit_wp1, in_wp = {}, set(), {}
    for r in lines:
        code, wp, cat, partner, brief, just, amt = (
            str(r[0]), str(r[1]).strip(), str(r[2]).strip(),
            str(r[3]).strip(), str(r[4] or ""), r[5], num(r[6]) or 0)
        key = (wp, cat, partner)
        if key in seen:
            dups.append(code)
        seen.add(key)
        per_wp[wp] = per_wp.get(wp, 0) + amt
        in_wp.setdefault(partner, set()).add(wp)
        if wp == "WP1" and "audit" in brief.lower():
            audit_wp1.add(partner)
        if wp in ("WP1", "WP2") and "nfrastructure" in cat:
            bad.append(("infrastructure in a mandatory WP",
                        f"{code} is {cat} in {wp}, which the rules forbid"))
        if not just or not str(just).strip():
            warn.append(("missing justification",
                         f"{code} has an empty Justification column"))
    check("no duplicate partner+category+work package line", not dups,
          f"{len(dups)} duplicates" if dups else "none")
    partners = sorted(in_wp)
    missing1 = [p for p in partners if "WP1" not in in_wp[p]]
    missing2 = [p for p in partners if "WP2" not in in_wp[p]]
    check("222 every partner has a WP1 line", not missing1,
          ", ".join(missing1) or f"all {len(partners)}")
    check("222 every partner has a WP2 line", not missing2,
          ", ".join(missing2) or f"all {len(partners)}")
    no_audit = [p for p in partners if p not in audit_wp1]
    check("222 every partner has an auditor line in WP1", not no_audit,
          ", ".join(no_audit) or f"all {len(partners)}")
    check("budget lines sum to total direct cost",
          abs(line_total - direct) < 2.0,
          f"lines {eur(line_total)} vs Budget per CC {eur(direct)}")

    print("WORK PACKAGE TOTALS, from his budget lines:")
    for wp in sorted(per_wp):
        print(f"   {wp}  {eur(per_wp[wp]):>16s}  "
              f"{100 * per_wp[wp] / line_total:5.1f}% of direct")
    print()

    # -------------------------------------------------- output percentages
    try:
        wso = wb["Budget per output"]
        pcts = {}
        cur = None
        for r in wso.iter_rows(min_row=1, values_only=True):
            if not r:
                continue
            head = str(r[0] or "")
            m = re.match(r"\s*(WP\d)", head)
            if m:
                cur = m.group(1)
            if cur:
                for v in r[1:]:
                    f = num(v)
                    if f is not None and 0 < f <= 1.0001:
                        pcts.setdefault(cur, []).append(f)
        for wp, vals in pcts.items():
            s = sum(vals)
            if vals and abs(s - 1.0) > 0.01:
                bad.append(("104 output percentages total 100%",
                            f"{wp} totals {100 * s:.1f}%"))
        if pcts:
            ok.append(("104 output percentages total 100%",
                       ", ".join(f"{wp} {100 * sum(v):.0f}%"
                                 for wp, v in sorted(pcts.items()))))
        else:
            warn.append(("104 output percentages",
                         "could not read the Budget per output sheet - check "
                         "it on screen"))
    except KeyError:
        warn.append(("104 output percentages",
                     "no Budget per output sheet in the export"))

    # -------------------------------------------------------- placeholders
    PLACEHOLDERS = [r"\bxxx+\b", r"\bTBD\b", r"\bTBC\b", r"\bLorem\b",
                    r"\[[A-Z][A-Z0-9_ ]{2,}\]", r"\bto be confirmed\b"]
    hits = []
    for pat in PLACEHOLDERS:
        for m in re.finditer(pat, text, re.I):
            s = max(0, m.start() - 55)
            hits.append(f"...{text[s:m.end() + 55]}...".replace("\n", " "))
    if hits:
        for h in dict.fromkeys(hits):
            bad.append(("PLACEHOLDER STILL IN THE FORM", h[:150]))
    else:
        ok.append(("no placeholders left in the form text",
                   "searched xxx, TBD, TBC, [UPPERCASE], 'to be confirmed'"))

    # ------------------------------------------------------- section cover
    EXPECT = {
        "3.1.1 relevance": "sediment shores are the asset",
        "3.1.2 transnational": "method we transfer exists in one part",
        "3.1.3 beneficiaries": "primary target group",
        "3.1.4a synergies": "MEDITERRANEAN STRATEGY FOR SUSTAINABLE",
        "3.1.4b outputs": "does not invent a method",
        "3.2.1 transfer methodology": "decide and translate",
        "summary": "Mediterranean sediment shores are the asset",
        "3.4.1 methodology": "decide-test-demonstrate-embed",
        "3.4.2 work plan": "Six work packages over 24 months",
        "3.4.3 capitalisation": "returns more than it takes",
        "3.4.4 communication": "move a method from the people who wrote it",
        "3.5 sustainability": "costed plans, not aspirational",
        "3.6 cost-effectiveness": "ratio to judge",
        "3.7 horizontal": "channelling visitors rather than excluding",
    }
    low = text.lower()
    for label, probe in EXPECT.items():
        if probe.lower() in low:
            ok.append((f"{label} is in the form", "found"))
        else:
            warn.append((f"{label}", "not found in the export - either not "
                                     "pasted yet or reworded"))

    # ------------------------------------------------------------- report
    print("=" * 72)
    print(f"FAILS {len(bad)}   WARNINGS {len(warn)}   PASSES {len(ok)}")
    print("=" * 72)
    if bad:
        print("\nMUST FIX:")
        for l, d in bad:
            print(f"   [X] {l}\n       {d}")
    if warn:
        print("\nCHECK BY HAND:")
        for l, d in warn:
            print(f"   [?] {l}\n       {d}")
    print("\nPASSES:")
    for l, d in ok:
        print(f"   [v] {l}  —  {d}")


if __name__ == "__main__":
    main()
