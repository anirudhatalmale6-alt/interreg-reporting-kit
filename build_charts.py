#!/usr/bin/env python3
"""Charts for EITHER programme. Change one word at the top.

WHAT YOU ASKED: "what changes do I have to do on the script?"

Answer: nothing, once the differences live in a table instead of scattered
through the code. Set PROGRAMME = "MINORMED" and run it. Everything that
differs between the two is in CONFIG below, and nothing else in the file
mentions either programme by name.

BUT THERE IS A PROBLEM IN YOUR MINOR-MED TRACKER, AND IT MATTERS MORE THAN
THE SCRIPT.

It still has the single "Budget ex VAT" column, which is the bug you found in
the EPIRUS tracker on 20 September and which we fixed there by splitting the
column. In MINOR-MED it is still there:

    19 contracts span more than one deliverable row
    ALL 19 carry the IDENTICAL contract total on every one of those rows

    sum by row      1,563,874.23
    sum deduped       607,029.04
    overstated by     956,845.19

So if this script simply added up the column, every MINOR-MED chart would
show roughly two and a half times the real value, with no error anywhere. The
tell is that 1,563,874.23 is larger than the whole project budget, and a
procurement plan cannot exceed the project it belongs to.

WHY THE SCRIPT DOES NOT DETECT THIS AUTOMATICALLY

I tried. Detecting "all rows of a contract carry the same amount, therefore it
is a contract total" is WRONG on your EPIRUS tracker: 12 contracts there span
several rows and 3 of them genuinely have equal per-deliverable shares. Auto
deduplication would have silently deleted 10,451.61 of real money.

So the grain is DECLARED in CONFIG, and then VERIFIED against the data. If you
declare "contract" and the rows disagree, or declare "deliverable" and every
multi-row contract turns out identical, the script stops and tells you. State
the fact, then check the fact.
"""
import collections
import datetime as dt

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import openpyxl
from matplotlib.ticker import FuncFormatter

from _shared import HERE, newest

# ============================================================== CHANGE THIS
PROGRAMME = "EPIRUS"          # "EPIRUS" or "MINORMED"
LANG = "EL"                   # "EN" or "EL"
# ==========================================================================
# ... or do not change it at all, and say which one you want on the command
# line, which is safer when you are producing both reports in the same hour:
#
#     python3 build_charts.py MINORMED
#     python3 build_charts.py EPIRUS EN
#
# Editing a constant at the top of a file means the file now REMEMBERS which
# programme you last ran. Come back to it in three weeks, run it, and it
# quietly builds the other project's charts into filenames that look right.
# An argument cannot do that, because it is gone when the run ends.
import sys as _sys

for _a in _sys.argv[1:]:
    if _a.upper() in ("EPIRUS", "MINORMED"):
        PROGRAMME = _a.upper()
    elif _a.upper() in ("EN", "EL"):
        LANG = _a.upper()
    else:
        _sys.exit(f"do not know what {_a!r} means. Pass EPIRUS or MINORMED, "
                  f"and optionally EN or EL.")

CONFIG = {
    "EPIRUS": dict(
        name="EPIRUSMEDEYE", mis="6010984", prefix="EP",
        pattern="*Tracker*EPIRUS*.xlsx", sheet="EPIRUS-MEDEYE",
        budget_heads=("THIS deliverable",),
        budget_grain="deliverable",      # the column really is per-deliverable
        total_af=1_362_373.17,
        project=(dt.date(2025, 10, 13), dt.date(2027, 10, 13)),
        period=(dt.date(2026, 5, 26), dt.date(2026, 9, 26)),
        # The period opens on the day the REMACO contract was signed, which
        # is why the cumulative chart draws that line and labels it as an
        # engagement start.
        period_is_contract_date=True,
        # names that must NEVER appear in this programme's charts
        forbidden=["MINOR-MED", "MINORMED", "6010952", "Vlore", "Vlora",
                   "Filiates", "ATHENA"],
    ),
    "MINORMED": dict(
        name="MINOR-MED", mis="6010952", prefix="MM",
        pattern="*Tracker*MINORMED*.xlsx", sheet="MINOR-MED",
        budget_heads=("Budget ex VAT",),
        budget_grain="contract",         # ONE amount repeated across rows
        # FILLED IN FROM THE SOURCE, 28 September, not from my notes.
        #
        # 887.786,27 is the total eligible cost printed on technical sheet
        # version 1.2, approved 02/09/2026, and it is the SAME figure on
        # versions 1.0 and 1.1 - no modification has moved a single budget
        # cell on this project. Verified by read_td.py, which checks the
        # partner rows, the cost categories, the work packages and the
        # analytical table against each other, 39 cross-checks per version.
        #
        # I had this number in my working notes already and refused to use it
        # until the PDF arrived, because a total eligible cost is what every
        # percentage in the report is divided by. A remembered denominator
        # that is even slightly wrong makes every figure in the report wrong
        # in the same direction, which is the hardest kind of error to spot.
        total_af=887_786.27,
        project=(dt.date(2026, 2, 13), dt.date(2028, 2, 13)),
        # Reporting: 5 reports over 20 months from 24/06/2026, same period
        # for the Joint Secretariat and for EDEYPY, so four months each.
        # This is report 1 of 5.
        period=(dt.date(2026, 6, 24), dt.date(2026, 10, 24)),
        # Nothing was signed on 24/06/2026 here. It is only where the
        # reporting calendar starts, so the chart must not call it a contract
        # start the way the EPIRUSMEDEYE one legitimately does.
        period_is_contract_date=False,
        forbidden=["EPIRUSMEDEYE", "EPIRUS-MEDEYE", "6010984", "Ioannina",
                   "Dropull"],
    ),
}
CFG = CONFIG[PROGRAMME]

NAVY, TEAL, AMBER, RED, GREY = "#1F385E", "#2E7059", "#C08428", "#B4553C", "#8A99AB"
PLUM = "#6B3F6B"     # BLOCKED, so it cannot be confused with NOT STARTED
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.edgecolor": "#B8C4D2", "axes.labelcolor": NAVY,
    "text.color": "#222222", "figure.dpi": 150,
})
eur = FuncFormatter(lambda v, _: f"{v:,.0f}")

STATUS_KEY = {
    "COMPLETED": "done",   "ΟΛΟΚΛΗΡΩΜΕΝΟ": "done",
    "IN DELAY": "late",    "ΣΕ ΚΑΘΥΣΤΕΡΗΣΗ": "late",
    "NOT STARTED": "none", "ΔΕΝ ΞΕΚΙΝΗΣΕ": "none",
    "IN PROGRESS": "going", "ΣΕ ΕΞΕΛΙΞΗ": "going",
    "IN REVIEW": "review", "ΣΕ ΑΞΙΟΛΟΓΗΣΗ": "review",
    "BLOCKED": "blocked",  "ΜΠΛΟΚΑΡΙΣΜΕΝΟ": "blocked",
}

T = {
    "EN": {"cum": "{p}, cumulative value under signed contract",
           "eur": "EUR excluding VAT",
           # TWO DIFFERENT LINES, BECAUSE IT IS TWO DIFFERENT DATES.
           # On EPIRUSMEDEYE the period opens on the day the REMACO contract
           # was signed, which is the whole reason that chart is worth
           # drawing. On MINOR-MED it is just the reporting period start and
           # no contract begins there - calling it an engagement start would
           # be a plain untruth on the face of a chart going to a Ministry.
           "start_contract": "engagement begins {d}",
           "start_period": "reporting period {d}",
           "before": "contracted before:", "during": "contracted during:",
           "after": "signed after period end:",
           "of_total": "of {t:,.0f} signed in total",
           "status": "{p}, procurement value by status, EUR excluding VAT",
           "rows": "contracts", "row1": "contract",
           "partner": "{p}, value by beneficiary, signed against outstanding",
           "signed_lbl": "under signed contract",
           "unsigned_lbl": "not yet contracted", "signed_pc": "signed",
           "delay": "{p}, delay profile: {n} contracts behind the planned dates",
           "delay_y": "days behind the planned date",
           "done": "COMPLETED", "late": "IN DELAY", "none": "NOT STARTED",
           "going": "IN PROGRESS", "review": "IN REVIEW", "blocked": "BLOCKED"},
    "EL": {"cum": "{p}, σωρευτική αξία υπογεγραμμένων συμβάσεων",
           "eur": "EUR χωρίς ΦΠΑ",
           "start_contract": "έναρξη σύμβασης {d}",
           "start_period": "έναρξη περιόδου αναφοράς {d}",
           "before": "συμβασιοποιήθηκε πριν:",
           "during": "συμβασιοποιήθηκε εντός:",
           "after": "μετά τη λήξη της περιόδου:",
           "of_total": "από {t:,.0f} συνολικά",
           "status": "{p}, αξία συμβάσεων κατά κατάσταση, EUR χωρίς ΦΠΑ",
           "rows": "συμβάσεις", "row1": "σύμβαση",
           "partner": "{p}, αξία κατά δικαιούχο, υπογεγραμμένες έναντι εκκρεμών",
           "signed_lbl": "υπογεγραμμένες συμβάσεις",
           "unsigned_lbl": "μη συμβασιοποιημένες",
           "signed_pc": "υπογεγραμμένο",
           "delay": "{p}, προφίλ καθυστερήσεων: {n} συμβάσεις σε καθυστέρηση",
           "delay_y": "ημέρες καθυστέρησης",
           "done": "ΟΛΟΚΛΗΡΩΜΕΝΟ", "late": "ΣΕ ΚΑΘΥΣΤΕΡΗΣΗ",
           "none": "ΔΕΝ ΞΕΚΙΝΗΣΕ", "going": "ΣΕ ΕΞΕΛΙΞΗ",
           "review": "ΣΕ ΑΞΙΟΛΟΓΗΣΗ", "blocked": "ΜΠΛΟΚΑΡΙΣΜΕΝΟ"},
}
TXT = T[LANG]


def title(key, **kw):
    """Build a title and REFUSE to emit the other programme's name."""
    s = TXT[key].format(p=CFG["name"], **kw)
    bad = [f for f in CFG["forbidden"] if f.lower() in s.lower()]
    assert not bad, f"the other programme leaked into a chart title: {bad}"
    return s


def load():
    src = newest(CFG["pattern"])
    ws = openpyxl.load_workbook(src, data_only=True)[CFG["sheet"]]
    heads = {c: str(ws.cell(row=3, column=c).value or "").replace("\n", " ").strip()
             for c in range(1, ws.max_column + 1)}

    def col(*n):
        return next((c for c, v in heads.items()
                     if any(x.lower() in v.lower() for x in n)), None)
    C = {"ben": col("Beneficiary"), "no": col("Contract #", "α/α"),
         "budget": col(*CFG["budget_heads"]), "asig": col("ACTUAL signature"),
         "status": col("STATUS"), "late": col("Days late"),
         "deliv": col("Deliverable")}
    missing = [k for k, v in C.items() if v is None]
    assert not missing, (f"cannot find {missing} in {src.name}. Headings: "
                         f"{list(heads.values())[:12]}")

    rows = []
    for r in range(4, ws.max_row + 1):
        if not ws.cell(row=r, column=C["ben"]).value:
            continue
        g = lambda k: ws.cell(row=r, column=C[k]).value
        b = g("budget")
        st = str(g("status") or "").strip()
        rows.append({
            "ben": str(g("ben")).strip(), "no": str(g("no") or "").strip(),
            "budget": float(b) if isinstance(b, (int, float)) else 0.0,
            "asig": g("asig") if isinstance(g("asig"), dt.datetime) else None,
            "status": st, "key": STATUS_KEY.get(st.upper()),
            "deliv": str(g("deliv") or "").strip(),
            "late": g("late") if isinstance(g("late"), (int, float)) else 0,
        })
    unknown = sorted({r["status"] for r in rows if r["key"] is None})
    assert not unknown, (f"status {unknown} not in STATUS_KEY. Add it in BOTH "
                         f"languages or every chart mis-colours and undercounts.")
    return rows, src


def grain_check(rows):
    """Verify the data matches the grain declared in CONFIG, then aggregate.

    Returns a list of (key, ben, budget, asig, late) at the correct grain.
    """
    groups = collections.defaultdict(list)
    for r in rows:
        groups[(r["ben"], r["no"])].append(r)
    multi = {k: v for k, v in groups.items() if len(v) > 1}
    identical = {k: v for k, v in multi.items()
                 if len({round(x["budget"], 2) for x in v}) == 1}

    if CFG["budget_grain"] == "contract":
        assert len(identical) == len(multi), (
            f"{PROGRAMME} is declared contract-grain but {len(multi)-len(identical)} "
            f"of {len(multi)} multi-row contracts have DIFFERING amounts. The "
            f"column may be per-deliverable after all. Stopping.")
        # WHEN ROWS OF ONE CONTRACT DISAGREE ABOUT STATUS, DO NOT PICK THE
        # FIRST ONE. Four of your 34 MINOR-MED contracts carry contradictory
        # statuses across their rows, and one of them is recorded as both
        # COMPLETED and NOT STARTED, which cannot both be true. Taking v[0]
        # would silently choose whichever row happened to come first.
        #
        # Status is DERIVED from the dates, so a contract with rows in
        # different states means partners filled a signature date on one
        # deliverable row and left it blank on another. That is a data problem
        # to chase, not something to average away.
        #
        # So: take the WORST status, which makes the chart understate rather
        # than overstate - the right direction for a report to a Ministry -
        # and print the conflicts so you can go and fix them.
        SEVERITY = ["blocked", "none", "late", "review", "going", "done"]
        unit, conflicts, datefix = [], [], []
        for (ben, no), members in groups.items():
            keys = {m["key"] for m in members}
            pick = dict(members[0])
            if len(keys) > 1:
                worst = min(keys, key=SEVERITY.index)
                conflicts.append((ben, no, sorted(keys), worst))
                pick["key"] = worst
            # THE SAME DEFECT AS THE STATUS, IN THE DATE COLUMNS.
            #
            # Taking members[0] for the status was wrong and I fixed it. I
            # then left every DATE on members[0], which has exactly the same
            # failure: one contract, several deliverable rows, and the
            # signature date typed on the second row rather than the first.
            # members[0] then reports that contract as never signed, so its
            # value drops out of "contracted during the period" and the
            # report UNDERSTATES what the partners achieved.
            #
            # A contract has one signature date, so where the rows carry
            # several the earliest is taken and the disagreement is printed.
            for fld in ("asig", "apub", "aoff"):
                vals = sorted({m[fld] for m in members if m.get(fld)},
                              key=lambda d: getattr(d, "date", lambda: d)())
                if vals:
                    pick[fld] = vals[0]
                    if len(vals) > 1:
                        datefix.append((ben, no, fld, vals))
                    elif not members[0].get(fld):
                        datefix.append((ben, no, fld, vals))
                else:
                    pick[fld] = None
            # The largest delay across the contract's rows, for the same
            # reason: a contract is as late as its latest part.
            pick["late"] = max((m["late"] or 0) for m in members)
            unit.append(pick)
        if conflicts:
            print(f"  !! {len(conflicts)} contract(s) have CONTRADICTORY "
                  f"statuses across their rows. Worst taken; chase these:")
            for ben, no, keys, worst in conflicts:
                print(f"     {ben[:28]:28s} #{no:3s} {keys} -> using {worst}")
        if datefix:
            print(f"  !! {len(datefix)} date(s) were NOT on the contract's "
                  f"first row, or differ between its rows. Earliest taken:")
            for ben, no, fld, vals in datefix:
                shown = ", ".join(f"{v:%d/%m/%Y}" for v in vals)
                print(f"     {ben[:28]:28s} #{no:3s} {fld:5s} {shown}")
        rowsum = sum(r["budget"] for r in rows)
        kept = sum(r["budget"] for r in unit)
        print(f"  grain=contract: counted each contract ONCE. "
              f"row-sum {rowsum:,.2f} would have overstated by "
              f"{rowsum - kept:,.2f}")
    else:
        assert len(identical) < len(multi) or not multi, (
            f"{PROGRAMME} is declared deliverable-grain but EVERY multi-row "
            f"contract carries an identical amount, which is what a CONTRACT "
            f"total looks like. Check the column before trusting this.")
        unit = rows
        print(f"  grain=deliverable: counted every row. "
              f"{len(identical)} of {len(multi)} multi-row contracts happen to "
              f"have equal shares, which is why this is not auto-detected.")
    return unit


def save(fig, name):
    path = HERE / f"chart_{CFG['prefix']}_{name}.png"
    fig.tight_layout()
    fig.savefig(path, facecolor="white")
    plt.close(fig)
    w, h = fig.get_size_inches() * fig.dpi
    assert w < 2000 and h < 2000, f"{name} is {int(w)}x{int(h)}"
    print(f"  {path.name}  {int(w)}x{int(h)}")


def chart_status(unit):
    v, n = collections.Counter(), collections.Counter()
    for r in unit:
        v[r["key"]] += r["budget"]
        n[r["key"]] += 1
    order = sorted(v, key=lambda k: v[k])
    col = {"done": TEAL, "late": AMBER, "none": RED, "going": NAVY,
           "review": GREY, "blocked": PLUM}
    fig, ax = plt.subplots(figsize=(9.5, 3.8))
    bars = ax.barh([TXT[k] for k in order], [v[k] for k in order],
                   color=[col[k] for k in order], height=0.6)
    tot = sum(v.values())
    for b, k in zip(bars, order):
        ax.text(b.get_width() + tot * 0.012, b.get_y() + b.get_height() / 2,
                f"{v[k]:,.0f}   {v[k]/tot:.1%}   "
                f"({n[k]} {TXT['row1'] if n[k] == 1 else TXT['rows']})",
                va="center", fontsize=9)
    ax.set_xlim(0, tot * 0.80)
    ax.set_title(title("status"), color=NAVY, fontsize=11.5,
                 fontweight="bold", loc="left")
    ax.xaxis.set_major_formatter(eur)
    ax.grid(axis="x", color="#EDF1F6")
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    save(fig, "status")


def chart_partner(unit):
    done, rest = collections.Counter(), collections.Counter()
    for r in unit:
        (done if r["key"] == "done" else rest)[r["ben"]] += r["budget"]
    bens = sorted(set(done) | set(rest), key=lambda b: done[b] + rest[b])
    short = [b.replace("General Hospital of ", "").replace("Municipality of ", "")
             .replace("Regional Council of ", "").replace(" (PB1) EDEYPY", " EDEYPY")[:26]
             for b in bens]
    fig, ax = plt.subplots(figsize=(9.8, 3.9))
    ax.barh(short, [done[b] for b in bens], color=TEAL, height=0.58,
            label=TXT["signed_lbl"])
    ax.barh(short, [rest[b] for b in bens], left=[done[b] for b in bens],
            color="#D9E2EC", height=0.58, label=TXT["unsigned_lbl"])
    for i, b in enumerate(bens):
        t = done[b] + rest[b]
        ax.text(t * 1.01, i,
                f"{t:,.0f}   {done[b]/t if t else 0:.0%} {TXT['signed_pc']}",
                va="center", fontsize=9)
    ax.set_xlim(0, max(done[b] + rest[b] for b in bens) * 1.34)
    ax.set_title(title("partner"), color=NAVY, fontsize=11.5,
                 fontweight="bold", loc="left")
    ax.xaxis.set_major_formatter(eur)
    ax.legend(frameon=False, fontsize=9, loc="lower right")
    ax.grid(axis="x", color="#EDF1F6")
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    save(fig, "partner")


def chart_delay(unit):
    late = sorted((r for r in unit if r["late"] > 0), key=lambda r: -r["late"])
    if not late:
        print("  no contracts behind plan, delay chart skipped")
        return
    labels = [f"{r['ben'].split()[0]} {r['deliv'].split('|')[0].strip()[:6]}"
              for r in late]
    fig, ax = plt.subplots(figsize=(10, 4.4))
    ax.bar(range(len(late)), [r["late"] for r in late], color=AMBER, width=0.72)
    ax.set_xticks(range(len(late)))
    ax.set_xticklabels(labels, rotation=90, fontsize=7)
    ax.set_ylabel(TXT["delay_y"])
    ax.set_title(title("delay", n=len(late)), color=NAVY, fontsize=11.5,
                 fontweight="bold", loc="left")
    ax.grid(axis="y", color="#EDF1F6")
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    save(fig, "delay")


def chart_cumulative(unit):
    if not CFG["period"]:
        print("  cumulative chart SKIPPED: this programme's reporting period "
              "is not set in CONFIG. Fill in period=(start, end).")
        return
    p0, p1 = CFG["period"]
    signed = sorted((r for r in unit if r["asig"]), key=lambda r: r["asig"])
    if not signed:
        print("  no signature dates, cumulative chart skipped")
        return
    xs, ys, run = [], [], 0.0
    for r in signed:
        run += r["budget"]
        xs.append(r["asig"].date())
        ys.append(run)
    # A REPORTING PERIOD HAS TWO ENDS AND I ONLY USED ONE.
    #
    # This was "during = everything signed on or after p0", with no upper
    # bound, so a contract dated after the period CLOSED was counted as
    # signed inside it. On EPIRUSMEDEYE that happened to be harmless - no row
    # is dated past 26/09/2026, so the 456.746,77 figure is unaffected. On
    # MINOR-MED it was not harmless: one ATHENA contract dated 05/01/2027 was
    # being added to the value "contracted within the reporting period".
    #
    # An overstatement of signed value in a report to a Managing Authority is
    # the worst direction for this error to run, and nothing on the chart
    # would have shown it - the number simply looked bigger.
    before = sum(r["budget"] for r in signed if r["asig"].date() < p0)
    during = sum(r["budget"] for r in signed if p0 <= r["asig"].date() <= p1)
    after = sum(r["budget"] for r in signed if r["asig"].date() > p1)

    # AND A SECOND THING, WHICH IS IN THE DATA RATHER THAN THE CODE.
    #
    # The column is headed ACTUAL signature. An actual date cannot be in the
    # future. Where it is, a partner has entered a PLANNED date in the actual
    # column, and every figure derived from it describes a contract that does
    # not exist yet. So it is named and excluded from nothing - it is left in
    # the chart, because it is what the tracker says - but it is printed
    # loudly so it is chased before the report quotes the total.
    today = dt.date.today()
    ahead = [r for r in signed if r["asig"].date() > today]
    if ahead:
        print(f"  !! {len(ahead)} contract(s) carry an ACTUAL signature date "
              f"in the FUTURE, totalling {sum(r['budget'] for r in ahead):,.2f}.")
        print(f"     An ACTUAL date cannot be later than today ({today:%d/%m/%Y}). "
              f"These are planned dates in the actual column:")
        for r in sorted(ahead, key=lambda r: r["asig"]):
            print(f"     {r['ben'][:30]:<30s} {r['asig'].date():%d/%m/%Y} "
                  f"{r['budget']:>12,.2f}")
        print("     Do NOT describe these as signed. Either the partner "
              "corrects the column or they are reported as planned.")
        overlap = sum(r["budget"] for r in ahead
                      if p0 <= r["asig"].date() <= p1)
        if overlap:
            print(f"     {overlap:,.2f} of that sits INSIDE the reporting "
                  f"period and is already inside the 'during' figure, so the "
                  f"defensible signed-within-period value today is "
                  f"{during - overlap:,.2f}, not {during:,.2f}.")

    fig, ax = plt.subplots(figsize=(10, 5.2))
    ax.step(xs, ys, where="post", color=NAVY, linewidth=2.2)
    ax.fill_between(xs, ys, step="post", color=NAVY, alpha=0.10)
    ax.axvline(p0, color=TEAL, linewidth=1.8, linestyle="--")
    ax.axvline(p1, color=GREY, linewidth=1.2, linestyle=":")
    key = "start_contract" if CFG["period_is_contract_date"] else "start_period"
    lines = [TXT[key].format(d=p0.strftime("%d/%m/%Y")),
             f"{TXT['before']}  {before:>12,.0f}",
             f"{TXT['during']}  {during:>12,.0f}"]
    if after:
        lines.append(f"{TXT['after']}  {after:>12,.0f}")
    lines.append(TXT["of_total"].format(t=run))
    # The curved leader from the box to the period line crossed the data in
    # the rendered PNG - visible only in the PNG, never in the code. A short
    # straight leader to the axis, and the box parked in whichever top corner
    # the curve is not using.
    ax.annotate("\n".join(lines),
                xy=(p0, 0), xytext=(0.985 if before == 0 else 0.02, 0.97),
                textcoords="axes fraction", va="top",
                ha="right" if before == 0 else "left",
                fontsize=9.5, color=TEAL,
                bbox=dict(boxstyle="round,pad=0.5", facecolor="white",
                          edgecolor=TEAL, linewidth=1))
    ax.set_title(title("cum"), color=NAVY, fontsize=11.5, fontweight="bold",
                 loc="left")
    ax.set_ylabel(TXT["eur"])
    ax.yaxis.set_major_formatter(eur)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%Y"))
    ax.grid(axis="y", color="#EDF1F6")
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    save(fig, "cumulative")
    print(f"  contracted before the period {before:,.2f} | during {during:,.2f}")


def main():
    print(f"{CFG['name']}, MIS {CFG['mis']}, language {LANG}")
    rows, src = load()
    unit = grain_check(rows)
    total = sum(r["budget"] for r in unit)
    print(f"  {len(unit)} units at grain '{CFG['budget_grain']}', "
          f"{total:,.2f} ex VAT")
    if CFG["total_af"]:
        print(f"  Article 3.7 threshold, 20% of {CFG['total_af']:,.2f} = "
              f"{CFG['total_af']*0.20:,.2f}")
    else:
        print("  total eligible cost NOT SET for this programme, so no 20% "
              "figure. Fill in total_af in CONFIG.")
    chart_cumulative(unit)
    chart_status(unit)
    chart_partner(unit)
    chart_delay(unit)


if __name__ == "__main__":
    main()
