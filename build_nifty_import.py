#!/usr/bin/env python3
"""Turn a procurement tracker into a CSV that Nifty's spreadsheet import accepts.

    python3 build_nifty_import.py EPIRUS
    python3 build_nifty_import.py MINORMED

YOUR QUESTION: "if I want to import EPIRUSMEDEYE and MINORMED in NIFTY how easy
it is?"

Easy, by the spreadsheet route, and the free tier fits exactly: two active
projects, which is two programmes, and 500 tasks against the 81 we would
actually create.

Of the nine importers Nifty offers - Spreadsheet, Asana, Basecamp, ClickUp,
Todoist, Jira, Monday, Trello, Wrike - only the first is any use to us,
because the data does not live in a project tool, it lives in your tracker.

WHAT BECOMES A TASK, AND WHY

One task per CONTRACT POSITION, not per deliverable and not per work package.
A contract position is the thing that has a planned date, slips, and can be
chased. Work packages do not slip; procurements do. And your tracker already
holds "Planned publication" and "Planned signature" for every row, which are
exactly the start and due dates a Gantt needs - so the chart draws itself
rather than being drawn by hand.

Grouping is by BENEFICIARY, so each partner opens Nifty and sees its own list
instead of scrolling past everyone else's.

THE WARNING THAT MATTERS MORE THAN THE SCRIPT

This creates a SECOND COPY of the truth. The tracker on SharePoint is what the
partners update; Nifty would be a parallel list that starts identical and
drifts from the first day somebody closes a task in one and not the other.
Then you have two answers to "how many contracts are signed" and no way to
tell which is right.

So use Nifty for what it is good at - your own scheduling, the reporting
deadlines, the Gantt, the things only you track - and keep the procurement
tracker as the single source for procurement. If you want Nifty to show
procurement, RE-IMPORT from this script rather than editing inside Nifty. A
regenerated list is always right; an edited copy is right for about a week.

ON THE COLUMN NAMES - NOW CONFIRMED FROM YOUR OWN SCREENSHOT

The first version guessed at the headings and said so. You sent the mapping
screen, so they are no longer a guess. Nifty's template fields are:

    Task Name*, Task Description, List, Status, Task Assignees, Task Tags,
    Due Date, Start Date, Reminder, Story Points, Completed, Archived,
    Subtasks

All eight of my columns auto-mapped, and the MINOR-MED file validated at
"All 34, Invalid 0". The headings below now use Nifty's own words exactly -
"List" rather than "Task Group", "Task Assignees" rather than "Assignee" -
so the mapping is a match rather than an inference.

One field was being left on the table: COMPLETED. It is separate from Status,
so a task whose Status reads "Completed" can still sit in the board as open
with a label on it. The export now sets both.
"""
import csv
import datetime as dt
import sys

from _shared import HERE, newest, read_tracker

from build_charts import CFG, PROGRAMME, STATUS_KEY

# Nifty's free tier, from the pricing panel you sent.
FREE_TASKS, FREE_PROJECTS, FREE_MB = 500, 2, 100

# The status words Nifty will see. Its own columns are usually To Do / In
# Progress / Done, so the mapping step will ask - these are chosen to make
# that choice obvious rather than to match any particular tool.
# Kept to four clean values with NO COMMAS and no compound states. A status
# like "In Progress, LATE" survives the CSV but makes the mapping step offer a
# fifth column nobody wants, and the comma is a quoting accident waiting to
# happen in whatever reads it next. Lateness is a TAG, not a status: a
# contract is either in progress or not, and separately it is late or not.
NIFTY_STATUS = {"done": "Completed", "going": "In Progress",
                "review": "In Review", "late": "In Progress",
                "none": "To Do", "blocked": "Blocked"}


def d(v):
    """A date Nifty will parse. ISO, because every importer accepts it."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    return ""


def tags(r):
    """Clean tags. The tracker's cost-category cell is "24|Greek label" and
    slicing it blind gave tags like "24|Dapanes prosopikou (kat apokopi kost"
    - truncated mid-word, with a pipe in it. A tag that is unreadable is
    worse than no tag, because it still takes up the column."""
    out = []
    cat = str(r.get("cat") or "").strip()
    if cat:
        code = cat.split("|")[0].strip()
        out.append(f"CC{code}" if code.isdigit() else code[:24])
    proc = str(r.get("proc") or "").strip()
    if proc:
        # "OPEN above the limits (s. Guidelines)" -> "OPEN"
        # "IPA SERVICE >20000 & <40000" -> "IPA SERVICE >20000"
        # Truncating at a fixed width left tags ending in "&" and a trailing
        # space. Cut at a word boundary and strip the punctuation.
        t = proc.split("(")[0].split(" above")[0].split(" under")[0].strip()
        if len(t) > 20:
            t = t[:20].rsplit(" ", 1)[0]
        out.append(t.rstrip(" &,-;").strip())
    if r["late"]:
        out.append("LATE")
    if r["blocked"] in ("YES", "ΝΑΙ", "NAI"):
        out.append("BLOCKED")
    return [t for t in out if t]


def main():
    src = newest(CFG["pattern"])
    rows = read_tracker(src, CFG["sheet"])

    # Deliverable grain for EPIRUS, contract grain for MINOR-MED - the same
    # distinction build_charts makes, and for the same reason: on MINOR-MED
    # the budget column is a CONTRACT total repeated across its rows, so one
    # task per row would create 76 tasks for 34 contracts and trebled money.
    #
    # For Nifty the money does not matter, but the duplication does: 42 of
    # those 76 tasks would be the same procurement listed again.
    if CFG["budget_grain"] == "contract":
        seen, unit = set(), []
        for r in rows:
            key = (r["ben"], r["no"])
            if key in seen:
                continue
            seen.add(key)
            unit.append(r)
        print(f"  contract grain: {len(rows)} rows collapsed to {len(unit)} "
              f"contracts, so {len(rows) - len(unit)} duplicate tasks avoided")
    else:
        unit = rows
        print(f"  deliverable grain: {len(unit)} positions, each its own task")

    out = HERE / f"nifty_{CFG['prefix']}_{CFG['name']}_tasks.csv"
    n_dated = 0
    with open(out, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Task Name", "Task Description", "List", "Status",
                    "Task Assignees", "Task Tags", "Due Date", "Start Date",
                    "Completed"])
        for r in unit:
            key = STATUS_KEY.get(str(r["status"]).strip().upper())
            assert key or not str(r["status"]).strip(), (
                f"row {r['row']}: status {r['status']!r} is not in STATUS_KEY. "
                f"Add it there rather than here, so the charts and this "
                f"import never disagree about what a status means.")
            # PLANNED publication and signature, not actual. A Gantt is a plan;
            # the actual dates are what you compare it against later. Both
            # come from _shared.TRACKER_COLUMNS so this script and the charts
            # can never disagree about which column is which.
            start, due = d(r.get("ppub")), d(r.get("psig"))
            # THE NAME HAS TO BE UNIQUE, because at deliverable grain one
            # contract spans several rows and the "what it buys" text is
            # identical on each. Two tasks called "1 Consulting services and
            # coordination efforts" are indistinguishable in a task list, and
            # whoever closes one has no idea which.
            name = f"{r['no']} {str(r.get('buys') or '')}".strip()[:110]
            deliv = str(r.get("deliv") or "").split("|")[0].strip()
            if CFG["budget_grain"] == "deliverable" and deliv:
                name = f"{name}  [{deliv}]"
            desc = " | ".join(x for x in [
                f"Deliverable: {r.get('deliv')}" if r.get("deliv") else "",
                f"Procedure: {r.get('proc')}" if r.get("proc") else "",
                f"Cost category: {r.get('cat')}" if r.get("cat") else "",
                f"Budget ex VAT: {r['budget']:,.2f}" if r["budget"] else "",
                f"Days late: {int(r['late'])}" if r["late"] else "",
                f"Blocked: {r['com']}" if r["com"] else "",
            ] if x)
            w.writerow([name or f"Contract {r['no']}", desc, r["ben"],
                        NIFTY_STATUS.get(key, "To Do"), r["ben"],
                        ";".join(tags(r)), due, start,
                        "true" if key == "done" else "false"])
            if start or due:
                n_dated += 1

    print(f"\n[written] {out.name}")
    print(f"  {len(unit)} tasks, {n_dated} with at least one date")
    if len(unit) > FREE_TASKS:
        print(f"  !! {len(unit)} tasks exceeds the free tier's {FREE_TASKS}")
    else:
        print(f"  free tier allows {FREE_TASKS} tasks, so this uses "
              f"{len(unit) / FREE_TASKS:.0%} of it")
    print(f"  two programmes = {FREE_PROJECTS} active projects = the whole "
          f"free allowance. AMMOS would need a paid plan or a deleted project.")
    return out


if __name__ == "__main__":
    main()
