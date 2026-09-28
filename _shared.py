#!/usr/bin/env python3
"""The five things every one of your scripts needs.  Import, do not copy.

WHY THIS FILE EXISTS

Every script I have sent you repeats the same five pieces of machinery: the
dash rule, the report writer, finding spreadsheet columns by heading, reading
a tracker into rows, and the guard that refuses to ship a document mentioning
the wrong programme.  Copying them into each script means fixing every bug
five times - which already happened once, when the dash rule was wrong in two
files and right in the third.

Put this file beside your scripts and start each one with:

    from _shared import clean, Report, locate, read_tracker, guard

Then there is one copy of each rule, and fixing it fixes everything.
"""
import os
import pathlib
import re

import openpyxl

# ---------------------------------------------------------------- 1. THE DASH
# A dash is a RANGE only when DIGITS sit on both sides: "2021-2027", "0.2-1",
# "M1-M24".  Anything else is punctuation and becomes a comma and a space,
# because "--" reads as unfinished in a document going to a Ministry.
#
# My first version accepted any alphanumeric on either side, which quietly
# turned the title "AMMOS - WORK PACKAGE STRUCTURE" into "AMMOS-WORK PACKAGE
# STRUCTURE".  Word to word is punctuation; number to number is a range.
_RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(text):
    """Punctuation dashes to ', '.  Range dashes preserved."""
    t = str(text)
    t = _RANGE.sub("\x00", t)              # protect ranges
    # NOTE: this was t.replace("--", ", ") and the self-test below caught it.
    # On "a -- b" that produced "a ,  b" - a space BEFORE the comma and two
    # after - because the surrounding spaces survived. Swallow the whitespace
    # with the dash. A self-test that finds a bug in five lines of code has
    # already paid for itself.
    # Only a LONE pair of hyphens is punctuation. A run of three or more is a
    # separator rule, and the first version turned "-"*78 into ", , , , , ..."
    # right across the top of a generated report. Look at your output, not just
    # your regex.
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")             # restore ranges as plain hyphens
    return re.sub(r",\s*,", ",", t)


# -------------------------------------------------------------- 2. THE REPORT
class Report:
    """Collect every line, print it, and write the whole thing to a file.

    This is the whole of "exporting a report" in Python.  Instead of
    print(...) everywhere, you call out(...) and at the end out.save(path).
    Collect, then write.
    """

    def __init__(self):
        self.lines = []

    def __call__(self, text=""):
        t = clean(text)
        print(t)
        self.lines.append(t)

    def save(self, path):
        pathlib.Path(path).write_text("\n".join(self.lines) + "\n",
                                      encoding="utf-8")
        print(f"\n[written to {path}]")


# ------------------------------------------------------------- 3. THE COLUMNS
def locate(ws, wanted, hdr_row=3):
    """Find columns by HEADING, never by position.

    wanted is {key: (search term, another term)}.  Returns {key: column number}
    with None where nothing matched.

    This matters more than it looks.  Your two trackers have the same data in
    DIFFERENT columns, because one has the budget split into two columns and
    the other does not.  A script that says "column 7 is the budget" works on
    one file and silently reads the wrong number from the other.  A script that
    says "the column whose heading contains 'Budget'" works on both.
    """
    heads = {}
    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(row=hdr_row, column=c).value or "").replace("\n", " ")
        if v.strip():
            heads[c] = v.strip()
    found = {}
    for key, needles in wanted.items():
        found[key] = next((c for c, v in heads.items()
                           if any(n.lower() in v.lower() for n in needles)),
                          None)
    return found, heads


# -------------------------------------------------------------- 4. THE READER
TRACKER_COLUMNS = {
    "ben": ("Beneficiary",), "no": ("Contract #",),
    "buys": ("What the contract",), "proc": ("Procedure",),
    "deliv": ("Deliverable",), "cat": ("Cost category",),
    "budget": ("THIS deliverable", "Budget ex VAT"),
    "ctot": ("CONTRACT tender",),
    "apub": ("ACTUAL publication",), "aoff": ("ACTUAL offers",),
    "asig": ("ACTUAL signature",), "amt": ("Contract amount incl",),
    "preapp": ("Pre-approval obtained",), "blocked": ("Blocked",),
    "com": ("Comments",), "by": ("Updated by",),
    "status": ("STATUS",), "late": ("Days late",),
}


def read_tracker(path, sheet, hdr_row=3):
    """A procurement tracker sheet as a list of dicts, one per row."""
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    C, heads = locate(ws, TRACKER_COLUMNS, hdr_row)
    for must in ("ben", "status", "budget"):
        assert C[must], (f"cannot find the {must} column in {path}. "
                         f"Headings present: {list(heads.values())[:12]}")
    rows = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        if not ws.cell(row=r, column=C["ben"]).value:
            continue
        get = lambda k: (ws.cell(row=r, column=C[k]).value if C[k] else None)
        b = get("budget")
        row = {k: get(k) for k in C}
        row["row"] = r
        row["ben"] = str(get("ben")).strip()
        row["budget"] = float(b) if isinstance(b, (int, float)) else 0.0
        row["status"] = str(get("status") or "").strip()
        row["blocked"] = str(get("blocked") or "").strip().upper()
        row["com"] = str(get("com") or "").strip()
        row["late"] = get("late") if isinstance(get("late"), (int, float)) else 0
        rows.append(row)
    return rows


# --------------------------------------------------------------- 5. THE GUARD
# Names belonging to ONE programme only.  DIAGNOSIS NGO and EDEYPY are in both,
# so they are deliberately absent from both lists.
ONLY_EPIRUS = ["EPIRUSMEDEYE", "EPIRUS-MEDEYE", "6010984", "Ioannina",
               "Dropull", "13/10/2025", "13/10/2027"]
ONLY_MINORMED = ["MINOR-MED", "MINORMED", "MINOR MED", "6010952", "Vlore",
                 "Vlora", "Filiates", "ATHENA", "13/02/2026", "13/02/2028"]


def guard(path, forbidden):
    """Scan a finished .docx and DELETE it if a forbidden name appears.

    Call this as the last line of the script, after d.save(path).  It reads
    the file back, checks every paragraph and every table cell, and if the
    wrong programme is in there it removes the file and raises.

    Promising to be careful is worth nothing.  This makes the mistake
    impossible, and it exists because I made it: one document went out titled
    for one programme with a table row naming the other.
    """
    from docx import Document
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs) + "\n" + "\n".join(
        c.text for t in doc.tables for r in t.rows for c in r.cells)
    hits = [f for f in forbidden if f.lower() in text.lower()]
    if hits:
        os.remove(path)
        raise AssertionError(
            f"{os.path.basename(path)} mentions {hits}, which belongs to the "
            f"other programme. The file has been deleted, nothing was written.")
    print(f"[guard: clean against {len(forbidden)} forbidden names]")


# ----------------------------------------------------------- 6. FINDING FILES
# CAREFUL: HERE is the folder _SHARED.PY lives in, not the folder of the
# script that imported it. For every script sitting beside this file those are
# the same folder, which is why it reads as "here" and works.
#
# They are NOT the same for a script in a sub-folder. I converted the AMMOS
# scripts in epirus/ to use HERE and they immediately started looking for
# their spreadsheets one level up - the run failed with FileNotFoundError on
# mmm_outputs.xlsm, which was the good outcome. The bad outcome would have
# been a script that found a DIFFERENT file of the same name in the parent
# folder and reported on the wrong data without complaining.
#
# So a script that does not live beside _shared.py must use here(__file__).
HERE = pathlib.Path(__file__).resolve().parent


def here(script_file):
    """The folder of the script that calls this, not of _shared.py.

        BASE = str(here(__file__))

    Use this instead of HERE in any script kept in a sub-folder.
    """
    return pathlib.Path(script_file).resolve().parent


def newest(pattern, folder=None):
    """The most recently MODIFIED file matching a pattern, beside this script.

        src = newest("*EPIRUS*.xlsx")

    Why a pattern instead of a filename: every time you download the tracker
    from SharePoint you get a new name, or the same name with (1) on the end.
    A script with the filename written into it needs editing every single
    time, and the day you forget, it silently charts LAST month's data and
    every number in your report is quietly stale.

    A pattern picks up whatever you downloaded most recently, and prints which
    file it chose so you can see it was the one you meant.
    """
    base = pathlib.Path(folder) if folder else HERE
    hits = sorted(base.glob(pattern), key=lambda f: f.stat().st_mtime,
                  reverse=True)
    if not hits:
        raise FileNotFoundError(
            f"nothing matching {pattern!r} in {base}. "
            f"Download the export into that folder first.")
    if len(hits) > 1:
        print(f"[{len(hits)} files match {pattern!r}, using the newest]")
    print(f"[reading {hits[0].name}, modified "
          f"{__import__('datetime').datetime.fromtimestamp(hits[0].stat().st_mtime):%d/%m/%Y %H:%M}]")
    return hits[0]


if __name__ == "__main__":
    # a tiny self-test, so you can check the dash rule behaves
    cases = [("AMMOS — WORK PACKAGE", "AMMOS, WORK PACKAGE"),
             ("M1–M24", "M1-M24"),
             ("0.2–1 m", "0.2-1 m"),
             ("2021–2027", "2021-2027"),
             ("a -- b", "a, b")]
    for src, want in cases:
        got = clean(src)
        print(f"  {'ok  ' if got == want else 'FAIL'} {src!r} -> {got!r}")
