# Interreg reporting toolkit

Scripts for producing progress reports for EPIRUSMEDEYE (MIS 6010984) and
MINOR-MED (MIS 6010952) from the files the MIS and SharePoint already give
you, without retyping a number.

## Install once

    python -m pip install -r requirements.txt

and install poppler for `pdftotext` - see the notes at the bottom of
requirements.txt. Keep `_shared.py` in the same folder as the other scripts.

## What reads what

| script | reads | produces |
|---|---|---|
| `read_td.py` | a MIS technical sheet PDF | the approved budget as CSVs plus a summary, after 39 cross-checks |
| `diff_td.py` | two or more technical sheet versions | what each modification actually changed |
| `build_td_docx.py` | a MIS technical sheet PDF | the budget annex as a Greek .docx |
| `build_charts.py` | the procurement tracker .xlsx | four charts, EN or EL |
| `build_narrative.py` | the procurement tracker .xlsx | factual sentences, plus a prompt block of proven figures |

## Run

    python read_td.py report_140109.pdf
    python diff_td.py report_26346.pdf report_118212.pdf report_140109.pdf
    python build_td_docx.py report_140109.pdf
    python build_charts.py MINORMED
    python build_narrative.py MINORMED

Pass `EPIRUS` or `MINORMED`, and optionally `EN` or `EL`. With no argument at
all, `read_td.py` and `diff_td.py` pick up the PDFs sitting beside them.

## The one rule that matters

The technical sheet is the APPROVED BUDGET, not spending. Everything these
scripts produce from it says so. Spending comes from the MIS expenditure and
verification exports, which are a different file.

## Why the scripts stop instead of printing

Every table is checked against every other table that ought to agree with it,
and a document is refused rather than written when a check fails. A budget
annex with a plausible wrong number in it is worse than no annex, because the
wrong number gets believed. Read the lines the scripts print: they name the
file they chose and when it was last modified.
