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
| `compare_techsheets.py` | the three EPIRUSMEDEYE tech sheet .xlsx exports | `techsheet_comparison_report.txt` |
| `build_techsheet_report.py` | whatever `compare_techsheets.py` computed | the comparison as a .docx |

### The two that answer "how do I turn the export into a report"

`compare_techsheets.py` writes `techsheet_comparison_report.txt`, and
`build_techsheet_report.py` turns the same computed figures into a .docx. It
imports the comparison script rather than reading its text output, so no figure
is ever re-typed or re-parsed — if a number in the document is wrong, the
script that computed it is wrong, and both are in this folder.

    python compare_techsheets.py          # needs the three .xlsx exports here
    python build_techsheet_report.py      # writes the .docx

Put these three files in the folder first, exactly as the MIS names them:
`1.0-interreg_report26375.xlsx`, `1.1-interreg_report96606.xlsx`,
`2-interreg_report129557.xlsx`.

## ammos/

The AMMOS proposal builders. They read `_shared.py` from the folder above, so
keep the `ammos/` folder inside this one.

| script | produces |
|---|---|
| `build_budget_model.py` | the live budget workbook, with the rule checks as formulas |
| `build_concept_note.py` | the concept note |
| `build_wp_blended.py` | the work package structure |

These three need their source documents beside them in `ammos/` and they are
not in this repository, because they are the programme's files and not mine to
redistribute: `nextmed_call.txt`, `nextmed_guide.txt` and `mmm_outputs.xlsm`.
Each script asserts against them before it writes anything, so a missing file
fails loudly rather than producing a document with an unchecked claim in it.

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
