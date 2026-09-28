#!/usr/bin/env python3
"""Narrative from the tracker: the facts written as sentences, in EN or EL.

YOUR QUESTION: "how do you get a narrative report of the data? Script, prompt?"

BOTH, AND THE SPLIT IS THE WHOLE ANSWER.

There are two different kinds of narrative in a progress report and they need
two different tools.

  FACTUAL narrative
      "Twenty five of forty seven contract positions are behind their planned
      date, the longest by two hundred and seventy nine days."
      Every word of that is computable. A SCRIPT should write it, because a
      script cannot get it wrong, cannot get tired, and produces exactly the
      same sentence next quarter from next quarter's data.

  INTERPRETIVE narrative
      "The delay profile is concentrated in the two partners who had no
      procurement officer until July, which is why the recovery plan puts
      support there rather than adding deadlines."
      That is judgement. No script can produce it. A model can help you draft
      it, and you have to own it.

THE DANGEROUS THING, AND IT IS THE ONE PEOPLE DO

Handing raw data to a language model and asking it to write the factual
narrative. It will produce fluent, confident, well-structured sentences with
numbers that are subtly wrong, and you will not catch them BECAUSE they read
well. A number in a sentence gets far less scrutiny than a number in a table.

So: the script below writes the factual paragraphs, with every figure computed
from the file. Then it writes a PROMPT BLOCK containing those same verified
facts, which you paste into a model to get help with the interpretation. The
model never sees the spreadsheet and never does arithmetic. It only ever works
from numbers this script has already proven.

That is the pattern. Compute, then narrate, then interpret. Never let the
interpretation layer touch the arithmetic.
"""
import collections
import datetime as dt

from _shared import HERE, Report, newest

from build_charts import CFG, LANG, PROGRAMME, STATUS_KEY, grain_check, load


def pc(x, dp=1):
    """Percent in the right convention: 52,4% in Greek, 52.4% in English."""
    t = f"{x:.{dp}%}"
    return t.replace(".", ",") if LANG == "EL" else t


def money(x):
    """Greek convention: 1.007.999,71 rather than 1,007,999.71."""
    if LANG != "EL":
        return f"{x:,.2f}"
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","}))


OUT = HERE / f"narrative_{CFG['prefix']}_{LANG}.txt"

S = {
    "EN": {
        "head": "{p}, MIS {mis}: factual narrative for the reporting period",
        "scope": ("The procurement plan for {p} contains {n} {unit} with a "
                  "total value of {tot} excluding VAT, spread across "
                  "{np} beneficiary partners."),
        "status": ("At the close of the period, {done_n} {unit} with a value "
                   "of {done_v} are under signed contract, which is "
                   "{done_pc} of the tracked value. {late_n} are behind their "
                   "planned date and {none_n} have not begun."),
        "period": ("Of the {sig} under signed contract, {before} was "
                   "signed before {p0} and {during} within the reporting "
                   "period, so {share} of all procurement value recorded as "
                   "signed on this project was signed during these {months} "
                   "months."),
        "after": ("A further {v} carries a signature date later than the "
                  "close of the reporting period on {p1}. It is excluded from "
                  "the figure for the period and reported as a commitment of "
                  "the next one."),
        "ahead": ("WARNING, {n} contract position(s) worth {v} carry an "
                  "ACTUAL signature date later than today. An actual date "
                  "cannot lie in the future, so a planned date has been "
                  "entered in the actual column. Do not describe these as "
                  "signed until the partner corrects the entry."),
        "delay": ("{n} {unit} are behind the date set in the beneficiaries' "
                  "own procurement plans. The longest delay is {worst} days "
                  "({who}, deliverable {what}) and the median is {med} days. "
                  "These are measured against the partners' own planning "
                  "dates, not against the approved application form, which "
                  "assigns every work package the same start and end date as "
                  "the project itself and therefore provides no intermediate "
                  "deadline."),
        "blocked": ("{n} {unit} with a combined value of {v}, being {pc} "
                    "of the tracked value, are recorded as blocked by the "
                    "beneficiary."),
        "partner": "{ben}: {n} {unit}, {v}, {pc} of the tracked value.",
        "quality": ("Data completeness: partners have recorded {asig} "
                    "signature dates, {apub} publication dates and {aoff} "
                    "offer deadlines against {n} {unit}."),
        "unit_s": "contract position", "unit_p": "contract positions",
        "unit_gen": "contract positions",
        "months": "four",
    },
    "EL": {
        "head": "{p}, MIS {mis}: αφήγηση πραγματικών στοιχείων περιόδου αναφοράς",
        "scope": ("Το Πλάνο Συμβάσεων του {p} περιλαμβάνει {n} {unit} συνολικής "
                  "αξίας {tot} χωρίς ΦΠΑ, κατανεμημένες σε {np} "
                  "δικαιούχους εταίρους."),
        "status": ("Κατά τη λήξη της περιόδου, {done_n} {unit} αξίας "
                   "{done_v} βρίσκονται υπό υπογεγραμμένη σύμβαση, ήτοι "
                   "{done_pc} της παρακολουθούμενης αξίας. {late_n} "
                   "παρουσιάζουν καθυστέρηση έναντι της προγραμματισμένης "
                   "ημερομηνίας και {none_n} δεν έχουν ξεκινήσει."),
        "period": ("Από τις {sig} υπό υπογεγραμμένη σύμβαση, {before} "
                   "υπογράφηκαν πριν την {p0} και {during} εντός της "
                   "περιόδου αναφοράς. Συνεπώς το {share} της συνολικής αξίας "
                   "των καταγεγραμμένων ως υπογεγραμμένων συμβάσεων "
                   "υπογράφηκε εντός αυτών των {months} μηνών."),
        "after": ("Επιπλέον ποσό {v} φέρει ημερομηνία υπογραφής μεταγενέστερη "
                  "της λήξης της περιόδου αναφοράς την {p1}. Δεν "
                  "συμπεριλαμβάνεται στο ποσό της περιόδου και αναφέρεται ως "
                  "δέσμευση της επόμενης."),
        "ahead": ("ΠΡΟΣΟΧΗ: {n} θέση/θέσεις συμβάσεων αξίας {v} φέρουν "
                  "ΠΡΑΓΜΑΤΙΚΗ ημερομηνία υπογραφής μεταγενέστερη της "
                  "σημερινής. Πραγματική ημερομηνία δεν μπορεί να είναι "
                  "μελλοντική, άρα έχει καταχωρισθεί προγραμματισμένη "
                  "ημερομηνία στη στήλη της πραγματικής. Να μην αναφερθούν ως "
                  "υπογεγραμμένες πριν διορθωθεί η καταχώριση."),
        "delay": ("{n} {unit} παρουσιάζουν καθυστέρηση έναντι των ημερομηνιών "
                  "του Πλάνου Συμβάσεων των ίδιων των δικαιούχων. Η μεγαλύτερη "
                  "καθυστέρηση ανέρχεται σε {worst} ημέρες ({who}, παραδοτέο "
                  "{what}) και η διάμεσος σε {med} ημέρες. Η μέτρηση γίνεται "
                  "έναντι του προγραμματισμού των εταίρων και όχι του "
                  "εγκεκριμένου Τεχνικού Δελτίου, το οποίο αποδίδει σε κάθε "
                  "πακέτο εργασίας τις ίδιες ημερομηνίες έναρξης και λήξης με "
                  "το έργο και επομένως δεν παρέχει ενδιάμεση προθεσμία."),
        "blocked": ("{n} {unit} συνολικής αξίας {v}, ήτοι {pc} της "
                    "παρακολουθούμενης αξίας, έχουν καταγραφεί από τον "
                    "δικαιούχο ως μπλοκαρισμένες."),
        "partner": "{ben}: {n} {unit}, {v}, {pc} της παρακολουθούμενης αξίας.",
        "quality": ("Πληρότητα στοιχείων: οι εταίροι έχουν καταχωρίσει {asig} "
                    "ημερομηνίες υπογραφής, {apub} ημερομηνίες δημοσίευσης και "
                    "{aoff} προθεσμίες προσφορών επί συνόλου {n} {unit}."),
        "unit_s": "σύμβαση", "unit_p": "συμβάσεις", "unit_gen": "συμβάσεων",
        "months": "τεσσάρων",
    },
}
TX = S[LANG]


def u(n):
    return TX["unit_s"] if n == 1 else TX["unit_p"]


def main():
    out = Report()
    rows, src = load()
    unit = grain_check(rows)
    tot = sum(r["budget"] for r in unit)
    v = collections.Counter()
    n = collections.Counter()
    for r in unit:
        v[r["key"]] += r["budget"]
        n[r["key"]] += 1
    bens = collections.Counter()
    bn = collections.Counter()
    for r in unit:
        bens[r["ben"]] += r["budget"]
        bn[r["ben"]] += 1

    out("=" * 78)
    out(TX["head"].format(p=CFG["name"], mis=CFG["mis"]))
    out(f"source: {src.name}   generated: "
        f"{dt.date.today():%d/%m/%Y}   grain: {CFG['budget_grain']}")
    out("=" * 78)
    out()
    out("FACTUAL NARRATIVE, every figure computed from the file above")
    out("-" * 78)
    out(TX["scope"].format(p=CFG["name"], n=len(unit), unit=u(len(unit)),
                           tot=money(tot), np=len(bens)))
    out()
    out(TX["status"].format(
        done_n=n["done"], unit=u(n["done"]), done_v=money(v["done"]),
        done_pc=pc(v['done']/tot), late_n=n["late"], none_n=n["none"]))

    signed = [r for r in unit if r["asig"]]
    if signed and CFG["period"]:
        # BOUNDED AT BOTH ENDS. This said ">= p0" with no upper bound, so a
        # contract signed after the period closed was counted inside it - and
        # the sentence it produced claimed a share of value "signed within
        # these four months" that included money signed outside them. Same
        # bug as the cumulative chart, same fix, and the reason _shared.py
        # exists is so this is the last place it can hide.
        p0, p1 = CFG["period"]
        before = sum(r["budget"] for r in signed if r["asig"].date() < p0)
        during = sum(r["budget"] for r in signed
                     if p0 <= r["asig"].date() <= p1)
        after = sum(r["budget"] for r in signed if r["asig"].date() > p1)
        if after:
            out()
            out(TX["after"].format(v=money(after), p1=f"{p1:%d/%m/%Y}"))
        if before + during:
            out()
            out(TX["period"].format(
                sig=money(before + during + after), before=money(before),
                during=money(during), p0=f"{p0:%d/%m/%Y}",
                share=pc(during / (before + during + after), 0),
                months=TX["months"]))
        # A future-dated ACTUAL signature is not a rounding quibble: it is a
        # contract being reported as signed when it is not. Named in the
        # narrative itself, not only on the console, because the narrative is
        # the thing he reads before he files.
        ahead = [r for r in signed if r["asig"].date() > dt.date.today()]
        if ahead:
            out()
            out(TX["ahead"].format(n=len(ahead),
                                   v=money(sum(r["budget"] for r in ahead))))

    late = sorted((r for r in unit if r["late"] > 0), key=lambda r: -r["late"])
    if late:
        med = sorted(r["late"] for r in late)[len(late) // 2]
        out()
        out(TX["delay"].format(
            n=len(late), unit=u(len(late)), worst=int(late[0]["late"]),
            who=late[0]["ben"], what=late[0]["deliv"].split("|")[0].strip(),
            med=int(med)))

    if n["blocked"]:
        out()
        out(TX["blocked"].format(n=n["blocked"], unit=u(n["blocked"]),
                                 v=money(v["blocked"]), pc=pc(v['blocked']/tot)))

    out()
    out("BY BENEFICIARY")
    for b, val in bens.most_common():
        out("  " + TX["partner"].format(ben=b, n=bn[b], unit=u(bn[b]), v=money(val),
                                        pc=pc(val/tot)))

    # WHAT PARTNERS ACTUALLY FILLED IN - AT THE SAME GRAIN AS EVERY OTHER
    # SENTENCE, WHICH IT WAS NOT.
    #
    # This counted filled CELLS out of len(rows), while every other sentence
    # in the narrative counts CONTRACTS. On MINOR-MED that printed "30
    # signature dates against 76 contracts", because 76 is the raw row count
    # and 34 is the number of contracts. Anybody dividing those gets 39%
    # completeness when the real figure is very different, and the sentence
    # gives them no way to know. Two different units under one heading is the
    # same defect as two different numbers for one thing.
    #
    # So completeness is now counted over the SAME units the rest of the
    # narrative describes: a contract counts as having a date if any of its
    # rows carries one.
    out()
    filled = {k: sum(1 for r in unit if r.get(k)) for k in
              ("asig", "apub", "aoff")}
    out(TX["quality"].format(asig=filled["asig"], apub=filled["apub"],
                             aoff=filled["aoff"],
                             n=len(unit), unit=TX["unit_gen"]))

    # ------------------------------------------------ the prompt block
    out()
    out("=" * 78)
    out("PROMPT BLOCK, paste this into Claude for the INTERPRETIVE narrative")
    out("=" * 78)
    out("Use ONLY the figures below. Do not calculate anything, do not infer "
        "any number that is not written here, and if something you want to say "
        "needs a figure that is absent, say so instead of estimating it.")
    out()
    out(f"Programme: {CFG['name']}, MIS {CFG['mis']}")
    if CFG["period"]:
        out(f"Reporting period: {CFG['period'][0]:%d/%m/%Y} to "
            f"{CFG['period'][1]:%d/%m/%Y}")
    out(f"Tracked procurement value: {tot:,.2f} EUR excluding VAT across "
        f"{len(unit)} {u(len(unit))} and {len(bens)} partners")
    for k in ("done", "going", "review", "late", "none", "blocked"):
        if n[k]:
            out(f"  {k:8s} {n[k]:3d} {u(n[k]):18s} {v[k]:>12,.2f}  "
                f"{v[k]/tot:6.1%}")
    if late:
        out(f"Contracts behind plan: {len(late)}, worst {int(late[0]['late'])} "
            f"days, median {int(sorted(r['late'] for r in late)[len(late)//2])} days")
    if CFG["total_af"]:
        out(f"Total eligible cost: {CFG['total_af']:,.2f}; the Article 3.7 "
            f"20% verified-expenditure threshold is {CFG['total_af']*0.2:,.2f}")
        out("NOTE: signed contract value is a COMMITMENT and is NOT verified "
            "expenditure. Do not treat them as the same thing.")
    else:
        out("Total eligible cost: NOT SUPPLIED, so do not state any percentage "
            "of the project budget or any Article 3.7 position.")

    # THE CAVEATS HAVE TO TRAVEL WITH THE FIGURES.
    #
    # The factual narrative above states them, but the PROMPT BLOCK is the
    # part that gets pasted somewhere else, and it was going out carrying
    # only the clean numbers. A model handed a signed total with no warning
    # attached will write a confident sentence about money that is not
    # signed, and the warning that would have stopped it stayed behind in a
    # file nobody re-reads.
    if signed and CFG["period"]:
        out()
        out("CAVEATS THAT MUST SURVIVE INTO THE PROSE:")
        out(f"- Value signed WITHIN the period: {during:,.2f}. This is the "
            f"only figure that may be described as signed during the period.")
        if before:
            out(f"- Signed BEFORE the period opened: {before:,.2f}. Not this "
                f"period's achievement.")
        if after:
            out(f"- Carries a signature date AFTER the period closed on "
                f"{CFG['period'][1]:%d/%m/%Y}: {after:,.2f}. A commitment of "
                f"the next period, not this one.")
        if ahead:
            out(f"- {len(ahead)} position(s) worth "
                f"{sum(r['budget'] for r in ahead):,.2f} carry an ACTUAL "
                f"signature date in the FUTURE, which cannot be true. Treat "
                f"as NOT signed and say the entry needs correcting.")
            # AND THE TWO CAVEATS OVERLAP, WHICH IS THE EASY THING TO MISS.
            #
            # A future-dated signature can sit INSIDE the reporting period -
            # dated next week, still inside the window. So it is counted in
            # "signed within the period" AND flagged as not really signed,
            # and anyone quoting the first figure while nodding along to the
            # second overstates the period by the overlap. Give the net
            # number rather than leaving it to be worked out.
            p0d, p1d = CFG["period"]
            overlap = sum(r["budget"] for r in ahead
                          if p0d <= r["asig"].date() <= p1d)
            if overlap:
                out(f"- {overlap:,.2f} of that future-dated value falls "
                    f"INSIDE the period and is therefore already counted in "
                    f"the {during:,.2f} above. The figure defensible as "
                    f"signed within the period TODAY is "
                    f"{during - overlap:,.2f}. Use that one if the partners "
                    f"do not correct the tracker before filing.")
    out()
    out("Write three paragraphs for a progress report to the Managing "
        "Authority: what the period achieved, what is at risk, and what is "
        "proposed for the next period. Do not repeat the figures above as a "
        "list; use them as evidence for judgements. Flag anything you cannot "
        "support.")

    out.save(OUT)


if __name__ == "__main__":
    main()
