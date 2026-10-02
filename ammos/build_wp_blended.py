#!/usr/bin/env python3
"""AMMOS work packages — my compliance skeleton with Conides's science poured in.

WHY A BLEND RATHER THAN A CHOICE

Conides's structure has the science.  Mine has the call's financial
architecture.  Neither works alone: his would fail the 50% MPC rule
structurally and put WP2 with a partner instead of the Lead Partner; mine had
no tourism economics and less specific field method.  So nothing of his is
discarded here - every content line from his file is placed, and the three
things he had that I did not are absorbed:

  * the tourism exploitation analysis, willingness to visit and willingness
    to pay.  I under-weighted this badly.  The database we capitalise from is
    the SUSTAINABLE TOURISM DATABASE and the topic attached to SO 2.2 is
    "tourism and the green transition" - a valuation survey is the evidence
    that ties the science to the call.
  * the specific field method: sea level rise scenarios from 0.2 to 1 metre,
    biodiversity spot mapping, photographic guide.  His draft also named the
    Hakanson methodology for the coastal dynamics; that has been REMOVED, see
    the note below.
  * the Limnos geopark absorbing maintenance cost after closure, which is the
    strongest sustainability argument in either document.

THE SITE MISUNDERSTANDING, CORRECTED

He read my "one site" recommendation as a proposal to run a single-site
project.  It was not.  It was strictly about Acipayam choosing between Ucari
Lake alone and Ucari Lake plus the Canyon.  The project has THREE
demonstration territories and always did - Lemnos, Ucari Lake and the
Lebanese coastal site - which is also what the 50% MPC rule and the transfer
logic both require.  Lemnos is the source territory and the calibration
baseline; it cannot be dropped and nobody proposed dropping it.

WHY THE TITLE HAS TO CHANGE

The working title says "dune systems".  Acipayam's confirmed site is a lake
shore.  An evaluator who reads "dune systems" and then finds a lacustrine
pilot will record an internal inconsistency, and they will be right.  The
unifying physical object across all three sites is not a dune, it is a
SEDIMENT SHORE - marine, lacustrine and riparian.  AMMOS still works, because
ammos is sand, and sand is exactly what the three have in common.

BUDGET FIGURES IN THIS DOCUMENT ARE THE AGREED ONES, NOT MY PREFERRED ONES

I have twice now put illustrative per-partner figures into a document that
was then forwarded to partners, and twice had to unpick a negotiation against
numbers nobody agreed.  So the budget table here carries the figures agreed on
22 September and nothing else.  The realignment that Acipayam's confirmed
operational role would justify is set out in its own section, explicitly
marked as not agreed, so it cannot be mistaken for an allocation.
"""
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

# TWO STANDING RULES HE ASKED FOR, 23 September.
#
# 1. No "--" and no dash used as punctuation: he forwards these files to the
#    Ministry, the Joint Secretariat and the partners, and a dash reads as
#    unfinished in formal correspondence.  Punctuation dashes become a comma
#    and a space.  RANGES are left alone - "M1-M24", "0.2-1 m", "2021-2027"
#    and "13/02/2026 - 13/02/2028" must keep meaning "to" rather than "and",
#    so the rule only fires on a doubled hyphen or a dash with whitespace
#    beside it, never on a dash between two alphanumerics.
#
# 2. My commentary goes at the END, in one block he can delete before
#    forwarding.  Everything above that block is partner-facing; everything
#    inside it is for him.
RANGE = re.compile(r"(?<=\d)\s*[\u2013\u2014]\s*(?=[A-Za-z]?\d)")


# A dash is a RANGE only when DIGITS sit on both sides: "2021-2027",
# "0.2-1", "M1-M24".  The first attempt accepted any alphanumeric and so
# turned the title "AMMOS - WORK PACKAGE STRUCTURE" into "AMMOS-WORK".
# Word to word is punctuation; number to number is a range.
def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = t.replace("--", ", ")
    t = re.sub(r"\s*[\u2013\u2014]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)

import pathlib
# WAS a hard-coded folder on MY server, so this script only ran on my
# machine. Now it uses its OWN folder.
#
# Note here(__file__) and not _shared.HERE: HERE is the folder _shared.py
# lives in, which is one level UP from this file. Using it sent these scripts
# looking for mmm_outputs.xlsm in the parent folder. here() takes the folder
# of the script that asks.
import sys as _sys
_sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from _shared import here

BASE = str(here(__file__))
OUT = f"{BASE}/AMMOS_Work_Packages_v4_1300000_NO_LEVER_2026-09-30.docx"

TOTAL_ELIGIBLE = 1_300_000.00
MPC_FLOOR = 0.50
ORG_CAP = 0.35
MAX_TECHNICAL_WPS = 4
MAX_ACTIVITIES = 3
SUBCONTRACT = 20_000.00
THREE_OFFERS_FROM, OPEN_TENDER_FROM = 15_000.00, 30_000.00

# PARTNER NUMBERS, FIXED BY HIM ON 02 OCTOBER AND NOT NEGOTIABLE HERE.
#
# PP00 to PP05, in his order, because LCEC is already preparing its papers as
# PP05 and the official budget template references partners by "Partner N."
# throughout. A partner that has started filling forms under one number must
# not be renumbered by us: every cross-reference in every annex would break,
# and the one thing worse than an unnumbered partnership is two numberings.
#
# Note it starts at PP00 for the Lead Partner, not PP01.
PP = {"HCMR": "PP00", "UAEG": "PP01", "PIYA": "PP02",
      "ACIP": "PP03", "CNR": "PP04", "LCEC": "PP05"}

P = {
    "HCMR": ("HCMR", "Greece", "EUMC"),
    "UAEG": ("University of the Aegean", "Greece", "EUMC"),
    "PIYA": ("Pi Youth Association", "Türkiye", "MPC"),
    "ACIP": ("Acıpayam Municipality", "Türkiye", "MPC"),
    "LCEC": ("LCEC", "Lebanon", "MPC"),
    "CNR": ("ISAC, CNR, Cagliari Section", "Italy", "EUMC"),
}
assert set(PP) == set(P), (
    f"the partner-number map and the partner list disagree: "
    f"{set(PP) ^ set(P)}. Every partner needs a PP number and every PP "
    f"number needs a partner.")
assert len(set(PP.values())) == len(PP), "two partners share a PP number"
# 30 September, six partners, total eligible 1.300.000, NO 4.4.3(b).
#
# Total eligible per partner is direct + 0.30 x staff, because office and
# administration and travel and accommodation are both flat rates of 15% of
# staff. So the figures below land on: HCMR 350.000, UAEG 205.000,
# LCEC 300.000, PIYA 230.000, ACIP 135.000, CNR 80.000.
#
# These are the 22 September figures restored, plus CNR. The 28 September
# version had HCMR at 363.000 and the Aegean at 215.000, paid for by counting
# CNR's work in MPC territory toward the floor - which CNR then refused to
# sign. So the Greeks give the 24.000 back and nothing in the budget needs
# anybody to interpret a rule.
DIRECT = {"HCMR": 297_950.0, "UAEG": 187_000.0, "LCEC": 276_000.0,
          "PIYA": 215_000.0, "ACIP": 126_000.0, "CNR": 68_000.0}
STAFF = {"HCMR": 173_500.0, "UAEG": 60_000.0, "LCEC": 80_000.0,
         "PIYA": 50_000.0, "ACIP": 30_000.0, "CNR": 40_000.0}

# HAKANSON IS OUT, 01 OCTOBER.
#
# Conides's draft named the Hakanson methodology for the coastal dynamics
# analysis. I then moved that analysis onto CNR's line without checking they
# use it, and Dr Simeone told us plainly: "personally I don't know
# specifically the Hakanson method". That is the same error as the seven
# sheets, two days apart - assigning a NAMED thing to a partner without
# checking they do that named thing.
#
# It comes out for a better reason than his modesty. The method we need for
# the physical and sea-level-rise work is ALREADY INSIDE the output we are
# capitalising: MMM_IF02 sheets 2.4 (submerged topography, including sediment
# exchange between emerged and submerged beach) and 2.6 (meteo-marine
# forcings: wave setup, runup, wind setup, tide, and the effect of morphology
# and bathymetry on coastal flooding). Importing an unrelated external method
# into a project whose premise is transferring THIS output weakens criterion
# 4.3 and criterion 2.1 at once.
#
# AND THE APPLICATION FORM IS CONTRACTUALLY BINDING. Naming a method here and
# using a different one later is a formal modification, not a free choice.
# Simone's instinct to keep the method description general is right; the
# answer is to be specific about the SHEETS and the SELECTION CRITERIA and
# general about which method wins at which site.
#
# PENDING: Simone to name the two or three candidate approaches he would
# actually compare for SLR on a sandy shoreline. Until he does, the text says
# "selected per site under activity 3.1" and names no method.

# THE 4.4.3(b) COMPONENT IS ZERO, AS OF 30 SEPTEMBER.
#
# CNR refused to sign it and was right to: 4.4.3(b) wants budgeted DIRECT
# costs itemised in a dedicated table, travel is a flat rate the Guidelines
# say need not be documented at all, and an Italian researcher's salary is
# paid in Italy. Nothing was itemisable. The floor is now carried by the MPC
# partners themselves under 4.4.3(a), at 51,15% with 15.000 of margin.
#
# The two fields stay in the code rather than being deleted, because the
# check below tests BOTH denominators with and without them and prints both.
# A zero that is computed and shown is stronger than an absence.
CNR_MPC_DIRECT = 0.0
CNR_MPC_STAFF = 0.0

TITLES = [
    ("AMMOS — Adaptive Management of Mediterranean shOreline Sediments under "
     "tourism pressure", "RECOMMENDED",
     "Keeps the acronym, which partners already use, and replaces \"dune "
     "systems\" with the object all three sites actually share. Covers marine, "
     "lacustrine and riparian sediment shores without straining."),
    ("AMMOS — Adaptive Management of Mediterranean shOres: from Source to Sea",
     "if you want the litter argument in the title",
     "Puts Acıpayam's source-to-sea framing on the cover. Strong, but it "
     "foregrounds litter over erosion and resilience, which are the larger part "
     "of what we transfer."),
    ("AMMOS — Adaptive Management of Mediterranean dune systems under tOurism "
     "preSsure", "CURRENT — do not submit",
     "\"Dune systems\" contradicts a confirmed lacustrine pilot site. An "
     "evaluator will read the title, reach Ucarı Lake, and record an internal "
     "inconsistency in the proposal."),
]

SITES = [
    ("Lemnos — Gomati", "Greece", "HCMR", "Coastal dune system, marine",
     "The source territory and the calibration baseline. All seven AMMIRARE "
     "method sheets run here, including the two that travel nowhere else — "
     "Posidonia meadow and submerged topography. This is where the method is "
     "proven before it is adapted, and it is the reason the adaptation is "
     "credible at all."),
    ("Ucarı Lake Recreation Area", "Türkiye", "ACIP with PIYA",
     "Lacustrine sediment shore",
     "Unconsolidated sediment, sand and fine gravel banks, managed by the "
     "municipality, under heavy seasonal visitor concentration. Sheets 2.1 "
     "vegetation, 2.3 drone topography and 2.7 ecological status transfer "
     "directly; the COMMON litter protocol works on any shoreline. Posidonia "
     "and marine bathymetry are excluded as non-applicable — Acıpayam "
     "identified that themselves."),
    ("Lebanese coastal site — to be confirmed by LCEC", "Lebanon", "LCEC",
     "Coastal, sediment shore",
     "The second coastal territory and the Arabic-language route. LCEC's own "
     "national contribution already sits inside MMM_NX37, so the policy half "
     "of the transfer is theirs to extend rather than to start."),
    ("Acıpayam Canyon — Benlik/Çorlak fluvial corridor", "Türkiye",
     "documented, NOT delivered", "Riparian sediment banks",
     "A REPLICATION TARGET, not a fourth demonstration site. Baseline "
     "documented in WP6 so the proposal can show the method travels to a third "
     "shore typology, without taking on a site whose pressure is uncontrolled "
     "hiking, off-road driving and wild camping — which is far harder to "
     "manage inside 24 months than a municipality-run recreation area."),
]

# (id, title, level, lead, months, description, activities, who)
# Items marked [C] come from Conides's file and are placed, not invented.
WPS = [
    ("WP1", "Management and Coordination", "MANDATORY — Lead Partner", "HCMR",
     "M1–M24",
     "Coordination, financial management, monitoring and risk management. HCMR "
     "carries legal responsibility to the Managing Authority and cannot "
     "delegate the work package.",
     [("1.1", "Scientific responsible and roles; management committee with one "
              "representative and one alternate per partner; RACI table; "
              "management process  [C]"),
      ("1.2", "Financial management, reporting, first level control "
              "coordination, connection with the funder  [C]"),
      ("1.3", "Deliverable timing and quality check; monitoring, evaluation "
              "and risk management; optional external advisory committee  [C]")],
     [("HCMR", "Lead. Reporting to the MA, financial consolidation."),
      ("ALL", "One nominated member and one alternate each on the management "
              "committee.")]),

    ("WP2", "Communication and dissemination", "MANDATORY — Lead Partner",
     "HCMR", "M1–M24",
     "Communication across three countries and four languages. The work package "
     "is the Lead Partner's because FAQ 1.11 requires it; delivery sits where "
     "the reach is.",
     [("2.1", "Website with blog and password-protected partner area, "
              "integrated into the municipality, prefecture and regional "
              "websites; multilingual channels EL, TR, AR, EN  [C]"),
      ("2.2", "Newsletters to local sectoral press in every partner country — "
              "one at month 1 on objectives, one at month 24 on results; "
              "booklet for tourists; on-site interpretation signage  [C]"),
      ("2.3", "Local workshops — one stakeholder, one results — and the final "
              "transferability conference; one scientific paper drafted; "
              "interim and final reports at months 12 and 24  [C]")],
     [("HCMR", "Lead. Scientific content, the reports and the paper, "
               "answerability to the MA for the visibility obligations."),
      ("PIYA", "Activity lead for 2.1 and 2.3 — website, channels, campaign "
               "design and the organisation of events."),
      ("ALL", "National-language adaptation and local press in each "
              "territory.")]),

    ("WP3", "Characterising and adapting the method",
     "TECHNICAL — knowledge level", "HCMR", "M1–M12",
     "The capitalisation core. MMM_IF02 is understood, the three receiving "
     "territories are characterised with the SAME indicators collected the SAME "
     "way, and the method is rewritten so it works on a marine dune, a lake "
     "shore and a river bank. This is where Conides's dune study lives — with "
     "one change: it happens in three territories, not one.",
     [("3.1", "Transfer workshop: the AMMIRARE and COMMON outputs presented to "
              "the partnership, transfer requirements agreed, and the modular "
              "decision recorded — which of the seven method sheets applies at "
              "which site"),
      ("3.2", "Territorial baselines: GIS mapping of shore state, utilities and "
              "roads; sea level rise scenarios 0.2 m to 1 m; biodiversity state "
              "and GIS map of biodiversity spots with species identification; "
              "geomorphological analysis of coastline and climate change "
              "effects, by a method selected per site under activity 3.1 and "
              "drawn from MMM_IF02 sheets 2.3, 2.4 and 2.6; marine coastal "
              "survey 0–100 m with Posidonia GIS, sea bottom type and fishing "
              "activity (Lemnos only); photographic guide  [C]"),
      ("3.3", "Tourism and valuation baseline: visitor numbers and seasonality "
              "at each site, tourist survey on willingness to visit and "
              "willingness to pay, and the potential effects of tourism on "
              "water and energy  [C]")],
     [("HCMR", "Lead. Adaptation of MMM_IF02, the Lemnos baseline, the "
               "ecological and biological sheet transfer and the marine "
               "survey. "
               "Scientific coordination of cross-site comparability, and the "
               "valuation component under activity 3.3."),
      ("UAEG", "Spatial data standards and the GIS layer, so three national "
               "baselines are actually comparable rather than merely similar."),
      ("ACIP", "Baseline of the Ucarı Lake shore, including visitor numbers."),
      ("LCEC", "Baseline of the Lebanese site and the energy profile of "
               "existing visitor facilities."),
      ("PIYA", "Field data collection with the volunteer network."),
      ("CNR", "Activity lead, 3.1. Method owner: ISAC, CNR researchers "
              "authored the AMMIRARE coastal dynamics work this project "
              "capitalises, so the transfer workshop and the modular decision "
              "on which of the seven method sheets applies at which site are "
              "led by the people who wrote them. Quality assurance of all "
              "three territorial baselines: on site at Lemnos and at Ucarı "
              "Lake, and as structured remote review of LCEC's fieldwork for "
              "the Lebanese baseline, which the current security situation "
              "makes the responsible arrangement. The purpose either way is "
              "that the three baselines are comparable in fact and not only "
              "in format.")]),

    ("WP4", "Demonstrating it on three shores", "TECHNICAL — technical level",
     "UAEG", "M8–M22",
     "The adapted method applied at real sites in three countries, with the "
     "same indicators collected the same way so the comparison means "
     "something. Conides's intervention list, delivered at every site rather "
     "than only on Lemnos.",
     [("4.1", "Visitor channelling and shore protection at each site: border "
              "marking poles around the protected area, tourist information "
              "signs showing boundaries, value, important species and entry "
              "guidelines, and simple protective barriers around biodiversity "
              "spots; leaflets for hotels, rooms-to-let and municipal "
              "info-kiosks  [C]"),
      ("4.2", "Low-impact visitor access: demountable modular boardwalks, "
              "demarcated zones and fencing to prevent shoreline vegetation "
              "trampling and bank degradation"),
      ("4.3", "Shared web-GIS monitoring platform carrying the common protocol "
              "across all three territories, with its own data collection "
              "endpoint — the COMMON protocol's own pipeline ends at a third "
              "party's form and cannot be used to report project indicators")],
     [("UAEG", "Lead. Web-GIS platform, spatial monitoring, data "
               "comparability."),
      ("HCMR", "Design and scientific supervision of the Lemnos "
               "interventions."),
      ("ACIP", "Activity lead for 4.1 and 4.2 at Ucarı Lake: the municipality "
               "is the authority that manages the site and its visitors, which "
               "is what makes it the right partner and not merely a willing "
               "one."),
      ("LCEC", "Energy and low-carbon specification of visitor facilities "
               "across all three sites — lighting, off-grid services, seasonal "
               "demand."),
      ("PIYA", "Field data collection and monitoring with volunteers.")]),

    ("WP5", "Building the capacity to keep doing it",
     "TECHNICAL — societal level", "PIYA", "M6–M22",
     "Transferring the human half of the AMMIRARE and COMMON outputs: the "
     "people who will still be managing these shores after the project closes.",
     [("5.1", "Shore Custodians — MMM_IF01 Beach Custodians adapted to a lake "
              "shore and a coastal dune, run with schools in all three "
              "countries; teaching material from the previous project adapted "
              "to dunes and sediment shores  [C]"),
      ("5.2", "Site-manager training built on the COMMON Capacity Building "
              "Training Design Plan, MMM_NX38"),
      ("5.3", "Visitor awareness at the sites, adapting the BEach CLEAN "
              "decalogue MMM_NX35, which already exists in Arabic")],
     [("PIYA", "Lead. Beach Custodians is written for exactly this "
               "organisation — youth and training institutions, with an "
               "accredited volunteering programme behind it."),
      ("ACIP", "Hosts the Shore Custodians field activities at Ucarı Lake with "
               "local schools and youth groups, coordinated jointly with "
               "PIYA."),
      ("HCMR", "Scientific curriculum content."),
      ("LCEC", "Arabic adaptation and delivery in Lebanon."),
      ("UAEG", "The monitoring platform as a teaching tool.")]),

    ("WP6", "Making it stick", "TECHNICAL — regulatory level", "LCEC",
     "M12–M24",
     "The difference between a pilot and a transfer is whether anybody adopts "
     "it. This produces the instruments that outlive the funding and puts them "
     "in front of the bodies that can apply them.",
     [("6.1", "A costed management plan for each site, taken through "
              "stakeholder consultation so it is adopted rather than filed; "
              "protection scheme guidelines and tourist visit management "
              "schemes  [C]"),
      ("6.2", "The Lemnos route: discussions with the Limnos geopark to bring "
              "the area into their network and to cover maintenance costs "
              "through their activities after the project ends  [C]"),
      ("6.3", "Joint transferability guide and policy recommendations in "
              "English and Arabic, extending MMM_NX37; plus the documented "
              "baseline of the Acıpayam Canyon as the next transfer target, "
              "proving the method reaches a third shore typology")],
     [("LCEC", "Lead. A national agency under a ministry is the only partner "
               "with a policy channel of its own, and MMM_NX37 already carries "
               "a Lebanese national contribution."),
      ("ACIP", "Activity lead for 6.1 — a municipal authority formally "
               "adopting a transferred management plan is the clearest "
               "demonstration that the method survives the project."),
      ("HCMR", "Activity lead for 6.2, the geopark route and the scientific "
               "case."),
      ("ASSOC", "Ministry of Tourism, Ministry of Maritime Affairs and the "
                "municipalities as associates: the uptake route and the "
                "evidence for award criterion 3.2.")]),
]
WPIDS = [w[0] for w in WPS]

SPREAD = {
    "HCMR": {"WP1": 0.16, "WP2": 0.12, "WP3": 0.36, "WP4": 0.20, "WP5": 0.10,
             "WP6": 0.06},
    "UAEG": {"WP1": 0.05, "WP2": 0.05, "WP3": 0.18, "WP4": 0.52, "WP5": 0.14,
             "WP6": 0.06},
    "LCEC": {"WP1": 0.05, "WP2": 0.06, "WP3": 0.16, "WP4": 0.30, "WP5": 0.21,
             "WP6": 0.22},
    "PIYA": {"WP1": 0.05, "WP2": 0.22, "WP3": 0.06, "WP4": 0.15, "WP5": 0.44,
             "WP6": 0.08},
    "ACIP": {"WP1": 0.05, "WP2": 0.06, "WP3": 0.14, "WP4": 0.26, "WP5": 0.15,
             "WP6": 0.34},
    # CNR is the method owner, so its money sits where the method is
    # characterised and where the teams are trained, and almost nowhere else.
    # A partner spread thinly across all six work packages reads as a partner
    # added for its flag rather than for its work, and an evaluator scoring
    # criterion 3 will say so.
    "CNR": {"WP1": 0.04, "WP2": 0.04, "WP3": 0.46, "WP4": 0.16, "WP5": 0.26,
            "WP6": 0.04},
}

NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)
GREEN = RGBColor(0x2E, 0x70, 0x59)


def matrix():
    m = {}
    for c, sh in SPREAD.items():
        assert abs(sum(sh.values()) - 1.0) < 1e-9, f"{c} shares != 1"
        m[c] = {w: round(DIRECT[c] * s, 2) for w, s in sh.items()}
        drift = round(DIRECT[c] - sum(m[c].values()), 2)
        if drift:
            m[c]["WP4"] = round(m[c]["WP4"] + drift, 2)
    return m


def check(m):
    tech = [w for w in WPS if w[2].startswith("TECHNICAL")]
    assert len(tech) <= MAX_TECHNICAL_WPS, f"{len(tech)} technical WPs"
    mand = [w for w in WPS if w[2].startswith("MANDATORY")]
    assert [w[0] for w in mand] == ["WP1", "WP2"], "WP1+WP2 must be the pair"
    for w in mand:
        assert w[3] == "HCMR", f"{w[0]} not with the Lead Partner (FAQ 1.11)"
    for w in WPS:
        assert len(w[6]) <= MAX_ACTIVITIES, f"{w[0]} has {len(w[6])} activities"
    leads = {w[3] for w in WPS}
    acts = {c for w in WPS for c, t in w[7] if "ctivity lead" in t}
    silent = set(P) - leads - acts
    assert not silent, f"these partners lead nothing: {sorted(silent)}"

    elig = {k: DIRECT[k] + 0.30 * STAFF[k] for k in DIRECT}
    td, te = sum(DIRECT.values()), sum(elig.values())
    assert abs(te - TOTAL_ELIGIBLE) < 0.5, f"{te:,.2f} != {TOTAL_ELIGIBLE:,.2f}"
    mpc_e = sum(elig[k] for k in P if P[k][2] == "MPC")
    mpc_d = sum(DIRECT[k] for k in P if P[k][2] == "MPC")

    # BOTH DENOMINATORS, AND NOW BOTH NUMERATORS TOO.
    #
    # The Guidelines say total eligible costs; the application template asks
    # for direct. I once switched to direct, read 49,61% as compliant on the
    # wrong basis, and only caught it by testing both - so both stay.
    #
    # On THIS split the two disagree in the opposite direction, which is
    # exactly why the pair is worth keeping. Partner location alone gives
    #     51,03% on direct costs        PASSES
    #     49,50% on total eligible      FAILS
    # A check written against direct only would have reported this budget as
    # compliant, and the rule that actually applies is the one it fails.
    #
    # So the 4.4.3(b) component is added to each numerator with the
    # arithmetic that denominator requires.
    mpc_e_lever = round(CNR_MPC_DIRECT + 0.30 * CNR_MPC_STAFF, 2)
    assert CNR_MPC_DIRECT <= DIRECT["CNR"] and CNR_MPC_STAFF <= STAFF["CNR"], (
        "the MPC-implemented part of CNR's budget cannot exceed CNR's budget")
    for label, num, lever, den in (("eligible", mpc_e, mpc_e_lever, te),
                                   ("direct", mpc_d, CNR_MPC_DIRECT, td)):
        bare, full = num / den, (num + lever) / den
        assert full >= MPC_FLOOR, (
            f"MPC on {label} is {full:.2%} even with 4.4.3(b), below the "
            f"{MPC_FLOOR:.0%} floor")
        if bare < MPC_FLOOR:
            print(f"  !! MPC on {label}: partners alone {bare:.2%} is BELOW "
                  f"the floor. With 4.4.3(b) {full:.2%}. This split is "
                  f"compliant ONLY if CNR works on the Turkish and Lebanese "
                  f"sites.")
        else:
            print(f"  MPC on {label}: partners alone {bare:.2%}, "
                  f"with 4.4.3(b) {full:.2%}")
    big = max(elig, key=elig.get)
    assert elig[big] / te <= ORG_CAP and DIRECT[big] / td <= ORG_CAP
    assert THREE_OFFERS_FROM <= SUBCONTRACT < OPEN_TENDER_FROM
    # every site must have an owner, and Lemnos must still be in the project
    owners = " ".join(s[2] for s in SITES)
    assert "HCMR" in owners, "Lemnos has lost its owner"
    assert any("Lemnos" in s[0] for s in SITES), "Lemnos is missing"
    assert sum(1 for s in SITES if "NOT delivered" not in s[2]) == 3, \
        "there must be exactly three DELIVERED demonstration territories"
    print(f"{len(tech)} technical WPs; WP1+WP2 with the LP; "
          f"3 delivered sites + 1 replication target")
    # This line used to print the partner-only shares with no mention of the
    # lever, immediately under the warning saying the partner-only share is
    # below the floor. Two numbers for one thing, three lines apart, and the
    # reassuring-looking one last.
    print(f"MPC including 4.4.3(b): {(mpc_e + mpc_e_lever)/te:.2%} of total "
          f"eligible / {(mpc_d + CNR_MPC_DIRECT)/td:.2%} of direct; "
          f"largest {big} {elig[big]/te:.2%} eligible")
    for w in WPIDS:
        t = sum(m[c][w] for c in P)
        print(f"  {w} {t:>11,.2f}  {t/td:6.2%}")


def main():
    m = matrix()
    check(m)
    td = sum(DIRECT.values())
    d = Document()
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(10.5)

    def h(t, size=13, colour=NAVY, space=10, after=4):
        p = d.add_paragraph()
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(clean(t))
        r.bold = True
        r.font.size = Pt(size)
        r.font.color.rgb = colour

    def para(t, bold=False, colour=None, size=10.5, after=5, italic=False):
        p = d.add_paragraph()
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(clean(t))
        r.bold, r.italic = bold, italic
        r.font.size = Pt(size)
        if colour is not None:
            r.font.color.rgb = colour

    def bul(items, size=10, colour=None):
        for t in items:
            p = d.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(3)
            r = p.add_run(clean(t))
            r.font.size = Pt(size)
            if colour is not None:
                r.font.color.rgb = colour

    def table(cols, rows, widths=None, size=9):
        tb = d.add_table(rows=1, cols=len(cols))
        tb.style = "Light Grid Accent 1"
        for i, x in enumerate(cols):
            tb.rows[0].cells[i].text = clean(x)
            for pp in tb.rows[0].cells[i].paragraphs:
                for rr in pp.runs:
                    rr.bold = True
                    rr.font.size = Pt(size)
        for row in rows:
            c = tb.add_row().cells
            for i, x in enumerate(row):
                c[i].text = clean(x)
                for pp in c[i].paragraphs:
                    for rr in pp.runs:
                        rr.font.size = Pt(size)
        return tb

    t = d.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(clean("AMMOS — WORK PACKAGE STRUCTURE"))
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = NAVY
    s = d.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = s.add_run(clean("Blended: the compliance structure with HCMR's science placed "
                  "inside it  ·  23 September 2026  ·  items marked [C] are "
                  "from the Conides file"))
    r.font.size = Pt(10.5)
    r.font.color.rgb = GREY

    h("Title")
    para("AMMOS  ,  Adaptive Management of Mediterranean shOreline Sediments "
         "under tourism pressure", bold=True, size=12, colour=NAVY)
    para("ammos = sand. The object all three demonstration territories share is "
         "a sediment shore: marine at Lemnos, lacustrine at Ucarı Lake, coastal "
         "in Lebanon, and riparian at the replication target.", size=10,
         colour=GREY)

    h("The three territories, and the fourth that is documented not delivered")
    table(["Site", "Country", "Owner", "Typology", "Role"],
          [(a, b, c, dd, e) for a, b, c, dd, e in SITES],
          size=8.5)

    h("The six work packages")
    for wid, title, level, lead, months, desc, acts, who in WPS:
        h(f"{wid} — {title}", size=12.5, space=12)
        para(f"{level}   ·   Lead: {P[lead][0]} ({P[lead][1]})   ·   {months}   "
             f"·   direct cost {sum(m[c][wid] for c in P):,.0f} "
             f"({sum(m[c][wid] for c in P)/td:.1%})",
             italic=True, colour=GREY, size=9.5, after=4)
        para(desc, size=10)
        para("Key activities", bold=True, size=10, after=2)
        for num, text in acts:
            p = d.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            rr = p.add_run(clean(num))
            rr.bold = True
            rr.font.size = Pt(9.5)
            r2 = p.add_run(clean("  " + text))
            r2.font.size = Pt(9.5)
        para("Who does what", bold=True, size=10, after=2)
        for code, text in who:
            label = ("All partners" if code == "ALL" else
                     "Associates" if code == "ASSOC" else P[code][0])
            p = d.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            rr = p.add_run(clean(label + ": "))
            rr.bold = True
            rr.font.size = Pt(9.5)
            r2 = p.add_run(clean(text))
            r2.font.size = Pt(9.5)

    h("Budget per work package — the shape the form wants")
    para("Figures agreed on 22 September. Direct costs only, because the two "
         "flat rates are calculated by the e-form from staff cost and cannot be "
         "distributed by hand.", size=10)
    rows = []
    for c in P:
        rows.append([P[c][0]] + [f"{m[c][w]:,.0f}" for w in WPIDS] +
                    [f"{DIRECT[c]:,.0f}",
                     f"{DIRECT[c] + 0.30*STAFF[c]:,.0f}"])
    rows.append(["TOTAL"] + [f"{sum(m[k][w] for k in P):,.0f}" for w in WPIDS] +
                [f"{td:,.0f}", f"{TOTAL_ELIGIBLE:,.0f}"])
    table(["Partner"] + WPIDS + ["Direct", "ELIGIBLE"], rows, size=8.5)

    h("Who leads what")
    rows = []
    for c in P:
        leads = [w[0] for w in WPS if w[3] == c]
        al = [w[0] for w in WPS for cc, tt in w[7]
              if cc == c and "ctivity lead" in tt]
        other = [w[0] for w in WPS if w[3] != c and w[0] not in al
                 and any(cc in (c, "ALL") for cc, _ in w[7])]
        # "—" AS THE EMPTY PLACEHOLDER WAS EATEN BY MY OWN DASH RULE.
        #
        # clean() turns a dash with whitespace either side into ", ", because
        # he asked for no dashes as punctuation. Applied to a lone em dash it
        # returns ", " - so three cells in this partner-facing table read
        # ", " instead of "none", and it shipped. The rule did exactly what it
        # was told; the placeholder was the wrong character to choose in a
        # document that strips that character.
        rows.append([f"{P[c][0]} ({P[c][1]})", ", ".join(leads) or "none",
                     ", ".join(al) or "none", ", ".join(other) or "none"])
    table(["Partner", "Leads", "Activity lead", "Contributes to"], rows, size=9)

    # THE AUTONOMOUS REGION OF SARDINIA CANNOT APPEAR IN THE PARTNER-FACING
    # DOCUMENT, IN ANY ROLE.
    #
    # RAS is the Managing Authority of Interreg NEXT MED and also holds the
    # audit function, so naming it as a partner, an associated partner or the
    # source of a letter of support puts the awarding and auditing authority
    # inside a proposal it assesses. This runs BEFORE the reviewer block,
    # because that block has to explain the rule and would trip a guard that
    # scanned the whole file.
    #
    # "RAS" needs word boundaries: as a bare substring it matches
    # "infrastructure", which appears in this document legitimately, and a
    # guard that fires on ordinary prose is a guard that gets switched off.
    body = "\n".join(p.text for p in d.paragraphs) + "\n" + "\n".join(
        c.text for tb in d.tables for r in tb.rows for c in r.cells)
    hits = [w for w in ("Autonomous Region of Sardinia", "Regione Autonoma")
            if w.lower() in body.lower()]
    if re.search(r"\bRAS\b", body):
        hits.append("RAS")
    assert not hits, (
        f"the partner-facing work package document names {hits}. The "
        f"Autonomous Region of Sardinia is this programme's MANAGING "
        f"AUTHORITY and audit authority and cannot be in the partnership.")

    # ---------------------------------------------------------- REVIEWER NOTES
    # He asked for my comments to sit at the END so he can read them, act on
    # them, delete the block and forward the rest.  Everything above this page
    # break is partner-facing.  Everything below it is for him alone.
    d.add_page_break()
    h("REVIEWER NOTES", size=15, colour=RED)
    para("This block is for you, not for the partners. Read it, decide the open "
         "points, then DELETE FROM THIS PAGE TO THE END before you forward the "
         "document. Nothing below this line is partner-facing.",
         bold=True, colour=RED)

    h("First, a correction: the project has three sites, not one", size=12)
    para("You read my \"one site\" recommendation as a proposal to run a "
         "single-site project. It was not, and I should have been clearer. It "
         "was strictly about Acıpayam choosing between Ucarı Lake alone and "
         "Ucarı Lake plus the Canyon.", bold=True, colour=RED)
    para("AMMOS has three demonstration territories and always has: Lemnos, "
         "Ucarı Lake and the Lebanese coastal site. Lemnos is not optional — it "
         "is the source territory and the calibration baseline, the only site "
         "where all seven AMMIRARE method sheets run, and the reason the "
         "adaptation to a lake shore is credible rather than asserted. Nobody "
         "proposed dropping it and the 50% MPC rule does not require dropping "
         "it.", colour=GREEN)

    h("The title has to change, and this is not cosmetic")
    para("The working title says \"dune systems\". Acıpayam's confirmed site is "
         "a lake shore. An evaluator who reads the title and then reaches Ucarı "
         "Lake will record an internal inconsistency, and they will be right. "
         "The object all three sites share is not a dune — it is a SEDIMENT "
         "SHORE, marine, lacustrine and riparian. AMMOS survives, because ammos "
         "is sand and sand is exactly what they have in common.")
    for title, verdict, why in TITLES:
        para(title, bold=True, size=10.5, after=1,
             colour=GREEN if verdict == "RECOMMENDED" else
             (RED if "do not" in verdict else NAVY))
        para(f"{verdict} — {why}", size=9.5, colour=GREY, after=6)

    h("What came from Conides's file, and what changed")
    bul([
        "Every content line in his file is placed in this structure. Nothing is "
        "discarded. Lines carrying [C] are his.",
        "ABSORBED, and better than what I had: the tourism exploitation "
        "analysis. Willingness to visit, willingness to pay, and the effects of "
        "tourism on water and energy. The database we capitalise from is the "
        "Sustainable Tourism Database and SO 2.2's topic is tourism and the "
        "green transition — a valuation survey is what ties the science to the "
        "call. I under-weighted it.",
        "ABSORBED: the specific field method — sea "
        "level rise scenarios 0.2 m to 1 m, biodiversity spot mapping with "
        "species identification, the photographic guide.",
        "ABSORBED and untouched: the Limnos geopark absorbing maintenance costs "
        "after closure. That is the strongest sustainability argument in either "
        "document.",
        "CHANGED: WP2 moves from Pi Youth to HCMR. FAQ 1.11 puts both mandatory "
        "work packages with the Lead Partner. Pi Youth keeps the delivery as "
        "named activity lead for 2.1 and 2.3.",
        "CHANGED: his dune study and his interventions happen in THREE "
        "territories, not only on Lemnos. In his file every field line sits "
        "with HCMR or the Aegean, both Greek, which would fail the 50% MPC rule "
        "structurally rather than marginally.",
        "CHANGED: Acıpayam appears. His file does not mention it once — not as "
        "a lead, not in a single content line. With exactly five organisations "
        "against a rule requiring five, a partner with no role is an "
        "eligibility risk, not a presentational one.",
    ])

    h("The realignment Acıpayam's confirmed role would justify — NOT AGREED",
      colour=RED)
    para("Acıpayam now has a real operational role: site manager at Ucarı Lake "
         "for visitor channelling, boardwalks and demarcation, the drone "
         "baseline, hosting Shore Custodians, and stakeholder adoption. At "
         "135.000 that role is underfunded relative to what it delivers.",
         size=10)
    para("These all pass every rule. The rules are not the constraint — the "
         "Lead Partner's own share is, and that is a conversation with Conides "
         "rather than a spreadsheet problem.", size=10, bold=True)

    # THIS TABLE USED TO BE FOUR ROWS OF TYPED NUMBERS AND IT WENT STALE THE
    # MOMENT THE SPLIT CHANGED.
    #
    # It said "as agreed 22 Sep, HCMR 400.000, MPC 51,15%" - all true on 22
    # September and all wrong on 28 September, when HCMR became 363.000 and
    # the MPC share started depending on 4.4.3(b). Hard-coded numbers in a
    # document generated by a script are the worst of both worlds: they look
    # computed and they are not. So every cell is derived now, and the base
    # row IS the current agreed split rather than a remembered one.
    elig_now = {k: DIRECT[k] + 0.30 * STAFF[k] for k in DIRECT}
    pool = round(elig_now["HCMR"] + elig_now["ACIP"], 2)
    mpc_fixed = round(elig_now["PIYA"] + elig_now["LCEC"]
                      + CNR_MPC_DIRECT + 0.30 * CNR_MPC_STAFF, 2)
    elke_pot = round(0.15 * STAFF["HCMR"], 2)      # CC2 is ELKE's only pot

    def gr2(x):
        return f"{x:,.0f}".replace(",", ".")

    opts = [("as agreed 30 Sep", elig_now["ACIP"]), ("A", 200_000.0),
            ("B", 230_000.0), ("C  (their request)", 270_000.0)]
    rows, elke_fail = [], []
    for label, acip in opts:
        hcmr = round(pool - acip, 2)
        mpc = round(mpc_fixed + acip, 2)
        elke = elke_pot / hcmr
        if elke < 0.065:
            elke_fail.append(label)
        rows.append([label, gr2(acip), gr2(hcmr),
                     f"{mpc / TOTAL_ELIGIBLE:.2%}".replace(".", ","),
                     f"{elke:.2%}".replace(".", ",")])
    table(["Option", "Acıpayam", "HCMR", "MPC partners, on total eligible",
           "ELKE"], rows, size=9)
    # ELKE is paid only out of CC2, which is a flat 15% of staff, so ELKE's
    # share is 0,15 x staff / that partner's own total eligible. It therefore
    # RISES as HCMR's budget falls with staff held constant - which means the
    # 28 September split, which cut HCMR, made ELKE's 6,5% easier to reach
    # rather than harder. Worth saying to HCMR explicitly.
    para(f"ELKE is paid only from cost category 2, a flat 15% of staff costs, "
         f"so its share is 0,15 x staff divided by HCMR's own total eligible: "
         f"{gr2(elke_pot)} over HCMR's budget. Because staff stays at "
         f"{gr2(STAFF['HCMR'])} while the budget falls, every option below "
         f"gives ELKE MORE than the 6,5% it asks for, not less. The "
         f"28 September reduction of HCMR to {gr2(elig_now['HCMR'])} takes "
         f"ELKE to {elke_pot / elig_now['HCMR']:.2%}.".replace(".", ",", 0),
         size=9.5, colour=GREY)
    assert not elke_fail, (
        f"these options put ELKE below its 6,5%: {elke_fail}. ELKE needs "
        f"staff to be at least 43,33% of HCMR's own eligible cost.")
    para("I would not sign option C. It puts the Lead Partner below LCEC while "
         "HCMR leads WP1, WP2, WP3 and most of WP6 content — an evaluator would "
         "read the budget against the work plan and see they disagree. A is "
         "defensible on the works; B is defensible if HCMR accepts it.", size=10)
    para("Ask Acıpayam for a bottom-up CC1 to CC6 breakdown against the actual "
         "works — boardwalk metres, signage units, fencing runs, drone and GPS, "
         "monitoring kit, staff months, audit — and give them an envelope of "
         "200.000 to 230.000. Do not negotiate against the 270.000: that figure "
         "came from my own illustrative draft and was never an allocation.",
         size=10, colour=RED)

    h("Classify the Ucarı works as CC5 equipment wherever it is honest")
    para("Acıpayam asked for CC4, infrastructure and works, for the boardwalks. "
         "CC4 carries obligations CC5 does not: the contribution must be repaid "
         "if within five years of final payment the asset's ownership changes "
         "in a way that gives undue advantage, or its nature, objectives or "
         "implementation conditions substantially change. Its eligible list is "
         "closed. Feasibility studies and environmental impact assessments must "
         "go under CC6, not CC4. Partners must guarantee building permits. And "
         "the programme runs environmental checks at step 2 of the evaluation — "
         "works in a fragile lacustrine zone will attract exactly that.",
         size=10)
    bul(["Demountable modular timber boardwalk sections, signage, fencing, "
         "demarcation posts, drone, GPS, monitoring kit  =  CC5 equipment",
         "Poured foundations, excavation, permanent built works  =  CC4 works"],
        size=10)
    para("Specifying demountable modular sections keeps the spend in CC5, "
         "avoids the five-year ownership trap, and takes the building permit off "
         "the critical path of a 24-month project. It also protects the WP6 "
         "geopark handover, which could otherwise trip the ownership condition. "
         "Acıpayam will procure whatever we classify, so this has to be decided "
         "here and not by them.", size=10, colour=GREEN)

    h("Management support, split across two partners on purpose")
    para("Activity 1.3 under HCMR is cost category 1, staff — no procurement of "
         "any kind, it is an appointment decision. The external monitoring and "
         f"data-governance contract sits under the Aegean in activity 4.3 at "
         f"EUR {SUBCONTRACT:,.0f} excluding VAT, procured by the Aegean so the "
         f"organisation awarding it is not the organisation whose payroll is "
         f"involved. The value sits above the "
         f"EUR {THREE_OFFERS_FROM:,.0f} direct-appointment line on purpose: "
         f"with a connected party bidding, a competed award is the defensible "
         f"one and a sole-source award is the weak point in the file.", size=10)
    para("Whoever bids for 4.3 must not appear on the partner or ASSOCIATED "
         "partner list — CC6 bars all three from subcontracting, and associated "
         "partners get added late for credibility. A name added in October makes "
         "the contract ineligible in November.", size=10, colour=RED)

    h("What still gates this")
    bul([
        "The Joint Secretariat answer on whether an inland lacustrine pilot site "
        "is eligible under SO 2.2. Questions close 14 October. Acıpayam asked "
        "for this themselves and they are right — it converts a judgement into "
        "a written answer we can put in front of a municipal council and an "
        "evaluator.",
        "Acıpayam's GIS coordinates, photographs and above all VISITOR NUMBERS. "
        "Without the numbers, \"acute tourism pressure\" is an assertion rather "
        "than a baseline indicator.",
        "LCEC's site, and whether it can fund its 8% co-financing — FAQ 3.19 "
        "lists national mechanisms for Greece and Türkiye and says nothing "
        "about Lebanon.",
        "Translation. MMM_IF02 is 68 pages of Italian and the whole project is "
        "the transfer of its contents to Turkish and Arabic speaking "
        "territories. 12.000 is earmarked inside HCMR's CC6; annexes 3.3 and "
        "3.4, the materials sheet and the field sheet, go first because they "
        "are what people hold in their hands.",
        "MMM_IF01 Beach Custodians has no standalone publication — the database "
        "link is a news article. The teaching material has to be requested from "
        "the AMMIRARE partnership, and WP5 is built on it. A letter of support "
        "from them would strengthen the proposal anyway.",
    ], colour=RED)

    d.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
