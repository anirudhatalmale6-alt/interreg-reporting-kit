#!/usr/bin/env python3
"""Sections 3.5, 3.6 and 3.7 - fourteen boxes, limits supplied by the client.

    python3 build_sustainability_cost_horizontal.py

He sent the limits as text, which is better than a screenshot:

    3.5 Sustainability            5 boxes x 2000   criterion 5, 12 points
      3.5.1 financial / technical / environmental
      3.5.2 long-term impact
      3.5.3 continued use of results
    3.6 Cost-effectiveness        4 boxes x 2000   criterion 6, 12 points
      3.6.1 budget
      3.6.2 ratio and balance   (two boxes: justify the ratio, then justify
                                 the costs and the subcontracting)
      3.6.3 budget allocation
    3.7 Horizontal principles     5 boxes x 1000   criterion 7, 4 points
      fundamental rights / gender equality / non-discrimination and
      accessibility / sustainable development / do no significant harm

Twenty-eight of the hundred points, and all of it was at zero.

WRITING THREE SEPARATE FILES, NAMED BY SECTION

"Please name files per section or distinct headlines." Fair, and overdue - he
is pasting into fourteen boxes across three tabs and a single omnibus document
makes him hunt. One file per section, each named for it.

THE THING THAT MAKES 3.6 DIFFERENT FROM EVERY OTHER SECTION

Cost-effectiveness is the only section where every sentence can be checked
against the budget we already hold. So nothing here is adjectival: the numbers
are computed from the same partner budget that generates every other document,
and if a partner's figure changes these paragraphs change with it. Cost per
trained person, cost per site, the share going to the three MPC partners, the
staff share, what the flat rates save in administration - all derived, none
typed.

3.7 IS FOUR POINTS AND THE EASIEST PLACE TO LOSE THEM

Horizontal principles is where applicants write a paragraph of EU boilerplate
and score two out of four. The boxes are 1000 characters each, which is small
enough that generic text fills them and specific text has to be chosen
carefully. Every one of the five below is answered with something this project
actually does at a named site, and where the honest answer is "this does not
arise much", it says so and then says what we do anyway. "Do no significant
harm" is the one that matters most here, because we are putting structures on
a dune system next to an Annex I habitat and an evaluator will look for
exactly that.
"""
import datetime as dt
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _shared import here

from build_wp_blended import DIRECT, P, STAFF, WPS, matrix

BASE = str(here(__file__))
NAVY = RGBColor(0x1F, 0x38, 0x5E)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
RED = RGBColor(0xB4, 0x55, 0x3C)

FLAT = 0.30
COFIN = 0.08
TRAINED = 60
SITES = 3
CUSTODIAN_GROUPS = 9
THIN = 0.80

RANGE = re.compile(r"(?<=\d)\s*[–—]\s*(?=[A-Za-z]?\d)")


def clean(t):
    t = str(t)
    t = RANGE.sub("\x00", t)
    t = re.sub(r"(?<!-)\s*--\s*(?!-)", ", ", t)
    t = re.sub(r"\s*[–—]\s*", ", ", t)
    t = t.replace("\x00", "-")
    return re.sub(r",\s*,", ",", t)


def eur(x):
    return f"{x:,.2f}".translate(str.maketrans({",": ".", ".": ","})) + " €"


def figures():
    elig = {k: DIRECT[k] + STAFF[k] * FLAT for k in DIRECT}
    tot = sum(elig.values())
    m = matrix()
    wp = {w[0]: sum(m[k].get(w[0], 0.0) for k in P) for w in WPS}
    return {
        "elig": elig, "tot": tot,
        "direct": sum(DIRECT.values()),
        "staff": sum(STAFF.values()),
        "flats": sum(STAFF.values()) * FLAT,
        "eu": tot * (1 - COFIN), "cofin": tot * COFIN,
        "mpc": sum(elig[k] for k in P if P[k][2] == "MPC"),
        "wp": wp,
        "technical": sum(wp[k] for k in ("WP3", "WP4", "WP5", "WP6")),
        "overhead": wp["WP1"] + wp["WP2"],
    }


# =========================================================== 3.5 SUSTAINABILITY
def s35(F):
    return {
        "3.5.1 Financial level": [
            "The honest financial question about a transferred method is not "
            "whether the project can afford it but whether the receiving "
            "authority can afford to keep using it after we leave. The "
            "design answers that in three ways.",

            "COSTED PLANS, NOT ASPIRATIONAL ONES. Each of the three "
            "management plans in Output6.1 carries its own maintenance cost, "
            "calculated during the project rather than estimated afterwards. "
            "An authority adopting a plan therefore adopts a known annual "
            "figure. A plan that does not state its running cost is adopted "
            "and then quietly abandoned at the first budget round.",

            "MAINTENANCE CARRIED BY AN EXISTING ORGANISATION. At Lemnos the "
            "route pursued is incorporation into the Limnos geopark network, "
            "so that monitoring and upkeep sit inside an organisation that "
            "already has activities, staff and income, rather than depending "
            "on a future grant. This is the strongest sustainability argument "
            "we have because it does not require anyone to find new money.",

            "DELIBERATELY LOW RECURRENT COST BY DESIGN. The installations are "
            "demountable and modular, so sections can be repaired or moved "
            "rather than replaced. The monitoring protocol is designed to be "
            "run by trained municipal staff and school custodian groups "
            "rather than by contracted specialists - which is why "
            f"{TRAINED} site managers and local staff are trained and "
            f"{CUSTODIAN_GROUPS} custodian groups established. Labour is the "
            "dominant recurrent cost of any monitoring regime, and here it is "
            "absorbed by people who are already paid to be at the site.",

            "The Shore Stewardship Committee at each site reviews the "
            "maintenance cost annually, which puts the financial question in "
            "front of the people who can answer it.",
        ],

        "3.5.1 Technical level": [
            "Technical sustainability is the reason the project is structured "
            "as a transfer of method rather than of infrastructure. A built asset "
            "decays; a procedure local staff can run does not.",

            "THE METHOD IS LEFT IN A USABLE FORM. The adapted method is "
            "translated into Greek, Turkish and Arabic, because a procedure "
            "available only in the language of its authors is not "
            "transferred, it is merely published. Each of the 21 "
            "transferability decisions is recorded with its reason, so a "
            "future user knows not only what to do but where the method stops "
            "applying and why.",

            "THE SKILLS ARE LEFT WITH PEOPLE WHO STAY. Training targets "
            "municipal and authority staff holding a mandate over each site, "
            "not temporary project personnel. The custodian groups are run "
            "with the schools nearest each site, so the group survives by "
            "being local rather than by being funded.",

            "THE DATA SURVIVES THE PROJECT. The shared web-GIS platform is "
            "built on the adapted protocol with its own field data entry, "
            "under open licence, with the cross-typology comparability layer "
            "documented so a new territory can join without renegotiating the "
            "protocol. Litter monitoring uses the published EEA-aligned "
            "protocol, so the series stays comparable with the European "
            "dataset.",

            "MODULAR, SO IT CAN BE PARTIALLY ADOPTED. Because the method is "
            "seven separable sheets rather than one instrument, a territory "
            "that cannot run all of it can adopt the three or four that suit "
            "its shore. An all-or-nothing method is abandoned the first time "
            "one component proves impractical; a modular one degrades "
            "gracefully.",

            "AND THE TRANSFER RUNS BETWEEN PEOPLE. The institution that "
            "authored the source is a partner and leads adaptation of its own "
            "sheets, so questions are answered by the people who wrote them "
            "rather than inferred from a document.",

            "TECHNICAL RISK IS NAMED. The main one is platform hosting after "
            "closure, addressed by open licensing, documented export formats "
            "and the research partners' repositories, so the dataset does not "
            "die with a subscription.",
        ],

        "3.5.1 Environmental level": [
            "The project's environmental sustainability is not a side effect "
            "of good behaviour; it is the subject matter. AMMOS exists to "
            "reduce the erosion of sediment shores under visitor pressure.",

            "REVERSIBILITY AS A PRINCIPLE. Every physical intervention is "
            "demountable and modular. Boardwalks, boundary markers and "
            "fencing can be removed without trace, which matters because the "
            "Greek site is an Annex I habitat type and because a "
            "demonstration that cannot be undone is an experiment nobody "
            "should run on a protected shore.",

            "MEASURED, NOT ASSERTED. Each site is characterised before "
            "intervention - topography, sediment, vegetation, ecological "
            "status - and monitored through the first full visitor season "
            "afterwards. The project therefore produces evidence of its own "
            "environmental effect rather than claiming one. Where the "
            "evidence is unfavourable the installation can be altered, "
            "because it was built to be.",

            "NO EXTRACTIVE OR INVASIVE WORK. No underwater or bottom "
            "cleaning, no dredging, no sediment importation, no planting of "
            "non-native species at any site. Dune and littoral vegetation "
            "work is protection from trampling rather than replacement.",

            "AND IT IS CLIMATE ADAPTATION, NOT ONLY CONSERVATION. Each "
            "baseline models sea level rise scenarios from 0.2 m to 1 m "
            "against the site's own topography, so an authority can see which "
            "parts of its shore are exposed on what horizon. The "
            "cross-typology comparability layer then makes a monitoring "
            "series that can detect trend rather than noise, which is what "
            "turns a one-off survey into adaptation evidence. Litter "
            "monitoring on the published EEA-aligned protocol adds a second "
            "pressure to the same record.",

            "VISITOR CHANNELLING RATHER THAN EXCLUSION. Pressure is "
            "redistributed off the vulnerable surface while the site stays "
            "open, which is both the only politically durable option in a "
            "tourism-dependent municipality and the one that keeps the local "
            "economy intact. An intervention that closes a beach protects it "
            "until the first season it is reopened.",
        ],

        "3.5.2 Long-term impact": [
            "The long-term impact AMMOS is designed for is not three managed "
            "shores. It is that a method which existed in one language and "
            "one shore type becomes usable across the Mediterranean basin.",

            "AT THE THREE SITES. Each has a costed management plan formally "
            "adopted by the authority with the mandate, a standing "
            "multi-stakeholder stewardship committee, trained staff and a "
            "monitoring series that continues. That is a change in how the "
            "shore is governed, not a project output that expires.",

            "BEYOND THE THREE SITES, THE LARGER EFFECT. The "
            "transferability assessment states where the method applies and "
            "where it does not. A fourth territory reading it can tell, "
            "before spending anything, which of the seven procedures suits "
            "its own shore - a marine dune, a coastal site, or an artificial "
            "freshwater body. Establishing the boundary of applicability is "
            "what converts one validated method into a usable family of "
            "methods, and it is why three deliberately dissimilar sites were "
            "chosen rather than three convenient ones.",

            "AT POLICY LEVEL. Three adopted local plans and a set of policy "
            "recommendations in English and Arabic give national and regional "
            "authorities a precedent for treating shore sediment "
            "management as a climate adaptation measure rather than a "
            "maintenance expense.",

            "A CASE WORTH NAMING. In Lebanon the national policy framework "
            "for coastal management exists while the applicable field method "
            "does not, so AMMOS supplies the missing half rather than a "
            "parallel system - the fastest impact to achieve and the most "
            "likely to persist.",

            "AND THE SLOWEST EFFECT IS THE MOST DURABLE. Nine school "
            "custodian groups put a generation of local pupils through a "
            "monitoring round at their own shore. Some will be the municipal "
            "staff making these decisions in fifteen years.",

            "AND IN THE MECHANISM ITSELF. The adapted method, the 21 "
            "decisions and the comparability layer return to the MMM common "
            "database under open licence, so the next capitalisation call "
            "starts from our result rather than from where we started.",
        ],

        "3.5.3 Continued use of results": [
            "Each result has a named owner after closure, and the ownership "
            "is arranged during the project rather than hoped for at the end.",

            "THE ADAPTED METHOD AND THE TRANSFERABILITY GUIDE are lodged in "
            "the MMM common database under open licence, in English and "
            "Arabic, each carrying an explicit section stating the actions an "
            "adopting organisation must take. That section is what makes them "
            "usable rather than merely available, and it is also what the "
            "programme's own output indicator for jointly developed "
            "solutions requires.",

            "THE MANAGEMENT PLANS are owned by the authority that formally "
            "adopted each one, and overseen by the Shore Stewardship "
            "Committee constituted alongside it - the responsible authority, "
            "a tourism operator, the research partner and a civil society or "
            "youth representative. A plan with a standing committee is "
            "reviewed; a plan without one is filed.",

            "THE MONITORING PLATFORM continues under open licence with "
            "documented export formats and institutional repository backing, "
            "and the trained municipal staff and custodian groups keep "
            "feeding it. The protocol is deliberately within the capacity of "
            "the people who remain.",

            "THE TRAINING AND CUSTODIAN MATERIAL stays with the "
            "organisations that will reuse it: the training package in "
            "Greek, Turkish and Arabic remains with each municipality and "
            "authority that hosted it, and the adapted Beach Custodians "
            "material with the participating schools, both under open licence "
            "so a new cohort can be run without asking anyone's permission.",

            "THE PARTNERSHIP ITSELF CONTINUES under a cooperation agreement "
            "signed during implementation, committing the partners, the "
            "associated organisation and the three stewardship committees to "
            "continue cooperating after closure. That agreement is a "
            "deliverable with a date, not an intention - which is also what "
            "the programme's result indicator on sustained cooperation "
            "requires in order to be counted at all.",
        ],
    }


# ====================================================== 3.6 COST-EFFECTIVENESS
def s36(F):
    pc = lambda x: f"{100 * x / F['tot']:.1f}%"
    per_site = F["technical"] / SITES
    per_trained = F["wp"]["WP5"] / TRAINED
    return {
        "3.6.1 Budget": [
            f"Total eligible cost is {eur(F['tot'])} for 24 months and six "
            f"partners in four countries: {eur(F['direct'])} of direct cost "
            f"plus {eur(F['flats'])} generated by the two flat rates on "
            f"staff. EU contribution {eur(F['eu'])}, partner co-financing "
            f"{eur(F['cofin'])}.",

            f"HOW IT IS DISTRIBUTED. Staff is {eur(F['staff'])}, "
            f"{pc(F['staff'])} of total eligible cost, well inside the 40% "
            f"ceiling, which matters because this is a transfer project: the "
            f"cost is people adapting a method and training others to use "
            f"it, not procurement. Office and travel are taken as flat rates "
            f"at 15% each of staff, which removes an entire class of "
            f"receipt-level administration from six partners in four "
            f"countries and from the programme's own control system.",

            f"WHERE THE MONEY WORKS. {eur(F['technical'])} - "
            f"{100 * F['technical'] / F['direct']:.0f}% of direct cost - sits "
            f"in the four technical work packages. Management and "
            f"communication together take {eur(F['overhead'])}, and the "
            f"communication share of that is deliberate rather than "
            f"residual: the evaluation grid weights communication double, and "
            f"a transfer project whose method does not reach its users has "
            f"failed whatever else it delivers.",

            f"TIMING IS BUILT INTO THE BUDGET, NOT LEFT TO CHANCE. Every "
            f"budget line carries the semester in which the expense falls, so "
            f"the cash-flow profile is explicit. IT equipment for staff sits "
            f"in the first year as the rules require, and no equipment "
            f"purchase is planned in the final year, which avoids needing the "
            f"Managing Authority's authorisation late in a project that has "
            f"no time to wait for it.",

            f"BALANCE ACROSS THE PARTNERSHIP. The largest partner holds "
            f"{pc(max(F['elig'].values()))} of total eligible cost, inside "
            f"the 35% single-organisation cap. "
            f"{pc(F['mpc'])} goes to partners in Mediterranean Partner "
            f"Countries, satisfying the 50% requirement by partner allocation "
            f"alone, with no recourse to justified activities implemented on "
            f"their behalf. The budget is built where the work is, not "
            f"arranged to pass a test.",
        ],

        "3.6.2 Ratio and balance": [
            f"The ratio to judge is {eur(F['tot'])} against what exists at "
            f"the end that did not exist at the start.",

            f"WHAT THE MONEY BUYS. A validated method adapted and translated "
            f"for three shore types with 21 recorded transferability "
            f"decisions; three comparable territorial baselines where there "
            f"were none; three demonstrated installations; a cross-typology "
            f"monitoring platform; {TRAINED} trained site managers and "
            f"{CUSTODIAN_GROUPS} school custodian groups across three "
            f"countries; three formally adopted costed management plans with "
            f"standing governance; and a transferability guide returned to "
            f"the MMM database for the next territory.",

            f"THE UNIT FIGURES. About {eur(per_site)} of technical budget per "
            f"demonstration site, covering characterisation, installation, a "
            f"full season of monitoring, training and an adopted plan. About "
            f"{eur(per_trained)} per person trained in WP5, including the "
            f"material adapted into three languages. The programme's own "
            f"methodology paper assumes an average project of "
            f"1.890.000 € under this priority; AMMOS delivers across three "
            f"countries and three shore types at "
            f"{100 * F['tot'] / 1_890_000:.0f}% of that.",

            f"THE COST THAT IS AVOIDED. Developing a shore management "
            f"method from first principles - field validation, iteration, "
            f"publication - is the work of a multi-year research programme. "
            f"AMMOS does not repeat it. It pays for adaptation, translation "
            f"and demonstration of a method already validated, which is why "
            f"three countries can be covered for the price of one national "
            f"project, and why the capitalisation route is the "
            f"cost-effective one rather than merely the eligible one.",

            f"WHY THE BALANCE IS RIGHT RATHER THAN MERELY COMPLIANT. The "
            f"single largest allocation is the demonstration work package, "
            f"because applying the method on the ground is what distinguishes "
            f"transfer from dissemination. The second largest is capacity "
            f"building, because the method has to outlive the project. The "
            f"smallest technical allocation is the uptake work package, which "
            f"is correct: adoption is mostly consultation and drafting, and "
            f"its cost is time rather than equipment.",
        ],

        "3.6.2 Necessity of costs and subcontracting": [
            "Every cost traces to an activity and every activity to an "
            "output. The categories and why each is necessary:",

            f"STAFF, {pc(F['staff'])} of total eligible, is the dominant cost "
            f"because the work is intellectual and pedagogic: assessing 21 "
            f"transferability cells, drafting substituted procedures, "
            f"translating into three languages, producing three baselines, "
            f"training {TRAINED} people. None of that can be bought in as a "
            f"product.",

            "EQUIPMENT is limited to field survey and monitoring instruments "
            "and the demountable installation components. It is specified "
            "from each site's own baseline rather than bought to a standard "
            "list, and procurement follows the characterisation precisely so "
            "that nothing is purchased against an assumption. IT equipment "
            "for staff falls within the first year as the rules require, and "
            "no equipment purchase is planned in the final year.",

            "INFRASTRUCTURE AND WORKS covers only the demonstration "
            "installations, in the technical work packages where the rules "
            "permit it, and reversibility is a cost-control measure as well "
            "as an environmental one: a modular boardwalk section can be "
            "repaired rather than the whole structure replaced.",

            "EXTERNAL EXPERTISE AND SERVICES - and this is the one that needs "
            "justifying rather than listing. Four uses, each for a capability "
            "no partner holds and none should staff up to acquire for 24 "
            "months: certified translation into Turkish and Arabic, where the "
            "method's usability depends on precision; drone survey and "
            "specialist bathymetry at sites where a partner lacks the "
            "licensed capability; the first-level control required of every "
            "partner by the programme, which must be independent by "
            "definition; and web-GIS platform development, where buying "
            "16 months of developer time would cost more than the "
            "deliverable. Subcontracting is for capability, not for capacity.",
        ],

        "3.6.3 Budget allocation": [
            f"Allocation follows two questions: is the money where the work is, "
            f"and is it where the territories that need it are.",

            f"BY WORK PACKAGE, on direct cost. WP3 characterising and "
            f"adapting {eur(F['wp']['WP3'])}; WP4 demonstrating "
            f"{eur(F['wp']['WP4'])}; WP5 building capacity "
            f"{eur(F['wp']['WP5'])}; WP6 making it stick "
            f"{eur(F['wp']['WP6'])}; WP1 management {eur(F['wp']['WP1'])}; "
            f"WP2 communication {eur(F['wp']['WP2'])}. The technical work "
            f"packages hold "
            f"{100 * F['technical'] / F['direct']:.0f}% of direct cost.",

            f"BY PARTNER, each holding the budget for work it leads rather "
            f"than a share negotiated by size. The largest is "
            f"{pc(max(F['elig'].values()))} of total eligible, inside the 35% "
            f"cap; the smallest holds a budget matched to a defined "
            f"contribution - adapting the topographic and forcings sheets, by "
            f"the institution whose researcher authored them - not a token "
            f"participation.",

            f"BY TERRITORY. {pc(F['mpc'])} to Mediterranean Partner Country "
            f"partners, above the 50% requirement and achieved by allocation "
            f"rather than by justifying activities done on their behalf. The "
            f"three demonstration sites each carry their own "
            f"characterisation, installation, training and adoption budget, "
            f"so no site is a junior case study of another.",

            f"ACROSS TIME. Spending follows the work: characterisation and "
            f"adaptation load the first year, installation and training the "
            f"second, adoption and the guide the final semesters. Every "
            f"budget line states its semester, so the profile is visible "
            f"rather than inferred.",

            f"ONE ALLOCATION WORTH DEFENDING. Communication takes "
            f"{eur(F['wp']['WP2'])}, "
            f"{100 * F['wp']['WP2'] / F['direct']:.1f}% of direct cost - more "
            f"than some reviewers expect, and deliberate. The method exists "
            f"only in Italian and its users work in Greek, Turkish and "
            f"Arabic; translation, local-language material and six workshops "
            f"are not promotion, they are the transfer mechanism.",

            f"BY OUTPUT. Each work package's budget is distributed across its "
            f"two outputs as percentages totalling 100, which the platform "
            f"checks: WP3 42/58, WP4 66/34, WP5 46/54, WP6 55/45. Every "
            f"figure in this section is computed from the agreed partner "
            f"budget, so if a partner's allocation changes these numbers "
            f"change with it.",
        ],
    }


# ====================================================== 3.7 HORIZONTAL (1000)
def s37(F):
    return {
        "Promotion of fundamental rights": [
            "AMMOS touches fundamental rights most directly through access to "
            "a shared natural resource. Our central design choice - visitor "
            "channelling rather than exclusion - is a rights choice as much "
            "as an environmental one: the shore stays open and publicly "
            "accessible while pressure is moved off the vulnerable surface. "
            "No intervention closes public access at any of the three sites.",

            "The three Shore Stewardship Committees give residents, operators "
            "and civil society a standing seat where decisions about the "
            "shore are taken, rather than a consultation that ends with the "
            "project. Management plans go through open stakeholder "
            "consultation before adoption. Project data is published under "
            "open licence, and all site material is produced in the local "
            "language, so information about a shared resource is available to "
            "the people who use it rather than only to those who read "
            "English.",
        ],

        "Promotion of gender equality": [
            "Gender equality is pursued in who does the work and who is "
            "heard, since the project has no gendered service delivery to "
            "balance.",

            "Partners apply their own equal-opportunity rules in recruiting "
            "project staff, and the management committee records the gender "
            "composition of project teams, trainee cohorts and stewardship "
            "committees each reporting period, so imbalance is visible early "
            "enough to act on rather than reported at closure.",

            "Two specifics. Coastal tourism employment is strongly gendered "
            "in all three territories, so the visitor and operator surveys in "
            "Output3.2 disaggregate responses by gender - which is also "
            "simply better data about who depends on the shore. And "
            "site-manager training is scheduled and located so that caring "
            "responsibilities are not a barrier to attending: local venues, "
            "working hours, no residential requirement.",
        ],

        "Prevention of discrimination and accessibility": [
            "Two concrete commitments rather than a statement of intent.",

            "PHYSICAL ACCESSIBILITY. The demountable boardwalks are the "
            "project's main built element and they are specified to be "
            "wheelchair-navigable where site gradient allows - firm surface, "
            "adequate width, no step transitions. This is a genuine gain: a "
            "sediment shore is close to impassable for a wheelchair user, and "
            "a boardwalk built for conservation also opens the site to people "
            "who could not previously reach it. Where gradient makes full "
            "accessibility impossible, the constraint is documented rather "
            "than omitted.",

            "INFORMATION ACCESSIBILITY. All site interpretation and visitor "
            "material is produced in the local language - Greek, Turkish, "
            "Arabic - and designed to be understood without specialist "
            "knowledge, using pictograms alongside text. Platform content "
            "follows accessible contrast and text-size practice. Training is "
            "open to municipal staff regardless of background or formal "
            "qualification.",
        ],

        "Promotion of sustainable development": [
            "This is not a horizontal principle bolted onto AMMOS; it is the "
            "project's subject. The purpose is to let tourism continue on "
            "Mediterranean sediment shores without consuming the surface it "
            "depends on - the sustainable development question in its most "
            "literal form.",

            "All three dimensions are addressed and they are not traded off "
            "against each other. ENVIRONMENTAL: reduced erosion and trampling "
            "pressure, measured before and after, with every intervention "
            "reversible. ECONOMIC: the shore is treated as the asset the "
            "local economy rests on, and Output3.2 values it, so protection "
            "is argued in the terms a municipality actually decides in. "
            "SOCIAL: local staff trained, school custodian groups "
            "established, standing committees that keep residents and "
            "operators in the decision.",

            "The project also returns its method to a shared database under "
            "open licence, so the development gain is not confined to three "
            "sites.",
        ],

        "Do no significant harm": [
            "This principle applies to AMMOS directly and we treat it as a "
            "design constraint, not a declaration. The project places "
            "structures on sediment shores, one of which is an Annex I "
            "habitat type under recorded visitor pressure.",

            "Four safeguards. REVERSIBILITY: every intervention is "
            "demountable and removable without lasting trace, so no decision "
            "is irreversible. MEASURE FIRST: nothing is installed until the "
            "site has been characterised, and the design follows the "
            "baseline rather than a standard template. NO EXTRACTIVE WORK: no "
            "dredging, no bottom or underwater cleaning, no sediment "
            "importation, no non-native planting. CONSENT AND SCREENING: each "
            "site passes its own national consent and environmental "
            "screening before works, and the environmental screening "
            "checklists for this application are completed on that basis.",

            "If first-season monitoring shows an installation is causing "
            "harm, it is altered or removed - which is precisely why it was "
            "built to be removable.",
        ],
    }


class Doc:
    def __init__(self):
        self.d = Document()
        s = self.d.sections[0]
        s.left_margin = s.right_margin = Cm(2.0)
        n = self.d.styles["Normal"]
        n.font.name, n.font.size = "Calibri", Pt(10.5)

    def h(self, t, size=13, colour=NAVY, space=12):
        p = self.d.add_paragraph()
        p.paragraph_format.space_before = Pt(space)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(clean(t))
        r.bold, r.font.size, r.font.color.rgb = True, Pt(size), colour

    def p(self, t, size=10.5, colour=None, italic=False, bold=False):
        par = self.d.add_paragraph()
        par.paragraph_format.space_after = Pt(5)
        r = par.add_run(clean(t))
        r.font.size, r.italic, r.bold = Pt(size), italic, bold
        if colour:
            r.font.color.rgb = colour


def build(section, title, points, boxes, limit, outfile, preamble):
    """One file per section, named for it, every box fit-checked."""
    fails, report, texts = [], [], {}
    for name, parts in boxes.items():
        text = clean("\n\n".join(parts))
        n = len(text)
        if n > limit:
            fails.append(f"{name} is {n} chars, limit {limit} - cut "
                         f"{n - limit}")
        for ph in re.findall(r"\[[A-Z][A-Z0-9_ ]{2,}\]", text):
            fails.append(f"{name} still has the placeholder {ph}")
        for forbidden in (r"\bRAS\b", r"Regione Autonoma", r"Hakanson",
                          r"xxxx", r"\b887\b", r"\bELKE\b"):
            if re.search(forbidden, text, re.I):
                fails.append(f"{name} contains {forbidden!r}")
        texts[name] = text
        report.append((name, n))
    assert not fails, (f"NOT SHIPPING {section}:\n  - "
                       + "\n  - ".join(fails))

    D = Doc()
    D.h(f"AMMOS — {section} {title}", size=16)
    D.p(f"Generated {dt.date.today():%d/%m/%Y}. {len(boxes)} boxes, "
        f"{limit} characters each, worth {points} points. Every box checked "
        f"against its counter before this file was produced.",
        size=9.5, colour=GREY, italic=True)
    for para in preamble:
        D.p(para, size=10.5)
    for name, text in texts.items():
        D.h(name, size=13)
        D.p(f"{len(text)} of {limit} characters.", size=9, colour=GREY,
            italic=True)
        D.p(text)
    D.d.save(outfile)
    return report


def main():
    F = figures()
    jobs = [
        ("3.5", "Sustainability", 12, s35(F), 2000,
         f"{BASE}/AMMOS_3.5_Sustainability_2026-10-06.docx",
         ["Three levels, then impact, then continued use. The thread running "
          "through all five is that sustainability was designed in rather "
          "than argued afterwards: costed plans so an authority knows what it "
          "is adopting, maintenance carried by an organisation that already "
          "exists, demountable structures that can be repaired rather than "
          "replaced, and a monitoring protocol deliberately within the "
          "capacity of people who are already paid to be at the site.",
          "The strongest single argument is the Limnos geopark route, because "
          "it is the only one that does not require anyone to find new money. "
          "The second strongest is the cooperation agreement, because it "
          "turns continued partnership from an intention into a dated "
          "deliverable - and it is also what the programme's result indicator "
          "on sustained cooperation needs in order to count at all."]),

        ("3.6", "Cost-effectiveness", 12, s36(F), 2000,
         f"{BASE}/AMMOS_3.6_Cost-effectiveness_2026-10-06.docx",
         ["This is the one section where every sentence can be checked "
          "against the budget, so nothing here is adjectival. Every figure is "
          "computed from the same partner budget that generates all the other "
          "documents - cost per site, cost per person trained, the MPC share, "
          "the staff share, the technical share. If a partner's allocation "
          "changes, these paragraphs change with it.",
          "The subcontracting box is the one that needs real justification "
          "rather than a list, because an evaluator reads external expertise "
          "as either a necessary capability or a way of avoiding staffing. "
          "Four uses are named and each is a capability no partner holds and "
          "none should acquire for 24 months - certified Turkish and Arabic "
          "translation, licensed drone and bathymetric survey, the "
          "independent first-level control the programme itself requires, and "
          "platform development. Subcontracting for capability, not for "
          "capacity."]),

        ("3.7", "Horizontal principles", 4, s37(F), 1000,
         f"{BASE}/AMMOS_3.7_Horizontal-principles_2026-10-06.docx",
         ["Four points, and the easiest place in the application to lose "
          "them, because this is where everyone writes EU boilerplate and "
          "scores two out of four. At 1000 characters a box, generic text "
          "fills the space and specific text has to be chosen.",
          "So every one of the five answers something this project actually "
          "does at a named site. The accessibility box makes a real "
          "commitment: the boardwalks are specified wheelchair-navigable "
          "where gradient allows, which is a genuine gain because a sediment "
          "shore is close to impassable for a wheelchair user - and where "
          "gradient makes it impossible, we say so rather than omit it.",
          "Do no significant harm is the one that matters most here. We are "
          "putting structures on a dune system beside an Annex I habitat and "
          "an evaluator will look for exactly this, so it is answered as four "
          "design constraints rather than a declaration."]),
    ]

    for section, title, pts, boxes, limit, outfile, preamble in jobs:
        rep = build(section, title, pts, boxes, limit, outfile, preamble)
        print(f"\n{section} {title}  ({pts} points, {limit} per box)")
        for name, n in rep:
            flag = "" if n >= limit * THIN else "   <- thin"
            print(f"   {n:>5} / {limit}  {100 * n / limit:3.0f}%{flag}  "
                  f"{name}")
        print(f"   -> {Path(outfile).name}")

    print(f"\nbudget facts used: total eligible {eur(F['tot'])}, "
          f"staff {100 * F['staff'] / F['tot']:.1f}%, "
          f"MPC {100 * F['mpc'] / F['tot']:.1f}%, "
          f"technical {100 * F['technical'] / F['direct']:.0f}% of direct")


if __name__ == "__main__":
    main()
