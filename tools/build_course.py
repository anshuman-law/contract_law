from __future__ import annotations

import html
import json
import re
import shutil
import sys
import urllib.parse
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SOURCE_INDEX = Path("/Users/anshuman/Downloads/index (1).html")
SOURCE_MODULE_ONE = Path("/Users/anshuman/Downloads/contractos-promise-machine.html")
SOURCE_CURRICULUM = Path(
    "/Users/anshuman/Desktop/LM Academy/Gamified Contract Law/"
    "Gamified Introduction to Indian Contract Law - Master Curriculum.docx"
)

WORLDS = {
    1: "Why Contracts Bind",
    2: "The Formation Game",
    3: "The Price of a Promise",
    4: "Who Can Contract?",
    5: "Was the Choice Really Free?",
    6: "Agreements the Law Will Not Enforce",
    7: "What Exactly Did We Agree To?",
    8: "Performing and Escaping the Bargain",
    9: "When the Deal Breaks",
    10: "Benefits Beyond the Bargain",
}

WORLD_RANGES = [(1, 5), (6, 16), (17, 23), (24, 27), (28, 37), (38, 45), (46, 51), (52, 58), (59, 64), (65, 69)]


def world_for(module_id: int) -> int:
    for index, (start, end) in enumerate(WORLD_RANGES, 1):
        if start <= module_id <= end:
            return index
    raise ValueError(module_id)


# Each fact pattern is deliberately short. The game asks the learner to identify
# the legally decisive fact before choosing a response and replaying the problem.
FACTS = {
    2: ("A food-delivery platform requires every home chef to accept a clause allowing a large unilateral fine and a complete ban on court proceedings.", "whether the bargain is lawful and whether a consumer term is materially one-sided", "The clause is separately negotiated between two well-advised businesses and preserves arbitration."),
    3: ("A sister promises to pay her brother's rent during his studies. The messages record monthly amounts, dates, repayments and repeated reliance.", "the objective context showing whether legal consequences were intended", "The promise is a casual assurance made at a family dinner, with no reliance or recorded terms."),
    4: ("A valid import contract is followed by a new legal prohibition that makes the promised shipment unlawful.", "when the defect arose and whether enforceability existed before it", "The shipment was prohibited on the day the parties first agreed."),
    5: ("A municipal supplier receives an email award from an officer, but no contract is executed in the constitutionally required manner by an authorised person.", "compliance with Article 299 form and authority requirements", "The written contract is executed in the name of the proper government and signed by an authorised officer."),
    6: ("A Bengaluru electronics shop places a laptop in its window with a price tag and stock number. A customer says the display itself accepted her money.", "whether the display shows final willingness to be bound or invites customers to make offers", "The advertisement promises one identified laptop to the first person who pays online before noon."),
    7: ("A company publicly offers a reward to anyone who returns a lost prototype after following stated verification steps.", "a definite public proposal accepted by performance of its conditions", "The company communicates withdrawal before the claimant begins performance."),
    8: ("A driver finds and returns a missing bag, then learns that its owner had advertised a reward.", "knowledge of the proposal while performing the requested act", "The driver learns of the reward before locating the owner and completes the return in response to it."),
    9: ("A buyer replies, 'I accept if you reduce the price by ten per cent.'", "whether assent is absolute and unqualified or introduces a new condition", "The buyer asks, 'Would you consider a lower price?' while separately accepting the original offer without qualification."),
    10: ("A passenger boards a city bus, remains aboard after seeing the fare chart and accepts the ticket.", "conduct objectively signifying assent or performance of the proposal's conditions", "The passenger boards after the conductor clearly says the service is cancelled and no ticket will be issued."),
    11: ("A seller posts an offer, the buyer posts acceptance, and later messages seek to withdraw each communication.", "the separate statutory completion points against proposer and acceptor", "The parties negotiate by a live phone call instead of post."),
    12: ("An offer requires acceptance by courier. The offeree sends a clear email acceptance and the offeror continues performance without promptly objecting.", "the prescribed mode and whether the proposer insists on it within a reasonable time", "The offeree does nothing and relies only on the offeror's statement that silence will count."),
    13: ("An acceptance letter and a revocation message travel in opposite directions before the stated deadline.", "when each communication becomes complete under sections 4 and 5", "A phone revocation reaches the offeree before the acceptance is dispatched."),
    14: ("Two parties orally agree to sell an identified apartment and assume section 10 makes every contract informal.", "whether another law requires writing, signature, stamping or registration for this transaction", "The transaction is an ordinary service contract concluded by authenticated electronic messages."),
    15: ("A tender offer states a closing date, depends on a licence and remains unanswered while the market changes.", "expiry of time, failure of a condition, rejection, or death or insanity known to the offeree", "No period is stated, the condition is met and acceptance follows within a commercially reasonable time."),
    16: ("Signed heads of terms say 'subject to contract' and leave the final price mechanism unresolved.", "objective finality and certainty of essential matters", "A later signed version removes the qualification and supplies an objective market-price mechanism."),
    17: ("A manufacturer promises a rebate if a retailer keeps a display in place for three months at the manufacturer's request.", "an act, abstinence or promise at the promisor's desire", "The retailer installed the display voluntarily before any request or promise."),
    18: ("A grandmother transfers property to her daughter on the daughter's promise that the granddaughter will pay an annuity.", "consideration may move from the promisee or another person, but must be at the promisor's desire", "A stranger pays money without any request from the promisor."),
    19: ("After stock is saved from a flood, the owner promises a bonus to the worker who acted before any request was made.", "whether the earlier act was requested by the promisor rather than merely voluntary", "The owner requested the emergency work first and promised payment after it was completed."),
    20: ("A friend promises a gift next month without receiving any act, abstinence or promise in return.", "consideration or every element of a recognised section 25 exception", "Close relatives record and register a promise made on account of natural love and affection."),
    21: ("A debtor orally promises to pay a debt already barred by limitation.", "the writing and signature required for the time-barred debt exception", "The debtor signs a written promise identifying the time-barred debt and the amount to be paid."),
    22: ("An owner in financial distress sells land for far below market value after advice and full disclosure.", "inadequacy is evidence about free consent, not an automatic rule of invalidity", "The low price is accompanied by domination of will and an unfair advantage taken by a trusted adviser."),
    23: ("A creditor agrees to accept Rs 70,000 in full satisfaction of a Rs 100,000 debt.", "the promisee's statutory power to remit or accept another satisfaction", "The parties substitute a new debtor and a new obligation for the original contract."),
    24: ("A seventeen-year-old signs a high-interest personal loan and misstates age on the form.", "majority and competency at the time of contracting", "A guardian arranges necessary medical treatment for the minor from the minor's property."),
    25: ("After turning eighteen, a borrower says she ratifies the same loan agreement signed while she was a minor.", "a void minor's agreement is not revived merely by later ratification", "After majority, the parties make a fresh agreement supported by fresh consideration."),
    26: ("During a severe manic episode, a buyer cannot understand the transaction or form a rational judgment about its effect.", "functional capacity at the precise time of contracting", "The same buyer contracts during a medically supported lucid interval and understands the transaction."),
    27: ("A luxury watch is supplied on credit to a minor who already owns several watches.", "whether the goods are necessaries suited to the person's condition in life and actual requirements", "Emergency medical treatment is supplied when it is genuinely needed."),
    28: ("Buyer and seller use the same ship name, but each has a different ship in mind.", "agreement upon the same thing in the same sense", "Both communications identify the vessel by a unique registration number."),
    29: ("A buyer contracts because of an innocent false statement about a machine's output.", "the free-consent defect and whether it caused consent", "The buyer proves she knew the true output and would have contracted on the same terms."),
    30: ("A warehouse refuses to release goods unless their owner signs an unrelated settlement.", "an unlawful threat or unlawful detention of property used to cause agreement", "The warehouse exercises a valid contractual lien limited to unpaid storage charges."),
    31: ("A trusted doctor pressures a distressed patient to sell valuable land cheaply to the doctor.", "domination of will combined with an unfair advantage", "The patient receives independent advice, negotiates freely and obtains a fair price."),
    32: ("A spiritual adviser receives a grossly one-sided transfer from a dependent follower.", "a dominant relationship plus an unconscionable appearance that may shift the evidential burden", "The donor is independent, advised and enters an ordinary transaction at market value."),
    33: ("A builder promises a completion date while internal records show the builder never intended or was able to meet it.", "dishonest intention at formation rather than a later failure alone", "The promise was honest when made, but an unforeseen lawful delay later causes breach."),
    34: ("A seller gives a technically true answer that omits a fact needed to prevent the answer from being misleading.", "a duty to speak, a half-truth, or circumstances in which silence is equivalent to speech", "The seller remains silent about an ordinary market risk in an arm's-length transaction with no duty to disclose."),
    35: ("A dealer states the wrong mileage while honestly relying on a forged service record.", "knowledge, belief and duty that distinguish fraud from misrepresentation", "Messages show the dealer knew the odometer had been altered."),
    36: ("Both parties contract for cargo that, unknown to both, had already been destroyed.", "a shared mistake about an existing fact essential to the agreement", "Both parties are correct about the cargo's existence but mistaken only about its quality."),
    37: ("Only the buyer miscalculates the profitability of an otherwise clear purchase.", "the general rule for unilateral factual mistake and the narrow identity or document exceptions", "The error concerns the very nature of a document signed through a fundamental misdescription."),
    38: ("A consultant is promised payment for arranging a bribe to obtain a public licence.", "separate examination of consideration, promised act and object under section 23", "The consultant is retained only for lawful compliance advice and transparent filing work."),
    39: ("A regulated sale is routed through a sham intermediary solely to evade a statutory prohibition.", "whether the arrangement defeats the purpose of law even without using prohibited words", "The transaction follows the licensed statutory route and makes all required disclosures."),
    40: ("A payment is promised for influencing a public appointment through private access.", "a recognised category of injury, immorality or public policy applied with judicial restraint", "A transparent donation supports a public training programme without buying an official decision."),
    41: ("One indivisible price covers lawful delivery and an unlawful smuggling promise.", "whether lawful promises are truly severable without rewriting the bargain", "The lawful and unlawful promises have separate prices and can operate independently."),
    42: ("An employment contract bars a worker from any trade anywhere in India for five years after leaving.", "whether the clause restrains lawful trade and whether a statutory exception applies", "A seller of goodwill accepts a limited local restraint connected to that sale."),
    43: ("A clause extinguishes every claim unless suit is filed within thirty days.", "section 28's rules on restraints, extinguishment and valid arbitration agreements", "A written clause submits defined disputes to arbitration without extinguishing substantive rights."),
    44: ("A supply agreement leaves price 'to be mutually agreed later' and provides no standard or mechanism.", "whether the missing matter is certain or objectively ascertainable", "Price is tied to a published commodity index on the delivery date."),
    45: ("Two people bet on a cricket result, each standing to win or lose with no other interest in the event.", "mutual chance of gain or loss compared with a genuine transaction dependent on a collateral event", "A crop supply contract allocates performance depending on rainfall measured by an official station."),
    46: ("A dealer states that a used car has never flooded, knowing the buyer treats that fact as essential.", "expertise, reliance, timing and importance in classifying the statement", "The statement is obvious sales praise with no factual content or demonstrated reliance."),
    47: ("A long-term power contract omits a minor operational step without which the express bargain cannot work.", "strict necessity and obviousness, not mere fairness or convenience", "The proposed term would improve profitability but is not needed to make the contract work."),
    48: ("A parking ticket refers to unusual liability terms only through a QR code shown after payment.", "reasonable notice and timing before or at formation, especially for unusual terms", "The unusual term is prominently displayed and acknowledged before checkout."),
    49: ("A seller's standard exclusion clause reasonably bears two meanings.", "text, contract as a whole, context, purpose and any residual ambiguity against the drafter", "The clause is clear, specific and individually negotiated."),
    50: ("A homebuyer contract lets the developer cancel freely while imposing a severe charge on the consumer for the same conduct.", "unfair consumer terms and statutory limits on exclusion or restraint", "Two advised businesses negotiate a balanced and proportionate liability cap."),
    51: ("An app pre-ticks a paid subscription and hides the recurring charge below the purchase button.", "clear affirmative assent and whether interface design impairs a free and informed choice", "The box is unticked, the recurring price is prominent and the user clicks a specific acceptance control."),
    52: ("Three contractors jointly promise delivery, but the final task depends on one contractor's personal artistic skill.", "joint obligations, personal performance and accepted third-party performance", "The remaining task is standard work that the promisee validly accepts from a competent third person."),
    53: ("A seller tenders only half the goods at the wrong warehouse after business hours.", "complete and unconditional tender at the proper time, place and manner", "The seller tenders the exact goods correctly, but the buyer refuses access to the delivery bay."),
    54: ("A contractor demands payment although the contract requires a certified first milestone before that instalment.", "the agreed or naturally necessary order of reciprocal promises and readiness", "The owner prevents access needed to complete the milestone."),
    55: ("A caterer delivers after the wedding ceremony, while a warehouse contractor finishes slightly late without defeating the commercial purpose.", "whether time was intended to be essential and the effect of accepting late performance", "The promisee accepts late performance but gives no notice of an intention to claim delay compensation."),
    56: ("A debtor owing three loans marks a transfer 'for Loan B'.", "the debtor's express appropriation before the creditor's choice or the statutory default", "The debtor gives no indication and the creditor promptly appropriates the payment to a lawful due debt."),
    57: ("All parties agree that a new company will replace the original contractor and assume a newly defined obligation.", "mutual agreement to novation or alteration, distinguished from remission or extra time", "The creditor merely extends the due date without replacing the parties or core obligation."),
    58: ("A unique event hall burns down without either party's fault before the booked performance.", "initial or supervening impossibility or unlawfulness, without treating hardship alone as frustration", "Costs double, but the promised work remains physically and legally possible."),
    59: ("A supplier unequivocally states one month before delivery that it will not perform.", "repudiation and the promisee's election to terminate or keep the contract alive", "The due date passes without delivery, creating an actual breach."),
    60: ("A buyer claims cover-purchase cost and proven lost margin after non-delivery.", "the position performance would have produced, supported by proof and compensatory principle", "The buyer adds a punitive multiplier unrelated to proved contractual loss."),
    61: ("A buyer claims an unusual downstream shutdown loss that was never disclosed to the seller.", "causation, ordinary course or contemplation, proof and reasonable mitigation", "The special production purpose and likely shutdown loss were clearly communicated before contracting."),
    62: ("A contract states Rs 10 lakh for any delay, including a delay of one day with no material effect.", "reasonable compensation subject to the stipulated ceiling, not automatic recovery of the named sum", "Loss is difficult to quantify and the amount is a proportionate genuine pre-estimate, subject to the statutory test."),
    63: ("After breach, a buyer gives the required written notice, waits thirty days and hires a substitute supplier.", "the amended Specific Relief Act rules on substituted performance, notice and later specific performance", "The buyer hires a substitute immediately without the statutory notice process."),
    64: ("A deed records the wrong plot number, while a separate threatened breach may still be prevented.", "matching rectification, rescission, cancellation or injunction to the legal problem", "The instrument is accurate, but a party threatens a continuing breach of a negative obligation."),
    65: ("A contract benefits a niece who is not a party, although another relative supplied consideration.", "privity of contract remains distinct from India's wider rule on who may supply consideration", "The beneficiary is also the named promisee with an enforceable contractual right."),
    66: ("A marriage settlement creates an identified obligation for a named beneficiary who is not a signatory.", "a recognised trust, charge, family-arrangement or acknowledgment basis for third-party enforcement", "The claimant is only an incidental beneficiary with no trust, charge or acknowledged right."),
    67: ("A party rescinds a voidable exchange after receiving an advance and transferring goods.", "the different restoration routes under sections 64 and 65 and section 33 of the Specific Relief Act", "The transaction is a minor's void agreement, requiring careful use of the special restitution rules rather than automatic section 65 enforcement."),
    68: ("To prevent auction, an interested tenant pays municipal tax that the owner was legally bound to pay.", "an interested payment of another's legal liability or a lawful non-gratuitous act whose benefit is enjoyed", "A volunteer expressly makes the payment as a gift with no expectation of reimbursement."),
    69: ("A business makes the same tax payment twice because of a processing mistake.", "payment or delivery by mistake or coercion and any fact-specific restitution defence", "The payer acts voluntarily with full knowledge and no mistake or legally relevant pressure."),
}


AUTHORITIES = {
    2: "Central Inland Water Transport Corporation Ltd v Brojo Nath Ganguly (1986) 3 SCC 156",
    4: "Satyabrata Ghose v Mugneeram Bangur & Co, AIR 1954 SC 44",
    5: "Mulamchand v State of Madhya Pradesh, AIR 1968 SC 1218; State of West Bengal v B K Mondal & Sons, AIR 1962 SC 779",
    7: "Lalman Shukla v Gauri Dutt, (1913) ILR 35 All 489",
    8: "Lalman Shukla v Gauri Dutt, (1913) ILR 35 All 489",
    10: "Bhagwandas Goverdhandas Kedia v Girdharilal Parshottamdas, AIR 1966 SC 543",
    11: "Bhagwandas Goverdhandas Kedia v Girdharilal Parshottamdas, AIR 1966 SC 543",
    14: "Trimex International FZE Ltd v Vedanta Aluminium Ltd, (2010) 3 SCC 1",
    16: "Kollipara Sriramulu v T Aswatha Narayana, AIR 1968 SC 1028",
    17: "Durga Prasad v Baldeo, (1880) ILR 3 All 221",
    18: "Chinnaya v Ramaya, (1882) ILR 4 Mad 137",
    22: "Chidambara Iyer v P S Renga Iyer, AIR 1966 SC 193",
    23: "Kapurchand Godha v Mir Nawab Himayat Ali Khan Azamjah, AIR 1963 SC 250",
    24: "Mohori Bibee v Dharmodas Ghose, (1903) 30 IA 114",
    25: "Mohori Bibee v Dharmodas Ghose, (1903) 30 IA 114",
    30: "Chikkam Ammiraju v Chikkam Seshamma, (1917) ILR 41 Mad 33",
    31: "Raghunath Prasad v Sarju Prasad, AIR 1924 PC 60",
    32: "Subhas Chandra Das Mushib v Ganga Prosad Das Mushib, AIR 1967 SC 878",
    36: "Tarsem Singh v Sukhminder Singh, (1998) 3 SCC 471",
    38: "Gherulal Parakh v Mahadeodas Maiya, AIR 1959 SC 781",
    39: "Mannalal Khetan v Kedar Nath Khetan, (1977) 2 SCC 424",
    40: "Gherulal Parakh v Mahadeodas Maiya, AIR 1959 SC 781",
    42: "Gujarat Bottling Co Ltd v Coca Cola Co, (1995) 5 SCC 545; Niranjan Shankar Golikari v Century Spinning, AIR 1967 SC 1098",
    43: "Hakam Singh v Gammon (India) Ltd, (1971) 1 SCC 286",
    44: "Keshavlal Lallubhai Patel v Lalbhai Trikumlal Mills Ltd, AIR 1958 SC 512",
    45: "Gherulal Parakh v Mahadeodas Maiya, AIR 1959 SC 781",
    47: "Nabha Power Ltd v Punjab State Power Corporation Ltd, (2018) 11 SCC 508",
    48: "Bharati Knitting Co v DHL Worldwide Express Courier Division, (1996) 4 SCC 704",
    49: "General Assurance Society Ltd v Chandmull Jain, AIR 1966 SC 1644; Nabha Power Ltd v PSPCL, (2018) 11 SCC 508",
    50: "Pioneer Urban Land & Infrastructure Ltd v Govindan Raghavan, (2019) 5 SCC 725",
    51: "Trimex International FZE Ltd v Vedanta Aluminium Ltd, (2010) 3 SCC 1",
    54: "Nathulal v Phoolchand, AIR 1970 SC 546",
    55: "Hind Construction Contractors v State of Maharashtra, (1979) 2 SCC 70",
    57: "Lata Construction v Dr Rameshchandra Ramniklal Shah, (2000) 1 SCC 586",
    58: "Satyabrata Ghose v Mugneeram Bangur & Co, AIR 1954 SC 44; Energy Watchdog v CERC, (2017) 14 SCC 80",
    60: "Murlidhar Chiranjilal v Harishchandra Dwarkadas, AIR 1962 SC 366",
    61: "Murlidhar Chiranjilal v Harishchandra Dwarkadas, AIR 1962 SC 366",
    62: "Fateh Chand v Balkishan Dass, AIR 1963 SC 1405; Kailash Nath Associates v DDA, (2015) 4 SCC 136",
    63: "Katta Sujatha Reddy v Siddamsetty Infra Projects Pvt Ltd, (2023) 1 SCC 355",
    65: "M C Chacko v State Bank of Travancore, (1969) 2 SCC 343",
    66: "Khwaja Muhammad Khan v Husaini Begum, (1910) 37 IA 152; M C Chacko v State Bank of Travancore, (1969) 2 SCC 343",
    68: "State of West Bengal v B K Mondal & Sons, AIR 1962 SC 779",
    69: "Sales Tax Officer, Banaras v Kanhaiya Lal Mukundlal Saraf, AIR 1959 SC 135",
}


SOURCES = {
    "ICA": ("Indian Contract Act, 1872", "https://www.indiacode.nic.in/handle/123456789/2187"),
    "SRA": ("Specific Relief Act, 1963", "https://www.indiacode.nic.in/handle/123456789/17493?locale=en"),
    "CPA": ("Consumer Protection Act, 2019", "https://www.indiacode.nic.in/handle/123456789/16939"),
    "IT Act": ("Information Technology Act, 2000", "https://www.indiacode.nic.in/handle/123456789/13683"),
    "Constitution": ("Constitution of India", "https://legislative.gov.in/constitution-of-india/"),
    "Majority Act": ("Majority Act, 1875", "https://www.indiacode.nic.in/handle/123456789/2284"),
    "Arbitration": ("Arbitration and Conciliation Act, 1996", "https://www.indiacode.nic.in/handle/123456789/1978"),
    "Trusts Act": ("Indian Trusts Act, 1882", "https://www.indiacode.nic.in/handle/123456789/2327"),
    "TPA": ("Transfer of Property Act, 1882", "https://www.indiacode.nic.in/handle/123456789/2338?view_type=browse"),
    "Dark Patterns": ("Guidelines for Prevention and Regulation of Dark Patterns, 2023", "https://consumeraffairs.nic.in/acts-and-rules/consumer-protection/consumer-protection"),
}


def clean_visible(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("—", " - ").replace("\u00a0", " ")).strip()


def extract_modules() -> list[dict]:
    doc = Document(SOURCE_CURRICULUM)
    lines = [clean_visible(p.text) for p in doc.paragraphs if p.text.strip()]
    modules: list[dict] = []
    pattern = re.compile(r"^(\d{2})\s+(.+)$")
    for index, line in enumerate(lines):
        match = pattern.match(line)
        if not match:
            continue
        module_id = int(match.group(1))
        if not 1 <= module_id <= 69:
            continue
        anchor = lines[index + 1].removeprefix("INDIAN STATUTORY/LEGAL ANCHOR ")
        doctrine = lines[index + 2].removeprefix("DOCTRINAL PROBLEM ")
        mechanic_line = lines[index + 3].removeprefix("GAME MECHANIC ")
        game, _, mechanic = mechanic_line.partition(": ")
        modules.append({
            "id": module_id,
            "world": world_for(module_id),
            "title": clean_visible(match.group(2)),
            "anchor": clean_visible(anchor),
            "doctrine": clean_visible(doctrine),
            "game": clean_visible(game),
            "mechanic": clean_visible(mechanic),
        })
    if [m["id"] for m in modules] != list(range(1, 70)):
        raise RuntimeError("Curriculum extraction did not produce Modules 1 through 69")
    return modules


def source_links(anchor: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    for marker, source in SOURCES.items():
        if marker in anchor or (marker == "Arbitration" and "Arbitration Act" in anchor):
            found.append(source)
    if not any(name == "Indian Contract Act, 1872" for name, _ in found):
        found.insert(0, SOURCES["ICA"])
    return found


def authority_link(authority: str) -> str:
    query = authority.split(";")[0].split(", AIR")[0].split(", (")[0]
    return "https://indiankanoon.org/search/?formInput=" + urllib.parse.quote(query)


def module_html(module: dict, all_modules: list[dict]) -> str:
    module_id = module["id"]
    scenario, key_fact, replay = FACTS[module_id]
    authority = AUTHORITIES.get(module_id, "The statutory text is the primary authority for this module; applications remain fact-sensitive.")
    links = source_links(module["anchor"])
    source_markup = "".join(
        f'<a href="{html.escape(url)}" target="_blank" rel="noopener noreferrer">{html.escape(name)}</a>'
        for name, url in links
    )
    if module_id in AUTHORITIES:
        source_markup += f'<a href="{html.escape(authority_link(authority))}" target="_blank" rel="noopener noreferrer">Locate the leading authority</a>'
    previous_href = "module-01.html" if module_id == 2 else f"module-{module_id - 1:02d}.html"
    next_href = "index.html" if module_id == 69 else f"module-{module_id + 1:02d}.html"
    next_label = "Return to course map" if module_id == 69 else f"Continue to Module {module_id + 1:02d}"
    decisive_options = [
        (clean_visible(key_fact), "That is the legally decisive inquiry for this file."),
        ("the label or heading chosen by one party, without checking the facts", "A party's label can be evidence, but it cannot replace the statutory test."),
        ("commercial desirability alone, regardless of the statutory elements", "Commercial sense matters to strategy, but it cannot displace the governing legal elements."),
    ]
    shift = module_id % 3
    decisive_options = decisive_options[shift:] + decisive_options[:shift]
    correct_fact = [option for option, _ in decisive_options].index(clean_visible(key_fact))
    data = {
        "correctFact": correct_fact,
        "factFeedback": [feedback for _, feedback in decisive_options],
        "correctStrategy": 1,
        "strategyFeedback": [
            "Too early. A conclusion without the decisive facts risks applying the right provision to the wrong legal category.",
            "Strong method. Preserve the facts, apply the statutory elements and state any case-law qualification before choosing the remedy or result.",
            "Too broad. Indian contract law rarely supports rejecting a transaction without first identifying the governing rule and material facts.",
        ],
        "correctReplay": 0,
        "replayFeedback": [
            "Correct. The changed fact is legally material, so the analysis must be performed again rather than copied from the first file.",
            "Look again. The replay changes the fact that the module identifies as legally decisive.",
            "The change may have commercial effects, but it also alters the legal classification required by this module.",
        ],
    }
    title = html.escape(module["title"])
    anchor = html.escape(module["anchor"])
    doctrine = html.escape(module["doctrine"])
    game = html.escape(module["game"])
    mechanic = html.escape(module["mechanic"])
    scenario_html = html.escape(clean_visible(scenario))
    replay_html = html.escape(clean_visible(replay))
    authority_html = html.escape(clean_visible(authority))
    fact_buttons = "".join(
        f'<button class="choice" type="button" data-group="fact" data-value="{i}"><span>{chr(65+i)}</span>{html.escape(option)}</button>'
        for i, (option, _) in enumerate(decisive_options)
    )
    return f'''<!doctype html>
<html lang="en" data-visualize-standalone>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="referrer" content="no-referrer">
<title>Module {module_id:02d} | {title}</title>
<style>
:root {{
  color-scheme: light dark;
  --background: light-dark(rgb(255 255 255), rgb(24 24 24));
  --foreground: light-dark(rgb(26 28 31), rgb(255 255 255));
  --card: color-mix(in oklab, var(--foreground) 5%, transparent);
  --primary: light-dark(rgb(51 156 255), rgb(131 195 255));
  --primary-foreground: light-dark(rgb(255 255 255), rgb(13 13 13));
  --secondary: light-dark(rgb(255 255 255 / 96%), rgb(54 54 54 / 96%));
  --muted: color-mix(in srgb, var(--foreground) 10%, transparent);
  --muted-foreground: light-dark(rgb(26 28 31 / 62%), rgb(255 255 255 / 68%));
  --destructive: light-dark(rgb(186 38 35), rgb(250 92 88));
  --border: light-dark(rgb(26 28 31 / 12%), rgb(255 255 255 / 14%));
  --input: light-dark(rgb(26 28 31 / 18%), rgb(255 255 255 / 20%));
  --ring: light-dark(rgb(51 156 255), rgb(131 195 255 / 76%));
  --green: light-dark(rgb(0 145 58), rgb(74 211 127));
  --viz-series-1: var(--primary);
  --viz-series-2: light-dark(#f3883b, #f59a56);
  font-family: "Courier New", ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
}}
* {{ box-sizing: border-box; }}
[hidden] {{ display:none !important; }}
html {{ background: var(--background); }}
body {{ margin:0; padding:12px; color:var(--foreground); background:var(--background); font-size:15px; line-height:1.5; }}
button, a {{ font:inherit; }} button {{ color:inherit; }}
.shell {{ position:relative; width:min(1040px,100%); margin:0 auto; padding:14px; overflow:hidden; border:1px solid var(--border); border-radius:10px; background:var(--card); }}
.scanlines {{ position:absolute; inset:0; z-index:5; pointer-events:none; opacity:.025; background:repeating-linear-gradient(to bottom,transparent 0,transparent 3px,var(--foreground) 4px); }}
.topbar {{ position:relative; z-index:1; display:flex; gap:12px; align-items:center; justify-content:space-between; padding:0 0 12px; border-bottom:1px solid var(--border); }}
.brand,.eyebrow {{ font-weight:500; letter-spacing:.08em; text-transform:uppercase; }}
.status,.role,.lede {{ color:var(--muted-foreground); }}
.status {{ font-size:.86rem; }}
.meter {{ position:relative; z-index:1; height:9px; margin-top:10px; border:1px solid var(--border); background:repeating-linear-gradient(to right,transparent 0,transparent calc(20% - 4px),var(--background) calc(20% - 4px),var(--background) 20%); overflow:hidden; }}
.meter span {{ display:block; width:0; height:100%; background:var(--primary); transition:width .3s steps(5,end); }}
.hero {{ position:relative; z-index:1; display:grid; gap:9px; padding:24px 4px 22px; border-bottom:1px solid var(--border); }}
.eyebrow {{ color:var(--foreground); font-size:.78rem; }}
h1,h2,h3,p {{ margin:0; }} h1,h2,h3,strong {{ font-weight:500; }}
h1 {{ max-width:820px; font-size:clamp(1.65rem,3.8vw,2.55rem); line-height:1.12; }}
h2 {{ margin:6px 0 12px; font-size:clamp(1.35rem,3vw,2rem); line-height:1.15; }}
h3 {{ margin-bottom:9px; font-size:1rem; }}
p {{ line-height:1.6; }} .lede {{ max-width:780px; }}
.chips {{ display:flex; flex-wrap:wrap; gap:8px; margin-top:7px; }}
.chip {{ padding:5px 9px; border:1px solid var(--border); background:color-mix(in srgb,var(--muted) 55%,transparent); font-size:.78rem; }}
.screen {{ position:relative; z-index:1; min-height:390px; }}
.stage {{ display:none; padding:24px 4px 20px; }}
.stage.active {{ display:block; animation:screen-enter .28s steps(4,end) both; }}
.stage-grid {{ display:grid; grid-template-columns:minmax(0,1.2fr) minmax(250px,.8fr); gap:22px; align-items:start; }}
.docket {{ margin:14px 0; padding:14px 0 14px 14px; border-left:3px solid var(--viz-series-2); border-block:1px solid var(--border); background:color-mix(in srgb,var(--viz-series-2) 10%,transparent); }}
.role {{ font-size:.86rem; }}
.rule-card {{ display:grid; grid-template-columns:52px minmax(0,1fr); gap:8px 13px; padding:14px; border:1px solid var(--border); background:color-mix(in srgb,var(--muted) 55%,transparent); }}
.rule-card::before {{ content:"··\\A⌣"; white-space:pre; display:grid; place-items:center; grid-row:1 / span 3; width:52px; height:52px; padding-top:5px; border:2px solid var(--border); background:color-mix(in srgb,var(--primary) 15%,transparent); text-align:center; line-height:1.1; animation:guide-bob 900ms steps(2,end) infinite alternate; }}
.rule-card h3::before {{ content:"[ GUIDE BOT ]"; display:block; margin-bottom:7px; color:var(--primary); font-size:.78rem; letter-spacing:.08em; animation:cursor-blink 700ms steps(2,end) infinite; }}
.choices {{ display:grid; gap:9px; margin:18px 0; }}
.choice {{ display:flex; gap:12px; width:100%; padding:12px; text-align:left; border:1px solid var(--border); border-radius:8px; background:var(--secondary); cursor:pointer; line-height:1.45; }}
.choice:hover,.choice:focus-visible {{ border-color:var(--primary); outline:2px solid var(--ring); outline-offset:1px; }}
.choice.selected {{ border-color:var(--primary); color:var(--primary-foreground); background:var(--primary); }}
.choice span {{ display:grid; place-items:center; flex:0 0 28px; height:28px; border:1px solid currentColor; font-weight:500; }}
.controls {{ display:flex; flex-wrap:wrap; gap:9px; margin-top:18px; }}
.btn {{ display:inline-flex; align-items:center; justify-content:center; min-height:40px; padding:9px 14px; border:1px solid var(--input); border-radius:8px; color:var(--foreground); background:var(--secondary); text-decoration:none; cursor:pointer; }}
.btn:hover,.btn:focus-visible {{ border-color:var(--primary); outline:2px solid var(--ring); outline-offset:1px; }}
.btn.primary {{ border-color:var(--foreground); color:var(--background); background:var(--foreground); font-weight:500; }}
.btn:disabled {{ opacity:.4; cursor:not-allowed; }}
.feedback {{ min-height:54px; margin-top:14px; padding:12px 0 12px 14px; border-left:3px solid var(--border); background:color-mix(in srgb,var(--muted) 45%,transparent); }}
.feedback.good {{ border-color:var(--green); }} .feedback.bad {{ border-color:var(--destructive); }}
.law-grid {{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; margin:18px 0; }}
.law-box {{ padding:13px; border:1px solid var(--border); background:color-mix(in srgb,var(--muted) 48%,transparent); }}
.law-box strong {{ display:block; margin-bottom:7px; color:var(--primary); letter-spacing:.04em; }}
.sources {{ display:flex; flex-wrap:wrap; gap:8px; margin-top:14px; }}
.sources a {{ color:var(--foreground); padding:7px 9px; border:1px solid var(--border); text-underline-offset:3px; }}
.result {{ display:inline-block; margin:18px 0 8px; padding:10px 14px; border:3px double var(--green); color:var(--green); font-size:clamp(1.25rem,3vw,1.8rem); letter-spacing:.06em; }}
.nav {{ position:relative; z-index:1; display:flex; justify-content:space-between; gap:10px; padding-top:14px; border-top:1px solid var(--border); }}
.sr-only {{ position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; border:0; }}
@keyframes screen-enter {{ from {{ opacity:0; transform:translateX(8px); }} }}
@keyframes cursor-blink {{ 50% {{ opacity:.45; }} }}
@keyframes guide-bob {{ to {{ transform:translateY(-3px); }} }}
@media (max-width:700px) {{ body {{ padding:6px; }} .shell {{ padding:10px; }} .stage-grid,.law-grid {{ grid-template-columns:1fr; }} .topbar {{ align-items:flex-start; }} .nav {{ flex-wrap:wrap; }} }}
@media (prefers-reduced-motion:reduce) {{ *,*::before,*::after {{ animation-duration:.01ms!important; transition-duration:.01ms!important; scroll-behavior:auto!important; }} }}
</style>
</head>
<body>
<main class="shell" id="module-root">
  <div class="scanlines" aria-hidden="true"></div>
  <header class="topbar">
    <div><div class="brand">LAW MATTERS // CASE LAB {module_id:02d}</div><div class="status">{html.escape(WORLDS[module['world']])}</div></div>
    <div class="status" aria-live="polite">Stage <strong id="stage-count">1/5</strong> | Reasoning <strong id="score">0</strong>/3</div>
  </header>
  <div class="meter" aria-label="Module progress"><span id="meter-fill"></span></div>
  <section class="hero">
    <div class="eyebrow">Module {module_id:02d} | {game} | 5 to 10 minutes</div>
    <h1>{title}</h1>
    <p class="lede">{mechanic}</p>
    <div class="chips"><span class="chip">Indian law</span><span class="chip">{anchor}</span><span class="chip">Changed-fact replay</span></div>
  </section>
  <section class="screen" aria-label="Interactive module">
    <section class="stage active" data-stage="0">
      <div class="stage-grid">
        <div><div class="eyebrow">Briefing</div><h2>Enter the {game}</h2><p>{doctrine}</p><div class="controls"><button class="btn primary" type="button" data-next>Open the case file</button></div></div>
        <aside class="rule-card"><h3>Your role</h3><p>You are the legal decision-maker. Protect the client's position without skipping the statutory test.</p><p class="role">Credit is awarded for identifying the decisive fact, choosing a defensible method and updating the answer when a material fact changes.</p></aside>
      </div>
    </section>
    <section class="stage" data-stage="1">
      <div class="eyebrow">Case file 1 | Fact diagnosis</div><h2>Which fact controls the legal route?</h2>
      <div class="docket"><strong>New file</strong><p>{scenario_html}</p></div>
      <div class="choices" role="group" aria-label="Choose the decisive fact">{fact_buttons}</div>
      <button class="btn primary" type="button" data-check="fact" disabled>Lock the decisive fact</button>
      <div class="feedback" id="feedback-fact" role="status">Select one fact before locking your answer.</div>
      <div class="controls"><button class="btn primary" type="button" data-next hidden>Choose a response</button></div>
    </section>
    <section class="stage" data-stage="2">
      <div class="eyebrow">Case file 1 | Strategy</div><h2>What is the most defensible next move?</h2>
      <div class="choices" role="group" aria-label="Choose legal strategy">
        <button class="choice" type="button" data-group="strategy" data-value="0"><span>A</span>Announce a final legal result immediately from the transaction label.</button>
        <button class="choice" type="button" data-group="strategy" data-value="1"><span>B</span>Preserve the evidence, test every statutory element and state any case-law qualification before selecting the result or remedy.</button>
        <button class="choice" type="button" data-group="strategy" data-value="2"><span>C</span>Ignore the legal classification and select only the commercially attractive outcome.</button>
      </div>
      <button class="btn primary" type="button" data-check="strategy" disabled>Test the strategy</button>
      <div class="feedback" id="feedback-strategy" role="status">Choose the method you would defend in a written opinion.</div>
      <div class="controls"><button class="btn primary" type="button" data-next hidden>Reveal the changed fact</button></div>
    </section>
    <section class="stage" data-stage="3">
      <div class="eyebrow">Replay | One fact changes</div><h2>Must the legal analysis change?</h2>
      <div class="docket"><strong>Changed fact</strong><p>{replay_html}</p></div>
      <div class="choices" role="group" aria-label="Classify the changed fact">
        <button class="choice" type="button" data-group="replay" data-value="0"><span>A</span>Yes. Reapply the legal test because the changed fact is material.</button>
        <button class="choice" type="button" data-group="replay" data-value="1"><span>B</span>No. Copy the first answer without further analysis.</button>
        <button class="choice" type="button" data-group="replay" data-value="2"><span>C</span>It changes only the commercial story and can never affect legal classification.</button>
      </div>
      <button class="btn primary" type="button" data-check="replay" disabled>Run the replay</button>
      <div class="feedback" id="feedback-replay" role="status">Treat every changed fact as a fresh legal question.</div>
      <div class="controls"><button class="btn primary" type="button" data-next hidden>Open the law desk</button></div>
    </section>
    <section class="stage" data-stage="4">
      <div class="eyebrow">Law desk | Verified source route</div><h2>Debrief the decision</h2>
      <div class="law-grid">
        <div class="law-box"><strong>Statutory anchor</strong><p>{anchor}</p></div>
        <div class="law-box"><strong>Doctrinal rule</strong><p>{doctrine}</p></div>
        <div class="law-box"><strong>Authority for study</strong><p>{authority_html}</p></div>
        <div class="law-box"><strong>Targeted misconception</strong><p>A legal label or attractive outcome does not substitute for the statutory elements, material facts and any controlling authority.</p></div>
      </div>
      <p class="role">Source check: statutory anchors were checked against the official India Code or responsible ministry source on 31 July 2026. Case references are learning signposts, not a substitute for checking the latest authorised report, subsequent treatment and any applicable state amendment.</p>
      <div class="sources">{source_markup}</div>
      <div class="result" id="final-result">Reasoning score: 0/3</div>
      <p id="final-copy">Complete all three decisions to calculate your result.</p>
      <div class="controls"><button class="btn" type="button" id="restart">Replay this module</button><a class="btn primary" href="{next_href}">{html.escape(next_label)}</a></div>
    </section>
  </section>
  <nav class="nav" aria-label="Course navigation"><a class="btn" href="{previous_href}">Previous module</a><a class="btn" href="index.html">Course map</a></nav>
</main>
<script>
(() => {{
  "use strict";
  const DATA = {json.dumps(data, ensure_ascii=False)};
  const root = document.getElementById("module-root");
  const stages = Array.from(root.querySelectorAll("[data-stage]"));
  const selections = {{ fact: null, strategy: null, replay: null }};
  const completed = {{ fact: false, strategy: false, replay: false }};
  let stage = 0;
  let score = 0;
  function showStage(index) {{
    stage = index;
    stages.forEach((node, i) => node.classList.toggle("active", i === index));
    document.getElementById("stage-count").textContent = `${{index + 1}}/5`;
    document.getElementById("meter-fill").style.width = `${{(index + 1) * 20}}%`;
    document.getElementById("score").textContent = String(score);
    if (index === 4) {{
      document.getElementById("final-result").textContent = `Reasoning score: ${{score}}/3`;
      document.getElementById("final-copy").textContent = score === 3 ? "Excellent. You identified the decisive fact, used a defensible method and updated the analysis on replay." : score === 2 ? "Good. Review the feedback once more, then replay for a complete legal-reasoning chain." : "Replay recommended. Focus on statutory elements before conclusions.";
    }}
    if (index > 0) root.querySelector(".screen").scrollIntoView({{ behavior: "smooth", block: "start" }});
  }}
  root.addEventListener("click", (event) => {{
    const next = event.target.closest("[data-next]");
    if (next) {{ showStage(Math.min(4, stage + 1)); return; }}
    const choice = event.target.closest(".choice[data-group]");
    if (choice) {{
      const group = choice.dataset.group;
      if (completed[group]) return;
      selections[group] = Number(choice.dataset.value);
      root.querySelectorAll(`.choice[data-group="${{group}}"]`).forEach(node => node.classList.toggle("selected", node === choice));
      const check = root.querySelector(`[data-check="${{group}}"]`);
      if (check) check.disabled = false;
      return;
    }}
    const check = event.target.closest("[data-check]");
    if (check) {{
      const group = check.dataset.check;
      if (completed[group] || selections[group] === null) return;
      const correct = DATA[`correct${{group[0].toUpperCase() + group.slice(1)}}`];
      const selected = selections[group];
      const ok = selected === correct;
      if (ok) score += 1;
      completed[group] = true;
      document.getElementById("score").textContent = String(score);
      const feedback = document.getElementById(`feedback-${{group}}`);
      feedback.textContent = DATA[`${{group}}Feedback`][selected];
      feedback.classList.add(ok ? "good" : "bad");
      root.querySelectorAll(`.choice[data-group="${{group}}"]`).forEach(node => node.disabled = true);
      check.disabled = true;
      const section = check.closest(".stage");
      const nextButton = section.querySelector("[data-next]");
      if (nextButton) nextButton.hidden = false;
    }}
  }});
  document.getElementById("restart").addEventListener("click", () => window.location.reload());
  showStage(0);
}})();
</script>
</body>
</html>
'''.replace("—", " - ")


def build_index() -> None:
    text = SOURCE_INDEX.read_text(encoding="utf-8")
    text = text.replace("contract-law-course-map.html", "index.html")
    text = text.replace("`modules/module-${padded(module.id)}.html`", "`module-${padded(module.id)}.html`")
    text = text.replace("modules/module-${padded(moduleId)}.html", "module-${padded(moduleId)}.html")
    text = text.replace("1 PLAYABLE · 68 PLANNED", "69 PLAYABLE")
    text = text.replace(
        'if (module.id === 1) {\n        link.dataset.live = "true";\n      } else {\n        link.dataset.planned = "true";\n        link.dataset.moduleId = String(module.id);\n      }',
        'link.dataset.live = "true";'
    )
    text = text.replace(
        'game.textContent = `${module.game}${module.id === 1 ? " · PLAYABLE" : " · PAGE PLANNED"}`;',
        'game.textContent = `${module.game} · PLAYABLE`;'
    )
    text = text.replace(
        'els.routeNotice.textContent = "Module 01 opens the completed game. Other milestone links identify their planned page route and corresponding curriculum entry.";',
        'els.routeNotice.textContent = "Every milestone opens its completed game page.";'
    )
    text = text.replace("—", " - ")
    (ROOT / "index.html").write_text(text, encoding="utf-8")


def build_module_one() -> None:
    text = SOURCE_MODULE_ONE.read_text(encoding="utf-8").replace("—", " - ")
    (ROOT / "module-01.html").write_text(text, encoding="utf-8")


def main() -> None:
    from strategic_course import build_strategic_course

    build_strategic_course()


if __name__ == "__main__":
    main()
