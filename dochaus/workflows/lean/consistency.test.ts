import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync, writeFileSync } from "node:fs"
import { finalBodyParagraphs, shownText } from "./checks"
import { consistencyFindings, resolvedReferences, type PackDocument, type Para } from "./consistency"
import { finalGate, type GateResult } from "./gate"
import { runLean } from "./runner"
import { leanFile, readLeanJson, readOpenItems } from "./state"
import { AGENT, NO_FINDINGS, SOURCE, cleanupMatters, drafter, draftingNotes, matter, noChanges, pidIn, sameSession, scriptedModel, type Reply } from "./test-support"

// Cross-references and defined terms of a final body, by code. The fixtures
// are a neutral, invented supply agreement and pack; each case changes one
// thing about it.

const doc = (lines: string[]): Para[] => lines.map((text, index) => ({ pid: `p${index + 1}`, text }))
const check = (input: Parameters<typeof consistencyFindings>[0]) => consistencyFindings({ form: "agreement", ...input })
const details = (input: Parameters<typeof consistencyFindings>[0]) => check(input).map(row => row.detail)
const pack = (title: string, lines: string[], earlier = true): PackDocument => ({ title, paras: doc(lines), earlier })

const PARTIES = 'This agreement is between Northfield Packaging Ltd (the "Supplier") and Eastmere Foods Ltd (the "Customer").'
const DEFINITIONS = [
  "1. Definitions",
  '1.1 "Goods" means the printed cartons described in Schedule 1.',
  '1.2 "Price" means the price for the Goods set out in Schedule 2.',
  '1.3 "Delivery Date" means the date stated in an order.',
]
const SUPPLY = ["2. Supply", "2.1 The Supplier shall supply the Goods to the Customer on each Delivery Date.", "2.2 Clause 2.1 does not oblige the Customer to place any order."]
const PAYMENT = ["3. Price and payment", "3.1 The Customer shall pay the Price within thirty days of invoice.", "3.2 Late payment carries interest under clause 3.1 and clauses 2.1 and 2.2."]
const LIABILITY = ["4. Liability", "4.1 Subject to clause 5, the liability of each party is limited to the Price paid in the preceding year."]
const TERMINATION = ["5. Termination", "5.1 Either party may terminate by notice if the other commits a material breach of clause 3 or clause 4."]
const SCHEDULES = ["Schedule 1 – Goods", "Printed cartons as specified in each order.", "Schedule 2 – Price", "The Price is stated in each order."]
const BASE = ["Supply Agreement", PARTIES, ...DEFINITIONS, ...SUPPLY, ...PAYMENT, ...LIABILITY, ...TERMINATION, ...SCHEDULES]
const TITLE = "Supply Agreement"

describe("cross-references after changes", () => {
  test("an unchanged agreement with valid references, statute and other instruments has no findings", () => {
    const final = [...BASE.slice(0, -4),
      "5.2 Nothing in this agreement limits liability under section 2 of the Sale of Goods Act 1979, Article 6 GDPR or section 5 CA 2006.",
      "5.3 The rights of the Customer under clause 14 of the Framework Agreement are not affected, nor is clause 14 itself.",
      "5.4 [Further termination rights under clause 99 to be confirmed.]",
      "5.5 Clauses 2 to 4 (inclusive) and clause 3.1 (Payment) survive termination.",
      ...SCHEDULES]
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([])
  })

  test("a reference to a deleted clause is reported, naming the clause that was there", () => {
    const final = BASE.filter(line => !LIABILITY.includes(line))
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([
      "Clause 5.1 refers to clause 4, but no clause in the final text is numbered 4. There was a clause 4 before the changes.",
    ])
  })

  test("a deleted clause whose number the next clause takes leaves its references pointing at the wrong clause", () => {
    const final = [...BASE.filter(line => !LIABILITY.includes(line) && !TERMINATION.includes(line)),
      "4. Termination", "4.1 Either party may terminate by notice if the other commits a material breach of clause 3 or clause 4."]
    final.push(...final.splice(final.indexOf(SCHEDULES[0]!), SCHEDULES.length))
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([
      'Clause 4.1 refers to clause 4. Before the changes that reference was to "Liability", which is no longer in the text; clause 4 is now "Termination", which was clause 5 before. Check that the reference points where it should.',
    ])
  })

  test("an inserted clause renumbers the clauses after it under references carried from before", () => {
    const final = ["Supply Agreement", PARTIES, ...DEFINITIONS, ...SUPPLY,
      "3. Quality", "3.1 The Goods shall conform to Schedule 1.",
      "4. Price and payment", "4.1 The Customer shall pay the Price within thirty days of invoice.", "4.2 Late payment carries interest under clause 3.1 and clauses 2.1 and 2.2.",
      "5. Liability", "5.1 Subject to clause 5, the liability of each party is limited to the Price paid in the preceding year.",
      "6. Termination", "6.1 Either party may terminate by notice if the other commits a material breach of clause 3 or clause 4.",
      ...SCHEDULES]
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([
      'Clause 4.2 refers to clause 3.1. Before the changes that reference was to "The Customer shall pay the Price within thirty days of invo…", which is now clause 4.1; clause 3.1 is now "The Goods shall conform to Schedule 1.". Check that the reference points where it should.',
      'Clause 5.1 refers to clause 5. Before the changes that reference was to "Termination", which is now clause 6; clause 5 is now "Liability". Check that the reference points where it should.',
      'Clause 6.1 refers to clause 3. Before the changes that reference was to "Price and payment", which is now clause 4; clause 3 is now "Quality". Check that the reference points where it should.',
      'Clause 6.1 refers to clause 4. Before the changes that reference was to "Liability", which is now clause 5; clause 4 is now "Price and payment". Check that the reference points where it should.',
    ])
  })

  test("a live cross-reference is read as pointing at its field's target: a stale cached number is reported, a current one is not checked again", () => {
    const linked = "4.2 Late payment carries interest under clause 3.1 and clauses 2.1 and 2.2."
    const final = ["Supply Agreement", PARTIES, ...DEFINITIONS, ...SUPPLY,
      "3. Quality", "3.1 The Goods shall conform to Schedule 1.",
      "4. Price and payment", "4.1 The Customer shall pay the Price within thirty days of invoice.", linked,
      "5. Liability", "5.1 Subject to clause 5, the liability of each party is limited to the Price paid in the preceding year.",
      "6. Termination", "6.1 Either party may terminate by notice if the other commits a material breach of clause 3 or clause 4.",
      ...SCHEDULES]
    // The field on "3.1" still points at the payment clause, now 4.1; Word
    // shows 4.1 once its fields are updated.
    const link = (label: string, to: string, from = 0) => ({ start: linked.indexOf(label, from), length: label.length, to })
    const paras = doc(final).map(row => row.text === linked ? { ...row, links: [link("3.1", "4.1"), link("2.1", "2.1"), link("2.2", "2.2", linked.indexOf("2.1"))] } : row)
    expect(details({ title: TITLE, final: paras, baselines: [doc(BASE)], others: [] })).toEqual([
      "Clause 4.2 shows clause 3.1, but its cross-reference field points at what is now clause 4.1. Check which clause is meant; updating the document's fields (F9 in Word) shows 4.1.",
      'Clause 5.1 refers to clause 5. Before the changes that reference was to "Termination", which is now clause 6; clause 5 is now "Liability". Check that the reference points where it should.',
      'Clause 6.1 refers to clause 3. Before the changes that reference was to "Price and payment", which is now clause 4; clause 3 is now "Quality". Check that the reference points where it should.',
      'Clause 6.1 refers to clause 4. Before the changes that reference was to "Liability", which is now clause 5; clause 4 is now "Price and payment". Check that the reference points where it should.',
    ])
    // A field on "clause 4" pointing at what is now clause 4 settles that reference.
    const terminate = final.find(line => line.startsWith("6.1 "))!
    const settled = paras.map(row => row.text === terminate ? { ...row, links: [{ start: terminate.lastIndexOf("4"), length: 1, to: "4" }] } : row)
    expect(details({ title: TITLE, final: settled, baselines: [doc(BASE)], others: [] })).not.toContainEqual(expect.stringContaining("refers to clause 4."))
  })

  test("a renumbered sub-clause is reported where the old number is gone, with its new number", () => {
    const final = BASE.map(line => line.startsWith("2.2 ") ? line.replace("2.2 ", "2.3 ") : line)
    const found = check({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })
    expect(found).toEqual([{
      pid: `p${final.indexOf(PAYMENT[2]!) + 1}`,
      detail: "Clause 3.2 refers to clause 2.2, but no clause in the final text is numbered 2.2. What was clause 2.2 before the changes is now clause 2.3.",
    }])
  })

  test("a reference to a missing schedule, and to a missing clause of another pack document, is reported", () => {
    const final = [...BASE.slice(0, -4), "5.2 The service levels in Schedule 3 and clause 7 of the Service Agreement apply.", ...SCHEDULES]
    const service = pack("Service Agreement", ["Service Agreement", "1. Services", "1.1 The services are listed in the order.", "2. Fees", "2.1 Fees are stated in the order."])
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [service] })).toEqual([
      "Clause 5.2 refers to Schedule 3, but the final text has no Schedule 3.",
      'Clause 5.2 refers to clause 7 of "Service Agreement", but no clause in that document\'s final text is numbered 7.',
    ])
    const longer = pack("Service Agreement", ["Service Agreement", "1. Services", "1.1 The services are listed in the order.", "7. Service levels", "7.1 Service levels are stated in the order.", "8. Fees"])
    const fixed = final.map(line => line.replace("Schedule 3", "Schedule 2"))
    expect(details({ title: TITLE, final: doc(fixed), baselines: [doc(BASE)], others: [longer] })).toEqual([])
  })

  test("a schedule heading the changes close with a marker and a full stop heads its schedule where it is a heading: marked as one, or after the clauses with the original's heading", () => {
    // Schedule headings with no separator after the number, as an original may write them.
    const schedules = ["Schedule 1 Goods and packaging", "Printed cartons as specified in each order.", "Schedule 2 Price", "The Price is stated in each order."]
    const before = [...BASE.slice(0, -4), "5.2 The service levels in Schedule 3 apply.", ...schedules]
    const close = (heading: string, marker: string) => before.map(line => line === heading ? `${line} ${marker}.` : line)
    const marked = (lines: string[], heading: string) => doc(lines).map(para => para.text.startsWith(`${heading} [`) ? { ...para, heading: true as const } : para)
    const missing3 = "Clause 5.2 refers to Schedule 3, but the final text has no Schedule 3."

    const first = close(schedules[0]!, "[open item: confirm the carton specification]")
    expect(details({ title: TITLE, final: marked(first, schedules[0]!), baselines: [doc(before)], others: [] })).toEqual([missing3])
    // Unmarked, the first schedule's line reads as a sentence, even with the
    // original's heading: only a marked heading ends the clauses.
    expect(details({ title: TITLE, final: doc(first), original: doc(before), baselines: [doc(before)], others: [] })).toContain(
      "Clause 1.1 refers to Schedule 1, but the final text has no Schedule 1. Schedule 1 was in the text before the changes.")

    const second = close(schedules[2]!, "[open item: confirm the price list]")
    expect(details({ title: TITLE, final: doc(second), original: doc(before), baselines: [doc(before)], others: [] })).toEqual([missing3])
    expect(details({ title: TITLE, final: doc(second), baselines: [doc(before)], others: [] })).toContain(
      "Clause 1.2 refers to Schedule 2, but the final text has no Schedule 2. Schedule 2 was in the text before the changes.")
    // Live cross-references follow the text alone unless the line is marked.
    const schedule2 = (paras: Para[]) => resolvedReferences(TITLE, paras, [], "agreement").filter(ref => ref.kind === "schedule" && ref.label === "2").map(ref => ref.target)
    expect(schedule2(doc(second))).toEqual([])
    expect(schedule2(marked(second, schedules[2]!))).toEqual([second.indexOf(`${schedules[2]} [open item: confirm the price list].`)])
  })

  test("a sentence, list line or table cell closed by a marker and a full stop heads no schedule, among the clauses or after them", () => {
    // The original's schedule headings are given and marked, so each case
    // has every input that can make a line a heading.
    const headings = new Set(SCHEDULES.filter(line => line.startsWith("Schedule")))
    const marked = (lines: string[]) => doc(lines).map(para => headings.has(para.text) ? { ...para, heading: true as const } : para)
    const findings = (final: string[], original = BASE) => details({ title: TITLE, final: marked(final), original: doc(original), baselines: [doc(original)], others: [] })
    const without2 = BASE.slice(0, -2)
    const gone2 = ["Clause 1.2 refers to Schedule 2, but the final text has no Schedule 2. Schedule 2 was in the text before the changes.",
      "Schedule 1 refers to Schedule 2, but the final text has no Schedule 2. Schedule 2 was in the text before the changes."]
    for (const sentence of [
      "Schedule 2 The Price list will be provided by the Supplier [open item: confirm the price list].",
      "Schedule 2 Pricing [insert pricing table].",
      "Schedule 2 Price [insert]. Payable monthly [insert].",
      "Schedule 2 Price Is To Be Attached [open item: confirm the price list]?",
      "Schedule 2 Part B applies to all orders [insert].",
    ]) expect(findings([...without2, sentence])).toEqual(gone2)
    expect(findings([...BASE.slice(0, -4), "5.2 The service levels in Schedule 3 apply.", ...SCHEDULES,
      "Schedule 3 Service Levels will be agreed between the parties [open item: confirm the levels]."])).toEqual([
      "Clause 5.2 refers to Schedule 3, but the final text has no Schedule 3.", "Schedule 2 refers to Schedule 3, but the final text has no Schedule 3."])
    // Among the clauses: a paragraph of clause 4, table cells, and a list of the schedules.
    expect(findings([...BASE.slice(0, 13), "Schedule 4 Service Credits apply to late delivery [open item: confirm whether retained].", ...BASE.slice(13)])).toEqual([
      "Clause 4 refers to Schedule 4, but the final text has no Schedule 4."])
    expect(findings([...BASE.slice(0, 6), "Schedule", "Description", "Schedule 2 Pricing [tbc].", "Schedule 1 Goods [tbc].", ...BASE.slice(6, -2)])).toEqual([
      "Clause 1.2 refers to Schedule 2, but the final text has no Schedule 2. Schedule 2 was in the text before the changes.",
      "Clause 1.3 refers to Schedule 2, but the final text has no Schedule 2. Schedule 2 was in the text before the changes."])
    const clauses = [...SUPPLY, ...PAYMENT, ...LIABILITY, "5. Termination", "5.1 Either party may terminate by notice if the other commits a material breach of clause 3 or clause 9."]
    const listed = ["Supply Agreement", PARTIES, "1. Interpretation", "1.1 The Schedules form part of this agreement:", "Schedule 1 Goods [insert description];",
      "Schedule 2 Price [insert price list];", "Schedule 3 Service Levels [to be agreed].", '1.2 "Goods" means the printed cartons described in Schedule 1.', ...clauses, ...SCHEDULES]
    expect(findings(listed, listed)).toEqual(["Clause 1.1 refers to Schedule 3, but the final text has no Schedule 3.",
      "Clause 5.1 refers to clause 9, but no clause in the final text is numbered 9."])
    const particulars = ["Supply Agreement", PARTIES, "Item", "Detail", "Goods", "Schedule 1 Goods [insert description].", "Price", "Schedule 2 Price [insert].",
      "1. Interpretation", '1.1 "Goods" means the printed cartons described in Schedule 1.', ...clauses, ...SCHEDULES]
    expect(findings(particulars, particulars)).toEqual(["Clause 5.1 refers to clause 9, but no clause in the final text is numbered 9."])
  })

  test("a schedule's reference to its own paragraphs, and a reference to a lettered clause, are not checked as the agreement's clauses", () => {
    // The agreement has five clauses: "clause 7 of this Schedule" is the
    // schedule's own numbering and "clause 6A" an inserted clause, not clause 6.
    // A schedule's plain "clause 8" is still the agreement's.
    const final = [...BASE.slice(0, -4), "5.2 Clause 6A applies after termination.",
      "Schedule 1 – Goods", "1. Specification", "1.1 Cartons meet clause 7 of this Schedule.", "1.2 Clause 8 applies.", ...SCHEDULES.slice(2)]
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([
      "Schedule 1 refers to clause 8, but no clause in the final text is numbered 8.",
    ])
  })

  test("a document whose clause numbers are not printed, and a letter, are not checked for internal references", () => {
    const unnumbered = BASE.map(line => line.replace(/^\d+(?:\.\d+)?\.?\s+/, ""))
    expect(details({ title: TITLE, final: doc(unnumbered), baselines: [], others: [] })).toEqual([])
    const letter = ["Dear Sirs", "1. We enclose the agreement. Clause 9 sets out the termination rights.", "2. Clause 12 deals with notices.", "3. Please sign and return it."]
    expect(details({ title: "Covering letter", form: "letter", final: doc(letter), baselines: [], others: [] })).toEqual([])
  })

  test("a document planned without a form, or with a form the plan does not know, is not checked", () => {
    const final = BASE.filter(line => !LIABILITY.includes(line))
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(BASE)], others: [] })).toHaveLength(1)
    expect(details({ title: TITLE, form: undefined, final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([])
    expect(details({ title: TITLE, form: "memo", final: doc(final), baselines: [doc(BASE)], others: [] })).toEqual([])
  })

  test("a document well past 10,000 characters is checked to its end", () => {
    const clauses = Array.from({ length: 40 }, (_, index) => index + 1).flatMap(number => [
      `${number}. Obligations part ${number}`,
      `${number}.1 The obligations in this part apply to each delivery made under this agreement and continue until this agreement ends, subject to clause ${Math.max(1, number - 1)}.`,
      `${number}.2 Each party shall perform the obligations in clause ${number}.1 with reasonable care and skill and in accordance with good industry practice.`,
    ])
    clauses.splice(clauses.length - 1, 0, "40.3 The notice periods in clause 41 apply to every notice under this part.")
    const final = ["Long Agreement", ...clauses]
    expect(final.join("\n").length).toBeGreaterThan(10_000)
    const found = check({ title: "Long Agreement", final: doc(final), baselines: [], others: [] })
    expect(found).toEqual([{ pid: `p${final.indexOf("40.3 The notice periods in clause 41 apply to every notice under this part.") + 1}`,
      detail: "Clause 40.3 refers to clause 41, but no clause in the final text is numbered 41." }])
  })
})

describe("defined terms after changes", () => {
  test("a party short name used as a defined term without a definition is reported", () => {
    const records = ["5.2 The Supplier shall keep records of each delivery.", "5.3 The Customer may inspect the records of the Supplier."]
    const named = [...BASE.slice(0, -4), ...records, ...SCHEDULES]
    expect(details({ title: TITLE, final: doc(named), baselines: [], others: [] })).toEqual([])
    const final = named.map(line => line === PARTIES ? 'This agreement is between Northfield Packaging Ltd and Eastmere Foods Ltd (the "Customer").' : line)
    expect(details({ title: TITLE, final: doc(final), baselines: [], others: [] })).toEqual([
      '"Supplier" is used as a defined term (3 times, first in clause 2.1) but is not defined in the document.',
    ])
  })

  test("a term defined twice, a definition never used, mixed case and a removed definition are each reported", () => {
    const final = [...BASE.slice(0, 6),
      '1.4 "Delivery Date" means the date agreed in writing.',
      '1.5 "Business Day" means a day other than a Saturday or Sunday.',
      ...BASE.slice(6, -4),
      "5.2 Notice of a changed Delivery date must be given in writing.",
      "5.3 Each Order is subject to this agreement.",
      ...SCHEDULES]
    const before = [...BASE.slice(0, 6), '1.4 "Order" means an order for Goods placed by the Customer.', ...BASE.slice(6)]
    expect(details({ title: TITLE, final: doc(final), baselines: [doc(before)], others: [] })).toEqual([
      'The term "Delivery Date" is defined more than once: in clause 1.3 and in clause 1.4.',
      'The defined term "Delivery Date" is written "Delivery date" in clause 5.2; check that the defined term is used consistently.',
      'The term "Business Day" is defined in clause 1.5 but is not used anywhere else in the document.',
      'The term "Order" is used in the final text (first in clause 5.3) but is no longer defined: before the changes it was defined in clause 1.4.',
    ])
  })

  test("a term two pack documents define differently is reported in the later one; a consistent pack is not", () => {
    const supply = pack(TITLE, BASE.map(line => line.startsWith("1.2 ") ? '1.2 "Price" means ten thousand pounds (£10,000) per order.' : line))
    const guarantee = (supplier: string, price: string) => doc(["Guarantee", `This guarantee is given by Hollins Group Ltd (the "Guarantor") to Eastmere Foods Ltd (the "Customer") for the obligations of ${supplier} (the "Supplier").`,
      "1. Definitions", `1.1 "Price" means ${price} per order.`, "1.2 \"Demand\" means a written demand by the Customer.",
      "2. Guarantee", "2.1 The Guarantor guarantees payment by the Supplier of every sum due to the Customer, including the Price, within five days of a Demand."])
    expect(details({ title: "Guarantee", final: guarantee("Westbrook Packaging Ltd", "twelve thousand pounds (£12,000)"), baselines: [], others: [supply] })).toEqual([
      'The term "Supplier" is defined differently in this document (the paragraph beginning "This guarantee is given by Hollins Group Ltd (the "Guaranto…") and in "Supply Agreement" (the paragraph beginning "This agreement is between Northfield Packaging Ltd (the "Su…").',
      'The term "Price" is defined differently in this document (clause 1.1) and in "Supply Agreement" (clause 1.2).',
    ])
    expect(details({ title: "Guarantee", final: guarantee("NORTHFIELD PACKAGING LIMITED", "ten thousand pounds (£10,000)"), baselines: [], others: [supply] })).toEqual([])
  })

  test("ordinary words in a definition, a defined term at the start of a sentence and an explained term are not reported", () => {
    const final = [...BASE.slice(0, 6),
      '1.4 "Bank Consent" means the bank consent to the assignment.',
      ...BASE.slice(6, -4),
      "5.2 Delivery date changes need the Bank Consent. The Effective Date, being the date of signature, starts this agreement.",
      "5.3 The Supplier shall notify the Customer of the Effective Date and of any delay after the Effective Date.",
      ...SCHEDULES]
    expect(details({ title: TITLE, final: doc(final), baselines: [], others: [] })).toEqual([])
  })
})

// Through the real coordinator; only the model is scripted. The changes stage
// renumbers a sub-clause that a later clause refers to.
const AGREEMENT = [
  "# Supply agreement", "", "## 1 Delivery", "", "1.1 The supplier shall deliver the goods within thirty days of the order date.", "",
  "1.2 Risk in the goods passes on delivery under clause 1.1.", "", "## 2 Payment", "", "2.1 Payment is due on delivery.", "",
  "2.2 Payment is made by bank transfer.", "", "## 3 Remedies", "", "3.1 Late payment under clause 2.2 carries interest.", "",
].join("\n")
const renumber: Reply = prompt => JSON.stringify({
  amendments: [{ pid: pidIn(prompt, "2\\.2 Payment is made"), updated: "2.3 Payment is made by bank transfer.", reason: "Leaves room for a new clause 2.2.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }],
  questions: [],
})

// The same agreement written with Word-numberable clause numbers; the changes
// restate one paragraph's number and delete a sub-clause a later clause cites.
const LISTED = AGREEMENT.replace(/^## (\d) /gm, "## $1. ")
const removeListed: Reply = prompt => JSON.stringify({
  amendments: [
    { pid: pidIn(prompt, "2\\.1 Payment is due"), updated: "2.1 Payment is due on delivery in full.", reason: "States that payment is in full.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] },
    { pid: pidIn(prompt, "2\\.2 Payment is made"), operation: "delete", updated: "", reason: "The terms do not settle the payment method.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] },
  ],
  questions: [],
})

// An agreement whose clause 2 lists six charges as an ordered list, and a
// clause that cites a clause 5 the agreement does not have.
const ORDERED = [
  "# Supply agreement", "", "## 1. Delivery", "", "1.1 The supplier shall deliver the goods within thirty days of the order date.", "",
  "## 2. Payment", "", "2.1 The customer shall pay:", "",
  "1. The price;", "2. The delivery charge;", "3. The insurance charge;", "4. The packaging charge;", "5. The storage charge; and", "6. Any duty.", "",
  "2.2 Late payment under clause 5 carries interest.", "",
].join("\n")

// A definition the changes take out while "the Specification" stays in use, and a
// clause referred to twice from one paragraph and then renumbered.
const DEFINED = [
  "# Supply agreement", "", "## 1 Definitions", "", '1.1 "Goods" means the printed cartons described in the Specification.', "",
  '1.2 "Specification" means the written description of the Goods agreed by the buyer.', "", '1.3 "Price" means the price stated in the Specification.', "",
  '1.4 "Delivery Date" means the date stated in the Specification.', "", "## 2 Payment", "", "2.1 The buyer shall pay the Price on the Delivery Date.", "",
  "2.2 Payment is made by bank transfer.", "", "## 3 Remedies", "", "3.1 Late payment under clause 2.2 carries interest, and clause 2.2 governs how that interest is paid.", "",
  "3.2 The Goods remain at the seller's risk until they meet the Specification.", "",
].join("\n")
const undefine: Reply = prompt => JSON.stringify({
  amendments: [
    { pid: pidIn(prompt, '1\\.2 "Specification" means'), updated: "1.2 Specifications are agreed in writing.", reason: "How specifications are agreed is stated instead.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] },
    { pid: pidIn(prompt, "2\\.2 Payment is made"), updated: "2.3 Payment is made by bank transfer.", reason: "Leaves room for a new clause 2.2.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] },
  ],
  questions: [],
})

describe("the check in a lean run", () => {
  afterEach(cleanupMatters)

  test("a reference the changes break becomes an apply open item that the drafting notes disclose and the gate settles", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.draft]: drafter(AGREEMENT), [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: sameSession(renumber), [AGENT.finalize]: draftingNotes })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const items = readOpenItems(directory).filter(item => item.kind === "consistency")
    expect(items.map(item => ({ stage: item.stage, detail: item.detail, document: item.ref.document }))).toEqual([{
      stage: "apply", document: "d1",
      detail: "Clause 3.1 refers to clause 2.2, but no clause in the final text is numbered 2.2. What was clause 2.2 before the changes is now clause 2.3.",
    }])
    expect(model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text).toContain(items[0]!.detail)
    expect(readFileSync(leanFile(directory, "drafting-notes.md"), "utf8")).toContain(`[ref: ${items[0]!.id}]`)
    const gate = readLeanJson(directory, "gate.json") as GateResult
    expect(gate.status).toBe("pass")
    expect(gate.items.find(item => item.id === items[0]!.id)).toMatchObject({ status: "disclosed", resolution: "cross-reference or defined-term problem named in the drafting notes" })

    writeFileSync(leanFile(directory, "drafting-notes.md"), "# Drafting notes\n\nNothing to report.\n")
    expect(finalGate(directory)).toMatchObject({ status: "blocked",
      blocking: [expect.objectContaining({ kind: "consistency", resolution: "the drafting notes do not name this cross-reference or defined-term problem" })] })
  }, 120_000)

  test("a drafted schedule heading the changes close with an open item still heads the schedule its clauses refer to", async () => {
    // Each part prints its own heading, as an outline's display headings have it.
    const parts: Record<string, string> = {
      "1 Delivery": "## 1 Delivery\n\n1.1 The supplier shall deliver the goods described in Schedule 1 within thirty days of the order date.\n\n1.2 The goods are packed as set out in Schedule 2.",
      "2 Payment": "## 2 Payment\n\n2.1 Payment is due on delivery.",
      "Schedule 1 Goods": "## Schedule 1 Goods\n\nPrinted cartons as specified in each order.",
    }
    const outline = ["Acting for: not stated", "# Supply agreement", "Form: agreement",
      ...Object.keys(parts).flatMap(heading => [`## ${heading}`, `Display heading: ${heading}`, `The ${heading} part.`])].join("\n")
    const draft: Reply = prompt => {
      const named = /^Draft part \d+ of \d+ of "[^"]*": (.+)$/m.exec(prompt)
      return named ? parts[named[1]!]! : outline
    }
    const flag: Reply = prompt => JSON.stringify({
      amendments: [{ pid: pidIn(prompt, "Schedule 1 Goods"), updated: "Schedule 1 Goods [open item: confirm the carton specification].",
        reason: "The carton specification is not settled.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }],
      questions: [],
    })
    const directory = matter()
    const model = scriptedModel({ [AGENT.draft]: draft, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: sameSession(flag), [AGENT.finalize]: draftingNotes })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).toContain("Schedule 1 Goods [open item: confirm the carton specification].")
    // Schedule 2, which the text never had, is still reported.
    expect(readOpenItems(directory).filter(item => item.kind === "consistency").map(item => item.detail)).toEqual([
      "Clause 1.2 refers to Schedule 2, but the final text has no Schedule 2.",
    ])
  }, 120_000)

  test("a body numbered by Word's list numbering is checked by the numbers its list gives it", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.draft]: drafter(LISTED), [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: sameSession(removeListed), [AGENT.finalize]: draftingNotes })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    // The numbers are list labels, not text; a replacement that restates its
    // label does not print it twice.
    const final = finalBodyParagraphs(directory, "d1")
    expect(final.find(row => row.text === "Payment is due on delivery in full.")?.label).toBe("2.1")
    // The live reference to the deleted clause names no clause now, and is reported.
    const marker = "[open item: refers to 2.2, which was removed]"
    expect(final.map(shownText)).toContain(`3.1 Late payment under clause ${marker} carries interest.`)
    expect(readOpenItems(directory).filter(item => item.ref.placeholder === marker).map(item => item.kind)).toEqual(["placeholder"])
    expect(readOpenItems(directory).filter(item => item.kind === "consistency")).toEqual([])
  }, 120_000)

  test("an ordered list inside a clause does not number clauses: a reference to a clause the list's numbers seem to give is still checked", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.draft]: drafter(ORDERED), [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: sameSession(noChanges), [AGENT.finalize]: draftingNotes })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const final = finalBodyParagraphs(directory, "d1")
    expect(final.find(row => row.text === "The storage charge; and")?.label).toBe("5.")
    expect(readOpenItems(directory).filter(item => item.kind === "consistency").map(item => item.detail))
      .toEqual([expect.stringContaining("Clause 2.2 refers to clause 5, but no clause in the final text is numbered 5.")])
  }, 120_000)

  test("each underlying problem becomes exactly one open item", async () => {
    const directory = matter()
    const model = scriptedModel({ [AGENT.draft]: drafter(DEFINED), [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: sameSession(undefine), [AGENT.finalize]: draftingNotes })
    const state = await runLean(directory, { sessions: model.port })

    expect(state.status).toBe("completed")
    const items = readOpenItems(directory)
    // The twice-cited clause is one item, and the lost definition is not
    // reported again as a term used without a definition.
    expect(items.filter(item => item.kind === "consistency").map(item => item.detail)).toEqual([
      "Clause 3.1 refers to clause 2.2, but no clause in the final text is numbered 2.2. What was clause 2.2 before the changes is now clause 2.3.",
      'The term "Specification" is used in the final text (first in clause 1.1) but is no longer defined: before the changes it was defined in clause 1.2.',
    ])
    expect(items.filter(item => item.kind !== "consistency" && /clause 2\.2|"Specification"/.test(item.detail))).toEqual([])
  }, 120_000)
})
