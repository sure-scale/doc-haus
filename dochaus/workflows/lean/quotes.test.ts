import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import path from "node:path"
import { checkDraftingNotes } from "./checks"
import { quoteFound, quoteMissIssue } from "./quotes"
import { SOURCE, cleanupMatters, matter } from "./test-support"

// Quotation matching against indexed source text: exact words, tolerant only of
// Markdown escapes in the source text and of curly versus straight quotes.

afterEach(cleanupMatters)

// Not String.raw: the transpiler may re-escape its non-ASCII characters.
const ESCAPED = "Clause 4\\.2 applies to \\*all\\* goods. The buyer\u2019s \u201Cacceptance\u201D is final."

describe("lean quotation matching", () => {
  test("an exact quotation matches", () => {
    const directory = matter()
    expect(quoteFound(directory, { quote: "within thirty days of the order date", source_id: SOURCE })).toBe(true)
  })

  test("a quotation matches a source whose text carries Markdown escapes and curly quotes", () => {
    const directory = matter(ESCAPED)
    expect(quoteFound(directory, { quote: "Clause 4.2 applies to *all* goods.", source_id: SOURCE })).toBe(true)
    expect(quoteFound(directory, { quote: `The buyer's "acceptance" is final.`, source_id: SOURCE })).toBe(true)
  })

  test("a misquotation does not match and the issue names the quotation and the document", () => {
    const directory = matter(ESCAPED)
    expect(quoteFound(directory, { quote: "Clause 4.2 applies to some goods.", source_id: SOURCE })).toBe(false)
    expect(quoteFound(directory, { quote: "Clause 4.2 applies to *all* goods.", source_id: "other.docx" })).toBe(false)
    const issue = quoteMissIssue(directory, "amendments[0] ([p1]).sources", { quote: "Clause 4.2 applies to some goods.", source_id: SOURCE })
    expect(issue).toContain(`quotes "Clause 4.2 applies to some goods." as from "${SOURCE}", but those words are not in that document.`)
  })
})

// A prose quotation may join passages of one source with an ellipsis. It is
// verified only when every passage is in that source, in the quoted order, and
// each is long enough to verify on its own; otherwise it is refused as a miss.
describe("lean elided quotation matching", () => {
  const TERMS = "The supplier shall deliver the goods within thirty days of the order date. Late delivery entitles the buyer to a credit. Payment is due in full on delivery of the goods."
  const notes = (quote: string, source = SOURCE) => `# Drafting notes\n\nThe terms provide "${quote}" [source: ${source}].`
  const issues = (directory: string, quote: string, source = SOURCE) => checkDraftingNotes(directory, notes(quote, source), []).issues

  test("passages of one source joined by an ellipsis, in order, are verified", () => {
    const directory = matter(TERMS)
    expect(issues(directory, "The supplier shall deliver the goods \u2026 Payment is due in full on delivery")).toEqual([])
    expect(issues(directory, "The supplier shall deliver the goods ... Payment is due in full on delivery")).toEqual([])
    expect(issues(directory, "... entitles the buyer to a credit. Payment is due in full")).toEqual([])
    expect(issues(directory, "Draft a short supply note \u2026 from the supply terms", "request")).toEqual([])
    // The lawyer's words are matched without regard to case, as quoteFound does.
    expect(issues(directory, "draft a short supply note \u2026 from the supply terms", "request")).toEqual([])
  })

  test("reordered, too short or absent passages are still refused with the plain miss", () => {
    const directory = matter(TERMS)
    const miss = (quote: string) => `The drafting notes quotes "${quote}" as from "${SOURCE}", but those words are not in that document.`
    for (const quote of [
      "Payment is due in full on delivery \u2026 The supplier shall deliver the goods",
      "The supplier shall deliver the goods \u2026 Payment is due",
      "The supplier shall deliver the goods \u2026 Payment is due within sixty days",
    ]) {
      const found = issues(directory, quote)
      expect(found).toHaveLength(1)
      expect(found[0]).toContain(miss(quote))
    }
    expect(issues(directory, "from the supply terms \u2026 Draft a short supply note", "request")).toHaveLength(1)
  })

  test("spaced, bracketed and sentence-ending ellipses separate passages under the same rules", () => {
    const directory = matter(TERMS)
    for (const mark of [" . . . ", " [\u2026] ", " [...] ", " [. . .] ", ". . . . "])
      expect(issues(directory, `The supplier shall deliver the goods${mark}Payment is due in full on delivery`)).toEqual([])
    expect(issues(directory, "The supplier shall deliver the goods . . . Payment is due")).toHaveLength(1)
    expect(issues(directory, "Payment is due in full on delivery [...] The supplier shall deliver the goods")).toHaveLength(1)
  })

  test("a passage must begin after the previous one ends and come from the same source", () => {
    const directory = matter(TERMS)
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    const schedule = "Late delivery entitles the buyer to a service credit of five percent."
    db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("delivery-schedule.docx", "1", schedule, 0, schedule.length)
    db.close()
    expect(issues(directory, "The supplier shall deliver the goods within \u2026 deliver the goods within thirty days")).toHaveLength(1)
    expect(issues(directory, "The supplier shall deliver the goods \u2026 entitles the buyer to a service credit")).toHaveLength(1)
    expect(issues(directory, "The supplier shall deliver the goods \u2026 entitles the buyer to a service credit", "delivery-schedule.docx")).toHaveLength(1)
    expect(issues(directory, "Late delivery entitles the buyer \u2026 a service credit of five percent", "delivery-schedule.docx")).toEqual([])
  })

  test("a recorded quotation field still needs the verbatim words", () => {
    const directory = matter(TERMS)
    expect(quoteFound(directory, { quote: "The supplier shall deliver the goods \u2026 Payment is due in full on delivery", source_id: SOURCE })).toBe(false)
  })
})

// An ellipsis may only shorten a passage, never turn it into words the source
// does not say: each passage begins and ends on whole words of the source, the
// omitted words hold no negation or qualification, and each omission stays a
// short run within one paragraph.
describe("lean elided quotation limits", () => {
  const LIMITS = [
    "The Supplier shall be liable for direct losses only and shall not be liable for indirect or consequential losses arising from late delivery.",
    "The Seller hereby agrees that it shall not indemnify the Buyer for third-party claims of any kind.",
    "The licence granted under this Agreement is non-exclusive and non-transferable.",
    "The Customer's liability under this Agreement is unlimited in amount for breach of confidentiality.",
  ].join("\n\n")
  const verified = (directory: string, quote: string) =>
    checkDraftingNotes(directory, `# Drafting notes\n\nThe terms provide "${quote}" [source: ${SOURCE}].`, []).issues.length === 0
  const PURCHASE = (omitted: string) => `The Distributor shall purchase the goods ${omitted} from the Manufacturer at the list price.`
  const PURCHASE_QUOTE = "The Distributor shall purchase the goods … from the Manufacturer at the list price"
  const NEUTRAL = "in the quantities stated in each purchase order and on the delivery dates agreed in writing by both parties"

  test("a passage that begins or ends inside a word of the source is refused", () => {
    const directory = matter(LIMITS)
    expect(verified(directory, "The licence granted under this Agreement is … exclusive and non-transferable.")).toBe(false)
    expect(verified(directory, "The Customer's liability under this Agreement is …limited in amount for breach of confidentiality.")).toBe(false)
    expect(verified(directory, "The licence granted under this Agreement is non … and non-transferable.")).toBe(false)
    expect(verified(directory, "The licence granted under this Agreement is … and non-transferable.")).toBe(true)
    // The start of a passage is checked on its own: "ir" is no prefix the
    // omission rule knows.
    expect(verified(matter("The licence granted under this Agreement is irrevocable and perpetual within the Territory."),
      "The licence granted under this Agreement is … revocable and perpetual within the Territory")).toBe(false)
    // Word text carries a non-breaking hyphen (U+2011) where the page shows "-";
    // a faithful quotation verifies with either hyphen.
    const nonBreaking = matter("The licence granted under this Agreement is non‑exclusive and is limited to the Territory.")
    expect(verified(nonBreaking, "… exclusive and is limited to the Territory")).toBe(false)
    for (const hyphen of ["‑", "-"])
      expect(verified(nonBreaking, `The licence granted under this Agreement is non${hyphen}exclusive … limited to the Territory`)).toBe(true)
  })

  test("an omission holding a negation or qualification is refused", () => {
    const directory = matter(LIMITS)
    expect(verified(directory, "The Supplier shall be liable … for indirect or consequential losses")).toBe(false)
    expect(verified(directory, "The Seller hereby agrees that it shall … indemnify the Buyer for third-party claims")).toBe(false)
    for (const omitted of ["not", "never", "only", "in no event", "nor", "neither", "none of", "without delay", "except samples", "unless rejected",
      "save as agreed", "subject to clause 4", "provided that stock lasts", "notwithstanding clause 4", "cannot", "won't",
      "other than samples", "excluding samples", "apart from samples", "otherwise than in bulk", "rather than samples", "solely", "exclusively",
      "if stock is available", "to the extent stock is available", "in the event stock is available"])
      expect([omitted, verified(matter(PURCHASE(omitted)), PURCHASE_QUOTE)]).toEqual([omitted, false])
    expect(verified(matter(PURCHASE(NEUTRAL)), PURCHASE_QUOTE)).toBe(true)
    // Whole words only: "gift" and "different" hold no "if".
    expect(verified(matter(PURCHASE("packed in gift boxes for the different regions")), PURCHASE_QUOTE)).toBe(true)
    // "No." before a number abbreviates "number"; it negates nothing.
    expect(verified(matter(PURCHASE("listed in schedule No. 4")), PURCHASE_QUOTE)).toBe(true)
  })

  test("an omission holding an amount is refused", () => {
    for (const omitted of ["less 5%", "less 5 %", "less 5 per cent", "less 5 percent", "up to £500", "up to $ 500", "up to €500", "up to ¥500"])
      expect([omitted, verified(matter(PURCHASE(omitted)), PURCHASE_QUOTE)]).toEqual([omitted, false])
    const cap = (amount: string) => `The Supplier's aggregate liability under this Agreement is capped at ${amount} the Charges paid in the twelve months before the claim.`
    const quote = "The Supplier's aggregate liability under this Agreement is capped at … the Charges paid in the twelve months before the claim"
    expect(verified(matter(cap("150% of")), quote)).toBe(false)
    expect(verified(matter(cap("the greater of £5,000,000 and")), quote)).toBe(false)
    expect(verified(matter(cap("the amount of")), quote)).toBe(true)
  })

  test("an omission that hands a sentence's verb to another subject or modal is refused", () => {
    expect(verified(matter("The Customer under this Agreement may request additional services from time to time. The Supplier shall provide the additional services within ten days."),
      "The Customer under this Agreement may … provide the additional services within ten days")).toBe(false)
    expect(verified(matter("The Buyer may inspect the goods on delivery and the Seller shall replace any defective goods within ten days of notice."),
      "The Buyer may inspect the goods … replace any defective goods within ten days of notice")).toBe(false)
    // A modal may be omitted where the next passage begins a sentence of its own.
    expect(verified(matter("The Buyer may inspect the goods on delivery and the Seller shall replace them. Notice of a defect is given in writing."),
      "The Buyer may inspect the goods … Notice of a defect is given in writing")).toBe(true)
  })

  test("an omission that drops a non- or un- prefix from the next word is refused", () => {
    for (const source of [
      "The licence granted under this Agreement is non exclusive and is limited to the Territory.",
      "The licence granted under this Agreement is non-\nexclusive and is limited to the Territory.",
      "The licence granted under this Agreement is un- conditional and is limited to the Territory.",
      // Word's non-breaking hyphen and hyphen, and an en dash, stand for "-".
      "The licence granted under this Agreement is non‑exclusive and is limited to the Territory.",
      "The licence granted under this Agreement is non‐exclusive and is limited to the Territory.",
      "The licence granted under this Agreement is non–exclusive and is limited to the Territory.",
    ]) expect([source, verified(matter(source), `The licence granted under this Agreement is … ${source.includes("un-") ? "conditional" : "exclusive"} and is limited to the Territory`)]).toEqual([source, false])
  })

  test("an omission may skip list items that Word writes as separate paragraphs", () => {
    // Converted Word text ends every paragraph, each list item included, with a blank line.
    const items = "The Supplier shall, at its own cost:\n\n(a) deliver the goods to the Site;\n\n(b) install the goods; and\n\n(c) invoice the Customer monthly in arrears.\n\n"
    expect(verified(matter(items), "The Supplier shall, at its own cost: … (c) invoice the Customer monthly in arrears")).toBe(true)
    // Without a list label to resume at, the omission still may not cross a paragraph.
    expect(verified(matter(items), "The Supplier shall, at its own cost: … invoice the Customer monthly in arrears")).toBe(false)
  })

  test("an omission longer than a short run or crossing a paragraph break is refused", () => {
    const long = Array.from({ length: 4 }, () => NEUTRAL).join(", ")
    expect(long.length).toBeGreaterThan(300)
    expect(verified(matter(PURCHASE(long)), PURCHASE_QUOTE)).toBe(false)
    expect(verified(matter(PURCHASE(`${NEUTRAL}, ${NEUTRAL}`)), PURCHASE_QUOTE)).toBe(true)
    const paragraphs = (gap: string) => `The Distributor shall purchase the goods each month.${gap}The price for the goods is the list price on the order date.`
    const quote = "The Distributor shall purchase the goods … The price for the goods is the list price"
    expect(verified(matter(paragraphs("\n\n")), quote)).toBe(false)
    expect(verified(matter(paragraphs("\r\n \r\n")), quote)).toBe(false)
    expect(verified(matter(paragraphs("\n")), quote)).toBe(true)
  })

  test("a later occurrence of a passage is tried when the first leaves too long an omission", () => {
    const far = Array.from({ length: 4 }, () => NEUTRAL).join(", ")
    const source = `The Distributor shall purchase the goods on the first order date, ${far}. ${PURCHASE("each month")}`
    expect(verified(matter(source), PURCHASE_QUOTE)).toBe(true)
    expect(verified(matter(source.replace(PURCHASE("each month"), "The goods come from the Manufacturer at the list price.")), PURCHASE_QUOTE)).toBe(false)
  })
})
