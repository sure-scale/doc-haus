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

  test("an omission holding a limit is refused", () => {
    const support = (omitted: string) => `The Supplier shall provide ${omitted} ten Business Days of support each month to the Customer at the Site.`
    const quote = "The Supplier shall provide … ten Business Days of support each month to the Customer at the Site"
    for (const omitted of ["up to", "at least", "a maximum of", "a minimum of"])
      expect([omitted, verified(matter(support(omitted)), quote)]).toEqual([omitted, false])
    expect(verified(matter(support("the Customer with")), quote)).toBe(true)
  })

  test("an omitted \"(if any)\" conditions nothing; any other \"if\" still does", () => {
    const expenses = (omitted: string) => `The Customer shall reimburse the Supplier for its expenses ${omitted} incurred in performing the Services at the Site.`
    const quote = "The Customer shall reimburse the Supplier for its expenses … incurred in performing the Services at the Site"
    expect(verified(matter(expenses("(if any)")), quote)).toBe(true)
    expect(verified(matter(expenses("(If Any)")), quote)).toBe(true)
    for (const omitted of [", if any,", "(if any goods are rejected)", "(if any) or (if agreed)"])
      expect([omitted, verified(matter(expenses(omitted)), quote)]).toEqual([omitted, false])
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

  test("an omission holding a currency code or a number grouped in thousands is refused", () => {
    const cap = (amount: string) => `The Supplier's aggregate liability under this Agreement is capped at ${amount} the Charges paid in the twelve months before the claim.`
    const quote = "The Supplier's aggregate liability under this Agreement is capped at … the Charges paid in the twelve months before the claim"
    for (const amount of ["the greater of GBP 500000 and", "the greater of USD1000 and", "the greater of 5,000,000 pounds and", "the greater of 25,000 and"])
      expect([amount, verified(matter(cap(amount)), quote)]).toEqual([amount, false])
    // A currency code is upper case: a short word before a number is no amount.
    expect(verified(matter(cap("the amount, for 30 days, of")), quote)).toBe(true)
  })

  test("an omission that hands a sentence's verb to another subject or modal is refused", () => {
    expect(verified(matter("The Customer under this Agreement may request additional services from time to time. The Supplier shall provide the additional services within ten days."),
      "The Customer under this Agreement may … provide the additional services within ten days")).toBe(false)
    expect(verified(matter("The Buyer may inspect the goods on delivery and the Seller shall replace any defective goods within ten days of notice."),
      "The Buyer may inspect the goods … replace any defective goods within ten days of notice")).toBe(false)
    // The omission may end on the modal and an adverb, a comma-set aside or a list label.
    for (const tail of ["shall promptly", "shall, at its own cost,", "shall (a)"])
      expect([tail, verified(matter(`The Buyer may inspect the goods on delivery and the Seller ${tail} replace any defective goods within ten days of notice.`),
        "The Buyer may inspect the goods … replace any defective goods within ten days of notice")]).toEqual([tail, false])
    // A modal may be omitted where the next passage begins a sentence of its own.
    expect(verified(matter("The Buyer may inspect the goods on delivery and the Seller shall replace them. Notice of a defect is given in writing."),
      "The Buyer may inspect the goods … Notice of a defect is given in writing")).toBe(true)
  })

  test("a modal earlier in the omission, in an aside or a date, hands over nothing", () => {
    for (const [source, quote] of [
      ["The Customer shall pay each valid invoice within 30 days of receipt (or such longer period as the parties may agree in writing) by bank transfer to the account notified by the Supplier.",
        "The Customer shall pay each valid invoice within 30 days of receipt … by bank transfer to the account notified by the Supplier"],
      ["The Supplier shall provide the Customer with such information about the Services as the Customer may reasonably require for the purposes of its audit.",
        "The Supplier shall provide the Customer with such information about the Services … for the purposes of its audit"],
      ["\"Confidential Information\" means all information disclosed by one party to the other, which shall include know-how and trade secrets, whether before or after the date of this Agreement.",
        "\"Confidential Information\" means all information disclosed by one party to the other … whether before or after the date of this Agreement"],
      ["A notice under subsection (1) must be in writing, must be in such form as may be prescribed, and must be served on every person with an interest in the land.",
        "A notice under subsection (1) must be in writing … and must be served on every person with an interest in the land"],
      ["The Share Purchase Agreement dated 14 May 2023 between the Sellers and the Buyer provides that the price is payable in cash on Completion.",
        "The Share Purchase Agreement … provides that the price is payable in cash on Completion"],
    ] as const) expect([quote, verified(matter(source), quote)]).toEqual([quote, true])
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

  test("a figure dash, en dash or minus sign stands for \"-\" as a hyphen does", () => {
    const licence = (prefix: string) => `The licence granted under this Agreement is ${prefix}exclusive and is limited to the Territory.`
    const quote = "The licence granted under this Agreement is … exclusive and is limited to the Territory"
    for (const prefix of ["non\u2012", "non\u2212", "non \u2013 ", "non - ", "un \u2013"])
      expect([prefix, verified(matter(licence(prefix)), quote)]).toEqual([prefix, false])
    expect(verified(matter(licence("non\u2013")), "… exclusive and is limited to the Territory")).toBe(false)
    // The em dash sets off an aside and joins no word.
    expect(verified(matter("The Supplier shall deliver the goods\u2014and any packaging\u2014to the Site at its own cost."),
      "The Supplier shall deliver the goods … to the Site at its own cost")).toBe(true)
    // A range keeps its dash on both sides of the comparison.
    expect(verified(matter("Clauses 1\u20135 of Schedule 2 apply to the Services and clauses 6\u20139 of Schedule 2 apply to the Goods."),
      "Clauses 1\u20135 of Schedule 2 apply … clauses 6\u20139 of Schedule 2 apply to the Goods")).toBe(true)
  })

  test("an omission may skip list items that Word writes as separate paragraphs", () => {
    // Converted Word text ends every paragraph, each list item included, with a blank line.
    const items = "The Supplier shall, at its own cost:\n\n(a) deliver the goods to the Site;\n\n(b) install the goods; and\n\n(c) invoice the Customer monthly in arrears.\n\n"
    expect(verified(matter(items), "The Supplier shall, at its own cost: … (c) invoice the Customer monthly in arrears")).toBe(true)
    // Without a list label to resume at, the omission still may not cross a paragraph.
    expect(verified(matter(items), "The Supplier shall, at its own cost: … invoice the Customer monthly in arrears")).toBe(false)
  })

  // Two clauses whose sub-clauses restart at the same label, written as Word
  // paragraphs, on one line, or one per line as in PDF text.
  const LABELS = {
    lower: (clause: number, item: number) => `(${"abc"[item - 1]})`,
    upper: (clause: number, item: number) => `(${["III", "IV"][item - 1]})`,
    letter: (clause: number, item: number) => `(${"AB"[item - 1]})`,
    decimal: (clause: number, item: number) => `4.${clause}.${item}`,
  }
  const clauses = (gap: string, label: (clause: number, item: number) => string) =>
    [`4.1 The Supplier shall, at its own cost:`, `${label(1, 1)} deliver the goods to the Site;`, `${label(1, 2)} install the goods at the Site.`,
      `4.2 The Customer shall:`, `${label(2, 1)} pay the Charges within 30 days of invoice.`].join(gap) + gap

  test("an omission may skip typed upper-case, roman and decimal list items", () => {
    for (const [name, label] of Object.entries(LABELS))
      expect([name, verified(matter(clauses("\n\n", label)), `The Supplier shall, at its own cost: … ${label(1, 2)} install the goods at the Site`)]).toEqual([name, true])
  })

  test("an omission may not hand a restarted list item to another party's modal", () => {
    for (const gap of ["\n\n", " ", "\n"]) for (const [name, label] of Object.entries(LABELS))
      expect([JSON.stringify(gap), name, verified(matter(clauses(gap, label)), `The Supplier shall, at its own cost: … ${label(2, 1)} pay the Charges within 30 days of invoice`)])
        .toEqual([JSON.stringify(gap), name, false])
    expect(verified(matter("The Customer may request additional services. The Supplier shall (a) provide the additional services within ten days."),
      "The Customer may request additional services. … (a) provide the additional services within ten days")).toBe(false)
  })

  test("a list label resumes a passage across a blank line only where it begins a line", () => {
    for (const verb of ["may reject", "rejects"])
      expect([verb, verified(matter(`The Supplier shall deliver the goods to the Site.\n\nThe Customer ${verb} goods that fail the tests in paragraph 4(c) within ten days of delivery.`),
        "The Supplier shall deliver the goods to the Site. … (c) within ten days of delivery")]).toEqual([verb, false])
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
