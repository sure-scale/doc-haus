import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
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

// The retry feedback for a quotation with an ellipsis in the middle: the words
// it leaves out ("near"), passages too far apart to read as one ("apart"), or
// the plain miss when a passage is not in the source.
const NEAR_TAIL = "Leaving words out can change who must do what or drop a condition, amount or limit, so quotations with words left out in the middle are not accepted. Quote one continuous passage that includes those words, without an ellipsis."
const near = (quote: string, omitted: string, source = SOURCE, where = "The drafting notes") =>
  `${where} quotes "${quote}" as from "${source}". The quoted words are in ${source === "request" ? "the request, a user update or one of the lawyer's answers" : "that document"}, but the "…" leaves out: "${omitted}". ${NEAR_TAIL}`
const apart = (quote: string, source = SOURCE) =>
  `The drafting notes quotes "${quote}" as from "${source}". The quoted parts are in that document, but in separate passages. Quote each passage as its own quotation, without an ellipsis.`
const miss = (quote: string, source = SOURCE) =>
  `The drafting notes quotes "${quote}" as from "${source}", but those words are not in that document. Copy the exact words with read-document (or check them with verify-quote), or drop the quotation marks and paraphrase.`
const short = (quote: string, source = SOURCE) =>
  `The drafting notes quotes "${quote}" as from "${source}". A passage around the "…" is too short to locate on its own. Quote one continuous passage without an ellipsis, or drop the quotation marks.`
const notes = (quote: string, source = SOURCE) => `# Drafting notes\n\nThe terms provide "${quote}" [source: ${source}].`
const issues = (directory: string, quote: string, source = SOURCE) => checkDraftingNotes(directory, notes(quote, source), []).issues

// A prose quotation may begin or end with an ellipsis: the passage left must be
// in the named source, long enough to verify on its own and cut on whole words.
// An ellipsis in the middle is refused, and the miss says what it leaves out.
describe("lean elided quotation matching", () => {
  const TERMS = "The supplier shall deliver the goods within thirty days of the order date. Late delivery entitles the buyer to a credit. Payment is due in full on delivery of the goods."
  const OMITTED = "within thirty days of the order date. Late delivery entitles the buyer to a credit."

  test("passages of one source joined by an ellipsis are refused with the words it leaves out", () => {
    const directory = matter(TERMS)
    for (const quote of ["The supplier shall deliver the goods … Payment is due in full on delivery", "The supplier shall deliver the goods ... Payment is due in full on delivery"])
      expect(issues(directory, quote)).toEqual([near(quote, OMITTED)])
    // The lawyer's words are located without regard to case, as quoteFound
    // matches them, and the miss shows them as written.
    for (const quote of ["Draft a short supply … from the supply terms", "draft a short supply … from the supply terms"])
      expect(issues(directory, quote, "request")).toEqual([near(quote, "note", "request")])
    // An ellipsis between passages that meet leaves nothing out.
    expect(issues(directory, "Draft a short supply note … from the supply terms", "request")).toEqual([
      `The drafting notes quotes "Draft a short supply note … from the supply terms" as from "request". The quoted words are one continuous passage of the request, a user update or one of the lawyer's answers; the "…" leaves nothing out. Quote the passage without an ellipsis.`,
    ])
  })

  test("the continuous passage verifies, as does each passage quoted on its own; deleting the ellipsis does not", () => {
    const directory = matter(TERMS)
    expect(issues(directory, `The supplier shall deliver the goods ${OMITTED} Payment is due in full on delivery`)).toEqual([])
    expect(checkDraftingNotes(directory, `# Drafting notes\n\nThe terms provide "The supplier shall deliver the goods" [source: ${SOURCE}] and "Payment is due in full on delivery" [source: ${SOURCE}].`, []).issues).toEqual([])
    const joined = "The supplier shall deliver the goods Payment is due in full on delivery"
    expect(issues(directory, joined)).toEqual([miss(joined)])
  })

  test("an ellipsis at the start or end verifies the passage it leaves", () => {
    const directory = matter(TERMS)
    for (const quote of ["... entitles the buyer to a credit. Payment is due in full", "The supplier shall deliver the goods within thirty days …",
      "… Late delivery entitles the buyer to a credit …", "[…] Late delivery entitles the buyer to a credit"])
      expect([quote, issues(directory, quote)]).toEqual([quote, []])
    for (const quote of ["Draft a short supply note …", "… a short supply note from the supply terms"])
      expect([quote, issues(directory, quote, "request")]).toEqual([quote, []])
    // Punctuation after an edge ellipsis ends the writer's sentence, not a passage.
    for (const quote of ["The supplier shall deliver the goods within thirty days ….", "… Late delivery entitles the buyer to a credit […]."])
      expect([quote, issues(directory, quote)]).toEqual([quote, []])
    // An ellipsis in the middle refuses the quotation even beside edge ellipses.
    const edged = "… shall deliver the goods … Payment is due in full on delivery …"
    expect(issues(directory, edged)).toEqual([near(edged, OMITTED)])
  })

  test("reordered or absent passages are still refused with the plain miss", () => {
    const directory = matter(TERMS)
    for (const quote of [
      "Payment is due in full on delivery … The supplier shall deliver the goods",
      "The supplier shall deliver the goods … Payment is due within sixty days",
    ]) expect(issues(directory, quote)).toEqual([miss(quote)])
    expect(issues(directory, "from the supply terms … Draft a short supply note", "request")).toEqual([
      `The drafting notes quotes "from the supply terms … Draft a short supply note" as from "request", but those words are not in the request, user updates or any one of the lawyer's answers. Copy the exact words with read-document (or check them with verify-quote), or drop the quotation marks and paraphrase.`,
    ])
  })

  test("a passage too short to locate on its own is refused as too short, not as missing", () => {
    const directory = matter(TERMS)
    for (const quote of ["The supplier shall deliver the goods … Payment is due", "The supplier … shall deliver the goods within thirty days",
      "The supplier shall deliver the goods . . . Payment is due", "… entitles the buyer", "Late delivery …", "Late delivery […]."])
      expect([quote, issues(directory, quote)]).toEqual([quote, [short(quote)]])
    expect(issues(directory, "Draft a … supply note", "request")).toEqual([short("Draft a … supply note", "request")])
    expect(quoteMissIssue(directory, "findings[0].quote", { quote: "Late delivery … a credit", source_id: SOURCE }, "json")).toBe(
      `findings[0].quote quotes "Late delivery … a credit" as from "${SOURCE}". A passage around the "…" is too short to locate on its own. The JSON quote field requires verbatim source words. Copy them with read-document or check them with verify-quote. Record paraphrases outside source-quote entries; removing JSON string delimiters does not make a paraphrase a valid quote.`)
  })

  test("a source that writes its own ellipsis differently asks for the exact characters", () => {
    for (const [written, quoted] of [["...", "…"], ["…", "..."], ["[...]", "…"]]) {
      const directory = matter(`The extract reads: The supplier shall deliver the goods ${written} Payment is due in full on delivery.`)
      const quote = `The supplier shall deliver the goods ${quoted} Payment is due in full on delivery`
      expect([written, issues(directory, quote)]).toEqual([written, [
        `The drafting notes quotes "${quote}" as from "${SOURCE}". The quoted words are in that document, but it writes the ellipsis as "${written}". Copy the exact characters, including that ellipsis.`,
      ]])
      expect([written, issues(directory, `The supplier shall deliver the goods ${written} Payment is due in full on delivery`)]).toEqual([written, []])
    }
  })

  test("a bound playbook refuses an ellipsis in the middle and verifies one at the start or end", () => {
    const workspace = process.env.WORKSPACE_ROOT
    const root = mkdtempSync(path.join(tmpdir(), "quote-playbooks-"))
    try {
      process.env.WORKSPACE_ROOT = root
      mkdirSync(path.join(root, ".playbooks", "playbook-supply-terms"), { recursive: true })
      writeFileSync(path.join(root, ".playbooks", "playbook-supply-terms", "SKILL.md"), TERMS)
      const directory = matter("Unrelated matter text.")
      const file = path.join(directory, "matter.json")
      writeFileSync(file, JSON.stringify({ ...(existsSync(file) ? JSON.parse(readFileSync(file, "utf8")) : {}), playbook: "playbook-supply-terms" }))
      const source = "playbook:playbook-supply-terms"
      const quote = "The supplier shall deliver the goods … Payment is due in full on delivery"
      expect(issues(directory, quote, source)).toEqual([near(quote, OMITTED, source)])
      expect(issues(directory, "… Payment is due in full on delivery", source)).toEqual([])
    } finally {
      if (workspace === undefined) delete process.env.WORKSPACE_ROOT
      else process.env.WORKSPACE_ROOT = workspace
      rmSync(root, { recursive: true, force: true })
    }
  })

  test("spaced, bracketed and sentence-ending ellipses separate passages under the same rules", () => {
    const directory = matter(TERMS)
    for (const mark of [" . . . ", " […] ", " [...] ", " [. . .] ", ". . . . "]) {
      const quote = `The supplier shall deliver the goods${mark}Payment is due in full on delivery`
      expect(issues(directory, quote)).toEqual([near(quote, OMITTED)])
    }
    expect(issues(directory, "Payment is due in full on delivery [...] The supplier shall deliver the goods")).toEqual([miss("Payment is due in full on delivery [...] The supplier shall deliver the goods")])
  })

  test("a passage must begin after the previous one ends and come from the same source", () => {
    const directory = matter(TERMS)
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    const schedule = "Late delivery entitles the buyer to a service credit of five percent."
    db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("delivery-schedule.docx", "1", schedule, 0, schedule.length)
    db.close()
    for (const [quote, source] of [
      ["The supplier shall deliver the goods within … deliver the goods within thirty days", SOURCE],
      ["The supplier shall deliver the goods … entitles the buyer to a service credit", SOURCE],
      ["The supplier shall deliver the goods … entitles the buyer to a service credit", "delivery-schedule.docx"],
    ] as const) expect(issues(directory, quote, source)).toEqual([miss(quote, source)])
    const quote = "Late delivery entitles the buyer … a service credit of five percent"
    expect(issues(directory, quote, "delivery-schedule.docx")).toEqual([near(quote, "to", "delivery-schedule.docx")])
  })

  test("a recorded quotation field still needs the verbatim words; its miss locates an elision too", () => {
    const directory = matter(TERMS)
    const quote = "The supplier shall deliver the goods … Payment is due in full on delivery"
    expect(quoteFound(directory, { quote, source_id: SOURCE })).toBe(false)
    expect(quoteMissIssue(directory, "findings[0].quote", { quote, source_id: SOURCE }, "json")).toBe(near(quote, OMITTED, SOURCE, "findings[0].quote"))
  })
})

// An ellipsis in the middle is refused whatever it leaves out: no list of words
// can tell which omissions change who must do what or drop a condition, amount
// or limit. The miss locates the passages so the author can quote one
// continuous passage, or each passage on its own.
describe("lean elided quotation refusals", () => {
  const LIMITS = [
    "The Supplier shall be liable for direct losses only and shall not be liable for indirect or consequential losses arising from late delivery.",
    "The Seller hereby agrees that it shall not indemnify the Buyer for third-party claims of any kind.",
    "The licence granted under this Agreement is non-exclusive and non-transferable.",
    "The Customer's liability under this Agreement is unlimited in amount for breach of confidentiality.",
  ].join("\n\n")
  const PURCHASE = (omitted: string) => `The Distributor shall purchase the goods ${omitted} from the Manufacturer at the list price.`
  const PURCHASE_QUOTE = "The Distributor shall purchase the goods … from the Manufacturer at the list price"
  const NEUTRAL = "in the quantities stated in each purchase order and on the delivery dates agreed in writing by both parties"

  test("a passage that begins or ends inside a word of the source is not found", () => {
    const directory = matter(LIMITS)
    for (const quote of ["The licence granted under this Agreement is … exclusive and non-transferable.",
      "The Customer's liability under this Agreement is …limited in amount for breach of confidentiality.",
      "The licence granted under this Agreement is non … and non-transferable."])
      expect(issues(directory, quote)).toEqual([miss(quote)])
    const whole = "The licence granted under this Agreement is … and non-transferable."
    expect(issues(directory, whole)).toEqual([near(whole, "non-exclusive")])
    // The start of a passage is checked on its own: "ir" is no prefix.
    const revocable = "The licence granted under this Agreement is … revocable and perpetual within the Territory"
    expect(issues(matter("The licence granted under this Agreement is irrevocable and perpetual within the Territory."), revocable)).toEqual([miss(revocable)])
    // Word text carries a non-breaking hyphen (U+2011) where the page shows "-";
    // a passage locates with either hyphen.
    const nonBreaking = matter("The licence granted under this Agreement is non‑exclusive and is limited to the Territory.")
    for (const hyphen of ["‑", "-"]) {
      const quote = `The licence granted under this Agreement is non${hyphen}exclusive … limited to the Territory`
      expect(issues(nonBreaking, quote)).toEqual([near(quote, "and is")])
    }
  })

  test("an ellipsis at the start or end may not cut a word at a hyphen or dash", () => {
    const licence = (grant: string) => matter(`The Licensor grants the Licensee ${grant} licence to use the Software in the Territory.`)
    const lead = "… exclusive licence to use the Software in the Territory"
    const trail = "The Licensor grants the Licensee a non …"
    // Hyphen, Word's hyphen and non-breaking hyphen, figure dash, en dash, minus sign.
    for (const dash of ["-", "‐", "‑", "‒", "–", "−"]) {
      const directory = licence(`a non${dash}exclusive`)
      expect([dash, issues(directory, lead), issues(directory, trail)]).toEqual([dash, [miss(lead)], [miss(trail)]])
      expect([dash, issues(directory, "… a non-exclusive licence to use the Software …")]).toEqual([dash, []])
    }
    expect(issues(licence("an exclusive"), lead)).toEqual([])
    // The em dash sets off an aside and joins no word.
    expect(issues(licence("a non—exclusive"), lead)).toEqual([])
  })

  test("an ellipsis at the start or end may not cut a word at an apostrophe", () => {
    for (const apostrophe of ["'", "’"]) {
      const directory = matter(`The obligations of the Customer${apostrophe}s affiliates under this Agreement end on termination.`)
      for (const quote of ["… s affiliates under this Agreement end on termination", "The obligations of the Customer …"])
        expect([apostrophe, quote, issues(directory, quote)]).toEqual([apostrophe, quote, [miss(quote)]])
      expect([apostrophe, issues(directory, "… affiliates under this Agreement end on termination")]).toEqual([apostrophe, []])
    }
  })

  test("an ellipsis at the start or end may not cut a number at its decimal point or thousands separator", () => {
    const directory = matter("The service credit is 2.5 per cent of the Charges. The fee payable under this Agreement is £1,000,000 per annum payable in advance. The number of licensed copies is 2. Payment is due on delivery.")
    for (const quote of ["… 5 per cent of the Charges", "The service credit is 2 …", "The service credit is 2. …", "The fee payable under this Agreement is £1 …",
      "… 000 per annum payable in advance", "… 000,000 per annum payable in advance"])
      expect([quote, issues(directory, quote)]).toEqual([quote, [miss(quote)]])
    for (const quote of ["… 2.5 per cent of the Charges", "The fee payable under this Agreement is £1,000,000 …", "The number of licensed copies is 2 …",
      "The number of licensed copies is 2. …"])
      expect([quote, issues(directory, quote)]).toEqual([quote, []])
  })

  test("an omission is refused whatever it leaves out, and the miss shows it", () => {
    const cap = (amount: string) => `The Supplier's aggregate liability under this Agreement is capped at ${amount} the Charges paid in the twelve months before the claim.`
    const CAP_QUOTE = "The Supplier's aggregate liability under this Agreement is capped at … the Charges paid in the twelve months before the claim"
    const support = (omitted: string) => `The Supplier shall provide ${omitted} ten Business Days of support each month to the Customer at the Site.`
    const SUPPORT_QUOTE = "The Supplier shall provide … ten Business Days of support each month to the Customer at the Site"
    for (const [source, quote, omitted] of [
      ...["not", "unless rejected", "up to £500", NEUTRAL, "listed in schedule No. 4"].map(omitted => [PURCHASE(omitted), PURCHASE_QUOTE, omitted]),
      [cap("the greater of £5,000,000 and"), CAP_QUOTE, "the greater of £5,000,000 and"],
      ...["at least", "the Customer with"].map(omitted => [support(omitted), SUPPORT_QUOTE, omitted]),
      [LIMITS, "The Supplier shall be liable … for indirect or consequential losses", "for direct losses only and shall not be liable"],
      [LIMITS, "The Seller hereby agrees that it shall … indemnify the Buyer for third-party claims", "not"],
      ["The Customer under this Agreement may request additional services from time to time. The Supplier shall provide the additional services within ten days.",
        "The Customer under this Agreement may … provide the additional services within ten days", "request additional services from time to time. The Supplier shall"],
      ["The Buyer may inspect the goods on delivery and the Seller shall replace them. Notice of a defect is given in writing.",
        "The Buyer may inspect the goods … Notice of a defect is given in writing", "on delivery and the Seller shall replace them."],
    ]) expect([omitted, issues(matter(source), quote)]).toEqual([omitted, [near(quote, omitted)]])
  })

  test("an omission that hands an obligation to another party is refused, and the miss shows the hand-off", () => {
    for (const [source, quote, omitted] of [
      ["Under this Agreement the Supplier shall deliver the Goods to the Site and the Customer shall pay the Price within thirty days of delivery.",
        "Under this Agreement the Supplier … shall pay the Price within thirty days of delivery", "shall deliver the Goods to the Site and the Customer"],
      ["The Secretary of State may make regulations under this section and the Authority must publish guidance on the operation of the scheme.",
        "The Secretary of State may make regulations … must publish guidance on the operation of the scheme", "under this section and the Authority"],
      ["The Landlord may enter the Premises to inspect them and the Tenant shall make good any damage within twenty working days.",
        "The Landlord may enter the Premises … shall make good any damage within twenty working days", "to inspect them and the Tenant"],
      ...["shall at its own cost", "shall (at its own cost)", "may elect to"].map(tail => [
        `The Buyer may inspect the goods on delivery and the Seller ${tail} replace any defective goods within ten days of notice.`,
        "The Buyer may inspect the goods … replace any defective goods within ten days of notice", `on delivery and the Seller ${tail}`]),
    ]) expect([omitted, issues(matter(source), quote)]).toEqual([omitted, [near(quote, omitted)]])
  })

  test("a harmless aside is refused the same way; the passage quoted whole verifies", () => {
    for (const [source, quote, omitted] of [
      ["The Customer shall pay each valid invoice within 30 days of receipt (or such longer period as the parties may agree in writing) by bank transfer to the account notified by the Supplier.",
        "The Customer shall pay each valid invoice within 30 days of receipt … by bank transfer to the account notified by the Supplier", "(or such longer period as the parties may agree in writing)"],
      ["The Share Purchase Agreement dated 14 May 2023 between the Sellers and the Buyer provides that the price is payable in cash on Completion.",
        "The Share Purchase Agreement … provides that the price is payable in cash on Completion", "dated 14 May 2023 between the Sellers and the Buyer"],
      ["The Customer shall reimburse the Supplier for its expenses (if any) incurred in performing the Services at the Site.",
        "The Customer shall reimburse the Supplier for its expenses … incurred in performing the Services at the Site", "(if any)"],
    ]) {
      const directory = matter(source)
      expect([omitted, issues(directory, quote)]).toEqual([omitted, [near(quote, omitted)]])
      expect([omitted, issues(directory, source.slice(0, -1))]).toEqual([omitted, []])
    }
  })

  test("a figure dash, en dash or minus sign locates as \"-\" does; an em dash is an aside", () => {
    const licence = (prefix: string) => `The licence granted under this Agreement is ${prefix}exclusive and is limited to the Territory.`
    const quote = "The licence granted under this Agreement is … exclusive and is limited to the Territory"
    for (const prefix of ["non‒", "non−"])
      expect([prefix, issues(matter(licence(prefix)), quote)]).toEqual([prefix, [miss(quote)]])
    const aside = "The Supplier shall deliver the goods … to the Site at its own cost"
    expect(issues(matter("The Supplier shall deliver the goods—and any packaging—to the Site at its own cost."), aside))
      .toEqual([near(aside, "—and any packaging—")])
    // A range keeps its dash on both sides of the comparison.
    const range = "Clauses 1–5 of Schedule 2 apply … clauses 6–9 of Schedule 2 apply to the Goods"
    expect(issues(matter("Clauses 1–5 of Schedule 2 apply to the Services and clauses 6–9 of Schedule 2 apply to the Goods."), range))
      .toEqual([near(range, "to the Services and")])
  })

  // Two clauses whose sub-clauses restart at the same label.
  const LABELS = {
    lower: (clause: number, item: number) => `(${"abc"[item - 1]})`,
    upper: (clause: number, item: number) => `(${["III", "IV"][item - 1]})`,
    letter: (clause: number, item: number) => `(${"AB"[item - 1]})`,
    decimal: (clause: number, item: number) => `4.${clause}.${item}`,
  }
  const clauses = (gap: string, label: (clause: number, item: number) => string) =>
    [`4.1 The Supplier shall, at its own cost:`, `${label(1, 1)} deliver the goods to the Site;`, `${label(1, 2)} install the goods at the Site.`,
      `4.2 The Customer shall:`, `${label(2, 1)} pay the Charges within 30 days of invoice.`].join(gap) + gap

  test("list items that Word writes as separate paragraphs are separate passages", () => {
    // Converted Word text ends every paragraph, each list item included, with a blank line.
    const items = matter("The Supplier shall, at its own cost:\n\n(a) deliver the goods to the Site;\n\n(b) install the goods; and\n\n(c) invoice the Customer monthly in arrears.\n\n")
    for (const quote of ["The Supplier shall, at its own cost: … (c) invoice the Customer monthly in arrears", "The Supplier shall, at its own cost: … invoice the Customer monthly in arrears"])
      expect(issues(items, quote)).toEqual([apart(quote)])
    expect(checkDraftingNotes(items, `# Drafting notes\n\nThe terms provide "The Supplier shall, at its own cost:" [source: ${SOURCE}] and "(c) invoice the Customer monthly in arrears" [source: ${SOURCE}].`, []).issues).toEqual([])
    for (const [name, label] of Object.entries(LABELS)) {
      const quote = `The Supplier shall, at its own cost: … ${label(1, 2)} install the goods at the Site`
      expect([name, issues(matter(clauses("\n\n", label)), quote)]).toEqual([name, [apart(quote)]])
    }
  })

  test("an omission that restarts a list for another party is refused with the words it leaves out", () => {
    for (const gap of [" ", "\n"]) for (const [name, label] of Object.entries(LABELS)) {
      const quote = `The Supplier shall, at its own cost: … ${label(2, 1)} pay the Charges within 30 days of invoice`
      const omitted = `${label(1, 1)} deliver the goods to the Site; ${label(1, 2)} install the goods at the Site. 4.2 The Customer shall:`
      expect([JSON.stringify(gap), name, issues(matter(clauses(gap, label)), quote)]).toEqual([JSON.stringify(gap), name, [near(quote, omitted)]])
    }
  })

  test("passages far apart or in separate paragraphs are refused as apart, without the words between", () => {
    const long = Array.from({ length: 4 }, () => NEUTRAL).join(", ")
    expect(long.length).toBeGreaterThan(400)
    expect(issues(matter(PURCHASE(long)), PURCHASE_QUOTE)).toEqual([apart(PURCHASE_QUOTE)])
    expect(issues(matter(PURCHASE(`${NEUTRAL}, ${NEUTRAL}`)), PURCHASE_QUOTE)).toEqual([near(PURCHASE_QUOTE, `${NEUTRAL}, ${NEUTRAL}`)])
    const paragraphs = (gap: string) => `The Distributor shall purchase the goods each month.${gap}The price for the goods is the list price on the order date.`
    const quote = "The Distributor shall purchase the goods … The price for the goods is the list price"
    for (const gap of ["\n\n", "\r\n \r\n"]) expect(issues(matter(paragraphs(gap)), quote)).toEqual([apart(quote)])
    expect(issues(matter(paragraphs("\n")), quote)).toEqual([near(quote, "each month.")])
  })

  test("the nearest occurrence of a passage is located", () => {
    const far = Array.from({ length: 4 }, () => NEUTRAL).join(", ")
    const source = `The Distributor shall purchase the goods on the first order date, ${far}. ${PURCHASE("each month")}`
    expect(issues(matter(source), PURCHASE_QUOTE)).toEqual([near(PURCHASE_QUOTE, "each month")])
    expect(issues(matter(source.replace(PURCHASE("each month"), "The goods come from the Manufacturer at the list price.")), PURCHASE_QUOTE)).toEqual([apart(PURCHASE_QUOTE)])
    // Of two near placements, the one leaving out least, whichever passage repeats.
    expect(issues(matter(`The Distributor shall purchase the goods each month. ${PURCHASE("in bulk")}`), PURCHASE_QUOTE)).toEqual([near(PURCHASE_QUOTE, "in bulk")])
    expect(issues(matter(`${PURCHASE("each month")} ${PURCHASE("in bulk")}`), PURCHASE_QUOTE)).toEqual([near(PURCHASE_QUOTE, "in bulk")])
  })
})
