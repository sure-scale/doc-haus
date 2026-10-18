import { afterEach, describe, expect, test } from "bun:test"
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

  test("a recorded quotation field still needs the verbatim words", () => {
    const directory = matter(TERMS)
    expect(quoteFound(directory, { quote: "The supplier shall deliver the goods \u2026 Payment is due in full on delivery", source_id: SOURCE })).toBe(false)
  })
})
