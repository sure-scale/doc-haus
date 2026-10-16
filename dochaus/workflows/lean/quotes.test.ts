import { afterEach, describe, expect, test } from "bun:test"
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
