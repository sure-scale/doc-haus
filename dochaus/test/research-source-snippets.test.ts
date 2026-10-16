import { describe, expect, test } from "bun:test"
import { displayedSourceSnippets, researchTerms, renderSourceSnippets, sourceSnippets } from "../lib/research-source-snippets"

describe("research source snippets", () => {
  test("splits human-entered terms without inventing new ones", () => {
    expect(researchTerms("assignment; client consent | fiduciary duty\narbitration")).toEqual([
      "assignment",
      "client consent",
      "fiduciary duty",
      "arbitration",
    ])
  })

  test("accepts array-shaped terms from tool calls", () => {
    expect(researchTerms(["assignment", "client consent"] as never)).toEqual(["assignment", "client consent"])
  })

  test("returns bounded snippets around requested terms", () => {
    const text = `Introductory navigation. ${"x ".repeat(600)} The operative source text says client consent is required before assignment. ${"y ".repeat(600)}`
    const snippets = sourceSnippets(text, ["client consent"], 2)
    expect(snippets).toHaveLength(1)
    expect(snippets[0].text).toContain("client consent is required")
    expect(snippets[0].text.length).toBeLessThan(2_000)
  })

  test("keeps complete boundary words and clips only tokens cut by the existing window", () => {
    const complete = "Not every request requires approval."
    for (const terms of [["request"], ["absent"]]) {
      expect(displayedSourceSnippets(complete, terms, 1)).toEqual([
        { term: terms[0] === "absent" ? "first-readable-text" : "request", text: complete, start: 0, end: complete.length },
      ])
    }
    const text = `${"abcdefghij ".repeat(150)}needle ${"klmnopqrst ".repeat(150)}`
    const at = text.indexOf("needle")
    const windowStart = at - 900
    const windowEnd = at + "needle".length + 900
    const snippet = sourceSnippets(text, ["needle"], 1)[0]!
    expect(snippet.start).toBeGreaterThan(windowStart)
    expect(snippet.end).toBeLessThan(windowEnd)
    expect(text[snippet.start - 1]).toMatch(/\s/)
    expect(text[snippet.end]).toMatch(/\s/)
    expect(text.slice(snippet.start, snippet.end)).toBe(snippet.text)
    expect(snippet.text).toContain("needle")
    expect(snippet.end - snippet.start).toBeLessThanOrEqual(windowEnd - windowStart)
  })

  test("renders XML-style source boundaries and missing terms", () => {
    const output = renderSourceSnippets({
      url: "https://example.gov/rule",
      contentType: "text/html",
      text: "The rule discusses termination rights.",
      terms: ["termination", "assignment"],
      maxSnippets: 4,
    })
    expect(output).toContain('<official_source url="https://example.gov/rule"')
    expect(output).toContain('<snippet index="1" matched="termination" locator="snippet-1">')
    expect(output).toContain("<term>assignment</term>")
  })

  test("covers every visible term in a shared snippet despite the display cap", () => {
    const output = renderSourceSnippets({
      url: "https://example.gov/source",
      contentType: "text/plain",
      text: "Amber includes cobalt and silver.",
      terms: ["amber", "cobalt", "silver", "absent"],
      maxSnippets: 1,
    })
    expect(output.match(/<snippet index=/g)).toHaveLength(1)
    expect(output).toContain("<missing_terms><term>absent</term></missing_terms>")
  })

  test("overlapping snippets do not make their other visible terms missing", () => {
    const output = renderSourceSnippets({
      url: "https://example.gov/source",
      contentType: "text/plain",
      text: `${Array.from({ length: 300 }, (_, index) => `prefix${index}`).join(" ")} amber ${"x ".repeat(40)} cobalt ${"y ".repeat(40)} silver ${"z ".repeat(40)} violet ${"tail ".repeat(300)}`,
      terms: ["amber", "cobalt", "silver", "violet"],
      maxSnippets: 2,
    })
    expect(output.match(/<snippet index=/g)).toHaveLength(2)
    expect(output).toContain("<missing_terms></missing_terms>")
  })

  test("matches visible phrases across case, spacing and compatible Unicode without accepting partial phrases", () => {
    const output = renderSourceSnippets({
      url: "https://example.gov/source",
      contentType: "text/plain",
      text: "ＡＰＰＲＯＶＡＬ\n\tprocess needs WRITTEN\u00a0NOTICE.",
      terms: ["approval process", "written notice", "approval policy"],
      maxSnippets: 1,
    })
    expect(output).toContain("<missing_terms><term>approval policy</term></missing_terms>")
  })

  test("keeps terms outside the capped displayed text missing even when they occur elsewhere in the source", () => {
    const output = renderSourceSnippets({
      url: "https://example.gov/source",
      contentType: "text/plain",
      text: `Amber record begins. ${"filler ".repeat(300)}Violet record ends.`,
      terms: ["amber", "violet"],
      maxSnippets: 1,
    })
    expect(output.match(/<snippet index=/g)).toHaveLength(1)
    expect(output).not.toContain("Violet record ends.")
    expect(output).toContain("<missing_terms><term>violet</term></missing_terms>")
  })
})
