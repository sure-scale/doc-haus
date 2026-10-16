// Optional-clause mechanics against an authored in-memory fixture — the operations
// draft-document runs (omit via deleteSection, keep via the marker-strip fill),
// exercised directly on a Docxodus session so no tool ctx is needed.
// Run from dochaus/: bun test templates/optional.test.ts
import { describe, expect, test } from "bun:test"
import { docxodus } from "../lib/docxodus"
import { renumberHeadings } from "../lib/markdown"
import { renderMarkdownToDocx } from "../lib/render"

// No generated demo document or client template is required on a fresh checkout.
// This minimal synthetic text exercises section/placeholder mechanics only.
const fixtureMarkdown = [
  "# Synthetic optional-clause example",
  ...Array.from({ length: 7 }, (_, index) => `## ${index + 1}. Example clause\n\nSynthetic clause text.`),
  "## 8. Governing Law\n\n[insert governing jurisdiction]",
  "## 9. Non-Solicitation [optional: non-solicitation]\n\nDo not solicit for employment for [insert non-solicit period in months] months.",
  "## 10. Entire Agreement\n\nSynthetic closing clause.",
].join("\n\n")
const { bytes: fixtureBytes } = await renderMarkdownToDocx(fixtureMarkdown)
const dx = await docxodus()

describe("optional clause markers", () => {
  test("template exposes the marker as an alternative_clause placeholder and its fill as an instruction", () => {
    const session = dx.openDocxSession(fixtureBytes, {})
    const placeholders = session.findPlaceholders()
    session.close()
    expect(
      placeholders.some((p) => p.kind === "alternative_clause" && p.match.text === "[optional: non-solicitation]"),
    ).toBe(true)
    expect(
      placeholders.some((p) => p.kind === "instruction" && p.match.text === "[insert non-solicit period in months]"),
    ).toBe(true)
  })

  test("omit removes the whole clause and leaves the rest intact", () => {
    const session = dx.openDocxSession(fixtureBytes, {})
    const anchor = session.findByText("[optional: non-solicitation]", { ignoreWhitespace: true })!
    expect(anchor.kind).toBe("h")
    session.deleteSection(anchor.id)
    // The markdown projection escapes punctuation ("Non\-Solicitation");
    // normalize so the assertions match the readable text.
    const text = session.project().markdown.replaceAll("\\", "")
    const placeholders = session.findPlaceholders().map((p) => p.match.text)
    session.close()
    expect(text).not.toContain("Non-Solicitation")
    expect(text).not.toContain("solicit for employment")
    expect(text).toContain("Entire Agreement")
    expect(placeholders).not.toContain("[insert non-solicit period in months]")
  })

  test("omit then renumber closes the heading-number hole", () => {
    const session = dx.openDocxSession(fixtureBytes, {})
    const anchor = session.findByText("[optional: non-solicitation]", { ignoreWhitespace: true })!
    session.deleteSection(anchor.id)
    renumberHeadings(session)
    const text = session.project().markdown.replaceAll("\\", "")
    session.close()
    expect(text).toContain("9. Entire Agreement")
    expect(text).not.toContain("10. Entire Agreement")
    expect(text).toContain("8. Governing Law")
  })

  test("keep strips the marker but leaves the clause", () => {
    const session = dx.openDocxSession(fixtureBytes, {})
    session.fillPlaceholders((p) => (p.match.text.startsWith("[optional:") ? "" : null), {
      kinds: dx.PlaceholderKinds.AlternativeClause,
      coalesceWhitespaceAroundEmptyFill: true,
    })
    const text = session.project().markdown.replaceAll("\\", "")
    const placeholders = session.findPlaceholders().map((p) => p.match.text)
    session.close()
    expect(text).toContain("Non-Solicitation")
    expect(text).not.toContain("[optional:")
    expect(placeholders.filter((t) => t.startsWith("[optional:"))).toEqual([])
    expect(placeholders).toContain("[insert non-solicit period in months]")
  })
})
