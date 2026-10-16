import { describe, expect, test } from "bun:test"
import { formatCitations, type DocumentCitation } from "../lib/citations"

describe("citation rendering", () => {
  test("clips long displayed excerpts while keeping citation context", () => {
    const citation: DocumentCitation = {
      documentName: "agreement.docx",
      docPath: "/tmp/agreement.docx",
      section: "Section 1",
      excerpt: "A".repeat(1400),
      charStart: 0,
      charEnd: 1400,
    }

    const rendered = formatCitations([citation])
    expect(rendered).toContain("[agreement.docx § Section 1]")
    expect(rendered).toContain("passage preview clipped at 1000 characters")
    expect(rendered).not.toContain("A".repeat(1200))
  })
})
