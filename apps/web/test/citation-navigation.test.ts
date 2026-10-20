import { describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import CitationView from "../src/components/CitationView"
import { citationDocument, citationPassage } from "../src/components/citation-target"
import type { Citation } from "../src/api/opencode/parts"

const document = { id: 1, name: "terms.docx", doc_path: "/matters/current/terms.docx", created_at: 1 }
const citation: Citation = {
  documentName: document.name, docPath: document.doc_path, section: "Fees", excerpt: "Payment is due on delivery.",
  charStart: 700, charEnd: 730, score: 1,
}

describe("source document resolution", () => {
  test("uses only a registered source in the current matter", () => {
    expect(citationDocument(citation, [document])).toBe(document)
    expect(citationDocument({ ...citation, docPath: document.name }, [document])).toBe(document)
    for (const docPath of ["/matters/other/terms.docx", "../terms.docx", "https://example.com/terms.docx", "file:///matters/current/terms.docx"]) {
      expect(citationDocument({ ...citation, docPath }, [document])).toBeUndefined()
    }
    expect(citationDocument(citation, [])).toBeUndefined()
    expect(citationDocument(citation, [document, document])).toBeUndefined()
    expect(citationDocument({ ...citation, documentName: "other.docx" }, [document])).toBeUndefined()
  })

  test("does not hand unsupported formats or path-like names to the viewer", () => {
    for (const name of ["terms.md", "notes.msg", "../terms.docx", "sub/terms.docx", "sub\\terms.docx", "sub/instructions.eml"]) {
      expect(citationDocument({ ...citation, documentName: name }, [{ ...document, name }])).toBeUndefined()
    }
    expect(citationDocument({ ...citation, documentName: "terms.PDF" }, [{ ...document, name: "terms.PDF" }])?.name).toBe("terms.PDF")
  })

  test("opens emails and spreadsheets, which ingest indexes and the assistant cites", () => {
    for (const name of ["Client instructions.eml", "Cap table.XLSX"]) {
      expect(citationDocument({ ...citation, documentName: name, docPath: name }, [{ ...document, name }])?.name).toBe(name)
    }
  })
})

describe("rendered citation controls", () => {
  test("renders a native keyboard button with a descriptive source and section", () => {
    const html = renderToStaticMarkup(createElement(CitationView, { citations: [citation], matterId: "current", documents: [document] }))
    expect(html).toContain('<button type="button" class="ref" aria-label="Open terms.docx, section Fees"')
    expect(html).toContain(citation.excerpt)
    expect(html).not.toContain(document.doc_path)
  })

  test("unresolvable citations remain readable and visibly unavailable", () => {
    for (const props of [{ documents: [] }, { documents: [document], matterId: undefined }]) {
      const html = renderToStaticMarkup(createElement(CitationView, { citations: [citation], matterId: "current", ...props }))
      expect(html).toContain("Source unavailable in this matter")
      expect(html).toContain(citation.excerpt)
      expect(html).not.toContain("<button")
    }
  })

  test("same-named sources at different paths are not deduplicated together", () => {
    const html = renderToStaticMarkup(createElement(CitationView, { citations: [citation, { ...citation, docPath: "/matters/other/terms.docx" }], matterId: "current", documents: [document] }))
    expect(html).toContain("Sources (2)")
    expect(html.match(/<button/g)?.length).toBe(1)
    expect(html).toContain("Source unavailable in this matter")
  })
})

describe("citation passage matching", () => {
  test("maps normalized whitespace back to rendered text rather than using ingest offsets", () => {
    const text = "\nBefore\nPayment\u00a0 is due\n on delivery.\nAfter"
    const result = citationPassage(text, citation.excerpt)
    expect(result.status).toBe("matched")
    if (result.status !== "matched") throw new Error("Expected match")
    expect(text.slice(result.start, result.end)).toBe("Payment\u00a0 is due\n on delivery.")
  })

  test("missing, empty and duplicate excerpts do not invent a precise location", () => {
    expect(citationPassage("Changed current document", citation.excerpt)).toEqual({ status: "unavailable" })
    expect(citationPassage("anything", " \n ")).toEqual({ status: "unavailable" })
    expect(citationPassage(`${citation.excerpt}\n${citation.excerpt}`, citation.excerpt)).toEqual({ status: "ambiguous" })
  })
})
