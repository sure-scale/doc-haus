import { describe, expect, test } from "bun:test"
import { distinctCitations } from "../src/components/CitationView"
import { citationDocument, citationPassage } from "../src/components/citation-target"
import type { Citation } from "../src/api/opencode/parts"
import { CLIENT_SUMMARY } from "../src/agents"
import { readTurn, toTurns } from "../src/components/chat/turn-model"

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
    for (const name of ["notes.mbox", "../terms.docx", "sub/terms.docx", "sub\\terms.docx", "sub/instructions.eml"]) {
      expect(citationDocument({ ...citation, documentName: name }, [{ ...document, name }])).toBeUndefined()
    }
    expect(citationDocument({ ...citation, documentName: "terms.PDF" }, [{ ...document, name: "terms.PDF" }])?.name).toBe("terms.PDF")
  })

  test("opens emails and spreadsheets, which ingest indexes and the assistant cites", () => {
    for (const name of ["Client instructions.eml", "Forwarded thread.msg", "Cap table.XLSX", "Chronology.md", "Record.MARKDOWN"]) {
      expect(citationDocument({ ...citation, documentName: name, docPath: name }, [{ ...document, name }])?.name).toBe(name)
    }
  })
})

describe("cited sources of an answer", () => {
  const quoted = (section: string, excerpt: string): Citation => ({ documentName: document.name, docPath: document.doc_path, section, excerpt, charStart: 0, charEnd: excerpt.length })
  const tool = (id: string, name: string, citations: Citation[] = []) => ({
    type: "tool" as const, id, sessionID: "ses_1", messageID: "msg_a", tool: name,
    state: { status: "completed" as const, input: {}, output: "", title: name, metadata: { citations } },
  })
  const parts = [
    tool("t1", "get-section"),
    tool("t2", "cite", [quoted("4.2", "The Customer may reject defective cartons within 5 days.")]),
    tool("t3", "cite", [quoted("7", "Either party may end this agreement on 30 days' notice.")]),
    { type: "text" as const, id: "t4", sessionID: "ses_1", messageID: "msg_a", text: "What this is [terms.docx § 4.2] [terms.docx § 7]" },
  ]

  test("quotations cite verified are the answer's Sources, settled and live", () => {
    const settled = toTurns([
      { info: { id: "msg_u", role: "user", agent: CLIENT_SUMMARY }, parts: [{ type: "text", id: "u", sessionID: "ses_1", messageID: "msg_u", text: "Explain terms.docx for the client." }] },
      { info: { id: "msg_a", role: "assistant" }, parts },
    ] as never, "")[1]!
    expect(settled.citations.map((c) => c.section)).toEqual(["4.2", "7"])
    const live = readTurn(new Map(parts.map((p) => [p.id, p])) as never, new Map([["msg_a", "assistant"]]), "", "ses_1", CLIENT_SUMMARY)
    expect(live.citations.map((c) => c.section)).toEqual(["4.2", "7"])
  })

  test("same-named sources at different paths are not deduplicated together", () => {
    const other = { ...citation, docPath: "/matters/other/terms.docx" }
    expect(distinctCitations([citation, other])).toHaveLength(2)
    expect(citationDocument(other, [document])).toBeUndefined()
  })

  test("for one section, verified quotations outrank search hits and distinct quotations stay listed", () => {
    const hit = { ...citation, section: "4.2", excerpt: "4.2 Rejection. The whole clause as retrieved.", score: 0.9 }
    const quote = quoted("4.2", "The Customer may reject defective cartons within 5 days.")
    for (const order of [[hit, quote], [quote, hit]]) {
      expect(distinctCitations(order).map((c) => c.excerpt)).toEqual([quote.excerpt])
    }

    const preambleHit = { ...citation, section: "Preamble", excerpt: "Whole preamble as retrieved.", score: 0.9 }
    const first = quoted("Preamble", "Gate 3 opens at 06:15 on weekdays.")
    const second = quoted("Preamble", "Synthetic test memo.")
    expect(distinctCitations([preambleHit, first, second, second]).map((c) => c.excerpt)).toEqual([first.excerpt, second.excerpt])
  })
})

describe("citation passage matching", () => {
  test("canonical text coordinates identify a repeated quotation and never reanchor a displaced range", () => {
    const text = `${citation.excerpt}\nA later report follows.\n${citation.excerpt}`
    const start = text.lastIndexOf(citation.excerpt)
    expect(citationPassage(text, citation.excerpt, { charStart: start, charEnd: text.length })).toEqual({ status: "matched", start, end: text.length })
    expect(citationPassage(text, citation.excerpt, { charStart: start - 1, charEnd: text.length - 1 })).toEqual({ status: "unavailable" })
    expect(citationPassage(text, citation.excerpt)).toEqual({ status: "ambiguous" })
  })
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
