import { afterEach, describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { INGEST_URL } from "../src/config"
import { compareUrl, fetchBlacklineBytes, fetchComparison, type ChangeHunk, type ComparisonJson } from "../src/api/ingest/compare"
import { choiceParams, compareCandidates } from "../src/components/CompareDialog"
import CompareDialog from "../src/components/CompareDialog"
import { ComparePanel, hunkLocation } from "../src/components/BlacklineViewer"
import CompareView from "../src/components/chat/CompareView"
import { toTurns } from "../src/components/chat/turn-model"
import DocumentUpload from "../src/components/DocumentUpload"
import DocumentViewer from "../src/components/DocumentViewer"
import type { Document } from "../src/api/ingest/documents"
import type { Part } from "../src/api/session-view"

const originalFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = originalFetch
})

const params = { left: "Our draft.docx", right: "Their draft (2).docx" }

describe("compare client", () => {
  test("builds the route with labels, every option flag and the format", () => {
    expect(compareUrl("m-1", { ...params, leftLabel: " Ours ", rightLabel: "Counterparty" })).toBe(
      `${INGEST_URL}/matters/m-1/compare?left=Our%20draft.docx&right=Their%20draft%20(2).docx&leftLabel=Ours&rightLabel=Counterparty&ignoreFormatting=1&moves=1&headers=1&format=docx`,
    )
    expect(compareUrl("m-1", { ...params, leftLabel: "  ", ignoreFormatting: false, detectMoves: false, headersFooters: false }, { format: "json" })).toBe(
      `${INGEST_URL}/matters/m-1/compare?left=Our%20draft.docx&right=Their%20draft%20(2).docx&ignoreFormatting=0&moves=0&headers=0&format=json`,
    )
    expect(compareUrl("m-1", params, { download: true })).toEndWith("&format=docx&download=1")
  })

  test("reads the comparison JSON and the blackline bytes from the same route", async () => {
    const urls: string[] = []
    globalThis.fetch = (async (input: string | URL | Request) => {
      urls.push(String(input))
      return String(input).includes("format=json") ? Response.json(comparison()) : new Response(new Uint8Array([80, 75, 3, 4]))
    }) as typeof fetch
    expect((await fetchComparison("m-1", params)).blacklineName).toBe("Blackline - Our draft vs Their draft.docx")
    expect(await fetchBlacklineBytes("m-1", params)).toEqual(new Uint8Array([80, 75, 3, 4]))
    expect(urls.map((url) => new URL(url).searchParams.get("format"))).toEqual(["json", "docx"])
  })

  test("a refused compare throws the route's plain message, never its code", async () => {
    globalThis.fetch = (async () =>
      Response.json({ error: "not_docx", message: "Scan.pdf is a PDF. Convert it to DOCX in the viewer first." }, { status: 400 })) as unknown as typeof fetch
    await expect(fetchComparison("m-1", params)).rejects.toThrow("Scan.pdf is a PDF. Convert it to DOCX in the viewer first.")
    await expect(fetchBlacklineBytes("m-1", params)).rejects.toThrow("Scan.pdf is a PDF. Convert it to DOCX in the viewer first.")
    globalThis.fetch = (async () => new Response("upstream failure", { status: 502 })) as unknown as typeof fetch
    await expect(fetchComparison("m-1", params)).rejects.toThrow("Could not compare the documents (502)")
    globalThis.fetch = (async () => Response.json({ error: "not_found" }, { status: 404 })) as unknown as typeof fetch
    await expect(fetchBlacklineBytes("m-1", params)).rejects.toThrow("Could not compare the documents (404)")
  })
})

const doc = (id: number, name: string, role: Document["role"] = "source"): Document => ({ id, name, doc_path: name, created_at: id * 1000, role })

describe("compare dialog", () => {
  const options = { ignoreFormatting: true, detectMoves: true, headersFooters: true }

  test("this file is the earlier version by default; swap moves it and its label to the later side", () => {
    const choice = { name: "Our draft.docx", other: "Their draft.docx", thisLabel: "Our draft", otherLabel: "Their draft", options }
    expect(choiceParams({ ...choice, thisSide: "left" })).toEqual({
      left: "Our draft.docx", right: "Their draft.docx", leftLabel: "Our draft", rightLabel: "Their draft", ...options,
    })
    expect(choiceParams({ ...choice, thisSide: "right" })).toEqual({
      left: "Their draft.docx", right: "Our draft.docx", leftLabel: "Their draft", rightLabel: "Our draft", ...options,
    })
  })

  test("unticked options reach the route as 0 and a cleared label falls back to the file name", () => {
    const next = choiceParams({
      name: "Our draft.docx", other: "Their draft.docx", thisSide: "left", thisLabel: " ", otherLabel: " Counterparty ",
      options: { ignoreFormatting: false, detectMoves: true, headersFooters: false },
    })
    expect(next).toEqual({ left: "Our draft.docx", right: "Their draft.docx", rightLabel: "Counterparty", ignoreFormatting: false, detectMoves: true, headersFooters: false })
    expect(new URL(compareUrl("m-1", next)).search).toBe(
      "?left=Our%20draft.docx&right=Their%20draft.docx&rightLabel=Counterparty&ignoreFormatting=0&moves=1&headers=0&format=docx",
    )
  })

  test("offers the matter's other Word files, sources and deliverables, newest first", () => {
    const documents = [doc(1, "Our draft.docx"), doc(2, "Scan.pdf"), doc(3, "Supply agreement - draft.docx", "deliverable"), doc(4, "Their draft.DOCX")]
    expect(compareCandidates(documents, "Our draft.docx").map((d) => d.name)).toEqual(["Their draft.DOCX", "Supply agreement - draft.docx"])
    const html = renderToStaticMarkup(createElement(CompareDialog, {
      matterId: "m-1", name: "Our draft.docx", documents, onUploaded: () => {}, onCompare: () => {}, onClose: () => {},
    }))
    expect(html.indexOf("Earlier version")).toBeLessThan(html.indexOf("Our draft.docx"))
    expect(html.indexOf("Our draft.docx")).toBeLessThan(html.indexOf("Later version"))
    expect([...html.matchAll(/<option value="([^"]+)"/g)].map(([, name]) => name)).toEqual(["Their draft.DOCX", "Supply agreement - draft.docx"])
    expect(html).toContain("Upload a file…")
    expect(html).toContain('accept=".docx"')
    for (const label of ["Ignore formatting-only changes", "Detect moved text", "Include headers and footers"]) expect(html).toContain(label)
    expect([...html.matchAll(/type="checkbox"[^>]*checked=""/g)]).toHaveLength(3)
    expect(html).toContain("author")
  })
})

const hunk = (id: string, untracked: boolean, patch: Partial<ChangeHunk> = {}): ChangeHunk => ({
  id, kind: "modified", untracked, location: { scope: "body", paragraph: 4, side: "right", heading: "2. Payment", number: "2.1" },
  before: "Payment is due within thirty days.", after: "Payment is due within sixty days.", ...patch,
})

function comparison(patch: Partial<ComparisonJson> = {}, right: Partial<ComparisonJson["right"]> = {}): ComparisonJson {
  return {
    left: { name: "Our draft.docx", label: "Our draft", trackedChanges: { count: 0, authors: [] }, pendingProposals: 0 },
    right: { name: "Their draft.docx", label: "Their draft", trackedChanges: { count: 3, authors: ["Pat Reviewer"] }, pendingProposals: 0, ...right },
    options: { ignoreFormatting: true, detectMoves: true, headersFooters: true },
    attribution: "authors",
    changes: [hunk("C1", false), hunk("C2", true, { kind: "inserted", before: undefined, after: "A new <b>late fee</b> applies." })],
    untracked: [hunk("U1", true, { kind: "inserted", before: undefined, after: "A new <b>late fee</b> applies." })],
    blacklineName: "Blackline - Our draft vs Their draft.docx",
    ...patch,
  }
}
const panel = (value: ComparisonJson) => renderToStaticMarkup(createElement(ComparePanel, { comparison: value }))

describe("blackline side panel", () => {
  test("untracked edits in a file with tracked changes come first, as a warning", () => {
    const html = panel(comparison())
    expect(html).toContain("compare-untracked is-warning")
    expect(html).toContain("1 edit in Their draft was made without tracked changes")
    expect(html.indexOf("Edits made without tracked changes")).toBeLessThan(html.indexOf("All changes"))
    expect(html.indexOf("U1")).toBeLessThan(html.indexOf("C1"))
    expect(html).toContain("Not tracked")
    expect(html).toContain("A new &lt;b&gt;late fee&lt;/b&gt; applies.")
    expect(html).not.toContain("<b>late fee</b>")
    expect(html).toContain("Their draft had 3 tracked changes (by Pat Reviewer); compared as accepted.")
    expect(html).not.toContain("Our draft had")
  })

  test("a later file without tracked changes says every difference was untracked, without a warning", () => {
    const html = panel(comparison({}, { trackedChanges: { count: 0, authors: [] } }))
    expect(html).toContain("compare-untracked is-info")
    expect(html).toContain("Their draft has no tracked changes, so every difference below was made without tracking.")
    expect(html).not.toContain("is-warning")
  })

  test("tracked changes with nothing outside them reassure instead of warn", () => {
    const html = panel(comparison({ untracked: [], changes: [hunk("C1", false)] }))
    expect(html).toContain("compare-untracked is-ok")
    expect(html).toContain("No edits outside the tracked changes.")
    expect(html).not.toContain("Not tracked")
  })

  test("the label fallback and inputs' own revisions and proposals are disclosed", () => {
    const html = panel(comparison({
      attribution: "label",
      left: { name: "Our draft.docx", label: "Our draft", trackedChanges: { count: 1, authors: [] }, pendingProposals: 2 },
    }))
    expect(html).toContain("Untracked edits are listed here only; in the Word file every change shows under Their draft.")
    expect(html).toContain("Our draft had 1 tracked change; compared as accepted.")
    expect(html).toContain("Our draft has 2 pending doc.haus proposals; they are not included.")
    expect(panel(comparison())).toContain("show under “Not tracked (Their draft)”")
  })

  test("locations read as heading, clause number or paragraph, and long text is clipped with a way to expand", () => {
    expect(hunkLocation(hunk("C1", false).location, comparison())).toBe("2. Payment · 2.1")
    expect(hunkLocation({ scope: "footer", paragraph: 2, side: "left" }, comparison())).toBe("Footer · paragraph 2 in Our draft")
    const html = panel(comparison({ untracked: [hunk("U1", true, { after: "word ".repeat(200) })] }))
    expect(html).toContain("…")
    expect(html).toContain("Show all")
  })
})

describe("chat compare card", () => {
  const meta = {
    compare: {
      left: "Our draft.docx", right: "Their draft.docx", leftLabel: "Our draft", rightLabel: "Their draft",
      options: { ignoreFormatting: true, detectMoves: false, headersFooters: true }, counts: { changes: 4, untracked: 1 },
    },
  }
  const turn = (metadata: Record<string, unknown>, status: "completed" | "running" = "completed") => {
    const part: Part = {
      type: "tool", id: "call", sessionID: "session", messageID: "msg", tool: "compare-versions",
      state: status === "completed" ? { status, input: {}, title: "Compared", output: "{}", metadata } : { status, input: {}, title: "Comparing" },
    }
    return toTurns([{ info: { id: "msg", role: "assistant" }, parts: [part] }], "")[0]!
  }

  test("reads the tool's compare metadata as the blackline to open", () => {
    expect(turn(meta).compares).toEqual([{
      left: "Our draft.docx", right: "Their draft.docx", leftLabel: "Our draft", rightLabel: "Their draft",
      ignoreFormatting: true, detectMoves: false, headersFooters: true, changes: 4, untracked: 1,
    }])
    expect(turn({ compare: { left: "Our draft.docx" } }).compares).toEqual([])
    expect(turn(meta, "running").compares).toEqual([])
  })

  test("the card names both files and both counts and opens the blackline", () => {
    const opened: unknown[] = []
    const html = renderToStaticMarkup(createElement(CompareView, { compares: turn(meta).compares!, onOpen: (value) => opened.push(value) }))
    expect(html).toContain("Compared Our draft.docx → Their draft.docx: 4 changes, 1 made without tracked changes")
    expect(html).toContain("Open blackline")
  })
})

describe("compare entry points", () => {
  const documents = [doc(1, "Our draft.docx"), doc(2, "Scan.pdf"), doc(3, "Supply agreement - draft.docx", "deliverable")]
  const compareButtons = (onCompare?: (name: string) => void) =>
    [...renderToStaticMarkup(createElement(DocumentUpload, { matterId: "m-1", documents, onUploaded: () => {}, onView: () => {}, onCompare }))
      .matchAll(/aria-label="Compare ([^"]+) with another version"/g)].map(([, name]) => name)

  test("each Word row in the documents drawer offers Compare; PDFs do not", () => {
    expect(compareButtons(() => {})).toEqual(["Supply agreement - draft.docx", "Our draft.docx"])
    expect(compareButtons()).toEqual([])
  })

  test("the viewer's More actions offers Compare with… for a Word file only", () => {
    const viewer = (name: string) => renderToStaticMarkup(createElement(DocumentViewer, { matterId: "m-1", name, onClose: () => {}, onCompare: () => {} }))
    expect(viewer("Our draft.docx")).toContain("Compare with…")
    expect(viewer("Scan.pdf")).not.toContain("Compare with…")
  })
})
