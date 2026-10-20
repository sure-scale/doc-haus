import { afterEach, describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { INGEST_URL } from "../src/config"
import { compareUrl, fetchBlacklineBytes, fetchComparison, type ChangeHunk, type ComparisonJson, type ProposalOutcome } from "../src/api/ingest/compare"
import { choiceParams, compareCandidates } from "../src/components/CompareDialog"
import CompareDialog from "../src/components/CompareDialog"
import { RevisionPresentation } from "docxodus"
import BlacklineViewer, { BLACKLINE_RENDER_OPTIONS, ComparePanel, changeBlock, excerpt, hunkLocation, saveDocx, wordDiff } from "../src/components/BlacklineViewer"
import CompareAnswer, { chatCompare } from "../src/components/chat/CompareAnswer"
import CompareView from "../src/components/chat/CompareView"
import { toTurns, type ListedChange } from "../src/components/chat/turn-model"
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
    expect(html).toContain("Our draft had 1 tracked change; the blackline compares its text with that change accepted.")
    expect(html).not.toContain("Our draft had 1 tracked change; compared as accepted.")
    expect(html).toContain("Our draft has 2 changes in doc.haus that are not in the file yet, so they are not included.")
    expect(html).not.toContain("pending change")
    expect(html).not.toContain("doc.haus proposal")
    expect(panel(comparison())).toContain("show under “Not tracked (Their draft)”")
    expect(panel(comparison({}, { pendingProposals: 1 }))).toContain(
      "Their draft has 1 change in doc.haus that is not in the file yet, so it is not included.",
    )
  })

  test("locations read as heading, clause number or paragraph, and long text is clipped with a way to expand", () => {
    expect(hunkLocation(hunk("C1", false).location, comparison())).toBe("2. Payment · 2.1")
    expect(hunkLocation({ scope: "footer", paragraph: 2, side: "left" }, comparison())).toBe("Footer · paragraph 2 in Our draft")
    expect(hunkLocation({ scope: "footer", paragraph: 0, side: "left" }, comparison())).toBe("Footer")
    expect(hunkLocation({ scope: "body", paragraph: 0, side: "right" }, comparison())).toBe("Location not found")
    const html = panel(comparison({ untracked: [hunk("U1", true, { after: "word ".repeat(200) })] }))
    expect(html).toContain("…")
    expect(html).toContain("Show all")
  })
})

// Our draft's own tracked changes and what Their draft did with each.
const proposal = (id: string, outcome: ProposalOutcome["outcome"], patch: Partial<ProposalOutcome> = {}): ProposalOutcome => ({
  id, kind: "inserted", text: "Interest accrues on late payments.", author: "Ann Lawyer", outcome,
  location: { scope: "body", paragraph: 5, side: "left", heading: "2. Payment", number: "2.2" }, ...patch,
})
const withProposals = (patch: Partial<ComparisonJson> = {}) => comparison({
  left: { name: "Our draft.docx", label: "Our draft", trackedChanges: { count: 3, authors: ["Ann Lawyer"] }, pendingProposals: 0 },
  untracked: [
    hunk("U1", true, { kind: "inserted", before: undefined, after: "A new late fee applies." }),
    hunk("U2", true, { before: "Payment is due within thirty days. Interest accrues on late payments.", after: "Payment is due within thirty days." }),
  ],
  proposals: [
    proposal("P1", "pending", { text: "Either party may terminate on notice.", location: { scope: "body", paragraph: 9, side: "left", heading: "5. Term" } }),
    proposal("P2", "accepted", { kind: "deleted", text: "Payment may be withheld.", location: { scope: "body", paragraph: 7, side: "left" } }),
    proposal("P3", "rejected", { untracked: "U2" }),
  ],
  ...patch,
})

describe("proposal outcomes", () => {
  test("no section when the earlier version proposed nothing", () => {
    for (const value of [comparison(), comparison({ proposals: [] })]) {
      const html = panel(value)
      expect(html).not.toContain("Tracked changes in")
      expect(html).not.toContain("compare-proposals")
    }
  })

  test("the section sits after the untracked edits and counts each outcome", () => {
    const html = panel(withProposals())
    expect(html).toContain("Tracked changes in Our draft")
    expect(html).toContain("3 tracked changes: 1 still pending, 1 accepted, 1 rejected or changed.")
    expect(html.indexOf("Edits made without tracked changes")).toBeLessThan(html.indexOf("Tracked changes in Our draft"))
    expect(html.indexOf("Tracked changes in Our draft")).toBeLessThan(html.indexOf("All changes"))
    expect(html.indexOf('id="compare-P1"')).toBeLessThan(html.indexOf('id="compare-P2"'))
    expect(html.indexOf('id="compare-P2"')).toBeLessThan(html.indexOf('id="compare-P3"'))
    const accepted = panel(withProposals({ proposals: [proposal("P1", "accepted"), proposal("P2", "accepted")] }))
    expect(accepted).toContain("2 tracked changes: 2 accepted.")
    expect(accepted).not.toContain("is-rejected")
  })

  test("each card names its outcome, place and proposed text; a rejected one points to its untracked edit", () => {
    const html = panel(withProposals())
    expect(html).toContain("Still pending")
    expect(html).toContain("Accepted")
    expect(html).toMatch(/Rejected or changed — not shown as a tracked change, see <button[^>]*>U2<\/button>/)
    expect(html).toContain('id="compare-U2"')
    expect(html).toContain('id="compare-P3" class="change-card compare-hunk compare-proposal is-rejected"')
    expect(html).toContain('id="compare-P1" class="change-card compare-hunk compare-proposal"')
    expect(html).toContain('id="compare-P2" class="change-card compare-hunk compare-proposal"')
    expect(html).toContain("5. Term · paragraph 9 in Our draft")
    expect(html).toContain("Paragraph 7 in Our draft")
    expect(html).toContain("<span>Insertion</span><span>Ann Lawyer</span>")
    expect(html).toContain("<span>Deletion</span><span>Ann Lawyer</span>")
    expect(html).not.toContain("Proposed")
    expect(html).toContain('<div class="compare-redline"><del>Payment may be withheld.</del></div>')
    expect(html).toContain('<div class="compare-redline"><ins>Either party may terminate on notice.</ins></div>')
  })

  test("a pending or accepted proposal whose paragraph has an untracked edit points to that edit too", () => {
    const html = panel(withProposals({
      proposals: [proposal("P1", "pending", { untracked: "U1" }), proposal("P2", "accepted", { untracked: "U2" }), proposal("P3", "accepted")],
    }))
    expect(html).toMatch(/Still pending<\/span><span[^>]*>Same paragraph has an untracked edit — see <button[^>]*>U1<\/button>/)
    expect(html).toMatch(/Accepted<\/span><span[^>]*>Same paragraph has an untracked edit — see <button[^>]*>U2<\/button>/)
    expect(html.match(/Same paragraph has an untracked edit/g)).toHaveLength(2)
    expect(html).not.toContain("is-rejected")
    expect(html).not.toContain("Removes tracked")
  })

  test("a pending proposal their tracked change reverses is counted and tagged apart", () => {
    const html = panel(withProposals({
      proposals: [
        proposal("P1", "pending"), proposal("P2", "countered"), proposal("P3", "accepted"), proposal("P4", "rejected", { untracked: "U2" }),
      ],
    }))
    expect(html).toContain("4 tracked changes: 1 still pending, 1 reversed by their tracked change, 1 accepted, 1 rejected or changed.")
    expect(html).toMatch(/<span>P2<\/span><span>Insertion<\/span><span>Ann Lawyer<\/span><span class="compare-outcome-tag">Still pending — their tracked change reverses it<\/span>/)
    expect(html).toContain('id="compare-P2" class="change-card compare-hunk compare-proposal"')
    const withEdit = panel(withProposals({ proposals: [proposal("P1", "countered", { untracked: "U1" })] }))
    expect(withEdit).toContain("1 tracked change: 1 reversed by their tracked change.")
    expect(withEdit).toMatch(/Still pending — their tracked change reverses it<\/span><span[^>]*>Same paragraph has an untracked edit — see <button[^>]*>U1<\/button>/)
  })

  test("the untracked edit that removed a proposal links back to it", () => {
    const card = (html: string, id: string) => html.slice(html.indexOf(`id="compare-${id}"`), html.indexOf("</div></div>", html.indexOf(`id="compare-${id}"`)))
    const one = panel(withProposals({
      proposals: [proposal("P1", "pending"), proposal("P2", "rejected", { untracked: "U2" }), proposal("P3", "accepted")],
    }))
    expect(card(one, "U2")).toMatch(/Removes tracked change <button[^>]*title="Show P2 under Tracked changes in Our draft"[^>]*>P2<\/button>/)
    expect(one).toContain('id="compare-P2" class="change-card compare-hunk compare-proposal is-rejected"')
    expect(card(one, "U1")).not.toContain("Removes tracked")
    expect(one.match(/Removes tracked/g)).toHaveLength(1)
    const several = panel(withProposals({
      proposals: [proposal("P2", "rejected", { untracked: "U2" }), proposal("P4", "pending", { untracked: "U2" }), proposal("P5", "rejected", { untracked: "U2" })],
    }))
    expect(card(several, "U2")).toMatch(/Removes tracked changes <button[^>]*>P2<\/button>, <button[^>]*>P5<\/button><\/span>/)
    expect(several).toContain("3 tracked changes: 1 still pending, 2 rejected or changed.")
  })

  test("the disclosure says the blackline compares accepted text and points to its tracked changes", () => {
    const html = panel(withProposals())
    expect(html).toContain(
      "Our draft had 3 tracked changes (by Ann Lawyer); the blackline compares its text with those changes accepted. “Tracked changes in Our draft” above shows what became of each in Their draft.",
    )
    expect(html).not.toContain("Our draft had 3 tracked changes (by Ann Lawyer); compared as accepted.")
  })
})

describe("redline as lawyers read it", () => {
  test("a card shows the changed words inline, in the page's order, not the whole paragraph struck and retyped", () => {
    const html = panel(comparison())
    expect(html).toContain('<div class="compare-redline">Payment is due within <ins>sixty</ins><del>thirty</del> days.</div>')
    expect(html).toContain('<div class="compare-redline"><ins>A new &lt;b&gt;late fee&lt;/b&gt; applies.</ins></div>')
    expect(html).not.toContain("change-old")
    expect(html).not.toContain("change-new")
  })

  test("words and the spaces between changed words read as one insertion and one deletion", () => {
    expect(wordDiff("pay within 30 days", "pay within 45 business days")).toEqual([
      { kind: "same", text: "pay within " }, { kind: "ins", text: "45 business" }, { kind: "del", text: "30" }, { kind: "same", text: " days" },
    ])
    expect(wordDiff("one two", "three four")).toEqual([{ kind: "ins", text: "three four" }, { kind: "del", text: "one two" }])
    expect(wordDiff("a b c d e", "a x c y e")).toEqual([
      { kind: "same", text: "a " }, { kind: "ins", text: "x" }, { kind: "del", text: "b" }, { kind: "same", text: " c " },
      { kind: "ins", text: "y" }, { kind: "del", text: "d" }, { kind: "same", text: " e" },
    ])
    expect(wordDiff("Fees are due.", "Fees are due.\nInterest accrues.")).toEqual([{ kind: "same", text: "Fees are due." }, { kind: "ins", text: "\nInterest accrues." }])
    const moved = panel(comparison({ changes: [hunk("C1", false, { kind: "moved", before: "Clause text.", after: "Clause text." })] }))
    expect(moved).toContain('<div class="compare-redline"><ins class="is-move">Clause text.</ins></div>')
  })

  test("long unchanged text around a change is cut to nearby words, with a way to see it all", () => {
    const words = (from: number, to: number) => Array.from({ length: to - from }, (_, i) => `w${from + i}`).join(" ")
    const pieces = wordDiff(`${words(0, 30)} old ${words(30, 60)} gone ${words(60, 90)}`, `${words(0, 30)} new ${words(30, 60)} here ${words(60, 90)}`)
    const short = excerpt(pieces)
    const text = short.map((piece) => piece.text).join("|")
    expect(text.startsWith("… w18 ")).toBe(true)
    expect(text).toContain("w29 |new|old| w30 ")
    expect(text).toContain(" w41 … w48 ")
    expect(text.endsWith(" w71…")).toBe(true)
    expect(excerpt(wordDiff("Fees are due.", "Fees are owed."))).toEqual(wordDiff("Fees are due.", "Fees are owed."))
  })

  test("the key names the marks in use; the untracked highlight only when the page shows it", () => {
    const key = (html: string) => html.slice(html.indexOf('class="compare-legend"'), html.indexOf("</div>", html.indexOf('class="compare-legend"')))
    expect(key(panel(comparison()))).toMatch(/redline-ins">Inserted.*redline-del">Deleted.*redline-move">Moved.*redline-untracked">Not tracked/)
    expect(key(panel(comparison({ attribution: "label" })))).not.toContain("Not tracked")
    expect(key(panel(comparison({ untracked: [] })))).not.toContain("Not tracked")
    expect(key(panel(comparison({ options: { ignoreFormatting: true, detectMoves: false, headersFooters: true } })))).not.toContain("Moved")
  })

  test("the page uses Word's markup, so it reads like the downloaded file", () => {
    expect(BLACKLINE_RENDER_OPTIONS.revisionPresentation).toBe(RevisionPresentation.Word)
    expect(BLACKLINE_RENDER_OPTIONS.renderTrackedChanges).toBe(true)
  })

  test("a change's place reads as a link to it in the blackline once the page is painted", () => {
    const linked = renderToStaticMarkup(createElement(ComparePanel, { comparison: comparison(), onShow: () => {} }))
    expect(linked).toContain('<button type="button" class="linklike compare-show" title="Show this change in the blackline">2. Payment · 2.1</button>')
    expect(panel(comparison())).not.toContain("compare-show")
    const proposals = renderToStaticMarkup(createElement(ComparePanel, { comparison: withProposals(), onShow: () => {} }))
    const card = proposals.slice(proposals.indexOf('id="compare-P1"'), proposals.indexOf('id="compare-P2"'))
    expect(card).not.toContain("compare-show")
  })
})

describe("finding a change in the painted blackline", () => {
  const block = (before: string, after = before, marked = before !== after) => ({ before, after, marked })
  const page = [
    block("Services Agreement"),
    block("Payment is due within thirty days."),
    block("Payment is due within thirty days.", "Payment is due within sixty days."),
    block("2.3 Interest accrues.", "2.3 "),
    block("Payment is due within thirty days."),
  ]

  test("the paragraph whose new text holds the change, preferring one with a revision", () => {
    expect(changeBlock(page, hunk("C1", false))).toBe(2)
    expect(changeBlock(page, hunk("C2", false, { kind: "deleted", after: undefined, before: "Interest accrues." }))).toBe(3)
  })

  test("repeated wording goes to the paragraph nearest the change's own place", () => {
    const repeated = [block("Notice."), block("x"), block("x"), block("Notice.")]
    const at = (paragraph: number) => hunk("C1", false, { kind: "inserted", before: undefined, after: "Notice.", location: { scope: "body", paragraph, side: "right" } })
    expect(changeBlock(repeated, at(1))).toBe(0)
    expect(changeBlock(repeated, at(4))).toBe(3)
  })

  test("an untracked edit read without their tracked change finds the revised paragraph sharing most of its words", () => {
    const mixed = [block("Payment terms apply."), block("The Customer shall pay the Fees within 30 days.", "The Customer shall promptly pay the Fees within 45 days.")]
    expect(changeBlock(mixed, hunk("U1", true, { after: "The Customer shall pay the Fees within 45 days." }))).toBe(1)
    expect(changeBlock(mixed, hunk("U1", true, { after: "Something else entirely." }))).toBeUndefined()
  })
})

describe("blackline download", () => {
  test("waits for the fetched blackline instead of linking to the route", () => {
    const html = renderToStaticMarkup(createElement(BlacklineViewer, { matterId: "m-1", params, onClose: () => {} }))
    expect(html).toMatch(/<button disabled="" title="Download the blackline/)
    expect(html).not.toContain("download=1")
  })

  test("saves the bytes already on screen as a Word file under the blackline's name", async () => {
    const anchor = { href: "", download: "", clicks: 0, click() { this.clicks++ } }
    const blobs: Blob[] = []
    const revoked: string[] = []
    const { createObjectURL, revokeObjectURL } = URL
    const original = globalThis.document
    URL.createObjectURL = (blob: Blob) => `blob:${blobs.push(blob)}`
    URL.revokeObjectURL = (url: string) => void revoked.push(url)
    globalThis.document = { createElement: () => anchor } as unknown as typeof globalThis.document
    try {
      saveDocx(new Uint8Array([80, 75, 3, 4]), "Blackline - Our draft vs Their draft.docx")
    } finally {
      Object.assign(URL, { createObjectURL, revokeObjectURL })
      globalThis.document = original
    }
    expect(anchor).toMatchObject({ href: "blob:1", download: "Blackline - Our draft vs Their draft.docx", clicks: 1 })
    expect(blobs[0]!.type).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    expect(new Uint8Array(await blobs[0]!.arrayBuffer())).toEqual(new Uint8Array([80, 75, 3, 4]))
    expect(revoked).toEqual(["blob:1"])
  })
})

describe("chat compare card", () => {
  const meta = {
    compare: {
      left: "Our draft.docx", right: "Their draft.docx", leftLabel: "Our draft", rightLabel: "Their draft",
      options: { ignoreFormatting: true, detectMoves: false, headersFooters: true }, counts: { changes: 4, untracked: 1 },
    },
  }
  const turnOf = (metadatas: Record<string, unknown>[], status: "completed" | "running" = "completed") => {
    const parts: Part[] = metadatas.map((metadata, i) => ({
      type: "tool", id: `call-${i}`, sessionID: "session", messageID: "msg", tool: "compare-versions",
      state: status === "completed" ? { status, input: {}, title: "Compared", output: "{}", metadata } : { status, input: {}, title: "Comparing" },
    }))
    return toTurns([{ info: { id: "msg", role: "assistant" }, parts }], "")[0]!
  }
  const turn = (metadata: Record<string, unknown>, status: "completed" | "running" = "completed") => turnOf([metadata], status)
  const hunks: ListedChange[] = [
    { id: "U1", kind: "modified", before: "Payment is due within thirty days.", after: "Payment is due within sixty days." },
    { id: "C1", kind: "inserted", after: "Delivery is complete on signed receipt." },
    { id: "C2", kind: "deleted", before: "Draft for discussion" },
  ]

  test("reads the tool's compare metadata as the blackline to open", () => {
    expect(turn(meta).compares).toEqual([{
      left: "Our draft.docx", right: "Their draft.docx", leftLabel: "Our draft", rightLabel: "Their draft",
      ignoreFormatting: true, detectMoves: false, headersFooters: true, changes: 4, untracked: 1, hunks: [],
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

  test("each page's changes join one comparison, once each, and malformed ones are skipped", () => {
    const junk = [{ id: "C9", kind: "rewritten", after: "x" }, { id: 3, kind: "inserted", after: "x" }, { id: "C8", kind: "inserted" }, "C7"]
    const { compares } = turnOf([{ ...meta, hunks: hunks.slice(0, 2) }, { ...meta, hunks: [hunks[1], hunks[2], ...junk] }])
    expect(compares).toHaveLength(1)
    expect(compares![0]!.hunks).toEqual(hunks)
    const other = { compare: { ...meta.compare, right: "Their draft (2).docx" }, hunks: [hunks[0]] }
    expect(turnOf([{ ...meta, hunks }, other]).compares!.map((compare) => compare.right)).toEqual(["Their draft.docx", "Their draft (2).docx"])
  })

  test("the chat answer reads as entries, each followed by its change's redline", () => {
    const answer = [
      "Version order taken from the file names.",
      "",
      "**Edits made without tracked changes**",
      "",
      "- **U1** · 4.2 Payment — Material — Favors the Supplier: payment moves out",
      "  without a tracked change.",
      "",
      "**Changes**",
      "",
      "**C1, C2** · Delivery — Minor — Favors neither party: wording tidied.",
      "",
      "**C9** · Not in this comparison.",
      "",
      "Minor: C2 (footer). P1 accepted.",
      "",
      "Overall the changes favour the Supplier.",
    ].join("\n")
    const compare = turn({ ...meta, hunks }).compares![0]!
    const html = renderToStaticMarkup(createElement(CompareAnswer, { text: answer, compare, onOpen: () => {} }))
    const at = (text: string) => {
      expect(html).toContain(text)
      return html.indexOf(text)
    }
    const order = [
      at("Version order taken"),
      at("<strong>U1</strong> · 4.2 Payment"),
      at("without a tracked change."),
      at("Payment is due within <ins>sixty</ins><del>thirty</del> days."),
      at("<strong>Changes</strong>"),
      at("<strong>C1, C2</strong>"),
      at(">C1</span>"),
      at("<ins>Delivery is complete on signed receipt.</ins>"),
      at(">C2</span>"),
      at("<del>Draft for discussion</del>"),
      at("<strong>C9</strong>"),
      at("Minor: C2 (footer)"),
      at("Overall the changes"),
    ]
    expect(order).toEqual([...order].sort((a, b) => a - b))
    expect(html).not.toContain("<li><strong>U1")
    expect(html.match(/class="compare-redline"/g)).toHaveLength(3)
    expect(html.match(/Show in blackline/g)).toHaveLength(3)
  })

  test("the chat answer draws redlines only for one comparison with changes listed", () => {
    expect(chatCompare(turn({ ...meta, hunks }).compares)?.right).toBe("Their draft.docx")
    expect(chatCompare(turn(meta).compares)).toBeUndefined()
    expect(chatCompare(turnOf([{ ...meta, hunks }, { compare: { ...meta.compare, right: "Their draft (2).docx" }, hunks }]).compares)).toBeUndefined()
    expect(chatCompare(undefined)).toBeUndefined()
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
