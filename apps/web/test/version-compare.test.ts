import { afterEach, describe, expect, test } from "bun:test"
import { INGEST_URL } from "../src/config"
import { compareUrl, fetchBlacklineBytes, fetchComparison, type ChangeHunk, type ComparisonJson } from "../src/api/ingest/compare"
import { choiceParams, compareCandidates } from "../src/components/CompareDialog"
import { RevisionPresentation } from "docxodus"
import { BLACKLINE_RENDER_OPTIONS, changeBlock, hunkLocation, hunkPieces, saveDocx } from "../src/components/BlacklineViewer"
import { excerpt, wordDiff } from "../src/components/redline"
import { answerSegments, chatCompare } from "../src/components/chat/CompareAnswer"
import { toTurns, type ListedChange } from "../src/components/chat/turn-model"
import type { Document } from "../src/api/ingest/documents"
import type { Part } from "../src/api/session-view"

const originalFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = originalFetch
})

const params = { left: "Our draft.docx", right: "Their draft (2).docx" }

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
    changes: [hunk("C1", false), hunk("C2", true, { kind: "inserted", before: undefined, after: "A new late fee applies." })],
    untracked: [hunk("U1", true, { kind: "inserted", before: undefined, after: "A new late fee applies." })],
    blacklineName: "Blackline - Our draft vs Their draft.docx",
    ...patch,
  }
}

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

describe("compare dialog model", () => {
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

  test("compareCandidates filters to other Word files, sources and deliverables, newest first", () => {
    const documents = [doc(1, "Our draft.docx"), doc(2, "Scan.pdf"), doc(3, "Supply agreement - draft.docx", "deliverable"), doc(4, "Their draft.DOCX")]
    expect(compareCandidates(documents, "Our draft.docx").map((d) => d.name)).toEqual(["Their draft.DOCX", "Supply agreement - draft.docx"])
  })
})

describe("redline and blackline helpers", () => {
  test("hunkLocation formats heading, clause number, scope and missing locations", () => {
    expect(hunkLocation(hunk("C1", false).location, comparison())).toBe("2. Payment · 2.1")
    expect(hunkLocation({ scope: "footer", paragraph: 2, side: "left" }, comparison())).toBe("Footer · paragraph 2 in Our draft")
    expect(hunkLocation({ scope: "footer", paragraph: 0, side: "left" }, comparison())).toBe("Footer")
    expect(hunkLocation({ scope: "body", paragraph: 0, side: "right" }, comparison())).toBe("Location not found")
  })

  test("words and spaces between changed words read as one insertion and one deletion", () => {
    expect(wordDiff("pay within 30 days", "pay within 45 business days")).toEqual([
      { kind: "same", text: "pay within " }, { kind: "ins", text: "45 business" }, { kind: "del", text: "30" }, { kind: "same", text: " days" },
    ])
    expect(wordDiff("one two", "three four")).toEqual([{ kind: "ins", text: "three four" }, { kind: "del", text: "one two" }])
    expect(wordDiff("a b c d e", "a x c y e")).toEqual([
      { kind: "same", text: "a " }, { kind: "ins", text: "x" }, { kind: "del", text: "b" }, { kind: "same", text: " c " },
      { kind: "ins", text: "y" }, { kind: "del", text: "d" }, { kind: "same", text: " e" },
    ])
    expect(wordDiff("Fees are due.", "Fees are due.\nInterest accrues.")).toEqual([{ kind: "same", text: "Fees are due." }, { kind: "ins", text: "\nInterest accrues." }])
  })

  test("long unchanged text around a change is cut to nearby words", () => {
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

  test("BLACKLINE_RENDER_OPTIONS renders Word tracked changes", () => {
    expect(BLACKLINE_RENDER_OPTIONS.revisionPresentation).toBe(RevisionPresentation.Word)
    expect(BLACKLINE_RENDER_OPTIONS.renderTrackedChanges).toBe(true)
  })

  test("hunkPieces trims surrounding whitespace for whole insertions and deletions", () => {
    expect(hunkPieces({ kind: "inserted", after: " It excludes 24 December." })).toEqual([{ kind: "ins", text: "It excludes 24 December." }])
    expect(hunkPieces({ kind: "deleted", before: ", without set-off \n" })).toEqual([{ kind: "del", text: ", without set-off" }])
  })

  test("saveDocx downloads bytes with DOCX MIME type and revokes the object URL", async () => {
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

describe("chat compare model", () => {
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
    { id: "P1", kind: "inserted", after: "Title passes on delivery.", outcome: "accepted" },
    { id: "P2", kind: "deleted", before: ", without set-off", outcome: "rejected", untracked: "U1" },
  ]

  test("reads the tool's compare metadata as the blackline to open", () => {
    expect(turn(meta).compares).toEqual([{
      left: "Our draft.docx", right: "Their draft.docx", leftLabel: "Our draft", rightLabel: "Their draft",
      ignoreFormatting: true, detectMoves: false, headersFooters: true, changes: 4, untracked: 1, hunks: [],
    }])
    expect(turn({ compare: { left: "Our draft.docx" } }).compares).toEqual([])
    expect(turn(meta, "running").compares).toEqual([])
  })

  test("each page's changes join one comparison, once each, and malformed ones are skipped", () => {
    const junk = [{ id: "C9", kind: "rewritten", after: "x" }, { id: 3, kind: "inserted", after: "x" }, { id: "C8", kind: "inserted" }, "C7"]
    const { compares } = turnOf([{ ...meta, hunks: hunks.slice(0, 2) }, { ...meta, hunks: [...hunks.slice(1), ...junk] }])
    expect(compares).toHaveLength(1)
    expect(compares![0]!.hunks).toEqual(hunks)
    const other = { compare: { ...meta.compare, right: "Their draft (2).docx" }, hunks: [hunks[0]] }
    expect(turnOf([{ ...meta, hunks }, other]).compares!.map((compare) => compare.right)).toEqual(["Their draft.docx", "Their draft (2).docx"])
  })

  test("chatCompare selects a comparison only when it is the single one with listed hunks", () => {
    expect(chatCompare(turn({ ...meta, hunks }).compares)?.right).toBe("Their draft.docx")
    expect(chatCompare(turn(meta).compares)).toBeUndefined()
    expect(chatCompare(turnOf([{ ...meta, hunks }, { compare: { ...meta.compare, right: "Their draft (2).docx" }, hunks }]).compares)).toBeUndefined()
    expect(chatCompare(undefined)).toBeUndefined()
  })

  test("answerSegments keeps Markdown examples, streamed fences, and qualified labels literal", () => {
    const examples = [
      "```text\n**C1** — Literal example, not a reviewed comparison.\n```",
      "    **C1** — Literal example, not a reviewed comparison.",
      "> **C1** — Literal example, not a reviewed comparison.",
      "`**C1** — Literal example, not a reviewed comparison.`",
      "**C1, `C2`** — Literal example, not a reviewed comparison.",
      "\\*\\*C1\\*\\* — Literal example, not a reviewed comparison.",
      "- Example:\n\n      **C1** — Literal example, not a reviewed comparison.",
    ]
    for (const example of examples) {
      const text = `Example only:\n\n${example}\n\nActual note.`
      expect(answerSegments(text)).toEqual([{ markdown: text }])
    }
    for (const tail of ["", "\n```", "\n```\nActual note."]) {
      const text = "```text\n**C1** — Example only." + tail
      expect(answerSegments(text)).toEqual([{ markdown: text }])
    }
    for (const label of ["C1 not C2", "C1 except C2", "C1 remains unapproved"]) {
      const text = `- **${label}** · The latter remains unapproved.`
      expect(answerSegments(text)).toEqual([{ markdown: text }])
    }
  })

  test("answerSegments parses genuine entries, multi-line continuations, and comma/dot/and ID lists while preserving source", () => {
    const text = [
      "1. **C1** · Delivery — Minor — Favors neither party: genuine note",
      "   continued with `literal **C2**`.",
      "",
      "```text",
      "**C2** — Example only.",
      "```",
      "Actual note after the example.",
      "",
      "- **C1,C99** · Unknown member preserves full source.",
      "",
      "**C2** · The actual deletion.",
    ].join("\n")
    const segments = answerSegments(text)
    const entries = segments.filter((s): s is Extract<(typeof segments)[number], { ids: string[] }> => "ids" in s)
    expect(entries).toEqual([
      {
        ids: ["C1"],
        note: "Delivery — Minor — Favors neither party: genuine note continued with `literal **C2**`.",
        source: "1. **C1** · Delivery — Minor — Favors neither party: genuine note\n   continued with `literal **C2**`.",
      },
      {
        ids: ["C1", "C99"],
        note: "Unknown member preserves full source.",
        source: "- **C1,C99** · Unknown member preserves full source.",
      },
      {
        ids: ["C2"],
        note: "The actual deletion.",
        source: "**C2** · The actual deletion.",
      },
    ])
    expect(segments.filter((s) => "markdown" in s).map((s) => s.markdown).join("\n")).toContain("**C2** — Example only.")

    for (const label of ["C1, C2", "C1 · C2", "C1 and C2"]) {
      expect(answerSegments(`**${label}** · Both reviewed changes.`)).toEqual([
        { ids: ["C1", "C2"], note: "Both reviewed changes.", source: `**${label}** · Both reviewed changes.` },
      ])
    }
  })
})
