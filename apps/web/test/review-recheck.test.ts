import { expect, test } from "bun:test"
import type { GridCellData, GridRecheck, RecheckPassages } from "../src/api/ingest/grid"
import type { Citation } from "../src/api/opencode/parts"
import { computeCell, describeRecheck, needsFill, recheckPrompt } from "../src/api/review-recheck"
import type { Part } from "../src/api/session-view"

const cite = (section: string, charStart: number, excerpt = `Clause ${section}.`): Citation => ({
  documentName: "Supply.docx", docPath: "/matter/Supply.docx", section, excerpt, charStart, charEnd: charStart + excerpt.length, score: 1,
})
const searched = (...citations: Citation[]) => ({ parts: [{ type: "tool", id: "tool", sessionID: "session", messageID: "message", tool: "search-document",
  state: { status: "completed", input: {}, output: "", title: "", metadata: { citations } } }] as Part[] })
const reply = (text: string) => ({ parts: [{ type: "text", text }] as Part[] })
const scope = { passages: 10, read: 2, candidates: 1, terms: ["liquidated damages", "liquidated", "damages"], channels: ["lexical", "phrase"] }
const input = { document: "Supply.docx", question: "What liquidated damages apply?", questionHash: "h1" }

// Records every call so a test can assert how many model jobs and searches ran.
function fake(first: Array<{ parts: Part[] }>, passages: RecheckPassages, second: Array<{ parts: Part[] }> = []) {
  const calls = { extract: [] as string[], passages: [] as unknown[], recheck: [] as string[] }
  const deps = {
    extract: async (prompt: string) => (calls.extract.push(prompt), first),
    passages: async (seen: unknown) => (calls.passages.push(seen), passages),
    recheck: async (prompt: string) => (calls.recheck.push(prompt), second),
  }
  return { calls, deps }
}

test("a cited answer is stored as before, with no re-check", async () => {
  const { calls, deps } = fake([searched(cite("10", 900)), reply("England and Wales\nSource: §10")], { citations: [], passages: "", scope })
  const cell = await computeCell(input, deps)
  expect(cell).toEqual({
    answer: "England and Wales",
    citation: { documentName: "Supply.docx", docPath: "/matter/Supply.docx", section: "10", excerpt: "Clause 10.", charStart: 900, charEnd: 910 },
    status: "filled",
    questionHash: "h1",
  })
  expect(calls.passages).toHaveLength(0)
  expect(calls.recheck).toHaveLength(0)
})

test("a not-addressed cell with no further matching passage records what was searched and makes no model call", async () => {
  const none = { ...scope, candidates: 0 }
  const { calls, deps } = fake([searched(cite("1", 0), cite("2", 50)), reply("Not addressed")], { citations: [], passages: "", scope: none })
  const cell = await computeCell(input, deps)
  expect(cell).toEqual({ answer: "Not addressed", status: "filled", questionHash: "h1", recheck: { outcome: "not_found", ...none } })
  // The deterministic pass excludes exactly the passages the first search returned.
  expect(calls.passages).toEqual([[{ docPath: "/matter/Supply.docx", charStart: 0, charEnd: 9 }, { docPath: "/matter/Supply.docx", charStart: 50, charEnd: 59 }]])
  expect(calls.extract).toHaveLength(1)
  expect(calls.recheck).toHaveLength(0)
})

test("one re-check adopts an answer whose source is one of the candidate passages", async () => {
  const found = cite("9", 400, "The Supplier shall pay liquidated damages of 2% per week.")
  const { calls, deps } = fake([searched(cite("1", 0)), reply("Not addressed")], { citations: [found], passages: "1. [Supply.docx § 9]\nThe Supplier shall pay liquidated damages of 2% per week.", scope },
    [reply("2% of the order value per week\nSource: §9")])
  const cell = await computeCell(input, deps)
  expect(cell).toEqual({
    answer: "2% of the order value per week",
    citation: { documentName: "Supply.docx", docPath: "/matter/Supply.docx", section: "9", excerpt: found.excerpt, charStart: 400, charEnd: found.charEnd },
    status: "filled",
    questionHash: "h1",
    recheck: { outcome: "found", ...scope },
  })
  expect(calls.recheck).toHaveLength(1)
  expect(calls.recheck[0]).toContain("1. [Supply.docx § 9]")
})

test("a planted wrong source label is rejected and the cell stays negative", async () => {
  const { calls, deps } = fake([reply("Not addressed")], { citations: [cite("9", 400)], passages: "passages", scope },
    [reply("5% of the order value\nSource: §12")])
  expect(await computeCell(input, deps)).toEqual({ answer: "Not addressed", status: "filled", questionHash: "h1", recheck: { outcome: "not_found", ...scope } })
  expect(calls.recheck).toHaveLength(1)
})

test("a re-check answer citing a passage that was not a candidate is rejected", async () => {
  // Even if the re-check session retrieved something itself, only the
  // deterministic pass's candidates can carry the adopted citation.
  const { deps } = fake([reply("Not addressed")], { citations: [cite("9", 400)], passages: "passages", scope },
    [searched(cite("3", 120)), reply("Price list\nSource: §3")])
  expect((await computeCell(input, deps)).citation).toBeUndefined()
})

test("an uncited answer is re-checked and keeps its wording when nothing verifies", async () => {
  const { calls, deps } = fake([searched(cite("1", 0)), reply("Two per cent per week\nSource: §Late delivery")], { citations: [], passages: "", scope: { ...scope, candidates: 0 } })
  const cell = await computeCell(input, deps)
  expect(cell.answer).toBe("Two per cent per week")
  expect(cell.citation).toBeUndefined()
  expect(cell.recheck?.outcome).toBe("not_found")
  expect(calls.passages).toHaveLength(1)
})

test("only blank, stale or never re-checked negative cells are filled; reviewed and re-checked cells are skipped", () => {
  const cited: GridCellData = { answer: "A", citation: { documentName: "d", docPath: "/d", section: "1", excerpt: "A", charStart: 0, charEnd: 1 }, status: "filled", questionHash: "h1" }
  const negative: GridCellData = { answer: "Not addressed", status: "filled", questionHash: "h1" }
  const rechecked: GridCellData = { ...negative, recheck: { outcome: "not_found", ...scope } }
  expect(needsFill(undefined, "h1")).toBe(true)
  expect(needsFill(cited, "h1")).toBe(false)
  expect(needsFill(cited, "h2")).toBe(true)
  expect(needsFill(rechecked, "h1")).toBe(false)
  expect(needsFill(rechecked, "h2")).toBe(true)
  // An older negative cell stored before re-checks existed recomputes once.
  expect(needsFill(negative, "h1")).toBe(true)
  for (const cell of [cited, negative, rechecked]) expect(needsFill({ ...cell, status: "reviewed" }, "h2")).toBe(false)
})

test("the re-check prompt names the document and question and carries the passages as supplied", () => {
  const prompt = recheckPrompt("Supply.docx", "What liquidated damages apply?", "PASSAGES BLOCK")
  expect(prompt).toStartWith('Document: "Supply.docx". Question: What liquidated damages apply?')
  expect(prompt).toContain("\n\nPASSAGES BLOCK")
})

test("the stored re-check reads as what was searched", () => {
  const record = (outcome: GridRecheck["outcome"], candidates: number, rest: Partial<GridRecheck> = {}): GridRecheck => ({ outcome, ...scope, candidates, ...rest })
  expect(describeRecheck(record("not_found", 0))).toBe(
    'Not found after re-check. Searched all 10 passages of this document for "liquidated damages", "liquidated", "damages" by exact terms: no passage beyond the 2 the first search returned mentions them.')
  expect(describeRecheck(record("not_found", 3))).toBe(
    'Not found after re-check. Searched all 10 passages of this document for "liquidated damages", "liquidated", "damages" by exact terms: 3 further passages mention them, but none supported a cited answer.')
  expect(describeRecheck(record("found", 1))).toBe(
    'Found on re-check. Searched all 10 passages of this document for "liquidated damages", "liquidated", "damages" by exact terms: the answer cites the one further passage that mentions them.')
  expect(describeRecheck(record("not_found", 0, { terms: [] }))).toBe("Not found after re-check. The question has no searchable terms, so the document was not searched again.")
  expect(describeRecheck(record("not_found", 0, { passages: 0, read: 0 }))).toBe("Not found after re-check. This document has no indexed passages to search.")
})
