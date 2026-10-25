import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { applyMarkupDecisions, MarkupRefusedError, type DocumentMarkup, type MarkupRevision } from "../src/api/ingest/markup"
import MarkupPanel, { carryDrafts, decisionLabel } from "../src/components/MarkupPanel"

const originalFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = originalFetch
})

const revision = (id: string, patch: Partial<MarkupRevision> = {}): MarkupRevision => ({
  id, kind: "insert", family: "content_insert", author: "Avery Lindqvist", date: null, text: "forty-five", scope: "body",
  locations: [{ paragraph: 1, context: "1. Supply. The Supplier shall deliver the Goods within forty-five days." }],
  status: "supported", diagnostic: null, findings: [], ...patch,
})
const markup = (sha256: string, revisions: MarkupRevision[]): DocumentMarkup => ({
  name: "Supply agreement.docx", version: { id: `v-${sha256}`, number: 1, sha256 }, current: true, revisions, comments: [], authors: [],
})

test("a refused batch keeps the changes it names, so the panel can mark their cards", async () => {
  const refuse = (body: object) => {
    globalThis.fetch = Object.assign(async () => Response.json(body, { status: 409 }), { preconnect: originalFetch.preconnect })
    return applyMarkupDecisions("m-1", "Supply agreement.docx", { base_sha256: "a", decisions: [{ id: "rev2-a", decision: "accept" }] }).catch((error: unknown) => error)
  }
  const merged = await refuse({ error: "Two changes became one when another change was resolved.", code: "decision_failed", id: "rev2-a", merged_with: ["rev2-b"] })
  expect(merged).toBeInstanceOf(MarkupRefusedError)
  expect(merged).toMatchObject({ status: 409, code: "decision_failed", ids: ["rev2-a", "rev2-b"], message: "Two changes became one when another change was resolved." })
  expect(await refuse({ error: "The file changed.", code: "base_changed" })).toMatchObject({ code: "base_changed", ids: [] })
})

test("decisions carry over to new bytes for the changes still listed, and the rest are counted", () => {
  const earlier = { base: "old", decisions: { "rev2-a": "accept", "rev2-b": "reject", "rev2-gone": "accept", "rev2-c": "hold" } }
  const reading = markup("new", [revision("rev2-a"), revision("rev2-b"), revision("rev2-c"), revision("rev2-word", { status: "unsupported" })])
  expect(carryDrafts(earlier, reading)).toEqual({ drafts: { "rev2-a": "accept", "rev2-b": "reject" }, carried: ["rev2-a", "rev2-b"], dropped: 1 })
  // The same bytes again keep their drafts without calling them carried over.
  expect(carryDrafts({ ...earlier, base: "new" }, reading)).toEqual({ drafts: { "rev2-a": "accept", "rev2-b": "reject" }, carried: [], dropped: 0 })
  // An unsupported change is resolved in Word, so a stored decision on it is not kept.
  expect(carryDrafts({ base: "old", decisions: { "rev2-word": "accept" } }, reading)).toEqual({ drafts: {}, carried: [], dropped: 1 })
})

test("each decision group names its change, author and an excerpt, and is one Tab stop", () => {
  const brk = revision("rev2-mark", { kind: "delete", family: "paragraph_mark", author: "Morgan Okafor", text: "¶", locations: [{ paragraph: 8, context: "7. Assignment. Neither party may assign" }] })
  const format = revision("rev2-format", { kind: "format", family: "properties_change", author: " ", text: "indirect" })
  expect(decisionLabel(revision("rev2-a"))).toBe("Inserted by Avery Lindqvist: “forty-five”")
  expect(decisionLabel(brk)).toBe("Paragraph break deleted by Morgan Okafor: “7. Assignment. Neither party may assign”")
  expect(decisionLabel(format)).toBe("Formatting changed with no author named: “indirect”")

  const html = renderToStaticMarkup(createElement(MarkupPanel, {
    matterId: "m-1", name: "Supply agreement.docx", markup: markup("a", [revision("rev2-a"), brk, format]), onApplied: () => {}, onRefetch: () => {},
  }))
  expect(html.match(/role="radiogroup"/g)).toHaveLength(3)
  expect(html.match(/role="radio"[^>]*tabindex="0"/g)).toHaveLength(3)
  expect(html.match(/role="radio"[^>]*tabindex="-1"/g)).toHaveLength(6)
  // A paragraph break has no text of its own; its card quotes the paragraph.
  expect(html).toContain("<q>7. Assignment. Neither party may assign</q>")
  expect(html).toContain("Accept keeps the new formatting; Reject restores the old.")
})
