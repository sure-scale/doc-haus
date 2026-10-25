import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { applyMarkupDecisions, MarkupRefusedError, type DocumentMarkup, type MarkupComment, type MarkupRevision } from "../src/api/ingest/markup"
import MarkupPanel, { carryDrafts, carryReplies, decisionLabel, type Reply } from "../src/components/MarkupPanel"

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

test("a reply carries over while its thread reads the same and has nothing newer, and must be approved again", () => {
  const comment = (id: number, author: string, text: string, patch: Partial<MarkupComment> = {}): MarkupComment => ({
    id, author, initials: null, date: null, text, parent: null, resolved: false, quote: "", paragraph: 1, findings: [], ...patch,
  })
  const root = comment(0, "Avery Lindqvist", "Our production lead time needs forty-five days.")
  const reply = (patch: Partial<Reply> = {}): Reply => ({
    text: "Agreed.", approved: true, open_items: [], root: { author: root.author, text: root.text }, seen: 2, ...patch,
  })
  const reading = { ...markup("new", []), comments: [root, comment(1, "Morgan Okafor", "Agreed if stock is confirmed.", { parent: 0 }), comment(5, "Morgan Okafor", "Closed.", { resolved: true })] }
  expect(carryReplies({ base: "old", replies: { 0: reply() } }, reading)).toEqual({ replies: { 0: reply({ approved: false }) }, carried: [0], dropped: 0 })
  // The same bytes keep the approval.
  expect(carryReplies({ base: "new", replies: { 0: reply() } }, reading)).toEqual({ replies: { 0: reply() }, carried: [], dropped: 0 })
  // A newer comment in the thread, an edited root, a resolved thread or a broken entry drops it.
  for (const stale of [{ 0: reply({ seen: 1 }) }, { 0: reply({ root: { author: root.author, text: "Thirty days." } }) }, { 5: reply() }, { 0: { text: "Agreed." } }])
    expect(carryReplies({ base: "old", replies: stale }, reading)).toEqual({ replies: {}, carried: [], dropped: 1 })
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

test("an edit inside a comment is placed in that comment's thread, and a list number is preferred to the paragraph's ordinal", () => {
  const comment: MarkupComment = {
    id: 4, author: "Morgan Okafor", initials: "MO", date: null, text: "Agreed only if the Buyer pays on time.", parent: null, resolved: null,
    quote: "Governing law", paragraph: 7, findings: [],
  }
  const inComment = revision("rev2-comment", { author: "Morgan Okafor", text: "only ", scope: "comment", locations: [{ paragraph: 1, context: "Agreed only if the Buyer pays on time.", comment: 4 }] })
  const term = revision("rev2-term", { text: "two", locations: [{ paragraph: 6, context: "Term. This Agreement continues for two years.", label: "5." }] })
  const html = renderToStaticMarkup(createElement(MarkupPanel, {
    matterId: "m-1", name: "Supply agreement.docx", markup: { ...markup("a", [inComment, revision("rev2-a"), term]), comments: [comment] }, onApplied: () => {}, onRefetch: () => {},
  }))
  expect(html).not.toContain("undefined")
  expect(html).toContain("<span>In a comment by Morgan Okafor</span>")
  // Word prints "5." before the clause; its paragraph ordinal is not shown.
  expect(html).toContain("<span>Clause 5</span>")
  expect(html).not.toContain("Paragraph 6")
  // Document order: the body's changes, then the comment's thread, then the edit inside it.
  const at = (text: string) => html.indexOf(text)
  expect(at("forty-five")).toBeLessThan(at("Clause 5"))
  expect(at("Clause 5")).toBeLessThan(at("markup-thread"))
  expect(at("markup-thread")).toBeLessThan(at("In a comment by Morgan Okafor"))
})

test("a move says where it left and where it landed, and a comment thread names its clause as a change card does", () => {
  const text = "Confidentiality. Each party keeps the other's information confidential."
  // Moved up: its places come in document order, destination first.
  const up = revision("rev2-up", { kind: "move", family: "move", text: `${text}¶`, locations: [
    { paragraph: 5, context: text, label: "4.", side: "to" },
    { paragraph: 7, context: text, label: "6.", side: "from" },
  ] })
  const unknown = revision("rev2-unknown", { kind: "move", family: "move", text: `${text}¶`, locations: [{ paragraph: 3, context: text }, { paragraph: 9, context: text }] })
  const comment: MarkupComment = {
    id: 2, author: "Morgan Okafor", initials: "MO", date: null, text: "Limit audits to once a year.", parent: null, resolved: null,
    quote: "the Supplier's records", paragraph: 6, label: "5.", findings: [],
  }
  const html = renderToStaticMarkup(createElement(MarkupPanel, {
    matterId: "m-1", name: "Supply agreement.docx", markup: { ...markup("a", [up, unknown]), comments: [comment] }, onApplied: () => {}, onRefetch: () => {},
  }))
  expect(html).toContain("Moved from clause 6 to clause 4")
  // Without sides the card keeps naming both places in document order.
  expect(html).toContain("Old and new places: paragraphs 3 and 9")
  expect(html).toContain("<span>Clause 5</span>")
  expect(html).not.toContain("Paragraph 6")
})
