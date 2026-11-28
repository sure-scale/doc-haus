import { afterEach, expect, test } from "bun:test"
import { applyMarkupDecisions, MarkupRefusedError, type DocumentMarkup, type MarkupComment, type MarkupDecision, type MarkupRevision } from "../src/api/ingest/markup"
import { authorTones, initialsOf } from "../src/components/comment-authors"
import { carryDrafts, carryReplies, decisionLabel, markupAuthors, type Reply } from "../src/components/MarkupPanel"

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

test("decisions carry over to new bytes for the changes still listed, and unsupported or removed ones are dropped", () => {
  const earlier = { base: "old", decisions: { "rev2-a": "accept", "rev2-b": "reject", "rev2-gone": "accept", "rev2-c": "hold" } }
  const reading = markup("new", [revision("rev2-a"), revision("rev2-b"), revision("rev2-c"), revision("rev2-word", { status: "unsupported" })])
  const kept: Record<string, MarkupDecision> = { "rev2-a": "accept", "rev2-b": "reject", "rev2-c": "hold" }
  expect(carryDrafts(earlier, reading)).toEqual({ drafts: kept, carried: ["rev2-a", "rev2-b", "rev2-c"], dropped: 1 })
  expect(carryDrafts({ ...earlier, base: "new" }, reading)).toEqual({ drafts: kept, carried: [], dropped: 0 })
  expect(carryDrafts({ base: "old", decisions: { "rev2-word": "accept" } }, reading)).toEqual({ drafts: {}, carried: [], dropped: 1 })
})

test("a reply carries over while its thread reads the same, clears approval on new bytes, and drops on edited, resolved, or detached threads", () => {
  const comment = (id: number, author: string, text: string, patch: Partial<MarkupComment> = {}): MarkupComment => ({
    id, author, initials: null, date: null, text, parent: null, resolved: false, quote: "", paragraph: 1, findings: [], ...patch,
  })
  const root = comment(0, "Avery Lindqvist", "Our production lead time needs forty-five days.")
  const reply = (patch: Partial<Reply> = {}): Reply => ({
    text: "Agreed.", approved: true, open_items: [], root: { author: root.author, text: root.text }, seen: 2, ...patch,
  })
  const reading = {
    ...markup("new", []),
    comments: [
      root,
      comment(1, "Morgan Okafor", "Agreed if stock is confirmed.", { parent: 0 }),
      comment(5, "Morgan Okafor", "Closed.", { resolved: true }),
      comment(6, "Morgan Okafor", "We use this waiver wording.", { detached: true }),
    ],
  }
  expect(carryReplies({ base: "old", replies: { 0: reply() } }, reading)).toEqual({ replies: { 0: reply({ approved: false }) }, carried: [0], dropped: 0 })
  expect(carryReplies({ base: "new", replies: { 0: reply() } }, reading)).toEqual({ replies: { 0: reply() }, carried: [], dropped: 0 })
  expect(carryReplies({ base: "old", replies: { 0: reply({ seen: 1 }) } }, reading))
    .toEqual({ replies: { 0: reply({ approved: false, changed: true }) }, carried: [0], dropped: 0 })

  const detachedReply: Reply = { text: "Agreed.", approved: true, open_items: [], root: { author: "Morgan Okafor", text: "We use this waiver wording." }, seen: 1 }
  for (const stale of [{ 0: reply({ root: { author: root.author, text: "Thirty days." } }) }, { 5: reply() }, { 6: detachedReply }, { 0: { text: "Agreed." } }]) {
    expect(carryReplies({ base: "old", replies: stale }, reading)).toEqual({ replies: {}, carried: [], dropped: 1 })
  }
})

test("decisionLabel names paragraph breaks, formatting changes, and inserted table rows without trailing paragraph marks", () => {
  const brk = revision("rev2-mark", { kind: "delete", family: "paragraph_mark", author: "Morgan Okafor", text: "¶", locations: [{ paragraph: 8, context: "7. Assignment. Neither party may assign" }] })
  const format = revision("rev2-format", { kind: "format", family: "properties_change", author: " ", text: "indirect" })
  const row = revision("rev2-row", { kind: "insert", family: "row_insert", text: "New key¶New value¶", locations: [{ paragraph: 21, context: "New key" }] })
  expect(decisionLabel(revision("rev2-a"))).toContain("forty-five")
  expect(decisionLabel(brk)).toContain("7. Assignment. Neither party may assign")
  expect(decisionLabel(format)).toContain("indirect")
  expect(decisionLabel(row)).toContain("New key New value")
  expect(decisionLabel(row)).not.toContain("¶")
})

test("each author keeps one colour and Word-style initials, told apart from the file's other authors", () => {
  const names = ["Avery Lindqvist", "Morgan Okafor", "Avery Lindqvist", "Synthetic Reviewing Lawyer", "Sam Lee", "Dana Wu", "Chris Park", "Robin Tse", "Jo Park"]
  const tone = authorTones(names, (name) => name === "Synthetic Reviewing Lawyer")
  expect(names.map(tone)).toEqual([0, 1, 0, "self", 2, 3, 4, 5, 0])
  expect(authorTones(["Morgan Okafor", " avery lindqvist", "Avery Lindqvist"])("AVERY LINDQVIST")).toBe(1)
  expect(initialsOf("Avery Lindqvist", null)).toBe("AL")
  expect(initialsOf("Morgan Okafor", "mo")).toBe("MO")
  expect(initialsOf("Synthetic Reviewing Lawyer", "a long given value")).toBe("SL")
  expect(initialsOf("   ", null)).toBe("?")
})

test("authors take their colours in the order their changes and comments first appear in the document", () => {
  const comment = (id: number, author: string, paragraph: number | null, parent: number | null = null): MarkupComment => ({
    id, author, initials: null, date: null, text: "Noted.", parent, resolved: null, quote: "", paragraph, findings: [],
  })
  const file = {
    ...markup("order", [revision("r1", { locations: [{ paragraph: 4, context: "4. Audit." }] }), revision("r2", { author: "Morgan Okafor", locations: [{ paragraph: 2, context: "2. Price." }] })]),
    comments: [comment(0, "Lee Chan", 1), comment(1, "Sam Rivera", 1, 0), comment(2, "Jo Park", null)],
  }
  expect(markupAuthors(file)).toEqual(["Lee Chan", "Sam Rivera", "Morgan Okafor", "Avery Lindqvist", "Jo Park"])
})
