import { afterEach, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { applyMarkupDecisions, MarkupRefusedError, type DocumentMarkup, type MarkupComment, type MarkupDecision, type MarkupRevision } from "../src/api/ingest/markup"
import { authorTones, initialsOf } from "../src/components/comment-authors"
import MarkupPanel, { carryDrafts, carryReplies, decisionLabel, markupAuthors, type Reply } from "../src/components/MarkupPanel"

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
  // Leave tracked is a decision too: a change left undecided is not.
  const kept: Record<string, MarkupDecision> = { "rev2-a": "accept", "rev2-b": "reject", "rev2-c": "hold" }
  expect(carryDrafts(earlier, reading)).toEqual({ drafts: kept, carried: ["rev2-a", "rev2-b", "rev2-c"], dropped: 1 })
  // The same bytes again keep their drafts without calling them carried over.
  expect(carryDrafts({ ...earlier, base: "new" }, reading)).toEqual({ drafts: kept, carried: [], dropped: 0 })
  // An unsupported change is resolved in Word, so a stored decision on it is not kept.
  expect(carryDrafts({ base: "old", decisions: { "rev2-word": "accept" } }, reading)).toEqual({ drafts: {}, carried: [], dropped: 1 })
})

test("a reply carries over while its thread reads the same, must be approved again, and is kept marked when the thread changed", () => {
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
  // A newer comment in the thread keeps the lawyer's words as a draft to check again.
  expect(carryReplies({ base: "old", replies: { 0: reply({ seen: 1 }) } }, reading))
    .toEqual({ replies: { 0: reply({ approved: false, changed: true }) }, carried: [0], dropped: 0 })
  // An edited root, a resolved thread or a broken entry drops it.
  for (const stale of [{ 0: reply({ root: { author: root.author, text: "Thirty days." } }) }, { 5: reply() }, { 0: { text: "Agreed." } }])
    expect(carryReplies({ base: "old", replies: stale }, reading)).toEqual({ replies: {}, carried: [], dropped: 1 })
})

test("each decision is a named radio group of its own, with nothing chosen yet, under tabs that count each kind", () => {
  const brk = revision("rev2-mark", { kind: "delete", family: "paragraph_mark", author: "Morgan Okafor", text: "¶", locations: [{ paragraph: 8, context: "7. Assignment. Neither party may assign" }] })
  const format = revision("rev2-format", { kind: "format", family: "properties_change", author: " ", text: "indirect" })
  expect(decisionLabel(revision("rev2-a"))).toBe("Inserted by Avery Lindqvist: “forty-five”")
  expect(decisionLabel(brk)).toBe("Paragraph break deleted by Morgan Okafor: “7. Assignment. Neither party may assign”")
  expect(decisionLabel(format)).toBe("Formatting changed with no author named: “indirect”")

  const html = renderToStaticMarkup(createElement(MarkupPanel, {
    matterId: "m-1", name: "Supply agreement.docx", markup: markup("a", [revision("rev2-a"), brk, format]), onApplied: () => {}, onRefetch: () => {},
  }))
  // One fieldset per change, named by it; native radios, one name per group, none checked.
  expect(html.match(/<fieldset class="markup-decision" aria-label="Your decision: [^"]+"/g)).toHaveLength(3)
  expect(html).toContain('aria-label="Your decision: Inserted by Avery Lindqvist: “forty-five”"')
  const radios = html.match(/<input type="radio"[^>]*>/g) ?? []
  expect(radios).toHaveLength(9)
  expect(new Set(radios.map((radio) => radio.match(/name="([^"]+)"/)?.[1])).size).toBe(3)
  expect(radios.some((radio) => radio.includes("checked"))).toBe(false)
  // Undecided shows as the card's edge and no pressed button; a screen reader is told.
  expect(html.match(/<span id="[^"]+-state" class="sr-only">Not yet decided<\/span>/g)).toHaveLength(3)
  expect(html.match(/<fieldset class="markup-decision" aria-label="[^"]+" aria-describedby="[^"]+-state"/g)).toHaveLength(3)
  // The tabs carry their counts; the Changes tab is open and tells a screen reader what it shows.
  expect(html).toMatch(/role="tab"[^>]*aria-selected="true"[^>]*>Changes <span class="markup-tab-count">3<\/span>/)
  expect(html).toMatch(/role="tab"[^>]*aria-selected="false"[^>]*>Comments <span class="markup-tab-count">0<\/span>/)
  expect(html).toContain('<p class="markup-showing sr-only" role="status">Showing all 3 changes</p>')
  // Accept all / Reject all waits, as Save does, until the settings say which edits are the lawyer's own.
  expect(html).not.toContain("Accept all")
  // A paragraph break has no text of its own; its card quotes the paragraph.
  expect(html).toContain("<q>7. Assignment. Neither party may assign</q>")
  expect(html).toContain("Accept keeps the new formatting; Reject restores the old.")
})

test("an inserted table row is badged as a table row, not a whole paragraph, and drops its closing mark", () => {
  const row = revision("rev2-row", { kind: "insert", family: "row_insert", text: "New key¶New value¶", locations: [{ paragraph: 21, context: "New key" }] })
  const html = renderToStaticMarkup(createElement(MarkupPanel, {
    matterId: "m-1", name: "Supply agreement.docx", markup: markup("a", [row]), onApplied: () => {}, onRefetch: () => {},
  }))
  expect(html).toContain('<span class="markup-kind-detail">Table row</span>')
  expect(html).not.toContain("Whole paragraph")
  expect(html).toContain(">New key¶New value</div>")
  expect(decisionLabel(row)).toBe("Table row inserted by Avery Lindqvist: “New key New value”")
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
  expect(html).toContain("<span class=\"markup-place-name\">In a comment by Morgan Okafor</span>")
  // Word prints "5." before the clause; its paragraph ordinal is not shown.
  expect(html).toContain("<span class=\"markup-place-name\">Clause 5</span>")
  expect(html).not.toContain("Paragraph 6")
  // Document order in the Changes tab: the body's changes, then the edit inside
  // the comment, which follows its thread's place; the thread is in Comments.
  const at = (text: string) => html.indexOf(text)
  expect(at("forty-five")).toBeLessThan(at("Clause 5"))
  expect(at("Clause 5")).toBeLessThan(at("In a comment by Morgan Okafor"))
  expect(at("In a comment by Morgan Okafor")).toBeLessThan(at("markup-thread"))
})

test("a move says where it left and where it landed, and a comment thread names its clause in its tooltip", () => {
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
  // Each card's place is its tooltip; the thread's anchor line is the quoted words alone.
  expect(html).toContain('title="Moved from clause 6 to clause 4"')
  expect(html).toMatch(/<li tabindex="-1" data-comment-id="2" title="Clause 5"/)
  expect(html).toContain('<q class="markup-quote">the Supplier&#x27;s records</q></span></p>')
  expect(html).not.toContain("Paragraph 6")
})

test("a detached comment shows why it takes no reply, offers no draft and keeps no stored reply", () => {
  const comment = (id: number, text: string, patch: Partial<MarkupComment> = {}): MarkupComment => ({
    id, author: "Morgan Okafor", initials: "MO", date: null, text, parent: null, resolved: false, quote: "", paragraph: 10, findings: [], ...patch,
  })
  const reading = { ...markup("a", []), comments: [comment(0, "Our lead time needs forty-five days.", { paragraph: 1 }), comment(5, "We use this waiver wording.", { detached: true })] }
  const html = renderToStaticMarkup(createElement(MarkupPanel, { matterId: "m-1", name: "Supply agreement.docx", markup: reading, onApplied: () => {}, onRefetch: () => {} }))
  const [open, detached] = html.split(/class="change-card markup-card markup-thread[^"]*"/).slice(1) as [string, string]
  expect(open).toContain("Draft reply")
  expect(detached).toContain("Text removed")
  expect(detached).toContain("The text this comment was on is no longer in the document, so it takes no reply here.")
  expect(detached).not.toContain("Your reply to")
  expect(html).toContain("Draft replies to 1 comment")
  // A reply kept for it from earlier bytes is not carried over.
  const reply: Reply = { text: "Agreed.", approved: true, open_items: [], root: { author: "Morgan Okafor", text: "We use this waiver wording." }, seen: 1 }
  expect(carryReplies({ base: "old", replies: { 5: reply } }, reading)).toEqual({ replies: {}, carried: [], dropped: 1 })
})

test("each author keeps one colour and Word-style initials, told apart from the file's other authors", () => {
  const names = ["Avery Lindqvist", "Morgan Okafor", "Avery Lindqvist", "Synthetic Reviewing Lawyer", "Sam Lee", "Dana Wu", "Chris Park", "Robin Tse", "Jo Park"]
  const tone = authorTones(names, (name) => name === "Synthetic Reviewing Lawyer")
  // As Word colours by author: in the order authors first appear, the lawyer's
  // own in the brand colour without taking a reviewer's, then round again.
  expect(names.map(tone)).toEqual([0, 1, 0, "self", 2, 3, 4, 5, 0])
  // A name repeated with other spacing or case is the same author.
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
  // The file lists its changes first, but Lee's comment on clause 1 and Sam's
  // reply come before Morgan's change to clause 2 and Avery's to clause 4; a
  // comment with no known place comes last.
  expect(markupAuthors(file)).toEqual(["Lee Chan", "Sam Rivera", "Morgan Okafor", "Avery Lindqvist", "Jo Park"])
})
