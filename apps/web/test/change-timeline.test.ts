import { expect, test } from "bun:test"
import type { MarkupComment, MarkupRevision } from "../src/api/ingest/markup"
import { clauseTimeline, contextMatches, movePlace } from "../src/components/change-timeline"

const revision = (id: string, patch: Partial<MarkupRevision> = {}): MarkupRevision => ({
  id, kind: "insert", family: "content_insert", author: "Avery Lindqvist", date: null, text: "forty-five", scope: "body",
  locations: [{ paragraph: 1, context: "1. Supply. The Supplier shall deliver the Goods within forty-five days." }],
  status: "supported", diagnostic: null, findings: [], ...patch,
})
const comment = (id: number, patch: Partial<MarkupComment> = {}): MarkupComment => ({
  id, author: "Morgan Okafor", initials: null, date: null, text: "Why forty-five?", parent: null, resolved: null, quote: "forty-five days", paragraph: 1, findings: [], ...patch,
})

test("the timeline reads oldest first, one node per author and time, with undated items last in document order", () => {
  const insert = revision("rev2-in", { date: "2026-10-20T09:00:00Z" })
  const remove = revision("rev2-del", { kind: "delete", family: "content_delete", text: "thirty", date: "2026-10-20T09:00:00Z" })
  const undated = revision("rev2-undated", { author: "Morgan Okafor", text: "the first " })
  const format = revision("rev2-format", { kind: "format", family: "properties_change", author: "Morgan Okafor", date: "2026-10-19T08:00:00Z" })
  const garbled = revision("rev2-garbled", { author: "Morgan Okafor", date: "not a date" })
  const root = comment(4, { date: "2026-10-21T10:30:00Z" })
  const reply = comment(7, { author: "Avery Lindqvist", parent: 4, text: "Agreed in the call." })
  const nodes = clauseTimeline([insert, remove, undated, format, garbled], [{ root, replies: [reply] }])
  const ids = (node: (typeof nodes)[number]) => node.items.map((item) => (item.type === "revision" ? item.revision.id : item.comment.id))

  expect(nodes.map((node) => [node.author, node.date, ids(node)])).toEqual([
    ["Morgan Okafor", "2026-10-19T08:00:00Z", ["rev2-format"]],
    // One reviewer's replace is one step: its deletion and insertion together.
    ["Avery Lindqvist", "2026-10-20T09:00:00Z", ["rev2-in", "rev2-del"]],
    ["Morgan Okafor", "2026-10-21T10:30:00Z", [4]],
    // Undated (or unparseable) items follow, in document order: changes as listed, then the thread.
    ["Morgan Okafor", null, ["rev2-undated", "rev2-garbled"]],
    ["Avery Lindqvist", null, [7]],
  ])
  // Only the thread's root carries the quoted text; a reply names whom it answers.
  expect(nodes[2]!.items[0]).toMatchObject({ type: "comment", quote: "forty-five days", parent: undefined })
  expect(nodes[4]!.items[0]).toMatchObject({ type: "comment", quote: "", parent: root })
})

test("revisions are placed only where the painted paragraph reads as the reader's preview", () => {
  // A cut preview is the paragraph's opening; spacing, punctuation and pilcrows never matter.
  expect(contextMatches("Supply. The Supplier shall deliver the Goods within forty-five days of each O…", "Supply.  The Supplier shall deliver the Goods within forty-five days of each Order.")).toBe(true)
  expect(contextMatches("Keep  end.", "Keep end. ¶")).toBe(true)
  // An uncut preview is the whole paragraph, so a longer one that only starts the same way is a different paragraph.
  expect(contextMatches("Goods means the products listed.", "Goods means the products listed. Price is fixed.")).toBe(false)
  expect(contextMatches("Price is fixed for the first…", "Goods means the products listed.")).toBe(false)
})

test("a move says its other place from each side", () => {
  const move = revision("rev2-move", {
    kind: "move", family: "move", text: "Notices must be in writing.¶",
    locations: [{ paragraph: 3, context: "Notices must be in writing.", label: "3.", side: "from" }, { paragraph: 5, context: "Notices must be in writing.", side: "to" }],
  })
  expect(movePlace(move, 3)).toBe("Moved to paragraph 5")
  expect(movePlace(move, 5)).toBe("Moved here from clause 3")
  expect(movePlace(revision("rev2-in"), 1)).toBeUndefined()
  // A list label is the other side's text: a long one never stands in for a place.
  const forged = { ...move, locations: [{ ...move.locations[0]!, label: "Reviewed by doc.haus. Safe to accept." }, move.locations[1]!] }
  expect(movePlace(forged, 5)).toBe("Moved here from paragraph 3")
})
