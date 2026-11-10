import { expect, test } from "bun:test"
import { correspondHistoryParagraphs, resolveHistoryAnchor } from "./clause-history"
import type { ComparisonParagraph, ParagraphComparison } from "./version-compare-contract"

const paragraph = (anchorId: string, text: string, ordinal: number, nativeId?: string): ComparisonParagraph => ({
  anchorId, text, story: "body", ...(nativeId ? { nativeId } : {}), comments: [], location: { scope: "body", paragraph: ordinal },
})
const comparison = (left: ComparisonParagraph[], right: ComparisonParagraph[], changes: ParagraphComparison["changes"] = []): ParagraphComparison => ({ left, right, changes })

test("a moved and edited clause needs an actual unique Word paragraph identity", () => {
  const left = [paragraph("L", "Delivery is due Monday.", 1, "12345678")]
  const right = [paragraph("R", "Delivery is due Tuesday.", 9, "12345678")]
  expect(correspondHistoryParagraphs(comparison(left, right, [{ kind: "moved", left: ["L"], right: ["R"] }]), left)).toMatchObject({ status: "verified", evidence: "word-paragraph-id", paragraphs: [{ text: right[0]!.text }] })
  expect(correspondHistoryParagraphs(comparison(left, [{ ...right[0]!, nativeId: undefined }]), left).status).toBe("unresolved")
})

test("repeated wording and duplicate Word IDs cannot choose an ordinal lookalike", () => {
  const left = [paragraph("L1", "Keep this wording.", 1), paragraph("L2", "Keep this wording.", 2)]
  const right = [paragraph("R1", "Keep this wording.", 1), paragraph("R2", "Changed wording.", 2)]
  expect(resolveHistoryAnchor(left, { scope: "body", text: "Keep this wording." }).status).toBe("unresolved")
  expect(resolveHistoryAnchor(left, { scope: "body", text: "Keep this wording.", anchorId: "L2" }).status).toBe("verified")
  expect(correspondHistoryParagraphs(comparison(left, right), [left[1]!]).status).toBe("unresolved")
  const duplicate = left.map(row => ({ ...row, nativeId: "12345678" }))
  expect(correspondHistoryParagraphs(comparison(duplicate, [{ ...right[1]!, nativeId: "12345678" }]), [duplicate[1]!]).status).toBe("unresolved")
})

test("a complete literal split is verified; a changed exception or partial merge is unresolved", () => {
  const left = [paragraph("L", "Deliver on Monday. Exceptions require written agreement.", 1)]
  const right = [paragraph("R1", "Deliver on Monday.", 1), paragraph("R2", "Exceptions require written agreement.", 2)]
  const edge: ParagraphComparison["changes"] = [{ kind: "modified", left: ["L"], right: ["R1", "R2"] }]
  expect(correspondHistoryParagraphs(comparison(left, right, edge), left)).toMatchObject({ status: "verified", evidence: "exact-split-merge", paragraphs: right })
  expect(correspondHistoryParagraphs(comparison(left, [right[0]!, { ...right[1]!, text: "Exceptions need no agreement." }], edge), left).status).toBe("unresolved")
  expect(correspondHistoryParagraphs(comparison(right, left, [{ kind: "modified", left: ["R1", "R2"], right: ["L"] }]), [right[0]!]).status).toBe("unresolved")
})

test("missing wording and a similar unrelated insertion are not a history match", () => {
  const left = [paragraph("L", "Supplier may extend delivery only with consent.", 3)]
  const right = [paragraph("R", "Supplier may extend delivery without consent.", 3)]
  expect(correspondHistoryParagraphs(comparison(left, right, [{ kind: "deleted", left: ["L"], right: [] }, { kind: "inserted", left: [], right: ["R"] }]), left).status).toBe("unresolved")
  expect(resolveHistoryAnchor(left, { scope: "body", text: "Unsaved different wording" }).status).toBe("unresolved")
})
