import { expect, test } from "bun:test"
import { Document, Packer, Paragraph } from "docx"
import { compareParagraphs } from "../../../dochaus/lib/version-compare"
import type { ReturnBaseContrast } from "../../../dochaus/lib/document-return-tie-break-contract"
import { checkBaseContrast, planBaseContrast, supportedBaseIds } from "./document-return-base-contrast"

const docx = (...text: string[]) => Packer.toBuffer(new Document({ sections: [{ children: text.map(item => new Paragraph(item)) }] }))
const excluded = "The transfer excludes the northern strip. The seller retains access, except for emergency use."
const included = "The transfer includes the northern strip. The seller retains no access rights."
const plan = async (left: Uint8Array, right: Uint8Array, incoming: Uint8Array) => {
  const [retained, leftIncoming, rightIncoming] = await Promise.all([compareParagraphs(left, right), compareParagraphs(left, incoming), compareParagraphs(right, incoming)])
  return planBaseContrast("old", "current", retained, leftIncoming, rightIncoming)
}
const answer = (part: ReturnType<typeof planBaseContrast>["parts"][number], support: string) => JSON.stringify({ support, reason: "Synthetic protocol judgment; human confirmation remains necessary.", limitations: [], evidence: { incomingQuote: part.incoming.text, leftQuote: part.left.text, rightQuote: part.right.text } })
const sources = [{ documentId: "doc", versionId: "old", current: false }, { documentId: "doc", versionId: "current", current: true }]
const judgments = sources.map(source => ({ versionId: source.versionId, relation: "same-document" }))
const model = { providerID: "protocol-fixture", id: "fast" }
function checked(contrastPlan: Awaited<ReturnType<typeof plan>>, supports: ("left" | "right" | "neither" | "uncertain")[]): ReturnBaseContrast {
  const parts = contrastPlan.parts.map((part, index) => ({ ...checkBaseContrast(answer(part, supports[index] ?? "uncertain"), part), model }))
  return { ...contrastPlan, coverage: { ...contrastPlan.coverage, processedPassages: parts.length }, parts }
}

test("complete retained version contrasts support older wording, reverse to current-only, and neutral new edits cannot invent a winner", async () => {
  const old = await docx("Agreement between Cedar and Riverstone.", excluded, "Completion is Monday.")
  const current = await docx("Agreement between Cedar and Riverstone.", included, "Completion is Tuesday.")
  const forward = await plan(old, current, await docx("Agreement between Cedar and Riverstone.", excluded, "Completion is Wednesday."))
  expect(forward.unresolved).toEqual([])
  expect(forward.parts).toHaveLength(2)
  expect(forward.parts[0]!.incoming.text).toBe(excluded)
  expect(forward.parts[0]!.left.text).toBe(excluded)
  expect(forward.parts[0]!.right.text).toBe(included)
  expect(supportedBaseIds(sources, sources, judgments, [checked(forward, ["left", "neither"])])).toEqual(["old"])
  const reversed = await plan(old, current, await docx("Agreement between Cedar and Riverstone.", included, "Completion is Wednesday."))
  expect(reversed.unresolved).toEqual([])
  expect(reversed.parts[0]!.incoming.text).toBe(included)
  expect(supportedBaseIds(sources, sources, judgments, [checked(reversed, ["right", "neither"])])).toEqual([])
  expect(supportedBaseIds(sources, sources, judgments, [checked(forward, ["neither", "neither"])])).toEqual([])
  expect(supportedBaseIds(sources, sources, judgments, [checked(forward, ["left", "right"])])).toEqual([])
  expect(supportedBaseIds(sources, sources, judgments, [checked(forward, ["left", "uncertain"])])).toEqual([])
  const shortened = JSON.parse(answer(forward.parts[0]!, "left")); shortened.evidence.leftQuote = "The transfer excludes the northern strip."
  expect(() => checkBaseContrast(JSON.stringify(shortened), forward.parts[0]!)).toThrow("complete exact quotation triple")
}, 30_000)

test("missing, repeated, split, oversized and identical retained correspondence never qualifies a base", async () => {
  const old = await docx("Agreement between Cedar and Riverstone.", excluded)
  const current = await docx("Agreement between Cedar and Riverstone.", included)
  for (const incoming of [await docx("Agreement between Cedar and Riverstone."), await docx(excluded, excluded), await docx("The transfer excludes the northern strip.", "The seller retains access, except for emergency use.")]) {
    const result = await plan(old, current, incoming)
    expect(result.unresolved.length).toBeGreaterThan(0)
    expect(supportedBaseIds(sources, sources, judgments, [checked(result, ["left"])])).toEqual([])
  }
  const long = excluded + " Complete qualified provisions apply. ".repeat(70)
  const large = await plan(await docx(long), await docx(long.replace("excludes", "includes")), await docx(long))
  expect(large.unresolved.join(" ")).toContain("bounded quotation")
  const identical = await plan(old, old, await docx(excluded))
  expect(identical.unresolved.join(" ")).toContain("No distinguishing retained wording")
  expect(supportedBaseIds(sources, sources, judgments, [checked(identical, [])])).toEqual([])
}, 30_000)

test("every retained rival needs directional support; tied older or omitted-current coverage cannot use current exclusion", async () => {
  const result = await plan(await docx(excluded), await docx(included), await docx(excluded))
  const contrast = checked(result, ["left"])
  const another = { documentId: "doc", versionId: "another-old", current: false }
  const all = [...sources, another], allJudgments = [...judgments, { versionId: another.versionId, relation: "same-document" }]
  expect(supportedBaseIds(all, all, allJudgments, [contrast])).toEqual([])
  const tie: ReturnBaseContrast = { leftVersionId: "old", rightVersionId: "another-old", coverage: { changedPassages: 0, processedPassages: 0 }, unresolved: ["No distinguishing retained wording."], parts: [] }
  const otherAgainstCurrent = { ...contrast, leftVersionId: "another-old" }
  expect(supportedBaseIds(all, all, allJudgments, [contrast, tie, otherAgainstCurrent])).toEqual([])
  expect(supportedBaseIds([sources[0]!], sources, [judgments[0]!], [])).toEqual([])
  expect(supportedBaseIds([sources[0]!], sources, [judgments[0]!], [contrast])).toEqual(["old"])
  expect(supportedBaseIds([{ ...sources[0]!, associationCorrected: true }], sources, [judgments[0]!], [contrast])).toEqual([])
  const missing = { ...contrast, coverage: { changedPassages: 2, processedPassages: 1 } }
  expect(supportedBaseIds(sources, sources, judgments, [missing])).toEqual([])
}, 30_000)
