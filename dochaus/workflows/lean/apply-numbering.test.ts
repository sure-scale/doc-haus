// Run from dochaus/: bun test workflows/lean/apply-numbering.test.ts
import { afterEach, expect, test } from "bun:test"
import { renderMarkdownToDocx } from "../../lib/render"
import { runApply } from "./apply-stage"
import { readParagraphs, shownText, withoutListLabel, type ParagraphRow } from "./checks"
import { writeLeanPlan } from "./plan"
import { paragraphRows } from "./stages"
import { leanFile, readLeanJson, readLeanState, writeLeanBytes, writeLeanText } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

const MARKDOWN = "# Supply Agreement\n\n## 1. Definitions\n\n1.1 In this Agreement the following terms apply.\n\n## 2. Supply\n\n2.1 The Supplier shall supply the Goods.\n\n2.2 Payment is due on delivery or invoice.\n\n2.3 Late payment under clause 2.2 carries interest.\n"

// Drafts `markdown`, splits the paragraph whose text starts with `starts` into
// `parts` through the apply step, and returns the drafted and applied rows.
async function applySplit(markdown: string, starts: string, parts: string[]) {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply Agreement", purpose: "", parts: [{ heading: "Supply", scope: "" }] }])
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(markdown)).bytes)
  const rows = await paragraphRows(leanFile(directory, "d1/body.docx"))
  writeLeanText(directory, "d1/paragraphs.json", `${JSON.stringify(rows, null, 2)}\n`)
  const target = rows.find(row => row.text.startsWith(starts))!
  const updated = parts.map(text => withoutListLabel(text, target.label))
  writeLeanText(directory, "d1/changes.json", JSON.stringify({
    amendments: [{ id: "a1", pid: target.pid, existing_sha256: target.sha256, operation: "split", updated, reason: "Separate the payment terms.", sources: [{ source_id: "terms.txt", quote: "Payment" }] }],
    explanations: [], questions: [],
  }))
  expect(await runApply(directory, readLeanState(directory)!, "d1")).toMatchObject({ status: "done" })
  return { directory, rows, target, applied: readLeanJson(directory, "d1/paragraphs-applied.json") as ParagraphRow[] }
}

// A clause split in two or more parts: the first keeps the clause's number;
// a later part opening with a deeper number of the clause list takes that
// level, and one without leaves the list, so the clauses after keep theirs.
test("the later parts of a split clause take a deeper level or leave the clause list", async () => {
  const { directory, rows, target, applied } = await applySplit(MARKDOWN, "Payment", ["2.2 Payment is due:", "(a) on delivery; or", "(b) on invoice.", "Interest runs daily."])
  expect(shownText(target)).toBe("2.2 Payment is due on delivery or invoice.")
  expect(applied.map(shownText)).toEqual([
    "Supply Agreement", "1. Definitions", "1.1 In this Agreement the following terms apply.", "2. Supply", "2.1 The Supplier shall supply the Goods.",
    "2.2 Payment is due:", "(a) on delivery; or", "(b) on invoice.", "Interest runs daily.", "2.3 Late payment under clause 2.2 carries interest.",
  ])
  expect(readParagraphs(directory, "d1")).toHaveLength(rows.length)
}, 60_000)

// A later part opening with the next number at the split clause's own level is
// a new sibling: the list numbers it, and the clauses after it renumber.
test("a later part of a split clause with the next number at its own level becomes the next clause", async () => {
  const sibling = await applySplit(MARKDOWN, "Payment", ["2.2 Payment is due on delivery.", "2.3 Late payment carries interest."])
  expect(sibling.applied.map(shownText).slice(5)).toEqual(["2.2 Payment is due on delivery.", "2.3 Late payment carries interest.", "2.4 Late payment under clause 2.2 carries interest."])
  expect(sibling.applied.find(row => row.text === "Late payment carries interest.")).toMatchObject({ label: "2.3", clause: true })
  const limbs = await applySplit("# Supply Agreement\n\n## 1. Supply\n\n1.1 Payment is due:\n\n(a) on delivery; or\n\n(b) on invoice.\n\n1.2 Interest runs daily.\n", "on delivery", ["(a) on delivery;", "(b) on acceptance; or"])
  expect(limbs.applied.map(shownText).slice(2)).toEqual(["1.1 Payment is due:", "(a) on delivery;", "(b) on acceptance; or", "(c) on invoice.", "1.2 Interest runs daily."])
}, 60_000)
