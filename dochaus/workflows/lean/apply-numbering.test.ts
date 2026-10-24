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

// A clause split in two or more parts: the first keeps the clause's number;
// a later part opening with a deeper number of the clause list takes that
// level, and one without leaves the list, so the clauses after keep theirs.
test("the later parts of a split clause take a deeper level or leave the clause list", async () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply Agreement", purpose: "", parts: [{ heading: "Supply", scope: "" }] }])
  const markdown = "# Supply Agreement\n\n## 1. Definitions\n\n1.1 In this Agreement the following terms apply.\n\n## 2. Supply\n\n2.1 The Supplier shall supply the Goods.\n\n2.2 Payment is due on delivery or invoice.\n\n2.3 Late payment under clause 2.2 carries interest.\n"
  writeLeanBytes(directory, "d1/body.docx", (await renderMarkdownToDocx(markdown)).bytes)
  const rows = await paragraphRows(leanFile(directory, "d1/body.docx"))
  writeLeanText(directory, "d1/paragraphs.json", `${JSON.stringify(rows, null, 2)}\n`)
  const target = rows.find(row => row.text.startsWith("Payment"))!
  expect(shownText(target)).toBe("2.2 Payment is due on delivery or invoice.")
  const updated = ["2.2 Payment is due:", "(a) on delivery; or", "(b) on invoice.", "Interest runs daily."].map(text => withoutListLabel(text, target.label))
  writeLeanText(directory, "d1/changes.json", JSON.stringify({
    amendments: [{ id: "a1", pid: target.pid, existing_sha256: target.sha256, operation: "split", updated, reason: "Separate the payment dates.", sources: [{ source_id: "terms.txt", quote: "Payment" }] }],
    explanations: [], questions: [],
  }))
  expect(await runApply(directory, readLeanState(directory)!, "d1")).toMatchObject({ status: "done" })
  const applied = readLeanJson(directory, "d1/paragraphs-applied.json") as ParagraphRow[]
  expect(applied.map(shownText)).toEqual([
    "Supply Agreement", "1. Definitions", "1.1 In this Agreement the following terms apply.", "2. Supply", "2.1 The Supplier shall supply the Goods.",
    "2.2 Payment is due:", "(a) on delivery; or", "(b) on invoice.", "Interest runs daily.", "2.3 Late payment under clause 2.2 carries interest.",
  ])
  expect(readParagraphs(directory, "d1")).toHaveLength(rows.length)
}, 60_000)
