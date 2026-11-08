import { Database } from "bun:sqlite"
import { afterEach, expect, spyOn, test } from "bun:test"
import path from "node:path"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { checkDraftingNotes } from "./checks"
import { writeLeanPlan } from "./plan"
import { quoteFound } from "./quotes"
import { writeLeanText } from "./state"
import { cleanupMatters, matter, SOURCE, SOURCE_TEXT } from "./test-support"

afterEach(cleanupMatters)

function supportedPlans() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Container handover", purpose: "Describe handling instructions.", parts: [{ heading: "Handling", scope: "Container instructions." }] }])
  const paragraphs = Array.from({ length: 24 }, (_, index) => {
    const text = `Keep container ${index + 1} labelled.`
    return { pid: `p${index.toString(16).padStart(12, "0")}`, text, sha256: paragraphSha256(text), draft_part: 0 }
  })
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(paragraphs))
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(paragraphs))
  writeLeanText(directory, "d1/change-plan.json", JSON.stringify({ shared_decisions: [], findings: [], questions: [],
    groups: paragraphs.map((row, index) => ({ id: `g${index}`, pids: [row.pid], goal: "Retain the stated handling instruction.",
      sources: [{ source_id: SOURCE, quote: SOURCE_TEXT }], limitations: [], resolves: [] })),
  }))
  return directory
}

function sourceReads(work: () => unknown) {
  const queries = spyOn(Database.prototype, "query")
  try {
    const result = work()
    return { result, reads: queries.mock.calls.filter(([sql]) => /SELECT section, text, char_start, char_end FROM chunks/.test(String(sql))).length }
  } finally { queries.mockRestore() }
}

test("one notes check verifies repeated plan source evidence once while preserving every prose check", () => {
  const directory = supportedPlans()
  const prose = `# Notes\n\nThe instruction is "${SOURCE_TEXT}" [source: ${SOURCE}].`
  const checked = sourceReads(() => checkDraftingNotes(directory, prose, [], { delegatedIDs: [] }))
  expect(checked.result).toEqual({ issues: [], openItems: [] })
  // Neither collecting ids twice nor revisiting the same source in many
  // groups requires reopening its exact quotation during this one check.
  expect(checked.reads).toBe(1)
  const invalid = checkDraftingNotes(directory, `# Notes\n\nDocument d1 leaves [INSERT COLOUR]. The instruction is "Discard the container." [source: ${SOURCE}]. In Lumen v. Solace, the custodian discards it.`, [], { delegatedIDs: ["q-inspection"] })
  for (const fragment of ["document ids", "placeholders", "those words are not in", "no matter source", "delegated questions"])
    expect(invalid.issues.some(issue => issue.includes(fragment))).toBe(true)
})

test("each check and each unscoped quote read observes changed canonical evidence", () => {
  const directory = supportedPlans()
  const prose = `# Notes\n\nThe instruction is "${SOURCE_TEXT}" [source: ${SOURCE}].`
  const quote = { source_id: SOURCE, quote: SOURCE_TEXT }
  expect(quoteFound(directory, quote)).toBe(true)
  expect(checkDraftingNotes(directory, prose, [], { delegatedIDs: [] }).issues).toEqual([])
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try { db.run("UPDATE chunks SET text = ?", ["The custodian must inspect the seal."]) } finally { db.close() }
  expect(quoteFound(directory, quote)).toBe(false)
  const checked = sourceReads(() => checkDraftingNotes(directory, prose, [], { delegatedIDs: [] }))
  expect(checked.reads).toBeGreaterThan(0)
  expect((checked.result as ReturnType<typeof checkDraftingNotes>).issues.some(issue => issue.includes("those words are not in"))).toBe(true)
  const item = { id: "review:omission:seal", kind: "omission" as const, stage: "review" as const, status: "open" as const,
    detail: "The seal design remains unknown.", ref: { repair_target: "memo" as const } }
  expect(checkDraftingNotes(directory, "# Notes\n\nInspect the seal.", [item], { delegatedIDs: [] }).issues.some(issue => issue.includes(item.id))).toBe(true)
})
