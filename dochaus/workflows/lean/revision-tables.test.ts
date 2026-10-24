import { afterEach, expect, test } from "bun:test"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, unzipSync } from "fflate"
import { pipeTable } from "../../lib/docx-tables"
import { liveText } from "../../lib/extract"
import { TABLE_FIXTURE, revisionTablesOriginalDocx } from "../../test/support/revision-tables-fixture"
import { readLeanEvents } from "./events"
import { runLean } from "./runner"
import type { GateResult } from "./gate"
import { leanFile, readLeanJson, readOpenItems, writeLeanText } from "./state"
import { AGENT, cleanupMatters, drafter, happyPath, matter, NO_FINDINGS, noChanges, scriptedModel, SOURCE } from "./test-support"

// A revision of a .docx whose Schedule 1 particulars, approval record and two
// signature grids are Word tables (test/support/revision-tables-fixture.ts),
// run through the lean workflow with a scripted drafter that writes the tables
// as Markdown tables, as the draft prompt shows them.

afterEach(cleanupMatters)

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const [particulars, approval, supplierGrid, buyerGrid] = TABLE_FIXTURE.tables as [string[][], string[][], string[][], string[][]]
const SUPPLIER = "Signed for and on behalf of Northwind Cartons Limited:"
const BUYER = "Signed for and on behalf of Bluebell Foods Limited:"

function revisedAgreement(options: { quantity?: string; tables?: false } = {}) {
  const schedule = particulars.map(row => row.map(cell => cell === "40,000 cartons" && options.quantity ? options.quantity : cell))
  return [
    `# ${TABLE_FIXTURE.title}`,
    "## 1. Supply",
    "This agreement is made between Northwind Cartons Limited (the Supplier) and Bluebell Foods Limited (the Buyer).",
    "The Supplier shall supply the Products in the quantities and at the prices set out in Schedule 1.",
    "## 2. Payment",
    `The Buyer shall pay each invoice within ${options.quantity ? "thirty" : "forty-five"} days of its date.`,
    "## 3. Term",
    "This agreement starts on the Commencement Date and continues for two years.",
    "## Schedule 1 – Particulars",
    pipeTable(schedule),
    ...options.tables === false ? [] : ["### Approval record", pipeTable(approval)],
    "### Execution",
    SUPPLIER,
    ...options.tables === false ? [] : [pipeTable(supplierGrid)],
    BUYER,
    ...options.tables === false ? [] : [pipeTable(buyerGrid)],
  ].join("\n\n") + "\n"
}

/** A revise run of the fixture with the scripted drafter's agreement; a part
 * prompt that asks for a table gets `retried`, when given. */
async function revise(request: string, agreement: string, retried?: string) {
  const original = revisionTablesOriginalDocx()
  // The matter's index holds the original's canonical text, as ingest records it.
  const scratch = matter()
  writeFileSync(path.join(scratch, SOURCE), original)
  const directory = matter(await liveText(path.join(scratch, SOURCE)))
  writeFileSync(path.join(directory, SOURCE), original)
  writeLeanText(directory, "request.md", request)
  const agent = drafter(agreement)
  const model = scriptedModel({ ...happyPath, [AGENT.review]: () => NO_FINDINGS, [AGENT.changes]: noChanges,
    [AGENT.draft]: (prompt, index) => {
      if (retried && prompt.includes("This part keeps the passage")) return retried
      const reply = agent(prompt, index)
      return prompt.includes("Draft part ") ? reply : reply.replace(`# ${TABLE_FIXTURE.title}`, `# ${TABLE_FIXTURE.title}\nMode: revise\nDeliverable: redline\nCompare against: ${SOURCE}\nFile: revised-agreement.docx`)
    } })
  const result = await runLean(directory, { sessions: model.port, registerDeliverable: async (dir, name, bytes) => { writeFileSync(path.join(dir, name), bytes); return name } })
  return { directory, original, result, model }
}

/** The body's tables as markup, without the paragraph pointers the workflow gives every non-empty paragraph. */
function tables(bytes: Uint8Array) {
  const document = new DOMParser().parseFromString(strFromU8(unzipSync(bytes)["word/document.xml"]!), "application/xml")
  const body = document.getElementsByTagNameNS(W, "body")[0]!
  for (const name of ["bookmarkStart", "bookmarkEnd"])
    Array.from(body.getElementsByTagNameNS(W, name)).forEach(mark => mark.parentNode!.removeChild(mark))
  return Array.from(body.childNodes).filter(node => node.nodeType === 1 && (node as Element).localName === "tbl")
    .map(node => new XMLSerializer().serializeToString(node).replace(/ xmlns:w="[^"]*"/, ""))
}

function file(directory: string, name: string) {
  return new Uint8Array(readFileSync(leanFile(directory, name)))
}

test("a revision that touches no table keeps every original table, empty signature cells included", async () => {
  const { directory, original, result, model } = await revise(TABLE_FIXTURE.untouched, revisedAgreement())
  expect(result.status).toBe("completed")
  const first = model.prompts.find(row => row.agent === AGENT.draft && row.text.includes("Draft part 1 of"))!.text
  expect(first).toContain("<original_tables>")
  expect(first).toContain(`Table 3, after "${SUPPLIER}":\n${pipeTable(supplierGrid)}`)
  expect(tables(file(directory, "d1/body.docx"))).toEqual(tables(original))
  expect(tables(file(directory, "d1/body-applied.docx"))).toEqual(tables(original))
  // The redline marks no change inside the tables.
  expect(tables(file(directory, "d1/redline.docx")).join("")).not.toMatch(/<w:(ins|del)\b/)
  expect(readOpenItems(directory).filter(item => item.ref.table !== undefined)).toEqual([])
  expect(readLeanEvents(directory).find(row => row.event === "tables")).toMatchObject({ outcome: "carried 4, edited 0, drafted 0, missing 0" })
}, 120_000)

test("a cell edit changes only that cell of the original's tables", async () => {
  const { directory, original, result } = await revise(TABLE_FIXTURE.cellEdit, revisedAgreement({ quantity: "45,000 cartons" }))
  expect(result.status).toBe("completed")
  const [schedule, ...others] = tables(file(directory, "d1/body-applied.docx"))
  const [before, ...unchanged] = tables(original)
  expect(schedule).toBe(before!.replace("40,000 cartons", "45,000 cartons"))
  expect(others).toEqual(unchanged)
  const redline = tables(file(directory, "d1/redline.docx"))
  expect(redline[0]).toMatch(/<w:del\b/)
  expect(redline[0]).toMatch(/<w:ins\b/)
  expect(redline.slice(1).join("")).not.toMatch(/<w:(ins|del)\b/)
  expect(readLeanEvents(directory).find(row => row.event === "tables")?.issues?.[0]).toContain("row 3, column 2 changed")
}, 120_000)

test("a draft that drops tables is asked for them once and the final text's missing tables are omissions", async () => {
  const dropped = revisedAgreement({ tables: false })
  const schedule = `## Schedule 1 – Particulars${dropped.split("## Schedule 1 – Particulars")[1]!}`.trim()
  const { directory, result, model } = await revise(TABLE_FIXTURE.untouched, dropped, schedule)
  expect(result.status).toBe("completed")
  const retries = model.prompts.filter(row => row.agent === AGENT.draft && row.text.includes("This part keeps the passage"))
  // One retry of the schedule part names both signature grids; the approval record, left out with its heading, is not asked for.
  expect(retries).toHaveLength(1)
  expect(retries[0]!.text).toContain(`This part keeps the passage "${SUPPLIER}"`)
  expect(retries[0]!.text).toContain(`This part keeps the passage "${BUYER}"`)
  expect(retries[0]!.text).toContain(pipeTable(buyerGrid))
  expect(retries[0]!.text).not.toContain("Approval record\"")
  expect(tables(file(directory, "d1/body-applied.docx"))).toHaveLength(1)
  // Each is an open item the drafting notes must name before the gate passes.
  const omissions = readOpenItems(directory).filter(item => item.ref.table !== undefined)
  expect(omissions.map(item => [item.stage, item.kind, item.ref.table])).toEqual([["apply", "omission", 2], ["apply", "omission", 3], ["apply", "omission", 4]])
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate.disclosed.filter(item => item.ref.table !== undefined).map(item => [item.ref.table, item.resolution])).toEqual([
    [2, "omission named in the drafting notes"], [3, "omission named in the drafting notes"], [4, "omission named in the drafting notes"]])
  expect(omissions[1]!.detail).toBe(`Table 3, after "${SUPPLIER}" of the original "${SOURCE}" (rows "Signature", "Name", "Office", "Date") is not in the final text; check whether the revision should keep it.`)
}, 120_000)
