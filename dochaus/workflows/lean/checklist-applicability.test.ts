import { afterEach, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { checklistFindingRecovery, type ChecklistCheckerInput, type ChecklistFinding, type ChecklistItem, type ChecklistProducerInput } from "./checklist-contract"
import { runChecklistCheckJob, runChecklistExtractionJob } from "./checklist-job"
import { checklistProducerInputs } from "./checklist-source-context"
import { writeLeanPlan } from "./plan"
import { writeLeanText } from "./state"
import { LEAN_MAX_RETRIES, type StageSessionPort } from "./stages"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

// Script only model assessments. These tests exercise real source preparation,
// quote validation and routing; they do not claim to test model reasoning.
function fixture(resolved: boolean) {
  const directory = mkdtempSync(path.join(tmpdir(), "checklist-applicability-"))
  directories.push(directory)
  writeLeanPlan(directory, [{ title: "Inspection terms", purpose: "Record the instructed inspection arrangement", file: "terms.docx", parts: [{ heading: "Inspection", scope: "Inspection arrangement" }] }])
  writeLeanText(directory, "request.md", resolved
    ? "Use signed-terms.txt as controlling for this document."
    : "The inspection method remains undecided. Flag the remote-inspection proposal in the document pending a decision.")
  const sources = [
    { document: "correspondence.txt", text: "Could inspections take place remotely?" },
    { document: "signed-terms.txt", text: resolved ? "Inspections must take place on site." : "The inspection method will be decided separately." },
  ]
  mkdirSync(path.join(directory, "sources"))
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.run("CREATE TABLE chunks (doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  for (const source of sources) {
    writeFileSync(path.join(directory, "sources", source.document), source.text)
    db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run(source.document, "1", source.text, 0, source.text.length)
  }
  db.close()
  writeFileSync(path.join(directory, "source-index.json"), JSON.stringify({ sources: sources.map(source => ({ document: source.document, extracted_text: `sources/${source.document}` })) }))
  const inputs = checklistProducerInputs(directory)
  const proposal = inputs[0]!.input
  const controlling = inputs[1]!.input.passages[0]!
  const item: ChecklistItem = {
    check: "Does the document flag the remote-inspection question as an open issue?",
    basis: "open_question", source_refs: proposal.passages, qualifications: [],
    target: { ...proposal.targets[0]!, treatment: "open_issue" },
  }
  const input: ChecklistCheckerInput = { item_id: "inspection-question", item, current_evidence: [{ document: "d1", locator: "paragraph 1", text: "Inspections must take place on site.", kind: "body" }], settled_instructions: proposal.instructions, retrieval: proposal.retrieval }
  const finding: ChecklistFinding = { item_id: input.item_id, status: resolved ? "not_applicable" : "unsatisfied", reason: resolved
    ? "The controlling terms settle this document's inspection method; no source requires retaining the earlier question. This does not establish acceptance or withdrawal of that question."
    : "The current human instruction requires disclosure of the undecided proposal, which the inspected paragraph omits.",
  evidence: input.current_evidence, source_refs: [...proposal.passages, controlling, ...proposal.instructions], inspection_scope: "Complete single-paragraph target and cited source passages.", limitation: "No approval is inferred." }
  return { directory, proposal, controlling, input, finding }
}

function port(reply: unknown) {
  const prompts: string[] = []
  const sessions: StageSessionPort = {
    async createSession() { return "session-applicability" },
    async prompt(input) { prompts.push(input.text); return JSON.stringify(reply) },
  }
  return { sessions, prompts, model: null }
}

test("resolved question supports no presentation item; a stale item can retain a cited not-applicable assessment without waiving operative terms", async () => {
  const { directory, proposal, controlling, input, finding } = fixture(true)
  const extraction = port({ items: [] })
  expect((await runChecklistExtractionJob(directory, extraction, proposal)).status).toBe("succeeded")
  const producerPacket: ChecklistProducerInput = JSON.parse(extraction.prompts[0]!.split("SOURCE_INPUT:\n\n")[1]!)
  expect(producerPacket.passages).toEqual(input.item.source_refs)
  expect(producerPacket.instructions).toEqual(input.settled_instructions)
  const check = port(finding)
  const reviewed = await runChecklistCheckJob(directory, check, input)
  expect(reviewed.status).toBe("succeeded")
  if (reviewed.status !== "succeeded") return
  expect(reviewed.value.source_refs).toContainEqual(controlling)
  expect(checklistFindingRecovery(input.item, reviewed.value)).toBe("none")
  expect(JSON.parse(check.prompts[0]!.split("CHECK_INPUT:\n\n")[1]!).item).toEqual(input.item)
  const operative: ChecklistItem = { ...input.item, check: "Does the document require inspections on site?", basis: "directive", source_refs: [controlling], target: { ...input.item.target, treatment: "operative" } }
  const operativeFinding = { ...finding, status: "satisfied" as const, reason: "Current text implements the controlling terms.", source_refs: [controlling] }
  expect((await runChecklistCheckJob(directory, port(operativeFinding), { ...input, item: operative })).status).toBe("succeeded")
  expect(checklistFindingRecovery(operative, { ...operativeFinding, status: "unsatisfied" })).toBe("repair")
})

test("unresolved proposal with an explicit flag instruction stays blocking until presented and then stays open without approval", async () => {
  const { directory, proposal, input, finding } = fixture(false)
  const item = { ...input.item, required_action: "flag" as const }
  const extraction = await runChecklistExtractionJob(directory, port({ items: [item] }), proposal)
  expect(extraction.status).toBe("succeeded")
  const omitted = { ...finding, action_status: "not_fulfilled" as const, open_decisions: ["Inspection method"] }
  const reviewed = await runChecklistCheckJob(directory, port(omitted), { ...input, item })
  expect(reviewed.status).toBe("succeeded")
  if (reviewed.status === "succeeded") expect(checklistFindingRecovery(item, reviewed.value)).toBe("repair")
  const evidence = [{ ...input.current_evidence[0]!, text: "[Open item: decide whether inspections may be remote.]" }]
  const presented = { ...omitted, status: "open" as const, action_status: "fulfilled" as const, reason: "The proposal is flagged; the decision is still outstanding.", evidence }
  const open = await runChecklistCheckJob(directory, port(presented), { ...input, item, current_evidence: evidence })
  expect(open.status).toBe("succeeded")
  if (open.status === "succeeded") {
    expect(open.value.status).toBe("open")
    expect(open.value.open_decisions).toEqual(["Inspection method"])
    expect(checklistFindingRecovery(item, open.value)).toBe("none")
  }
})

test("not-applicable status cannot bypass canonical source quote validation", async () => {
  const { directory, input, finding, controlling } = fixture(true)
  expect((await runChecklistCheckJob(directory, port(finding), input)).status).toBe("succeeded")
  const altered = port({ ...finding, source_refs: [{ ...controlling, quote: "Remote inspections have been approved." }] })
  const rejected = await runChecklistCheckJob(directory, altered, input)
  expect(rejected.status).toBe("failed")
  expect(rejected.attempts).toBe(LEAN_MAX_RETRIES + 1)
  if (rejected.status === "failed") expect(rejected.issues.some(issue => issue.includes("not in that document"))).toBe(true)
})
