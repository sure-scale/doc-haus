import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { paragraphSha256, withParagraphPointers } from "../../lib/drafting-amendment"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { renderMarkdownToDocx } from "../../lib/render"
import type { ChecklistCheckerInput, ChecklistProducerInput } from "./checklist-contract"
import { runChecklistCheckFlow, runChecklistExtractionFlow } from "./checklist-flow"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { factOpenItems, withFactOptions } from "./fact-answers"
import { finalGate } from "./gate"
import { leanSteps, writeLeanPlan } from "./plan"
import { provisionalFactMarkerItems } from "./provisional-fact-markers"
import { checklistOutputState, LEAN_PROMPT_VERSION, stageInputs, type StageSessionPort } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanState, readQuestions, updateLeanStage, writeLeanBytes, writeLeanState, writeLeanText, writeOpenItems, writeQuestions, type Question } from "./state"
import { cleanupMatters, matter } from "./test-support"

afterEach(cleanupMatters)

const marked = "The inventory reference remains unknown [open item: obtain the inventory reference]."

async function setup() {
  const directory = matter("The inventory reference is needed for the handover.", "Prepare an inventory handover.")
  writeLeanPlan(directory, [{ title: "Handover", purpose: "Record the inventory handover.", file: "handover.docx", parts: [{ heading: "Reference", scope: "Inventory reference." }] }])
  const offered = withFactOptions({ id: "q-reference", stage: "intake", kind: "fact", status: "open", question: "What is the inventory reference?", context: "The reference is not supplied." })
  const question: Question = { ...offered, status: "answered", answer: "Prepare the review draft while leaving this unknown and disclosing the outstanding check.", answer_class: "answered", fact_status: "provisional" }
  writeQuestions(directory, [question])
  writeOpenItems(directory, factOpenItems(directory))
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nThe inventory reference remains unknown and must be obtained before use. [ref: ${factOpenItems(directory)[0]!.id}]\n`)
  await body(directory, "d1", marked)
  return { directory, question }
}

async function body(directory: string, document: string, text: string) {
  const bytes = withParagraphPointers((await renderMarkdownToDocx(text)).bytes)
  writeLeanBytes(directory, `${document}/body.docx`, bytes)
  writeLeanBytes(directory, `${document}/body-applied.docx`, bytes)
  const rows = readDocxCommentParagraphs(bytes).map(row => ({ pid: row.pids[0]!, text: row.paragraph_text, sha256: paragraphSha256(row.paragraph_text) }))
  writeLeanText(directory, `${document}/paragraphs-applied.json`, JSON.stringify(rows))
  writeLeanText(directory, `${document}/paragraphs.json`, JSON.stringify(rows))
  const step = `${document}.apply` as "d1.apply"
  updateLeanStage(directory, readLeanState(directory)!, step, { status: "done", input_sha: hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)) })
}

function model(options: { notApplicable?: string } = {}) {
  const checked: ChecklistCheckerInput[] = []
  const extracted: ChecklistProducerInput[] = []
  let sessions = 0
  const port: StageSessionPort = {
    async createSession() { return `marker-${++sessions}` },
    async prompt(input) {
      if (input.agent === CHECKLIST_JOB_AGENTS.extract) {
        const packet: ChecklistProducerInput = JSON.parse(input.text.split("SOURCE_INPUT:\n\n")[1]!)
        extracted.push(packet)
        return JSON.stringify({ items: [{ check: "The handover records the supplied inventory context.", basis: "observation", source_refs: [packet.source.document === "request" ? packet.passages[0] : packet.instructions[0]], qualifications: [], target: { ...packet.targets[0], treatment: "operative" } }] })
      }
      const packet: ChecklistCheckerInput = JSON.parse(input.text.split("CHECK_INPUT:\n\n")[1]!)
      checked.push(packet)
      const notApplicable = packet.item.target.document === options.notApplicable || options.notApplicable === "*"
      return JSON.stringify({ item_id: packet.item_id, status: notApplicable ? "not_applicable" : "satisfied", reason: notApplicable ? "This separate document does not depend on the unknown reference." : "Every reference-dependent passage visibly preserves the unknown.",
        ...(packet.item.required_action ? { action_status: notApplicable ? "not_applicable" : "fulfilled" } : {}), open_decisions: [],
        evidence: packet.current_evidence.filter(row => row.kind === "body"), source_refs: packet.item.source_refs, inspection_scope: "The complete current target for this one question.", limitation: "The reference itself remains unknown; this checks only its handling." })
    },
  }
  return { deps: { sessions: port, model: null }, checked, extracted }
}

function fresh(directory: string) {
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done", input_sha: hashStageInputs("checklist-check", LEAN_PROMPT_VERSION, stageInputs(directory, "checklist-check")) })
}

const fact = (directory: string) => finalGate(directory).items.find(item => item.ref.question_id === "q-reference")!

test("typed provisional permission gets its own checked marker duty without becoming a fact answer", async () => {
  const { directory, question } = await setup()
  const script = model()
  expect((await runChecklistExtractionFlow(directory, script.deps)).status).toBe("succeeded")
  expect((await runChecklistCheckFlow(directory, script.deps)).status).toBe("succeeded")
  writeLeanText(directory, "checklist/report.md", "# Current checked items\n")
  fresh(directory)
  expect(checklistOutputState(directory).valid).toBe(true)
  expect(fact(directory).status).toBe("disclosed")
  expect(readQuestions(directory)[0]).toEqual(question)
  expect(script.checked.filter(row => row.item.required_action === "flag")).toHaveLength(1)
})

test("adding only marker checks reuses ordinary checks and preserves every completed prefix key and body byte", async () => {
  const { directory } = await setup()
  const script = model()
  await runChecklistExtractionFlow(directory, script.deps)
  await runChecklistCheckFlow(directory, script.deps)
  const markerIDs = provisionalFactMarkerItems(directory).map(row => row.id)
  const saved = readLeanJson(directory, "checklist/checks.json") as { expected_items: string[]; checks: { item_id: string }[] }
  // Saved pre-upgrade checks contain the ordinary items but no product marker checks.
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ ...saved,
    expected_items: saved.expected_items.filter(id => !markerIDs.includes(id)), checks: saved.checks.filter(row => !markerIDs.includes(row.item_id)) }))
  const steps = leanSteps(directory).slice(0, leanSteps(directory).indexOf("checklist-check"))
  const state = readLeanState(directory)!
  writeLeanState(directory, { ...state, stages: { ...state.stages, ...Object.fromEntries(steps.map(step => [step, { ...state.stages[step]!, status: "done" as const }])) } })
  const keys = () => steps.map(step => [step, hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step))])
  const before = keys()
  const extraction = readFileSync(leanFile(directory, "checklist/extraction.json"))
  const draft = readFileSync(leanFile(directory, "d1/body.docx"))
  const accepted = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  const checks = script.checked.length, extracted = script.extracted.length
  expect(checklistOutputState(directory).valid).toBe(false)
  expect((await runChecklistCheckFlow(directory, script.deps)).status).toBe("succeeded")
  expect(script.checked.slice(checks).map(row => row.item_id)).toEqual(markerIDs)
  expect(script.extracted).toHaveLength(extracted)
  expect(keys()).toEqual(before)
  expect(readFileSync(leanFile(directory, "checklist/extraction.json"))).toEqual(extraction)
  expect(readFileSync(leanFile(directory, "d1/body.docx"))).toEqual(draft)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(accepted)
})

test("an unscoped fact checks every body and needs positive markers as well as settled non-applicability elsewhere", async () => {
  const { directory } = await setup()
  writeLeanPlan(directory, [
    { title: "Handover", purpose: "Record the inventory handover.", file: "handover.docx", parts: [{ heading: "Reference", scope: "Inventory reference." }] },
    { title: "Schedule", purpose: "Record a separate meeting.", file: "schedule.docx", parts: [{ heading: "Meeting", scope: "Meeting arrangements." }] },
  ])
  await body(directory, "d1", marked)
  await body(directory, "d2", "The separate meeting takes place after the handover.")
  const script = model({ notApplicable: "d2" })
  await runChecklistExtractionFlow(directory, script.deps)
  await runChecklistCheckFlow(directory, script.deps)
  writeLeanText(directory, "checklist/report.md", "# Current checked items\n")
  fresh(directory)
  expect(provisionalFactMarkerItems(directory).map(row => row.item.target.document)).toEqual(["d1", "d2"])
  expect(fact(directory).status).toBe("disclosed")
  const saved = readLeanJson(directory, "checklist/checks.json") as { checks: { item_id: string; finding: { status: string; action_status?: string; open_decisions?: string[]; evidence: unknown[] } }[] }
  const ids = provisionalFactMarkerItems(directory).map(row => row.id)
  const reset = () => structuredClone(saved)
  const allNotApplicable = reset()
  for (const row of allNotApplicable.checks.filter(row => ids.includes(row.item_id))) Object.assign(row.finding, { status: "not_applicable", action_status: "not_applicable" })
  writeLeanText(directory, "checklist/checks.json", JSON.stringify(allNotApplicable)); fresh(directory)
  expect(fact(directory).status).toBe("open")
  const missingScope = reset()
  missingScope.checks = missingScope.checks.filter(row => row.item_id !== ids[1])
  writeLeanText(directory, "checklist/checks.json", JSON.stringify(missingScope)); fresh(directory)
  expect(fact(directory).status).toBe("open")
  const uncertainScope = reset()
  uncertainScope.checks.find(row => row.item_id === ids[1])!.finding.open_decisions = ["Whether this document depends on the unknown reference."]
  writeLeanText(directory, "checklist/checks.json", JSON.stringify(uncertainScope)); fresh(directory)
  expect(fact(directory).status).toBe("open")
  const unevidencedScope = reset()
  unevidencedScope.checks.find(row => row.item_id === ids[1])!.finding.evidence = []
  writeLeanText(directory, "checklist/checks.json", JSON.stringify(unevidencedScope)); fresh(directory)
  expect(fact(directory).status).toBe("open")
})

test("a scoped question cannot borrow another body's markers and adopted wording creates no human permission duty", async () => {
  const { directory, question } = await setup()
  writeQuestions(directory, [{ ...question, document: "d2" }])
  expect(provisionalFactMarkerItems(directory)).toEqual([])
  writeQuestions(directory, [{ ...question, answer_kind: "adopted" }])
  expect(provisionalFactMarkerItems(directory)).toEqual([])
  writeQuestions(directory, [{ ...question, fact_status: "wait" }])
  expect(provisionalFactMarkerItems(directory)).toEqual([])
  writeQuestions(directory, [{ ...question, pid: "p123456789abc", document: "d1" }])
  expect(provisionalFactMarkerItems(directory)).toEqual([])
})
