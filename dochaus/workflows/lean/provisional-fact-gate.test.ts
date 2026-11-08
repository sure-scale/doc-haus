import { afterEach, expect, test } from "bun:test"
import type { ChecklistFinding, ChecklistItem } from "./checklist-contract"
import { paragraphSha256 } from "../../lib/drafting-amendment"
import { factOpenItems, withFactOptions } from "./fact-answers"
import { finalGate } from "./gate"
import { writeLeanPlan } from "./plan"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { hashStageInputs, readLeanState, readQuestions, updateLeanStage, writeLeanBytes, writeLeanText, writeOpenItems, writeQuestions, type Question } from "./state"
import { cleanupMatters, matter } from "./test-support"
import { provisionalFactMarkerItems } from "./provisional-fact-markers"

afterEach(cleanupMatters)

const rows = [
  { pid: "p-seal", text: "The seal design remains unknown [open item: obtain the seal design from the custodian]." },
  { pid: "p-handover", text: "The handover awaits the same unknown design [open item: verify the seal design before relying on the handover]." },
]

function body(directory: string, paragraphs: { pid: string; text: string }[]) {
  writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify(paragraphs.map(row => ({ ...row, sha256: paragraphSha256(row.text) }))))
}

function setup(checklist = true) {
  const directory = matter("The custodian must identify the seal design.", "Prepare an archive handover.")
  writeLeanPlan(directory, [{ title: "Archive handover", purpose: "Describe the handover.", parts: [{ heading: "Seal", scope: "Seal design and its effect on handover." }] }])
  const offered = withFactOptions({ id: "q-seal", stage: "intake", kind: "fact", question: "What is the seal design?", context: "The design is not supplied.", status: "open" })
  const question: Question = { ...offered, status: "answered", answer_kind: "option", option_id: "o1", answer: offered.options![0]!.text, fact_status: "provisional" }
  writeQuestions(directory, [question])
  writeOpenItems(directory, factOpenItems(directory))
  writeLeanText(directory, "drafting-notes.md", `# Notes\n\nThe seal design remains unknown; obtain it from the custodian. Draft preparation is permitted, but sending or signature is not approved. [ref: ${factOpenItems(directory)[0]!.id}]\n`)
  body(directory, rows)
  writeLeanBytes(directory, "d1/body-applied.docx", Buffer.from("synthetic checked Word bytes"))
  const expected = provisionalFactMarkerItems(directory)[0]!
  const item: ChecklistItem = structuredClone(expected.item)
  const finding: ChecklistFinding = { item_id: expected.id, status: "satisfied", action_status: "fulfilled", reason: "Each dependent passage preserves the unknown design with an explicit marker.", evidence: rows.map(row => ({ document: "d1", locator: row.pid, text: row.text, kind: "body" as const })), source_refs: item.source_refs, inspection_scope: "Every current handover passage dependent on the seal design.", limitation: "The design remains unknown.", open_decisions: [] }
  const data = { directory, question, item, finding }
  if (checklist) save(data)
  return data
}

function fresh(directory: string) {
  updateLeanStage(directory, readLeanState(directory)!, "checklist-check", { status: "done", input_sha: hashStageInputs("checklist-check", LEAN_PROMPT_VERSION, stageInputs(directory, "checklist-check")) })
}

function save({ directory, item, finding }: ReturnType<typeof setup>, extra: { item: ChecklistItem; finding: ChecklistFinding }[] = []) {
  const entries = [{ item, finding }, ...extra]
  // Product handling duties are derived at check time, not source quotations
  // authored by extraction. A forged item under their reserved ID cannot replace one.
  const expected = provisionalFactMarkerItems(directory)[0]!
  const forged = JSON.stringify(item) === JSON.stringify(expected.item) ? [] : [{ id: finding.item_id, item }]
  writeLeanText(directory, "checklist/extraction.json", JSON.stringify({ expected_sources: ["answer"], sources: [{ key: "answer", status: "succeeded", items: [...forged, ...extra.map(row => ({ id: row.finding.item_id, item: row.item }))] }] }))
  writeLeanText(directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: entries.map(row => row.finding.item_id), checks: entries.map(row => ({ item_id: row.finding.item_id, input_sha: "focused-check", status: "succeeded", issues: [], finding: row.finding })) }))
  writeLeanText(directory, "checklist/report.md", "# Checked current marker duties\n")
  fresh(directory)
}

const fact = (directory: string) => finalGate(directory).items.find(item => item.ref.question_id === "q-seal")!

test("a current satisfied answer-derived marker duty recognizes descriptive markers without answering the fact or approving release", () => {
  const { directory, question } = setup()
  expect(rows.every(row => !row.text.includes(question.question))).toBe(true)
  expect(fact(directory)).toMatchObject({ status: "disclosed", resolution: "the fact remains unknown; marked in the body and named in the drafting notes" })
  expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
  expect(readQuestions(directory)[0]).toEqual(question)
})

test("the existing literal-placeholder route remains available without a checklist", () => {
  const { directory } = setup(false)
  body(directory, [{ pid: "p-literal", text: factOpenItems(directory)[0]!.ref.placeholder! }])
  expect(fact(directory).status).toBe("disclosed")
  expect(finalGate(directory).delivery).toBe("review_draft")
})

test("a satisfied duty for another fact, artifact or action cannot substitute for this fact's body markers", () => {
  const changes: Array<(data: ReturnType<typeof setup>) => void> = [
    data => { data.item.source_refs[0]!.locator = ".dochaus/lean/questions.json answer (q-other)" },
    data => { data.finding.source_refs[0]!.locator = ".dochaus/lean/questions.json answer (q-other)" },
    data => { data.item.source_refs[0]!.quote = "show a placeholder wherever it is needed" },
    data => { data.finding.source_refs[0]!.quote = `${data.question.answer!} The seal is already verified.` },
    data => { data.item.target.document = "d2" },
    data => { data.finding.evidence.forEach(row => { row.document = "notes" }) },
    data => { data.finding.evidence.forEach(row => { row.kind = "comment" }) },
    data => { data.finding.evidence = [] },
    data => { data.item.target.treatment = "operative" },
    data => { data.item.basis = "observation" },
    data => {
      data.item.check = "Do not imply authority to sign or send the handover."
      data.item.source_refs[0]!.quote = "This does not authorise sending, signing or serving the document."
      data.finding.source_refs[0]!.quote = data.item.source_refs[0]!.quote
      data.finding.evidence = [{ document: "d1", locator: "p-authority", text: "[open item: obtain signature approval]", kind: "body" }]
      body(data.directory, [{ pid: "p-authority", text: data.finding.evidence[0]!.text }])
    },
  ]
  for (const change of changes) {
    const data = setup()
    change(data)
    save(data)
    expect(fact(data.directory).status).toBe("open")
  }
})

test("every matching marker duty must succeed and its exact body evidence must remain marked and current", () => {
  const changes: Array<(data: ReturnType<typeof setup>) => void> = [
    data => { data.finding.status = "open" },
    data => { data.finding.status = "unsatisfied" },
    data => { data.finding.status = "not_applicable" },
    data => { data.finding.open_decisions = ["Whether all dependent passages are marked."] },
    data => { data.finding.action_status = "not_fulfilled" },
    data => { data.finding.evidence[0]!.text = "A marker that never appears [open item: obtain the seal design]." },
    data => {
      const unmarked = "The seal design is not supplied."
      data.finding.evidence[0]!.text = unmarked
      body(data.directory, [{ ...rows[0]!, text: unmarked }, rows[1]!])
    },
  ]
  for (const change of changes) {
    const data = setup()
    change(data)
    save(data)
    expect(fact(data.directory).status).toBe("open")
  }
  const data = setup()
  save(data, [{ item: structuredClone(data.item), finding: { ...structuredClone(data.finding), status: "open" } }])
  expect(fact(data.directory).status).toBe("open")
  const other = setup()
  save(other, [{ item: { ...structuredClone(other.item), target: { ...other.item.target, treatment: "operative" } }, finding: { ...structuredClone(other.finding), item_id: "operative-marker-duty", status: "unsatisfied" } }])
  expect(fact(other.directory).status).toBe("disclosed")
})

test("stale or incomplete checklist output cannot establish descriptive marker coverage", () => {
  const changes: Array<(data: ReturnType<typeof setup>) => void> = [
    data => { updateLeanStage(data.directory, readLeanState(data.directory)!, "checklist-check", { input_sha: "old" }) },
    data => { updateLeanStage(data.directory, readLeanState(data.directory)!, "checklist-check", { status: "pending" }) },
    data => { writeLeanBytes(data.directory, "d1/body-applied.docx", Buffer.from("changed Word bytes")) },
    data => { body(data.directory, [rows[1]!]) },
    data => { body(data.directory, [{ ...rows[0]!, text: `${rows[0]!.text} The seal design has been approved.` }, rows[1]!]) },
    data => { writeLeanText(data.directory, "checklist/checks.json", JSON.stringify({ extraction_complete: false, expected_items: ["marker-duty"], checks: [] })) },
    data => { writeLeanText(data.directory, "checklist/checks.json", JSON.stringify({ extraction_complete: true, expected_items: ["marker-duty"], checks: [{ item_id: "marker-duty", input_sha: "focused-check", status: "failed", issues: ["No valid check."] }] })) },
  ]
  for (const change of changes) {
    const data = setup()
    change(data)
    expect(fact(data.directory).status).toBe("open")
  }
})

test("a question scoped to an unmarked paragraph cannot borrow another paragraph's checked marker", () => {
  const data = setup()
  writeQuestions(data.directory, [{ ...data.question, pid: "p-unmarked" }])
  body(data.directory, [...rows, { pid: "p-unmarked", text: "The handover is ready." }])
  fresh(data.directory)
  expect(fact(data.directory).status).toBe("open")
})

test("notes alone, an unrelated marker, absent permission or changed answer evidence cannot disclose the fact", () => {
  const changes: Array<(data: ReturnType<typeof setup>) => void> = [
    data => { writeLeanText(data.directory, "drafting-notes.md", "# Notes\n\nThe handover is prepared.\n") },
    data => { body(data.directory, [{ pid: "p-unrelated", text: "[open item: insert the custodian's address]" }]) },
    data => { writeQuestions(data.directory, [{ ...data.question, fact_status: "wait" }]) },
    data => { writeQuestions(data.directory, [{ ...data.question, fact_status: "unknown" }]) },
    data => { writeQuestions(data.directory, [{ ...data.question, answer: "Prepare the draft with appropriate qualifications." }]) },
    data => { writeQuestions(data.directory, [{ ...data.question, answer_kind: "adopted" }]) },
  ]
  for (const change of changes) {
    const data = setup()
    change(data)
    fresh(data.directory)
    expect(fact(data.directory).status).toBe("open")
  }
})

test("the same express permission typed by the human is not restricted to an option click", () => {
  const data = setup()
  writeQuestions(data.directory, [{ ...data.question, answer_kind: undefined, option_id: undefined }])
  fresh(data.directory)
  expect(fact(data.directory).status).toBe("disclosed")
})
