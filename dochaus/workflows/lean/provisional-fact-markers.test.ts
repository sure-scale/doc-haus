import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { paragraphSha256, withParagraphPointers } from "../../lib/drafting-amendment"
import { renderMarkdownToDocx } from "../../lib/render"
import type { ChecklistCheckerInput, ChecklistFinding, ChecklistItem, ChecklistProducerInput } from "./checklist-contract"
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

describe("provisional fact gate", () => {
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
})

describe("provisional fact marker flow", () => {
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
})
