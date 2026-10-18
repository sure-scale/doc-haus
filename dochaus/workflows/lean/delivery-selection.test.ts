import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { readFileSync, writeFileSync } from "node:fs"
import path from "node:path"
import { renderMarkdownToDocx } from "../../lib/render"
import { memoNotesFile, readLeanDocuments } from "./deliverables"
import { checkDeliverySelections, selectedDeliveryPlan } from "./delivery-selection"
import { packageReviewPrompt } from "./package-review"
import { DOC_STAGES, readLeanPlan, writeLeanPlan } from "./plan"
import { readRepairRound, startRepairRound } from "./repairs"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { stagePrompt } from "./prompts"
import { leanFile, readLeanState, writeLeanState, writeLeanText } from "./state"
import { AGENT, DRAFT, SOURCE, SOURCE_TEXT, cleanupMatters, drafter, happyPath, matter, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

test("an explicit Word output title supplies the exact basename without inventing a name", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Terms", scope: "Preserve" }] }])
  const selection = { document: "d1" as const, role: "body" as const, file: "Supply Review.docx", instruction_quote: "" }
  for (const instruction of [
    "Create an editable Word draft in this matter titled Supply Review, with an open-items list.",
    'Provide the clean document with the title "Supply Review" for review.',
  ]) {
    writeLeanText(directory, "request.md", instruction)
    expect(checkDeliverySelections(directory, [{ ...selection, instruction_quote: instruction }])).toEqual([])
    for (const file of ["Guessed Review.docx", "Supply.docx", "Review.docx", "supply review.docx"])
      expect(checkDeliverySelections(directory, [{ ...selection, file, instruction_quote: instruction }]).length).toBeGreaterThan(0)
    // Same name, but a model-authored quote cannot supply the user's intent.
    expect(checkDeliverySelections(directory, [{ ...selection, instruction_quote: "Write a draft titled Supply Review." }]).length).toBeGreaterThan(0)
  }
  for (const instruction of [
    "Create a draft titled Supply Review Final.",
    "Create a draft titled Supply Review.pdf.",
  ]) {
    writeLeanText(directory, "request.md", instruction)
    expect(checkDeliverySelections(directory, [{ ...selection, instruction_quote: instruction }]).length).toBeGreaterThan(0)
  }
})

test("literal name grounding leaves output intent and roles to the semantic reviewer", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Terms", scope: "Preserve" }] }])
  for (const instruction of [
    "Deliver the annotated comparison as Supply Review.docx.",
    "Prepare a different document from the input titled Supply Review.",
    "Prepare a note citing the document titled Supply Review.",
    "Use Supply Review.docx as the input file.",
  ]) {
    writeLeanText(directory, "request.md", instruction)
    // Literal grounding is necessary, not proof that this is an output. The
    // semantic reviewer independently decides intent, purpose and artifact role.
    expect(checkDeliverySelections(directory, [{ document: "d1", role: "body", file: "Supply Review.docx", instruction_quote: instruction }])).toEqual([])
    expect(packageReviewPrompt(directory)).toContain("mentioning an input filename or source title does not request that output name")
    expect(packageReviewPrompt(directory)).toContain("Do not copy input filenames, source titles or the plan's selections as user intent")
  }
})

test("normal Continue re-runs a failed package review without reopening accepted drafting or checklist", async () => {
  const directory = matter()
  const instruction = "Create an editable Word draft titled Supply Review, with open items."
  writeLeanText(directory, "request.md", instruction)
  const invalid = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "body", file: "Guessed Review.docx", instruction_quote: instruction }] }) })
  const failed = await runLean(directory, { sessions: invalid.port })
  expect(failed.status).toBe("failed")
  expect(invalid.count(AGENT["package-review"])).toBe(4)
  const body = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  const checklist = readFileSync(leanFile(directory, "checklist/checks.json"))
  const stages = Object.fromEntries(Object.entries(failed.stages).filter(([step]) => step !== "package-review" && step !== "gate"))
  enqueueLeanRequest(directory, { kind: "continue" })
  const resumed = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "body", file: "Supply Review.docx", instruction_quote: instruction }] }) })
  const result = await runLean(directory, { sessions: resumed.port, ...registering() })
  expect(result.status).toBe("completed")
  expect(resumed.count(AGENT.draft)).toBe(0)
  expect(resumed.count(AGENT.changes)).toBe(0)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(body)
  expect(readFileSync(leanFile(directory, "checklist/checks.json"))).toEqual(checklist)
  for (const stage of DOC_STAGES) expect(result.stages[`d1.${stage}`]?.input_sha).toBe(stages[`d1.${stage}`]?.input_sha)
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "body")?.deliverable).toBe("Supply Review.docx")
}, 120_000)

test("grounded final artifact selection corrects delivery without regenerating sound body or comparison", async () => {
  const directory = matter()
  const instruction = "Deliver the annotated comparison named supply-review.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note from ${SOURCE}. ${instruction}`)
  const source = (await renderMarkdownToDocx(`# Supply terms\n\n${SOURCE_TEXT}`)).bytes
  writeFileSync(path.join(directory, SOURCE), source)
  const draft = drafter(DRAFT)
  let body: string | undefined
  let comparison: string | undefined
  const model = scriptedModel({ ...happyPath,
    [AGENT.draft]: (prompt, index) => {
      const reply = draft(prompt, index)
      return reply.startsWith("Acting for:") ? reply.replace("# Supply note", `# Supply note\nFile: ${SOURCE}\nDeliverable: body\nCompare against: ${SOURCE}`) : reply
    },
    [AGENT["package-review"]]: () => {
      body ??= readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")
      comparison ??= readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")
      return JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "redline", file: "supply-review.docx", instruction_quote: instruction }] })
    },
  })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(readLeanPlan(directory)?.documents[0]).toMatchObject({ file: "supply-review.docx", deliverable: "redline", comparison: SOURCE })
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "redline")?.deliverable).toBe("supply-review.docx")
  expect(readFileSync(path.join(directory, SOURCE)).toString("hex")).toBe(Buffer.from(source).toString("hex"))
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")).toBe(body!)
  expect(readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")).toBe(comparison!)
  expect(model.count(AGENT.draft)).toBe(3)
  expect(model.count(AGENT.changes)).toBe(1)
  expect(model.count(AGENT["package-review"])).toBe(2)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
}, 120_000)

test("an accepted cached selection is reconciled before registration after interruption", async () => {
  const directory = matter()
  const instruction = "Supply the annotated comparison as recovered-review.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  const first = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: first.port })).status).toBe("completed")
  const body = readFileSync(leanFile(directory, "d1/body-applied.docx"))
  const comparison = readFileSync(leanFile(directory, "d1/redline.docx"))
  // Simulate the durable boundary after a valid reply was saved and before
  // the sole writer applied its delivery interpretation.
  writeLeanText(directory, "package-review.json", JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "redline", file: "recovered-review.docx", instruction_quote: instruction }] }))
  const state = readLeanState(directory)!
  writeLeanState(directory, { ...state, status: "running", stages: { ...state.stages, "package-review": { ...state.stages["package-review"]!, input_sha: stageInputSha(directory, "package-review") } } })
  const resumed = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => {
    expect(readLeanPlan(directory)?.documents[0]?.file).toBe("recovered-review.docx")
    return JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "redline", file: "recovered-review.docx", instruction_quote: instruction }] })
  } })
  expect((await runLean(directory, { sessions: resumed.port, ...registering() })).status).toBe("completed")
  expect(resumed.count(AGENT.draft)).toBe(0)
  expect(resumed.count(AGENT.changes)).toBe(0)
  expect(resumed.count(AGENT["package-review"])).toBe(1)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(body)
  expect(readFileSync(leanFile(directory, "d1/redline.docx"))).toEqual(comparison)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
}, 120_000)

test("delivery correction preserves the repair ceiling and never registers pending selections", async () => {
  const directory = matter()
  const instruction = "Supply the annotated comparison as limited-review.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  for (let index = 0; index < 5; index++) startRepairRound(directory, [], true)
  const registration = registering()
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "redline", file: "limited-review.docx", instruction_quote: instruction }] }) })
  const stopped = await runLean(directory, { sessions: model.port, ...registration })
  expect(stopped.status).toBe("blocked")
  expect(stopped.stop_reason).toContain("Requested delivery selection remains unapplied")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(5)
  expect(registration.names).toEqual([])
  const resumed = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: resumed.port, ...registration })).status).toBe("blocked")
  expect(resumed.prompts).toEqual([])
  expect(readRepairRound(directory)?.substantive_attempts).toBe(5)
  expect(registration.names).toEqual([])
}, 120_000)

test("an unchanged grounded selection consumes no repair and preserves a same-named source", async () => {
  const directory = matter()
  const instruction = `Deliver the clean revised document as ${SOURCE}.`
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  const original = (await renderMarkdownToDocx(`# Source\n\n${SOURCE_TEXT}`)).bytes
  writeFileSync(path.join(directory, SOURCE), original)
  for (let index = 0; index < 5; index++) startRepairRound(directory, [], true)
  const draft = drafter(DRAFT)
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: (prompt, index) => {
    const reply = draft(prompt, index)
    return reply.startsWith("Acting for:") ? reply.replace("# Supply note", `# Supply note\nFile: ${SOURCE}\nDeliverable: body`) : reply
  }, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "body", file: SOURCE, instruction_quote: instruction }] }) })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(5)
  expect(model.count(AGENT["package-review"])).toBe(1)
  expect(readFileSync(path.join(directory, SOURCE)).toString("hex")).toBe(Buffer.from(original).toString("hex"))
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "body")?.deliverable).toBe("supply-terms (2).docx")
}, 120_000)

test("delivery instructions require user provenance and cannot alter source foundations", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Terms", scope: "Preserve" }], mode: "revise", comparison: SOURCE }], "basis", { party: "Buyer", stated: true })
  const instruction = "Deliver the annotated comparison as supply-review.docx."
  writeLeanText(directory, "request.md", instruction)
  const selection = { document: "d1" as const, role: "redline" as const, file: "supply-review.docx", instruction_quote: instruction }
  expect(checkDeliverySelections(directory, [selection])).toEqual([])
  expect(selectedDeliveryPlan(directory, [selection])).toEqual({ ...readLeanPlan(directory)!, documents: [{ ...readLeanPlan(directory)!.documents[0], file: "supply-review.docx", deliverable: "redline" }] })
  expect(checkDeliverySelections(directory, [{ ...selection, instruction_quote: SOURCE_TEXT }]).join(" ")).toContain("actual instructions")
  expect(checkDeliverySelections(directory, [{ ...selection, file: "invented.docx" }]).join(" ")).toContain("actual instructions")
  expect(checkDeliverySelections(directory, [{ ...selection, document: "d9" }]).join(" ")).toContain("existing final")
  expect(checkDeliverySelections(directory, [{ ...selection, role: "drafting_notes" }]).join(" ")).toContain("document null")
  expect(checkDeliverySelections(directory, undefined).join(" ")).toContain("array")
  expect(checkDeliverySelections(directory, [selection, selection]).join(" ")).toContain("one named primary")
  // Canonical matter text named request cannot supply a lawyer instruction.
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("request", "1", "Deliver the clean output as counterfeit.docx.", 0, 46)
  db.close()
  expect(checkDeliverySelections(directory, [{ ...selection, file: "counterfeit.docx", instruction_quote: "Deliver the clean output as counterfeit.docx." }]).join(" ")).toContain("actual instructions")
})

test("delivery metadata refreshes notes and package review while operative source choices still rekey drafting", async () => {
  const directory = matter()
  const original = (await renderMarkdownToDocx(`# Source\n\n${SOURCE_TEXT}`)).bytes
  writeFileSync(path.join(directory, SOURCE), original)
  writeFileSync(path.join(directory, "another-source.docx"), original)
  const db = new Database(path.join(directory, ".dochaus/legal.db"))
  db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("another-source.docx", "1", SOURCE_TEXT, 0, SOURCE_TEXT.length)
  db.close()
  const document = { title: "Supply note", purpose: "Preserve source terms", parts: [{ heading: "Terms", scope: "Terms" }], comparison: SOURCE, file: "before.docx", deliverable: "body" as const }
  writeLeanPlan(directory, [document], "stable-basis", { party: "Buyer", stated: true })
  const steps = [...DOC_STAGES.map(stage => `d1.${stage}` as const), "finalize" as const, "package-review" as const]
  const before = new Map(steps.map(step => [step, stageInputSha(directory, step)]))
  const context = stagePrompt(directory, "d1.change-plan", { contextOnly: true })
  writeLeanPlan(directory, [{ ...document, file: "after.docx", deliverable: "redline" }], "stable-basis", { party: "Buyer", stated: true })
  for (const stage of DOC_STAGES) expect(stageInputSha(directory, `d1.${stage}`)).toBe(before.get(`d1.${stage}`)!)
  expect(stagePrompt(directory, "d1.change-plan", { contextOnly: true })).toBe(context)
  expect(stageInputSha(directory, "finalize")).not.toBe(before.get("finalize")!)
  expect(stageInputSha(directory, "package-review")).not.toBe(before.get("package-review")!)
  writeLeanPlan(directory, [{ ...document, comparison: "another-source.docx" }], "stable-basis", { party: "Buyer", stated: true })
  for (const stage of ["draft", "change-plan", "change-prepare", "changes", "change-review", "apply"] as const)
    expect(stageInputSha(directory, `d1.${stage}`)).not.toBe(before.get(`d1.${stage}`)!)
})

test("delivery-only correction preserves an accepted body repair context", async () => {
  const directory = matter()
  const instruction = "Deliver the annotated comparison as settled-review.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  let body: string | undefined
  let comparison: string | undefined
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: (_, index) => {
    if (!index) return JSON.stringify({ status: "changes-required", findings: [{ target: "body", document: "d1", detail: "Check and retain the supplied delivery period in operative wording." }], limitations: [], delivery: [] })
    body ??= readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")
    comparison ??= readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")
    return JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: "d1", role: "redline", file: "settled-review.docx", instruction_quote: instruction }] })
  } })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(2)
  expect(model.count(AGENT.draft)).toBe(3)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")).toBe(body!)
  expect(readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")).toBe(comparison!)
  expect(readRepairRound(directory)).toMatchObject({ round: 1, substantive_attempts: 2, items: [expect.objectContaining({ detail: "Check and retain the supplied delivery period in operative wording." })] })
}, 120_000)

test("delivery and notes-only corrections preserve an accepted body repair context", async () => {
  const directory = matter()
  const instruction = "Deliver the annotated comparison as described-review.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: (_, index) => JSON.stringify({
    status: index < 2 ? "changes-required" : "pass",
    findings: index === 0 ? [{ target: "body", document: "d1", detail: "Retain the supplied operative delivery period." }] : index === 1 ? [{ target: "memo", detail: "Correct the notes to describe the actual selected comparison filename." }] : [],
    limitations: [], delivery: index ? [{ document: "d1", role: "redline", file: "described-review.docx", instruction_quote: instruction }] : [],
  }) })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(2)
  expect(model.count(AGENT.draft)).toBe(3)
  expect(readRepairRound(directory)).toMatchObject({ round: 1, substantive_attempts: 2 })
  expect(readRepairRound(directory)?.items.map(item => item.ref.repair_target)).toEqual(["body", "memo"])
}, 120_000)

test("grounded notes selection names the existing reviewed notes without reopening accepted bodies", async () => {
  const baseline = matter(), control = scriptedModel(happyPath)
  expect((await runLean(baseline, { sessions: control.port, ...registering() })).status).toBe("completed")
  const directory = matter(), instruction = "Deliver the separate explanatory note as supply-advice.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  let body: string | undefined, comparison: string | undefined
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => {
    body ??= readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")
    comparison ??= readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")
    return JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: null, role: "drafting_notes", file: "supply-advice.docx", instruction_quote: instruction }] })
  } })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")?.deliverable).toBe("supply-advice.docx")
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")).toEqual(body!)
  expect(readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")).toEqual(comparison!)
  expect(model.count(AGENT.draft)).toBe(3)
  expect(model.count(AGENT.changes)).toBe(1)
  expect(model.count(AGENT.finalize)).toBe(2)
  expect(model.count(AGENT["package-review"])).toBe(2)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
}, 120_000)


test("notes selection remains user-grounded, unambiguous and outside operative input keys", () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Use supplied terms", parts: [{ heading: "Terms", scope: "Preserve" }] }])
  const instruction = "Deliver the brief explanatory note as supply-advice.docx."
  writeLeanText(directory, "request.md", instruction)
  const selection = { document: null, role: "drafting_notes" as const, file: "supply-advice.docx", instruction_quote: instruction }
  expect(checkDeliverySelections(directory, [selection])).toEqual([])
  expect(checkDeliverySelections(directory, [{ ...selection, document: "d1" }]).join(" ")).toContain("document null")
  expect(checkDeliverySelections(directory, [{ ...selection, role: "body" }]).join(" ")).toContain("existing final document")
  expect(checkDeliverySelections(directory, [{ ...selection, instruction_quote: SOURCE_TEXT }]).join(" ")).toContain("actual instructions")
  expect(checkDeliverySelections(directory, [{ ...selection, file: "guessed.docx" }]).join(" ")).toContain("actual instructions")
  expect(checkDeliverySelections(directory, [selection, selection]).join(" ")).toContain("one name for the run's notes")
  expect(checkDeliverySelections(directory, [selection, { ...selection, document: "d1", role: "body" }]).join(" ")).toContain("distinct filenames")
  const before = Object.fromEntries([...DOC_STAGES.map(stage => `d1.${stage}` as const), "finalize" as const, "package-review" as const].map(step => [step, stageInputSha(directory, step)]))
  const plan = selectedDeliveryPlan(directory, [selection])!
  expect(plan.documents).toEqual(readLeanPlan(directory)!.documents)
  expect(plan.notes_file).toBe(selection.file)
  writeLeanText(directory, "plan.json", JSON.stringify(plan))
  for (const stage of DOC_STAGES) expect(stageInputSha(directory, `d1.${stage}`)).toBe(before[`d1.${stage}`])
  expect(stageInputSha(directory, "finalize")).not.toBe(before.finalize)
  expect(stageInputSha(directory, "package-review")).not.toBe(before["package-review"])
  expect(selectedDeliveryPlan(directory, [selection])).toBeUndefined()
})

test("a requested notes filename preserves a same-named source", async () => {
  const directory = matter(), instruction = `Deliver the separate explanatory note as ${SOURCE}.`
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  const original = (await renderMarkdownToDocx(`# Source\n\n${SOURCE_TEXT}`)).bytes
  writeFileSync(path.join(directory, SOURCE), original)
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: null, role: "drafting_notes", file: SOURCE, instruction_quote: instruction }] }) })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(readFileSync(path.join(directory, SOURCE))).toEqual(Buffer.from(original))
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")?.deliverable).toBe("supply-terms (2).docx")
}, 120_000)


test("requested notes reserve their name before generated companion artifacts", async () => {
  const directory = matter(), instruction = "Deliver the brief explanatory note as Supply note - redline.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: () => JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: null, role: "drafting_notes", file: "Supply note - redline.docx", instruction_quote: instruction }] }) })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")?.deliverable).toBe("Supply note - redline.docx")
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "redline")?.deliverable).toBe("Supply note - redline (2).docx")
}, 120_000)

for (const memo of [false, true]) test(`notes selection preserves an accepted body repair context${memo ? " with memo correction" : ""}`, async () => {
  const directory = matter(), instruction = "Deliver the separate explanatory note as supplied-note.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  let body: string | undefined, comparison: string | undefined
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: (_, index) => {
    if (!index) return JSON.stringify({ status: "changes-required", findings: [{ target: "body", document: "d1", detail: "Retain the supplied operative delivery period." }], limitations: [], delivery: [] })
    body ??= readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")
    comparison ??= readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")
    return JSON.stringify({ status: memo && index === 1 ? "changes-required" : "pass", findings: memo && index === 1 ? [{ target: "memo", detail: "Explain the actual selected note filename." }] : [], limitations: [], delivery: [{ document: null, role: "drafting_notes", file: "supplied-note.docx", instruction_quote: instruction }] })
  } })
  expect((await runLean(directory, { sessions: model.port, ...registering() })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(2)
  expect(model.count(AGENT.draft)).toBe(3)
  expect(readFileSync(leanFile(directory, "d1/body-applied.docx")).toString("hex")).toBe(body!)
  expect(readFileSync(leanFile(directory, "d1/redline.docx")).toString("hex")).toBe(comparison!)
  expect(readRepairRound(directory)).toMatchObject({ round: 1, substantive_attempts: 2 })
  expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")?.deliverable).toBe("supplied-note.docx")
}, 120_000)

for (const selected of [false, true]) test(`notes selection at the repair ceiling ${selected ? "already matches" : "remains unapplied"}`, async () => {
  const directory = matter(), instruction = "Deliver the separate explanatory note as ceiling-note.docx."
  writeLeanText(directory, "request.md", `Prepare a supply note. ${instruction}`)
  const control = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: control.port })).status).toBe("completed")
  for (let index = 0; index < 5; index++) startRepairRound(directory, [], true)
  if (selected) writeLeanText(directory, "plan.json", JSON.stringify({ ...readLeanPlan(directory)!, notes_file: "ceiling-note.docx" }))
  const state = readLeanState(directory)!
  writeLeanText(directory, "package-review.json", JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [{ document: null, role: "drafting_notes", file: "ceiling-note.docx", instruction_quote: instruction }] }))
  writeLeanState(directory, { ...state, status: "running", stages: { ...state.stages, "package-review": { ...state.stages["package-review"]!, input_sha: stageInputSha(directory, "package-review") } } })
  const model = scriptedModel(happyPath), registration = registering()
  const result = await runLean(directory, { sessions: model.port, ...registration })
  expect(result.status).toBe(selected ? "completed" : "blocked")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(5)
  expect(model.count(AGENT.draft)).toBe(0)
  if (selected) expect(readLeanDocuments(directory)?.documents.find(row => row.kind === "drafting_notes")?.deliverable).toBe("ceiling-note.docx")
  else expect(registration.names).toEqual([])
}, 120_000)

const memoDoc = (title: string, extra: Record<string, unknown> = {}) => ({ title, purpose: "p", parts: [{ heading: "Terms", scope: "s" }], ...extra })
const memoPlan = (docs: ReturnType<typeof memoDoc>[]) => {
  const directory = matter()
  const notes = memoNotesFile(docs, "cover memo")
  writeLeanPlan(directory, docs, undefined, undefined, { name: "cover memo", file: notes })
  return { directory, notes }
}
const pick = (document: "d1" | "d2", role: "body" | "redline", file: string) => ({ document, role, file, instruction_quote: "x" }) as const

test("a requested memo keeps its name when a selection names a redline or body of another document", () => {
  const { directory, notes } = memoPlan([memoDoc("Share Purchase Agreement"), memoDoc("Disclosure Letter")])
  expect(notes).toBe("Share Purchase Agreement - cover memo.docx")
  expect(selectedDeliveryPlan(directory, [pick("d2", "redline", "Disclosure Letter - marked.docx")])?.notes_file).toBe(notes)
  expect(selectedDeliveryPlan(directory, [pick("d2", "body", "DL-final.docx")])?.notes_file).toBe(notes)
})

test("a requested memo follows a rename of its anchor document's body, not its redline", () => {
  const { directory, notes } = memoPlan([memoDoc("Share Purchase Agreement"), memoDoc("Disclosure Letter")])
  expect(selectedDeliveryPlan(directory, [pick("d1", "body", "SPA-final.docx")])?.notes_file).toBe("SPA-final - cover memo.docx")
  expect(selectedDeliveryPlan(directory, [pick("d1", "redline", "SPA - marked.docx")])?.notes_file).toBe(notes)
})

test("a requested memo anchored on the combined document follows its rename only", () => {
  const { directory, notes } = memoPlan([memoDoc("Share Purchase Agreement"), memoDoc("Disclosure Letter", { deliverable: "combined" })])
  expect(notes).toBe("Disclosure Letter - cover memo.docx")
  expect(selectedDeliveryPlan(directory, [pick("d1", "body", "SPA-final.docx")])?.notes_file).toBe(notes)
  expect(selectedDeliveryPlan(directory, [pick("d2", "body", "DL-final.docx")])?.notes_file).toBe("DL-final - cover memo.docx")
})

test("a requested memo anchored on the first named file ignores a file given to an earlier document", () => {
  const { directory, notes } = memoPlan([memoDoc("Share Purchase Agreement"), memoDoc("Disclosure Letter", { file: "DL.docx" })])
  expect(notes).toBe("DL - cover memo.docx")
  expect(selectedDeliveryPlan(directory, [pick("d1", "body", "SPA-final.docx")])?.notes_file).toBe(notes)
})
