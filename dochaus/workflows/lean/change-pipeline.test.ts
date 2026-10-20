import { afterEach, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { acceptChangePlan, acceptChangeReview, checkChangePlan, checkChangeReview, projectedChangeParagraphs } from "./change-pipeline"
import { checkChanges } from "./change-checks"
import { readChangeList, readParagraphs } from "./checks"
import { readLeanEvents } from "./events"
import { finalGate } from "./gate"
import { enqueueLeanRequest } from "./requests"
import { readRepairRound } from "./repairs"
import { restoreLeanStage } from "./restore-stage"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, runLean, stageInputSha } from "./runner"
import { acceptStageReply, LEAN_CHILD_PERMISSIONS, stageOutputValid } from "./stages"
import { answerQuestion, leanFile, readLeanJson, readLeanState, readOpenItems, readQuestions, writeLeanText } from "./state"
import { AGENT, DRAFT, SOURCE, SUPPORT, changes, cleanupMatters, drafter, happyPath, listedFindingIDs, matter, pidOf, sameSession, scriptedModel, type Reply } from "./test-support"

afterEach(cleanupMatters)
const BODY = DRAFT.replace("Payment is due on delivery.", "Payment timing remains to be agreed.")
const PAYMENT = "Payment shall be due on delivery."
const DELIVERY = "The supplier shall deliver the goods within thirty days of the order date."
const PAYMENT_SUPPORT = [{ source_id: SOURCE, quote: "Payment is due on delivery." }]
// The delivery group amends for the review's delivery-period finding.
const plan: Reply = prompt => JSON.stringify({ shared_decisions: ["Use supplied terms; retain unrelated wording."], groups: [
  { id: "delivery", pids: [pidOf(prompt, "The supplier shall deliver the goods.")], goal: "State the source delivery period.", sources: SUPPORT, limitations: [], resolves: listedFindingIDs(prompt) },
  { id: "payment", pids: [pidOf(prompt, "Payment timing remains to be agreed.")], goal: "State agreed payment timing.", sources: PAYMENT_SUPPORT, limitations: [] },
], dismissed: [], questions: [] })
function preparing(badQuote = false): Reply {
  const contexts = new Map<string, string>()
  return (prompt, index, sessionID = "one") => {
    if (prompt.includes("<group_goal_and_evidence>")) contexts.set(sessionID, prompt)
    const context = contexts.get(sessionID)!
    const value = context.includes('"id": "delivery"') ? JSON.parse(changes(context, index)) : { amendments: [{
      pid: pidOf(context, "Payment timing remains to be agreed."), updated: PAYMENT, reason: "States agreed payment timing.", sources: PAYMENT_SUPPORT,
    }], questions: [] }
    if (badQuote) value.amendments[0].sources = [{ source_id: SOURCE, quote: "not present in the canonical source" }]
    return JSON.stringify(value)
  }
}
const reconcile: Reply = prompt => JSON.stringify({ amendments: [
  { pid: pidOf(prompt, "The supplier shall deliver the goods."), updated: DELIVERY, reason: "States source delivery period.", sources: SUPPORT },
  { pid: pidOf(prompt, "Payment timing remains to be agreed."), updated: PAYMENT, reason: "States agreed payment timing.", sources: PAYMENT_SUPPORT },
], questions: [] })
function replies() { return { ...happyPath, [AGENT.draft]: drafter(BODY), [AGENT["change-plan"]]: plan, [AGENT["change-prepare"]]: preparing(), [AGENT.changes]: reconcile } }
function projected(prompt: string) { return /<projected_full_body>\n([\s\S]*?)\n<\/projected_full_body>/.exec(prompt)![1] }

test("independent groups prepare concurrently with focused context; full text is reviewed before apply", async () => {
  const directory = matter(), model = scriptedModel(replies())
  let release!: () => void
  const both = new Promise<void>(resolve => { release = resolve }), started: string[] = []
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT["change-prepare"]) {
      started.push(input.sessionID)
      if (started.length === 2) release()
      await both
      expect(readLeanState(directory)?.stages["d1.change-prepare"]?.session_ids).toHaveLength(2)
      expect(input.text).not.toContain("<findings>")
      expect(input.text).not.toContain(input.text.includes('"id": "delivery"') ? "Payment timing remains to be agreed." : "The supplier shall deliver the goods.")
    }
    if (input.agent === AGENT.changes) {
      expect(input.text).toContain("<prepared_changes>")
      expect(input.text).not.toContain("<findings>")
      expect(input.text).toContain(DELIVERY)
      expect(input.text).toContain(PAYMENT)
    }
    if (input.agent === AGENT["change-review"]) {
      expect(existsSync(leanFile(directory, "d1/body-applied.docx"))).toBe(false)
      expect(projected(input.text)).toContain(DELIVERY)
      expect(projected(input.text)).toContain(PAYMENT)
    }
    return model.port.prompt(input)
  } }
  expect((await runLean(directory, { sessions: port })).status).toBe("completed")
  expect(new Set(started).size).toBe(2)
  model.sessions.filter(row => row.agent === AGENT["change-prepare"]).forEach(row => expect(row.permissions).toEqual(LEAN_CHILD_PERMISSIONS))
  expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)).toContain(PAYMENT)
}, 120_000)

test("canonical support and dependency admission reject only the mutated input", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const value = structuredClone(readLeanJson(directory, "d1/change-plan.json")) as { groups: { pids: string[]; sources: { quote: string }[] }[] }
  expect(checkChangePlan(directory, "d1", value)).toEqual([])
  value.groups[0].sources[0].quote = "absent from the canonical source"
  const sourceIssues = checkChangePlan(directory, "d1", value).join(" ")
  expect(sourceIssues).toContain("not in that document")
  expect(sourceIssues).toContain("The JSON quote field requires verbatim source words.")
  expect(sourceIssues).not.toContain("drop the quotation marks")
  value.groups[0].sources[0].quote = SUPPORT[0].quote
  expect(checkChangePlan(directory, "d1", value)).toEqual([])
  value.groups[1].pids.push(value.groups[0].pids[0])
  expect(checkChangePlan(directory, "d1", value).join(" ")).toContain("more than one group")
}, 120_000)

test("one wrapped JSON stage result is admitted without a formatting retry, with source checks unchanged", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const value = readLeanJson(directory, "d1/change-plan.json") as { groups: { sources: { quote: string }[] }[] }
  const reply = JSON.stringify(value)
  expect(acceptChangePlan(directory, "d1", reply).issues).toEqual([])
  const wrapped = `The supporting passages are verified.\n\n\`\`\`json\n${reply}\n\`\`\`\n`
  expect(acceptChangePlan(directory, "d1", wrapped).issues).toEqual([])
  expect(readLeanJson(directory, "d1/change-plan.json")).toEqual(value)
  const unsupported = structuredClone(value)
  unsupported.groups[0].sources[0].quote = "absent from the canonical source"
  expect(acceptChangePlan(directory, "d1", `Source check complete.\n\`\`\`json\n${JSON.stringify(unsupported)}\n\`\`\``).issues.join(" ")).toContain("not in that document")
  expect(acceptChangePlan(directory, "d1", `${wrapped}\n\`\`\`json\n${reply}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", JSON.stringify([value])).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", `Result:\n\`\`\`json\n${JSON.stringify([value])}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", `Result:\n\`\`\`json\n[${reply}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", `Result:\n[${reply}`).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", reply.slice(0, -1)).issues.length).toBeGreaterThan(0)
}, 120_000)

test("wrapped review results retain material findings and reject ambiguous alternatives", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const value = { status: "changes-required", findings: [{ pid: readParagraphs(directory, "d1")[0].pid, detail: 'Preserve the unresolved condition {delivery timing}.', target: "body" }], limitations: [] }
  const reply = JSON.stringify(value)
  expect(acceptChangeReview(directory, "d1", reply).issues).toEqual([])
  const wrapped = `Review complete.\n\`\`\`json\n${reply}\n\`\`\``
  expect(acceptChangeReview(directory, "d1", wrapped).issues).toEqual([])
  expect(readLeanJson(directory, "d1/change-review.json")).toEqual(value)
  expect(acceptChangeReview(directory, "d1", `${wrapped}\n${JSON.stringify({ status: "pass", findings: [], limitations: [] })}`).issues.length).toBeGreaterThan(0)
  expect(acceptChangeReview(directory, "d1", JSON.stringify([value])).issues.length).toBeGreaterThan(0)
  expect(acceptChangeReview(directory, "d1", `Result:\n\`\`\`json\n${JSON.stringify([value])}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangeReview(directory, "d1", `Result:\n\`\`\`json\n[${reply}\n\`\`\``).issues.length).toBeGreaterThan(0)
}, 120_000)

test("planner question pauses before preparation; canonical answer survives restoring the planner snapshot", async () => {
  const directory = matter(), question = "Which agreed delivery instruction should govern?"
  const model = scriptedModel({ ...replies(), [AGENT["change-plan"]]: prompt => {
    const value = JSON.parse(plan(prompt, 0))
    return JSON.stringify({ ...value, questions: [{ pid: value.groups[0].pids[0], question, kind: "instruction", context: "Conflicting material instructions need a decision." }] })
  } })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  expect(model.count(AGENT["change-prepare"])).toBe(0)
  expect(model.count(AGENT.changes)).toBe(0)
  const waiting = scriptedModel(replies())
  enqueueLeanRequest(directory, { kind: "continue" })
  expect((await runLean(directory, { sessions: waiting.port })).status).toBe("blocked")
  expect(waiting.prompts).toHaveLength(0)
  const asked = readQuestions(directory).find(row => row.question === question)!
  expect(asked.stage).toBe("change-plan")
  answerQuestion(directory, asked.id, { text: "Use the supplied delivery period." })
  restoreLeanStage(directory, "d1.change-plan")
  const resumed = scriptedModel(replies())
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.count(AGENT["change-plan"])).toBe(0)
  expect(resumed.count(AGENT["change-prepare"])).toBe(2)
}, 120_000)

test("queued update settles old workers and discards their entire preparation batch", async () => {
  const directory = matter(), model = scriptedModel(replies())
  let updated = false
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT["change-prepare"] && !updated) {
      updated = true
      enqueueLeanRequest(directory, { kind: "update", text: "Keep source-backed delivery and payment timing." })
    }
    return model.port.prompt(input)
  } }
  expect((await runLean(directory, { sessions: port })).status).toBe("completed")
  expect(model.count(AGENT["change-prepare"])).toBe(4)
  expect(readLeanEvents(directory).filter(row => row.stage === "change-prepare" && row.event === "done")).toHaveLength(1)
  expect(model.prompts.filter(row => row.agent === AGENT.changes)).toHaveLength(1)
  expect(model.prompts.find(row => row.agent === AGENT.changes)?.text).toContain("Keep source-backed")
}, 120_000)

test("a failed worker cannot publish partial preparation or reach reconcile/apply", async () => {
  const directory = matter(), model = scriptedModel(replies())
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT["change-prepare"] && input.text.includes('"id": "payment"')) throw new Error("simulated provider failure")
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions: port })
  expect(state.status).toBe("failed")
  expect(state.stop_reason).toContain("simulated provider failure")
  expect(existsSync(leanFile(directory, "d1/change-preparation.json"))).toBe(false)
  expect(model.count(AGENT.changes)).toBe(0)
  expect(existsSync(leanFile(directory, "d1/body-applied.docx"))).toBe(false)
}, 120_000)

test("projected repair retains prior amendment A while correcting missing effect B before apply", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT.changes]: sameSession((prompt, index) => index ? reconcile(prompt, index) : changes(prompt, index)),
    [AGENT["change-review"]]: prompt => JSON.stringify(projected(prompt).includes(PAYMENT) ? { status: "pass", findings: [], limitations: [] } : {
      status: "changes-required", findings: [{ pid: pidOf(prompt, "Payment timing remains to be agreed."), detail: "Restore accepted payment timing without losing delivery correction." }], limitations: [],
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(2)
  expect(model.count(AGENT["change-plan"])).toBe(2)
  expect(model.count(AGENT["change-prepare"])).toBe(4)
  expect(model.count(AGENT.review)).toBe(1)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
  const repair = model.prompts.filter(row => row.agent === AGENT.changes).at(-1)!.text
  expect(repair).toContain("<previous_changes>")
  expect(repair).toContain(DELIVERY)
  expect(repair).toContain("Restore accepted payment timing")
  expect(readChangeList(directory, "d1")?.amendments).toHaveLength(2)
}, 120_000)

test("material projected defects consume the shared bound and cannot pass through limitations", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: () => JSON.stringify({ status: "changes-required", findings: [{ detail: "Material operative defect remains." }], limitations: ["Recorded in notes"] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("blocked")
  expect(state.stop_reason).toContain(`${LEAN_MAX_SUBSTANTIVE_REPAIRS} substantive repair rounds`)
  expect(model.count(AGENT["change-review"])).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS + 1)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
  expect(existsSync(leanFile(directory, "d1/body-applied.docx"))).toBe(false)
  expect(model.count(AGENT.finalize)).toBe(0)
  expect(checkChangeReview(directory, "d1", { status: "pass", findings: [{ detail: "Still material" }], limitations: [] })).not.toEqual([])
}, 120_000)

test("a repaired intermediate withheld proposal is not falsely described as unapplied at the final gate", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-prepare"]]: preparing(true) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT["change-prepare"])).toBe(8)
  expect(readOpenItems(directory).filter(row => row.stage === "change-prepare" && row.kind === "unapplied_amendment")).toHaveLength(0)
  expect(finalGate(directory).delivery).toBe("execution_ready")
}, 120_000)

test("initial-draft projection requires one replacement/pid while legacy checker preserves overlaps", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const list = readChangeList(directory, "d1")!, paragraphs = readParagraphs(directory, "d1")
  const overlap = { amendments: [list.amendments[0], { ...list.amendments[0], updated: "The supplier must deliver the goods within thirty days of the order date." }], questions: [] }
  expect(checkChanges(directory, JSON.stringify(overlap), paragraphs).issues).toEqual([])
  expect((await acceptStageReply(directory, "d1.changes", JSON.stringify(overlap))).issues.join(" ")).toContain("one full replacement")
  const beforeOwnOutput = stageInputSha(directory, "d1.changes")
  writeLeanText(directory, "d1/changes.json", JSON.stringify({ ...list, ...overlap }))
  expect(stageInputSha(directory, "d1.changes")).toBe(beforeOwnOutput)
  expect(stageOutputValid(directory, "d1.changes")).toBe(false)
  expect(() => projectedChangeParagraphs(directory, "d1")).toThrow("one replacement")
  const keys = ["d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review"] as const
  const original = keys.map(key => stageInputSha(directory, key))
  const changed = structuredClone(readLeanJson(directory, "plan.json")) as { documents: { parts: { heading: string }[] }[] }
  changed.documents[0].parts[0].heading = "Delivery instructions"
  writeLeanText(directory, "plan.json", JSON.stringify(changed))
  keys.forEach((key, index) => expect(stageInputSha(directory, key)).not.toBe(original[index]))
}, 120_000)

test("a saved change list that addresses a paragraph the accepted draft lacks is refused by projection, the checker and the output check", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const list = readChangeList(directory, "d1")!
  const stale = JSON.stringify({ ...list, amendments: [{ ...list.amendments[0], pid: "p222222222222" }] })
  writeLeanText(directory, "d1/changes.json", stale)
  expect(() => projectedChangeParagraphs(directory, "d1")).toThrow("unknown initial paragraph")
  expect(checkChanges(directory, stale, readParagraphs(directory, "d1"), { document: "d1" }).issues.join(" ")).toContain("not a paragraph")
  expect(stageOutputValid(directory, "d1.changes")).toBe(false)
}, 120_000)

test("essential new authority issue from projected review revisits research before amendment work", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", findings: [], limitations: [] } : {
    status: "changes-required", findings: [{ target: "research", detail: "A new essential authority applicability question needs actual retrieval beyond current findings." }], limitations: [],
  }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.research)).toBe(2)
  expect(model.prompts.filter(row => row.agent === AGENT.research).at(-1)?.text).toContain("new essential authority")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
}, 120_000)

test("pre-apply and package semantic failures consume one shared repair allowance", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(),
    [AGENT["change-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", findings: [], limitations: [] } : { status: "changes-required", findings: [{ detail: "One material body correction." }], limitations: [] }),
    [AGENT["package-review"]]: () => JSON.stringify({ status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Explanation still disagrees with final wording." }], limitations: [] }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  expect(model.count(AGENT["change-review"])).toBe(2)
  expect(model.count(AGENT["package-review"])).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
}, 120_000)

test("parallel worker questions are unioned under preparation and pause reconciliation until answered", async () => {
  const directory = matter(), contexts = new Map<string, string>()
  const model = scriptedModel({ ...replies(), [AGENT["change-prepare"]]: (prompt, _, sessionID = "one") => {
    if (prompt.includes("<group_goal_and_evidence>")) contexts.set(sessionID, prompt)
    const context = contexts.get(sessionID)!, delivery = context.includes('"id": "delivery"')
    const original = delivery ? "The supplier shall deliver the goods." : "Payment timing remains to be agreed."
    const pid = pidOf(context, original), question = delivery ? "Which delivery instruction is authorised?" : "Which payment instruction is authorised?"
    return JSON.stringify({ amendments: [{ pid, updated: `${original} [open item: ${question}]`, reason: "Shows a newly identified essential unanswered decision.", sources: delivery ? SUPPORT : PAYMENT_SUPPORT }], questions: [{ pid, question, kind: "instruction", context: "A newly discovered instruction conflict requires the lawyer's decision." }] })
  } })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("blocked")
  const questions = readQuestions(directory).filter(row => row.status === "open")
  expect(questions).toHaveLength(2)
  expect(questions.every(row => row.stage === "change-prepare" && row.document === "d1")).toBe(true)
  expect(readOpenItems(directory).filter(row => row.stage === "change-prepare" && row.kind === "question")).toHaveLength(2)
  expect(model.count(AGENT.changes)).toBe(0)
  expect(existsSync(leanFile(directory, "d1/body-applied.docx"))).toBe(false)
  questions.forEach(row => answerQuestion(directory, row.id, { text: "Use the supplied source instruction." }))
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  expect(readQuestions(directory).filter(row => row.status === "open")).toHaveLength(0)
}, 120_000)

test("a prior final gate pass cannot conceal a material projected-review rejection", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  expect(finalGate(directory).status).toBe("pass")
  const key = stageInputSha(directory, "gate")
  writeLeanText(directory, "d1/change-review.json", JSON.stringify({ status: "changes-required", findings: [{ detail: "Material resulting wording conflicts with an agreed condition." }], limitations: ["Disclosed in old notes"] }))
  expect(stageInputSha(directory, "gate")).not.toBe(key)
  expect(finalGate(directory).status).toBe("blocked")
  expect(finalGate(directory).delivery).toBeNull()
  expect(finalGate(directory).blocking).toContainEqual(expect.objectContaining({ stage: "change-review" }))
}, 120_000)

test("new-pipeline saved changes are canonically rechecked before being reused", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  expect(stageOutputValid(directory, "d1.changes")).toBe(true)
  const list = readChangeList(directory, "d1")!
  list.amendments[0].sources[0].quote = "absent canonical words"
  writeLeanText(directory, "d1/changes.json", JSON.stringify(list))
  expect(stageOutputValid(directory, "d1.changes")).toBe(false)
  const restored = scriptedModel(replies())
  expect((await runLean(directory, { sessions: restored.port })).status).toBe("completed")
  expect(restored.count(AGENT.changes)).toBe(1)
  expect(stageOutputValid(directory, "d1.changes")).toBe(true)
}, 120_000)
