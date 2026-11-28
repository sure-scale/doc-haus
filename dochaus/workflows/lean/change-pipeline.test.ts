import { afterEach, describe, expect, spyOn, test } from "bun:test"
import { existsSync } from "node:fs"
import { acceptChangePlan, acceptChangeReview, changeReviewPrompt, checkChangePlan, checkChangeReview, deferredFindingItems, projectedChangeParagraphs, readChangePlan, reviewDispositionText, unadjudicatedFindings, writerGroup, writerPlan, type ChangeGroup, type ChangePlan, type ReviewDisposition } from "./change-pipeline"
import { checkChanges } from "./change-checks"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import type { ChecklistCheckerInput, ChecklistProducerInput } from "./checklist-contract"
import { mergeRepairAdjudications, type RepairAdjudication } from "./checklist-repair-adjudication"
import { parseJsonObject, readChangeList, readParagraphs } from "./checks"
import { readLeanEvents } from "./events"
import { finalGate, type GateResult } from "./gate"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan } from "./plan"
import { completedDocumentsText, stagePrompt } from "./prompts"
import { openItem } from "./quotes"
import { enqueueLeanRequest } from "./requests"
import { changesRepairSections, readRepairRound } from "./repairs"
import { restoreLeanStage } from "./restore-stage"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, runLean, stageInputSha } from "./runner"
import { acceptStageReply, disclosableItems, LEAN_CHILD_PERMISSIONS, stageOutputValid } from "./stages"
import * as leanState from "./state"
import { answerQuestion, leanFile, readLeanJson, readLeanState, readLeanText, readOpenItems, readQuestions, updateLeanStage, writeLeanText, writeOpenItems, writeQuestions, type Question, type QuestionAdvice } from "./state"
import { AGENT, APPLIED, DRAFT, DRAFTED, LETTER, SOURCE, SOURCE_TEXT, SUPPORT, changes, changesFirst, cleanupMatters, drafter, happyPath, listedFindingIDs, matter, pidOf, reviewFirst, sameSession, scriptedModel, type Reply } from "./test-support"

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

test("a preparation job returns amendments only for paragraphs that change, and a group with none still completes", async () => {
  const directory = matter(), texts: string[] = []
  const prepare = preparing()
  const model = scriptedModel({ ...replies(),
    [AGENT["change-prepare"]]: (prompt, index, sessionID) => {
      if (prompt.includes("<group_goal_and_evidence>")) texts.push(prompt)
      const value = JSON.parse(prepare(prompt, index, sessionID))
      return JSON.stringify(value.amendments.some((row: { updated: string }) => row.updated === PAYMENT) ? { amendments: [], questions: [] } : value)
    },
    [AGENT.changes]: prompt => JSON.stringify({ amendments: [{ pid: pidOf(prompt, "The supplier shall deliver the goods."), updated: DELIVERY, reason: "States source delivery period.", sources: SUPPORT }], questions: [] }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(texts).toHaveLength(2)
  const applied = readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)
  expect(applied).toContain(DELIVERY)
  expect(applied).toContain("Payment timing remains to be agreed.")
}, 120_000)

test("a body review finding names a paragraph of its owning draft; a null pid counts as none and a change id is refused", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const pid = readParagraphs(directory, "d1")[0]!.pid
  const review = (finding: Record<string, unknown>) => ({ status: "changes-required", findings: [{ detail: "A material correction remains.", ...finding }], limitations: [] })
  const missing = "Each body finding needs the pid of a paragraph in its owning document's initial draft: name the paragraph the correction amends or where the missing wording belongs."
  expect(checkChangeReview(directory, "d1", review({ pid }))).toEqual([])
  expect(checkChangeReview(directory, "d1", review({}))).toEqual([missing])
  expect(checkChangeReview(directory, "d1", review({ pid: null, target: "body" }))).toEqual([missing])
  expect(checkChangeReview(directory, "d1", review({ pid: "a4" }))).toEqual(['Finding pid "a4" is not a paragraph of its owning document\'s initial draft; a change id such as a4 is not a paragraph id.'])
  expect(checkChangeReview(directory, "d1", review({ pid: "" }))[0]).toStartWith('Finding pid ""')
  // Research needs no paragraph; its null pid is left out of the saved review.
  expect(checkChangeReview(directory, "d1", review({ target: "research" }))).toEqual([])
  expect(acceptChangeReview(directory, "d1", JSON.stringify(review({ pid: null, target: "research" }))).issues).toEqual([])
  expect(readLeanJson(directory, "d1/change-review.json")).toEqual(review({ target: "research" }))
  expect(acceptChangeReview(directory, "d1", JSON.stringify(review({ pid: null }))).issues).toEqual([missing])
}, 120_000)

// The writer repairs each listed item; the change review only checks whether
// the previous review's findings are still present in the projected body.
test("the change review reads repair items as findings to verify, without the writer's repair instructions or change list", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const pid = readParagraphs(directory, "d1")[0]!.pid
  const item = openItem("draft", "quote_miss", "The quoted figure is not in the supplied terms.", { document: "d1", pid })
  leanState.replaceStageOpenItems(directory, "draft", "d1", [item])
  const writer = changesRepairSections(directory, "d1").join("\n")
  expect(writer).toContain("<required_repairs>")
  expect(writer).toContain("<previous_changes>")
  const prompt = changeReviewPrompt(directory, "d1")
  expect(prompt).toContain("<previous_findings>")
  expect(prompt).toContain(`- ${item.id} [${pid}]: ${item.detail}`)
  for (const writing of ["<required_repairs>", "<previous_changes>"]) expect(prompt).not.toContain(writing)
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

test("a source name the change plan does not know is named with the names it accepts, and an adopted answer is quoted under its own name", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  // The lawyer delegated a question, so its recommended option was adopted.
  const adopted = "Add no separate late-delivery remedy."
  const advice: QuestionAdvice = { options: [{ id: "o1", text: adopted }, { id: "o2", text: "Reduce the price for late delivery." }], recommendation: { option: "o1", rationale: "The terms add no remedy.", basis: SUPPORT } }
  writeQuestions(directory, [...readQuestions(directory), { id: "q-late", stage: "intake", question: "Should late delivery carry a remedy?", context: "The terms are silent.", status: "open", kind: "judgment", advice }])
  answerQuestion(directory, "q-late", { adopt: true })
  const value = structuredClone(readLeanJson(directory, "d1/change-plan.json")) as { groups: { sources: { source_id: string; quote: string }[] }[]; findings?: unknown[] }
  expect(checkChangePlan(directory, "d1", value)).toEqual([])
  const named = `"request" for the request and the lawyer's answers, "adopted:q-late" for an adopted answer, or an approved research URL.`
  value.groups[0].sources[0].source_id = "terms.docx"
  expect(checkChangePlan(directory, "d1", value)).toEqual([`groups[0].sources[0].source_id "terms.docx" is not a known source name. Use the exact name of a matter document, ${named}`])
  value.groups[0].sources[0] = { source_id: SOURCE, quote: " " }
  expect(checkChangePlan(directory, "d1", value)).toEqual(["groups[0].sources[0].quote must be non-empty."])
  // Findings name their sources the same way.
  value.groups[0].sources[0] = SUPPORT[0]
  const finding = (source: { source_id: string; quote: string }) => ({ disposition: "deferred", kind: "information", detail: "The remedy for late delivery stays as adopted.", reason: "The lawyer delegated it.", documents: ["d1"], sources: [source], limitations: [] })
  value.findings = [finding({ source_id: "terms.docx", quote: adopted })]
  expect(checkChangePlan(directory, "d1", value)).toEqual([`findings[0].sources[0].source_id "terms.docx" is not a known source name. Use the exact name of a matter document, ${named}`])
  // The adopted words are not the lawyer's: under "request" they are refused and pointed to their own name, without the generic copy advice.
  value.findings = [finding({ source_id: "request", quote: adopted })]
  const asRequest = checkChangePlan(directory, "d1", value).join(" ")
  expect(asRequest).toContain(`Quote them under the source name "adopted:q-late".`)
  expect(asRequest).not.toContain("read-document")
  value.findings = [finding({ source_id: "adopted:q-late", quote: adopted })]
  expect(checkChangePlan(directory, "d1", value)).toEqual([])
  value.findings = [finding({ source_id: "adopted:q-late", quote: "Add a late-delivery remedy." })]
  expect(checkChangePlan(directory, "d1", value).join(" ")).toContain("not in the recommendation adopted for that question")

  // The change list's amendments accept the adopted answer the same way.
  const list = readChangeList(directory, "d1")!, paragraphs = readParagraphs(directory, "d1")
  const amended = (source: { source_id: string; quote: string }) => JSON.stringify({ amendments: [{ ...list.amendments[0], sources: [source] }], questions: [] })
  expect(checkChanges(directory, amended({ source_id: "adopted:q-late", quote: adopted }), paragraphs).issues).toEqual([])
  const changeIssues = checkChanges(directory, amended({ source_id: "request", quote: adopted }), paragraphs).issues.join(" ")
  expect(changeIssues).toContain(`Quote them under the source name "adopted:q-late".`)
  expect(changeIssues).not.toContain("read-document")
  expect(checkChanges(directory, amended({ source_id: "terms.docx", quote: adopted }), paragraphs).issues)
    .toContain(`amendments[0].sources[0].source_id "terms.docx" is not a known source name. Use the exact name of a matter document, ${named}`)
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
  // A second block of the same shape is the model's final answer; one of another shape leaves the reply ambiguous.
  expect(acceptChangePlan(directory, "d1", `${wrapped}\n\`\`\`json\n${reply}\n\`\`\``).issues).toEqual([])
  expect(acceptChangePlan(directory, "d1", `${wrapped}\n\`\`\`json\n{}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", JSON.stringify([value])).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", `Result:\n\`\`\`json\n${JSON.stringify([value])}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", `Result:\n\`\`\`json\n[${reply}\n\`\`\``).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", `Result:\n[${reply}`).issues.length).toBeGreaterThan(0)
  expect(acceptChangePlan(directory, "d1", reply.slice(0, -1)).issues.length).toBeGreaterThan(0)
}, 120_000)

test("a question naming an internal id is retried with how to say it in words, and asked in words on the last attempt", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(replies()).port })).status).toBe("completed")
  const pid = readParagraphs(directory, "d1").find(row => row.text === "Payment timing remains to be agreed.")!.pid
  const naming = { pid, question: `Should paragraph ${pid} keep its payment timing?`, kind: "instruction", context: "The answer changes d1." }
  const plain = { pid, question: 'Should "Payment timing remains to be agreed." keep its payment timing?', kind: "instruction", context: "The answer changes the supply note." }
  const value = readLeanJson(directory, "d1/change-plan.json") as Record<string, unknown>
  const planned = (question: Record<string, unknown>) => JSON.stringify({ ...value, questions: [question] })
  expect(acceptChangePlan(directory, "d1", planned(naming)).issues).toEqual([
    `questions[0].question names the internal id "${pid}". The lawyer sees only words: name the paragraph by its heading or quote its words.`,
    'questions[0].context names the internal id "d1". The lawyer sees only words: name the document by its title.',
  ])
  expect(acceptChangePlan(directory, "d1", planned(plain)).issues).toEqual([])
  const last = acceptChangePlan(directory, "d1", planned(naming), { final: true })
  expect(last.issues).toEqual([])
  const saved = (readLeanJson(directory, "d1/change-plan.json") as { questions: { question: string; context: string }[] }).questions[0]!
  for (const question of [last.questions![0]!, saved]) {
    expect(`${question.question} ${question.context}`).not.toMatch(new RegExp(`${pid}|\\bd1\\b`))
    expect(question.question).toContain("payment timing")
  }

  // The change list's questions are checked the same way, and on the last attempt (withholdUnverified) asked in words.
  const paragraphs = readParagraphs(directory, "d1")
  const changed = (question: Record<string, unknown>) => JSON.stringify({ amendments: [{ pid, updated: "Payment timing remains to be agreed. [open item: confirm the payment timing]", reason: "Payment timing needs the lawyer.", sources: [] }], questions: [question] })
  expect(checkChanges(directory, changed(naming), paragraphs, { document: "d1" }).issues).toEqual([
    `questions[0].question names the internal id "${pid}". The lawyer sees only words: name the paragraph by its heading or quote its words.`,
    'questions[0].context names the internal id "d1". The lawyer sees only words: name the document by its title.',
  ])
  expect(checkChanges(directory, changed(plain), paragraphs, { document: "d1" }).issues).toEqual([])
  const withheld = checkChanges(directory, changed(naming), paragraphs, { document: "d1", withholdUnverified: true })
  expect(withheld.issues).toEqual([])
  const asked = withheld.changes!.questions[0]!
  expect(`${asked.question} ${asked.context}`).not.toMatch(new RegExp(`${pid}|\\bd1\\b`))
  expect(asked.id).toBe(checkChanges(directory, changed(naming), paragraphs, { document: "d1", withholdUnverified: true }).changes!.questions[0]!.id)
}, 120_000)

test("a JSON syntax error is retried with where it is, so the model can find a stray bracket", () => {
  const directory = matter()
  const issue = (reply: string) => checkChanges(directory, reply, []).issues.join(" ")
  // One brace too many closes the object while its list is still open.
  expect(issue('{"amendments": [{"pid": "p1", "sources": [{"quote": "a"}]} }], "questions": []}'))
    .toContain('at character 59, "]" was expected to close the "amendments" list opened at character 15, but "}" was found.')
  expect(issue('{"amendments": [{"pid": "p1", "sources": [{"quote": "a"}]} }], "questions": []}')).toContain(">>>}],")
  expect(issue('{"amendments": [{"pid": "p1"}, "questions": []}')).toContain('":" was found where it cannot stand')
  expect(issue('{"amendments": [], "questions": [], "note": {"a": 1}')).toContain('the reply ends before "}" closes the object opened at character 0')
  // Brackets inside strings are text, not structure: the break is the stray word after them.
  const stray = '{"amendments": [{"pid": "p1", "reason": "see [3} and \\"}\\""}], "questions": [] x}'
  expect(issue(stray)).toContain(`at character ${stray.indexOf(" x}") + 1}, "," was expected between two values, but "x" was found.`)
  // Prose around the object does not move the place: it is counted in the reply.
  expect(issue(`Here are the changes:\n${stray}`)).toContain(`at character ${stray.indexOf(" x}") + 23}, "," was expected`)
})

test("every stage reply is read as one object: bare, fenced or in prose, never an ambiguous one", () => {
  const object = '{"amendments": [], "note": "kept"}'
  const read = (reply: string) => parseJsonObject(reply)
  for (const reply of [
    object,
    `Here is the result:\n\`\`\`json\n${object}\n\`\`\``,
    `\`\`\`json\n${object}\n\`\`\`\nThat is all.`,
    `\`\`\`JSON\n${object}\n\`\`\``,
    `The result is ${object} as asked.`,
  ]) expect(read(reply)).toEqual({ ok: true, value: { amendments: [], note: "kept" } })
  const refused = (reply: string) => { const parsed = read(reply); return parsed.ok ? "" : parsed.error }
  // A model that drafts its object and then corrects it leaves several blocks: the last is taken when every
  // block has the last's keys and no brace stands outside the fences; any other several are refused.
  const revised = '{"amendments": [{"pid": "p1"}], "note": "revised"}'
  expect(read(`First:\n\`\`\`json\n${object}\n\`\`\`\nOn reflection:\n\`\`\`json\n${revised}\n\`\`\``)).toEqual({ ok: true, value: { amendments: [{ pid: "p1" }], note: "revised" } })
  expect(refused(`\`\`\`json\n${object}\n\`\`\`\n\`\`\`json\n{}\n\`\`\``)).toContain("more than one fenced block")
  expect(refused(`\`\`\`json\n${object}\n\`\`\`\nKeep {this} aside.\n\`\`\`json\n${revised}\n\`\`\``)).toContain("more than one fenced block")
  expect(refused(`\`\`\`json\n${object}\n\`\`\`\n\`\`\`json\n${revised.slice(0, -1)}\n\`\`\``)).toContain("more than one fenced block")
  expect(refused(`Keep {this} aside.\n\`\`\`json\n${object}\n\`\`\``)).toContain("ambiguous")
  expect(refused(`Choose {"a": 1} or {"b": 2}.`)).toContain("not valid JSON")
  expect(refused(`[${object}]`)).toContain("not an object")
  expect(refused(`Result: [${object}]`)).toContain("not an object")
  expect(refused("No object here.")).toBe("No JSON object was found in the reply.")
  // A cut-off reply says where it ends.
  expect(refused('{"amendments": [')).toContain("the reply ends before")
})

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
  // With allowance left the finding starts a repair round: nothing is withheld or left open.
  expect(readChangeList(directory, "d1")?.withheld ?? []).toEqual([])
  expect(readOpenItems(directory).filter(row => row.stage === "change-review" || row.kind === "unapplied_amendment")).toEqual([])
}, 120_000)

test("material projected defects consume the shared bound, then are disclosed with a review draft rather than passing through limitations", async () => {
  const directory = matter()
  // The finding names the heading, a paragraph no amendment touches.
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: () => JSON.stringify({ status: "changes-required", findings: [{ pid: readParagraphs(directory, "d1")[0]!.pid, detail: "Material operative defect remains." }], limitations: ["Recorded in notes"] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("completed")
  expect(model.count(AGENT["change-review"])).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS + 1)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
  // No pending amendment is withheld for it, and it stays open for the lawyer.
  expect(readChangeList(directory, "d1")?.withheld ?? []).toEqual([])
  expect(readChangeList(directory, "d1")?.amendments).toHaveLength(2)
  const item = readOpenItems(directory).find(row => row.stage === "change-review")!
  expect(item).toMatchObject({ kind: "omission", detail: "Material operative defect remains.", ref: { document: "d1" } })
  expect(readLeanText(directory, "drafting-notes.md")).toContain(`[ref: ${item.id}]`)
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(gate.disclosed.map(row => row.id)).toContain(item.id)
  // A finding recorded nowhere still blocks the gate.
  writeOpenItems(directory, readOpenItems(directory).filter(row => row.id !== item.id))
  expect(finalGate(directory).status).toBe("blocked")
  expect(checkChangeReview(directory, "d1", { status: "pass", findings: [{ detail: "Still material" }], limitations: [] })).not.toEqual([])
}, 120_000)

test("a material finding on a pending amendment, once the repair allowance is used up, withholds that amendment for a review draft", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: prompt => JSON.stringify({ status: "changes-required",
    findings: [{ pid: pidOf(prompt, "Payment timing remains to be agreed."), detail: "The payment wording changes the agreed effect." }], limitations: [] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("completed")
  expect(model.count(AGENT["change-review"])).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS + 1)
  const pid = readParagraphs(directory, "d1").find(row => row.text === "Payment timing remains to be agreed.")!.pid
  const list = readChangeList(directory, "d1")!
  expect(list.amendments.map(row => row.updated)).toEqual([DELIVERY])
  expect(list.withheld).toEqual([expect.objectContaining({ amendment: expect.objectContaining({ pid, updated: PAYMENT }), issues: ["The payment wording changes the agreed effect."] })])
  const applied = readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)
  expect(applied).toContain(DELIVERY)
  expect(applied).toContain("Payment timing remains to be agreed.")
  expect(applied).not.toContain(PAYMENT)
  expect(readOpenItems(directory).filter(row => row.stage === "change-review")).toEqual([])
  const item = readOpenItems(directory).find(row => row.stage === "changes" && row.kind === "unapplied_amendment")!
  expect(item.ref).toMatchObject({ document: "d1", pid, amendment_id: list.withheld![0]!.amendment.id })
  expect(item.detail).toContain("The payment wording changes the agreed effect.")
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(gate.disclosed.map(row => row.id)).toContain(item.id)
}, 120_000)

test("a run stopped just after an exhausted change review is recorded as done keeps the amendment it withheld", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: prompt => JSON.stringify({ status: "changes-required",
    findings: [{ pid: pidOf(prompt, "Payment timing remains to be agreed."), detail: "The payment wording changes the agreed effect." }], limitations: [] }) })
  const update = leanState.updateLeanStage
  const stop = spyOn(leanState, "updateLeanStage").mockImplementation((...args) => {
    const state = update(...args)
    if (args[2] === "d1.change-review" && args[3].status === "done" && (readRepairRound(directory)?.substantive_attempts ?? 0) >= LEAN_MAX_SUBSTANTIVE_REPAIRS) throw new Error("stopped")
    return state
  })
  expect((await runLean(directory, { sessions: model.port })).stop_reason).toContain("stopped")
  stop.mockRestore()
  const reviews = model.count(AGENT["change-review"])
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT["change-review"])).toBe(reviews)
  expect(readChangeList(directory, "d1")?.withheld).toEqual([expect.objectContaining({ amendment: expect.objectContaining({ updated: PAYMENT }) })])
  expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)).not.toContain(PAYMENT)
  const item = readOpenItems(directory).find(row => row.stage === "changes" && row.kind === "unapplied_amendment")!
  expect((readLeanJson(directory, "gate.json") as GateResult).disclosed.map(row => row.id)).toContain(item.id)
}, 120_000)

test("a change review re-run after it withheld an amendment records that amendment once", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: prompt => JSON.stringify({ status: "changes-required",
    findings: [{ pid: pidOf(prompt, "Payment timing remains to be agreed."), detail: "The payment wording changes the agreed effect." }], limitations: [] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const recorded = () => readOpenItems(directory).filter(row => row.stage === "change-review" || row.kind === "unapplied_amendment")
  const before = recorded()
  expect(before).toEqual([expect.objectContaining({ stage: "changes", kind: "unapplied_amendment" })])
  const reviews = model.count(AGENT["change-review"])

  // With the allowance used up, a resume on unchanged inputs does not run the change review again.
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT["change-review"])).toBe(reviews)
  expect(recorded()).toEqual(before)
  expect(readChangeList(directory, "d1")?.withheld).toHaveLength(1)
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft" })
}, 120_000)

test("a change review finding on a paragraph the change checks withheld is recorded with its own text", async () => {
  const directory = matter()
  // Reconciliation keeps citing a passage the source lacks, so the change checks withhold the payment amendment.
  const contexts = new Map<string, string>()
  const unverified: Reply = (prompt, index, sessionID = "one") => {
    if (!contexts.has(sessionID)) contexts.set(sessionID, prompt)
    const value = JSON.parse(reconcile(contexts.get(sessionID)!, index))
    value.amendments[1].sources = [{ source_id: SOURCE, quote: "not present in the canonical source" }]
    return JSON.stringify(value)
  }
  const model = scriptedModel({ ...replies(), [AGENT.changes]: unverified, [AGENT["change-review"]]: prompt => JSON.stringify({ status: "changes-required",
    findings: [{ pid: pidOf(prompt, "Payment timing remains to be agreed."), detail: "The drafted payment wording leaves the due date open." }], limitations: [] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const pid = readParagraphs(directory, "d1").find(row => row.text === "Payment timing remains to be agreed.")!.pid
  // The change checks withheld the payment amendment, not the change review.
  expect(readChangeList(directory, "d1")?.withheld).toEqual([expect.objectContaining({ amendment: expect.objectContaining({ pid }) })])
  expect(readChangeList(directory, "d1")!.withheld![0]!.issues).not.toContain("The drafted payment wording leaves the due date open.")
  const item = readOpenItems(directory).find(row => row.stage === "change-review")!
  expect(item).toMatchObject({ kind: "omission", detail: "The drafted payment wording leaves the due date open.", ref: { document: "d1", pid } })
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(gate.disclosed.map(row => row.id)).toContain(item.id)
}, 120_000)

test("a new change review finding on an already withheld paragraph is recorded when the review runs again", async () => {
  const directory = matter()
  let detail = "The payment wording changes the agreed effect."
  const model = scriptedModel({ ...replies(), [AGENT["change-review"]]: prompt => JSON.stringify({ status: "changes-required",
    findings: [{ pid: pidOf(prompt, "Payment timing remains to be agreed."), detail }], limitations: [] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readOpenItems(directory).filter(row => row.stage === "change-review")).toEqual([])

  // With the allowance used up a resume on unchanged inputs does not run the
  // review again (a re-run is forced here); when it runs, its new finding is recorded.
  detail = "The drafted payment wording leaves the due date open."
  updateLeanStage(directory, readLeanState(directory)!, "d1.change-review", { status: "pending", input_sha: null })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const item = readOpenItems(directory).find(row => row.stage === "change-review")!
  expect(item).toMatchObject({ kind: "omission", detail, ref: { document: "d1" } })
  expect(readOpenItems(directory).filter(row => row.kind === "unapplied_amendment")).toHaveLength(1)
  const gate = readLeanJson(directory, "gate.json") as GateResult
  expect(gate).toMatchObject({ status: "pass", delivery: "review_draft" })
  expect(gate.disclosed.map(row => row.id)).toContain(item.id)
}, 120_000)

test("a repaired intermediate withheld proposal is not falsely described as unapplied at the final gate", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-prepare"]]: preparing(true) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT["change-prepare"])).toBe(8)
  expect(readOpenItems(directory).filter(row => row.stage === "change-prepare" && row.kind === "unapplied_amendment")).toHaveLength(0)
  expect(finalGate(directory).delivery).toBe("execution_ready")
}, 120_000)

test("a preparation job is shown its cited canonical passages, yet a quote outside the source is still refused and never kept", async () => {
  const directory = matter()
  const model = scriptedModel({ ...replies(), [AGENT["change-prepare"]]: preparing(true) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const prepare = model.prompts.filter(row => row.agent === AGENT["change-prepare"])
  const first = prepare.filter(row => row.text.includes("<group_goal_and_evidence>"))
  expect(first).toHaveLength(2)
  for (const row of first) {
    const passages = /<cited_source_passages>\n\n([\s\S]*?)\n\n<\/cited_source_passages>/.exec(row.text)?.[1] ?? ""
    expect(passages).toContain(`[${SOURCE} — `)
    expect(passages).toContain(SOURCE_TEXT)
    // Shared blocks come first; the group's own evidence follows them.
    expect(row.text.indexOf("<adjudicated_dependencies_not_independent_authority>")).toBeLessThan(row.text.indexOf("<cited_source_passages>"))
  }
  const retries = prepare.filter(row => row.text.startsWith("Your reply did not pass the change-prepare checks"))
  expect(retries.length).toBeGreaterThan(0)
  for (const row of retries) {
    expect(row.text).toContain('quotes "not present in the canonical source" as from')
    expect(row.text).toContain("The JSON quote field requires verbatim source words.")
  }
  const preparation = readLeanJson(directory, "d1/change-preparation.json") as { groups: { changes: { amendments: { sources: { quote: string }[] }[] } }[] }
  expect(preparation.groups.flatMap(row => row.changes.amendments).flatMap(row => row.sources).map(row => row.quote)).not.toContain("not present in the canonical source")
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
    [AGENT["change-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", findings: [], limitations: [] } : { status: "changes-required", findings: [{ pid: readParagraphs(directory, "d1")[0]!.pid, detail: "One material body correction." }], limitations: [] }),
    [AGENT["package-review"]]: () => JSON.stringify({ status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Explanation still disagrees with final wording." }], limitations: [] }),
  })
  // The allowance used up, the package review's finding goes with a review draft.
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft" })
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
  writeLeanText(directory, "d1/change-review.json", JSON.stringify({ status: "changes-required", findings: [{ pid: readParagraphs(directory, "d1")[0]!.pid, detail: "Material resulting wording conflicts with an agreed condition." }], limitations: ["Disclosed in old notes"] }))
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

// How a change plan adjudicates review findings by id: the resolves a focused
// checklist merge keeps, what acceptance saves and asks, what amendment writers
// are shown, and when the runner may take the focused path at all.
describe("finding adjudication", () => {
  const PID = "p111111111111", OTHER = "p222222222222"
  const FINDING_ID = /\bf-[0-9a-f]{10}\b/

  const group = (id: string, pids: string[], resolves?: string[]): ChangeGroup =>
    ({ id, pids, goal: `Goal ${id}.`, sources: SUPPORT, limitations: [], ...(resolves ? { resolves } : {}) })
  const disposition = (detail: string, resolves?: string[]): ReviewDisposition =>
    ({ disposition: "deferred", kind: "information", detail, reason: "Still open.", documents: ["d1"], sources: [], limitations: ["Unverified."], ...(resolves ? { resolves } : {}) })
  const question = (text: string): Question & { resolves?: string[] } =>
    ({ id: `q-${text.length}`, stage: "change-plan", question: text, context: "Why it matters.", pid: PID, status: "open", kind: "fact" })

  describe("a focused checklist merge keeps the review findings a plan resolves", () => {
    const previous: ChangePlan = { shared_decisions: [], groups: [group("old", [OTHER], ["f-000000000a"])], dismissed: [], questions: [],
      findings: [disposition("Prior timing assessment.", ["f-000000000b"]), disposition("Unrelated advice.", ["f-000000000c"])] }
    const superseding = (patch: Partial<RepairAdjudication>): RepairAdjudication => ({ shared_decisions: [], groups: [], dismissed: [], findings: [], questions: [],
      superseded_decisions: [], superseded_findings: [{ index: 0, reason: "The source calls for advice.", sources: SUPPORT }], ...patch })

    test("a superseded disposition's resolves pass to the first group, else finding, else question, of the same result", () => {
      const toGroup = mergeRepairAdjudications(previous, [superseding({ groups: [group("timing", [PID]), group("later", ["p333333333333"])], findings: [disposition("Also deferred.")] })])
      expect(toGroup.groups.map(row => row.resolves)).toEqual([["f-000000000b"], undefined])
      expect(toGroup.findings).toEqual([previous.findings![1]!, disposition("Also deferred.")])

      const toFinding = mergeRepairAdjudications(previous, [superseding({ findings: [disposition("Deferred again."), disposition("Second.")], questions: [question("Is approval given?")] })])
      expect(toFinding.findings!.map(row => row.resolves)).toEqual([["f-000000000c"], ["f-000000000b"], undefined])
      expect(toFinding.questions.slice(-1)[0]!.resolves).toBeUndefined()

      const toQuestion = mergeRepairAdjudications(previous, [superseding({ questions: [question("Is approval given?")] })])
      expect(toQuestion.questions.map(row => row.resolves)).toEqual([["f-000000000b"]])
    })

    test("a merged overlapping group keeps the union of its parts' resolves", () => {
      const merged = mergeRepairAdjudications(previous, [
        superseding({ groups: [group("timing", [PID], ["f-000000000d"])] }),
        superseding({ superseded_findings: [], groups: [group("dependent", [PID, OTHER], ["f-000000000e", "f-000000000d"])] }),
      ])
      expect(merged.groups).toHaveLength(1)
      expect(merged.groups[0]!.resolves!.toSorted()).toEqual(["f-000000000b", "f-000000000d", "f-000000000e"])
    })
  })

  describe("acceptChangePlan", () => {
    const OWN = "f-00000000a1", SECOND = "f-00000000a2"
    /** One planned document with one paragraph and a saved review of `findings`. */
    function reviewed(...findings: string[]) {
      const directory = matter()
      writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record the supply terms", file: "supply note.docx", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
      writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "initial", text: DRAFTED }]))
      writeLeanText(directory, "d1/review.json", JSON.stringify({ findings: findings.map((id, index) => ({ id, document: "d1", origin: "drafting", finding: `Finding ${index + 1}.`, sources: [] })) }))
      return directory
    }
    const reply = (patch: Partial<Record<keyof ChangePlan, unknown>>) => JSON.stringify({ shared_decisions: [], groups: [], findings: [], questions: [], ...patch })
    const ask = (text: string, resolves?: string[]) => ({ pid: PID, question: text, kind: "fact", context: "Why it matters.", ...(resolves ? { resolves } : {}) })

    test("a question's resolves are saved in the change plan only; the questions to ask carry none", () => {
      const directory = reviewed(OWN)
      const accepted = acceptChangePlan(directory, "d1", reply({ questions: [ask("Is approval given?", [OWN])] }))
      expect(accepted.issues).toEqual([])
      expect(accepted.questions).toHaveLength(1)
      expect(accepted.questions![0]).not.toHaveProperty("resolves")
      expect((readLeanJson(directory, "d1/change-plan.json") as ChangePlan).questions[0]!.resolves).toEqual([OWN])
    })

    test("a plan without dismissed is accepted, and each list of the wrong shape is named with the shape it takes", () => {
      const directory = reviewed(OWN)
      const findings = [disposition("Kept open.", [OWN])]
      expect(acceptChangePlan(directory, "d1", reply({ findings })).issues).toEqual([])
      const shaped = acceptChangePlan(directory, "d1", reply({ shared_decisions: [{ decision: "Keep the term." }], dismissed: [{ finding_id: OWN, reason: "Not needed." }], findings }))
      expect(shaped.issues).toEqual(["Return one JSON object with shared_decisions, groups, findings and questions lists. Its shared_decisions is the wrong shape; give a list of nonempty text, possibly empty. Its dismissed is the wrong shape; give a list of text, or leave it out."])
      expect(acceptChangePlan(directory, "d1", JSON.stringify({ shared_decisions: [], findings, questions: [] })).issues).toEqual(["Return one JSON object with shared_decisions, groups, findings and questions lists. Its groups is missing; give a list of change groups, possibly empty."])
    })

    test("an id that names no listed finding is dropped without a retry", () => {
      const directory = reviewed(OWN)
      expect(acceptChangePlan(directory, "d1", reply({ findings: [disposition("Kept open.", [OWN, "f-00000000ff"])] })).issues).toEqual([])
      expect((readLeanJson(directory, "d1/change-plan.json") as ChangePlan).findings![0]!.resolves).toEqual([OWN])
    })

    test("a listed finding left out is retried before the last attempt and saved on it", () => {
      const directory = reviewed(OWN, SECOND)
      const plan = reply({ findings: [disposition("Kept open.", [OWN])] })
      const retried = acceptChangePlan(directory, "d1", plan)
      expect(retried.issues).toHaveLength(1)
      expect(retried.issues[0]).toContain(`These review findings are not adjudicated: ${SECOND}.`)
      expect(retried.issues[0]).not.toContain("An answered question")
      expect(readLeanText(directory, "d1/change-plan.json")).toBeUndefined()
      expect(acceptChangePlan(directory, "d1", plan, { final: true }).issues).toEqual([])
      expect(readLeanText(directory, "d1/change-plan.json")).toBeDefined()
    })

    test("the answered-question sentence appears only when the answered question carried resolves", () => {
      const directory = reviewed(OWN, SECOND)
      writeQuestions(directory, [{ id: "q-answered", stage: "change-plan", document: "d1", pid: PID, question: "Is approval given?", context: "Why it matters.", status: "answered", answer: "Yes." }])
      const findings = [disposition("Kept open.", [OWN])]
      const carried = acceptChangePlan(directory, "d1", reply({ findings, questions: [ask("Is approval given?", [SECOND])] }))
      expect(carried.issues[0]).toContain(`not adjudicated: ${SECOND}.`)
      expect(carried.issues[0]).toContain("An answered question no longer adjudicates a finding")
      const bare = acceptChangePlan(directory, "d1", reply({ findings, questions: [ask("Is approval given?")] }))
      expect(bare.issues[0]).toContain(`not adjudicated: ${SECOND}.`)
      expect(bare.issues[0]).not.toContain("An answered question")
    })
  })

  describe("amendment writers are shown no finding ids", () => {
    test("writerPlan and writerGroup strip resolves and settled groups", () => {
      const plan: ChangePlan = { shared_decisions: ["Keep payment."], groups: [group("timing", [PID], ["f-000000000a"])], dismissed: [],
        findings: [disposition("Kept open.", ["f-000000000b"])], questions: [{ ...question("Is approval given?"), resolves: ["f-000000000c"] }], settled_groups: [group("old", [OTHER], ["f-000000000d"])] }
      const shown = JSON.stringify(writerPlan(plan))
      expect(shown).not.toContain("resolves")
      expect(shown).not.toContain("settled_groups")
      expect(shown).not.toMatch(FINDING_ID)
      expect(writerPlan(plan).groups).toEqual([group("timing", [PID])])
      expect(writerGroup(plan.groups[0]!)).toEqual(group("timing", [PID]))
    })

    test("in a run, only the change plan prompt lists finding ids; preparation, changes and change review see none", async () => {
      const directory = matter()
      // The plan amends for every listed finding in one group, so preparation runs.
      const plan: Reply = prompt => JSON.stringify({ shared_decisions: [], dismissed: [], findings: [], questions: [],
        groups: [{ id: "delivery", pids: [pidOf(prompt, DRAFTED)], goal: "State the delivery period.", sources: SUPPORT, limitations: [], resolves: listedFindingIDs(prompt) }] })
      const model = scriptedModel({ ...happyPath, [AGENT["change-plan"]]: plan })
      expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
      const texts = (agent: string) => model.prompts.filter(row => row.agent === agent).map(row => row.text)
      expect(texts(AGENT["change-plan"]).some(text => FINDING_ID.test(text))).toBe(true)
      for (const agent of [AGENT["change-prepare"], AGENT.changes, AGENT["change-review"]]) {
        expect([agent, texts(agent).length > 0]).toEqual([agent, true])
        expect([agent, texts(agent).filter(text => FINDING_ID.test(text))]).toEqual([agent, []])
      }
    }, 120_000)
  })

  describe("the focused checklist repair path", () => {
    const FOCUSED = "ONE supplied checklist assessment"
    // One request requirement the checker finds unsatisfied, so a checklist
    // repair round starts after the first pass. The focused adjudication defers
    // it; a full plan rejects every listed review finding.
    function replies(fail: { once: boolean }): Record<string, Reply> {
      return {
        ...happyPath,
        [AGENT["change-plan"]]: text => {
          if (text.includes(FOCUSED)) {
            if (fail.once) { fail.once = false; throw new Error("Scripted adjudication failure") }
            return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [], superseded_findings: [], superseded_decisions: [], findings: [
              { disposition: "deferred", kind: "information", detail: "The assessed requirement still needs a correction.", reason: "No supported correction.", documents: ["d1"], sources: [], limitations: ["No supported correction identified."] }] })
          }
          const resolves = listedFindingIDs(text)
          return JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [], findings: resolves.length ? [
            { disposition: "rejected", kind: "suggestion", detail: "The scripted plan adjudicates the review findings.", reason: "The scripted changes stage decides the amendments.", documents: ["d1"], sources: [], limitations: ["Scripted fixture."], resolves }] : [] })
        },
        [CHECKLIST_JOB_AGENTS.extract]: text => {
          const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!) as ChecklistProducerInput
          return JSON.stringify({ items: input.source.document === "request" ? [{ check: "The requested supply note is drafted.", basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } }] : [] })
        },
        [CHECKLIST_JOB_AGENTS.check]: text => {
          const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!) as ChecklistCheckerInput
          return JSON.stringify({ item_id: input.item_id, status: "unsatisfied", reason: "Focused result unsatisfied.", evidence: input.current_evidence,
            source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected." })
        },
      }
    }
    const focused = (model: ReturnType<typeof scriptedModel>) => model.prompts.filter(row => row.agent === AGENT["change-plan"] && row.text.includes(FOCUSED)).length

    test("a review finding the saved plan does not name sends the round's change plan to a full plan", async () => {
      const directory = matter()
      const fail = { once: true }
      const first = scriptedModel(replies(fail))
      expect((await runLean(directory, { sessions: first.port })).status).toBe("failed")
      // The saved plan named every finding, so the round took the focused path.
      expect(focused(first)).toBe(1)
      expect(unadjudicatedFindings(directory, "d1")).toEqual([])

      // A review record that changed since the plan: new finding text, new id.
      const NEW = "f-00000000b1"
      const review = readLeanJson(directory, "d1/review.json") as { findings: unknown[] }
      writeLeanText(directory, "d1/review.json", JSON.stringify({ findings: [...review.findings, { id: NEW, document: "d1", origin: "drafting", finding: "The payment paragraph needs a due date.", sources: [] }] }))
      expect(unadjudicatedFindings(directory, "d1").map(row => row.id)).toEqual([NEW])

      enqueueLeanRequest(directory, { kind: "continue" })
      const resumed = scriptedModel(replies(fail))
      expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
      expect(focused(resumed)).toBe(0)
      expect(resumed.prompts.some(row => row.agent === AGENT["change-plan"] && listedFindingIDs(row.text).includes(NEW))).toBe(true)
      expect(unadjudicatedFindings(directory, "d1")).toEqual([])
    }, 180_000)
  })
})

describe("findings handoff", () => {
  const findings: ReviewDisposition[] = [
    { disposition: "deferred", kind: "research", detail: "Check whether an external enforceability rule qualifies the supplied delivery obligation.", reason: "No verified authority resolves that proposition; obtain targeted research before relying on it.", documents: ["d1"], sources: [], limitations: ["Enforceability has not been verified."] },
    { disposition: "deferred", kind: "information", detail: "Confirm the recipient for delivery notices.", reason: "No recipient is supplied; complete this before issue.", documents: ["d1"], sources: [], limitations: ["Recipient not supplied."] },
    { disposition: "recommendation", kind: "suggestion", detail: "Consider a reminder before the supplied delivery deadline.", reason: "This is an optional practical proposal; the agreed period stays unchanged.", documents: ["d1"], sources: SUPPORT, limitations: [] },
    { disposition: "rejected", kind: "suggestion", detail: "Replace the supplied delivery period with immediate delivery.", reason: "The source specifies a period; no instruction supports shortening it.", documents: ["d1"], sources: SUPPORT, limitations: [] },
  ]
  const plan = (rows = findings): ChangePlan => ({ shared_decisions: [], groups: [], dismissed: [], questions: [], findings: structuredClone(rows) })

  function setup() {
    const directory = matter()
    writeLeanPlan(directory, [
      { title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Supplied obligation" }] },
      { title: "Client advice", purpose: "Explain the supply note and remaining choices", parts: [{ heading: "Advice", scope: "Consequences and remaining action" }] },
    ])
    writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: "p111111111111", sha256: "initial", text: "Delivery is due under the supplied terms." }]))
    writeLeanText(directory, "d1/paragraphs-applied.json", JSON.stringify([{ pid: "p222222222222", sha256: "final", text: "Delivery remains due under the supplied terms." }]))
    expect(acceptChangePlan(directory, "d1", JSON.stringify(plan())).issues).toEqual([])
    return directory
  }

  test("dispositions survive fresh later writers, notes and review without making optional or rejected suggestions open decisions", () => {
    const directory = setup()
    expect(readChangePlan(directory, "d1")?.findings).toEqual(findings)
    for (const prompt of [completedDocumentsText(directory, "d2"), stagePrompt(directory, "d2.draft"), stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
      findings.forEach(row => expect(prompt).toContain(row.detail))
      expect(prompt).toContain('"origin_document": "d1"')
      expect(prompt).toContain('"disposition": "rejected"')
      expect(prompt).toContain('"sources": []')
      expect(prompt).toContain(SUPPORT[0].quote)
    }
    // Scope names the operative body; its later advice still gets the findings.
    expect(stagePrompt(directory, "d2.draft")).not.toContain("p111111111111")
    expect(deferredFindingItems(directory)).toHaveLength(2)
    expect(disclosableItems(directory)).toEqual(deferredFindingItems(directory))
    expect(readOpenItems(directory)).toEqual([]) // no duplicate persisted ledger
    writeLeanText(directory, "drafting-notes.md", "# Notes\n\nNo remaining work.")
    expect(finalGate(directory).status).toBe("blocked")
    writeLeanText(directory, "drafting-notes.md", `# Notes\n\n${deferredFindingItems(directory).map(row => `${row.id}: ${row.detail}`).join("\n")}`)
    expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
  })

  test("admission is causal for evidence, scope, uncertainty and duplicate dispositions", () => {
    const directory = setup()
    expect(checkChangePlan(directory, "d1", plan())).toEqual([])
    const unsupported = plan()
    unsupported.findings![3].sources[0].quote = "Unsupported words absent from canonical text."
    expect(checkChangePlan(directory, "d1", unsupported).join(" ")).toContain("not in that document")
    const wrongScope = plan()
    wrongScope.findings![0].documents = ["d9"]
    expect(checkChangePlan(directory, "d1", wrongScope).join(" ")).toContain("planned document ids")
    const noLimitation = plan()
    noLimitation.findings![0].limitations = []
    expect(checkChangePlan(directory, "d1", noLimitation).join(" ")).toContain("evidence limitation")
    // A question's kind on a finding is named, with the kinds a finding takes.
    const questionKind = plan()
    questionKind.findings![0].kind = "judgment" as never
    expect(checkChangePlan(directory, "d1", questionKind).join(" ")).toContain('Its kind is "judgment"; use "information", "research" or "suggestion".')
    // Each malformed field is named with the shape it takes.
    const unnamed = plan()
    delete (unnamed.findings![1] as Partial<ReviewDisposition>).reason
    unnamed.findings![1].limitations = "Recipient not supplied." as never
    const message = checkChangePlan(directory, "d1", unnamed).join(" ")
    expect(message).toContain("Its reason is missing; give nonempty text.")
    expect(message).toContain("Its limitations is the wrong shape; give a list of nonempty text, possibly empty.")
    expect(message).not.toContain("Its detail")
    const group = { ...plan(), groups: [{ id: "delivery", pids: [], goal: "Keep the period.", sources: SUPPORT, limitations: [] }] }
    expect(checkChangePlan(directory, "d1", group).join(" ")).toContain("Its pids is the wrong shape; give a nonempty list of initial draft pids.")
    const duplicated = plan([...findings, { ...findings[0], detail: `  ${findings[0].detail.toUpperCase()} `, documents: ["d2"] }])
    expect(checkChangePlan(directory, "d1", duplicated).join(" ")).toContain("consolidate duplicate findings")
    expect(checkChangePlan(directory, "d1", plan())).toEqual([])
  })

  test("a finding without affected documents is routed to the document whose review raised it", () => {
    const directory = setup()
    const unrouted = plan()
    unrouted.findings![1].documents = []
    delete (unrouted.findings![2] as Partial<ReviewDisposition>).documents
    expect(acceptChangePlan(directory, "d2", JSON.stringify(unrouted)).issues).toEqual([])
    expect(readChangePlan(directory, "d2")?.findings?.map(row => row.documents)).toEqual([["d1"], ["d2"], ["d2"], ["d1"]])
  })

  test("changed dispositions invalidate dependent consumers while leaving every draft stable", () => {
    const directory = setup()
    const steps = ["d1.draft", "d2.draft", "d2.review", "d2.change-plan", "finalize", "package-review", "gate"] as const
    const before = steps.map(step => stageInputSha(directory, step))
    const amended = plan()
    amended.findings![2].reason = "Explain this as a reminder option without altering the operative period."
    expect(acceptChangePlan(directory, "d1", JSON.stringify(amended)).issues).toEqual([])
    expect(steps.map((step, index) => stageInputSha(directory, step) !== before[index])).toEqual([false, false, true, true, true, true, true])
  })

  test("legacy dismissal retains review navigation without being treated as evidence-backed rejection", async () => {
    const directory = setup()
    const reviewed = await acceptStageReply(directory, "d1.review-independent", JSON.stringify({ findings: [
      { pid: "p111111111111", origin: "law", balance: "neutral", finding: "Further authority is needed before relying on enforceability.", sources: [] },
    ] }))
    expect(reviewed.issues).toEqual([])
    const id = (readLeanJson(directory, "d1/review-independent.json") as { findings: { id: string }[] }).findings[0]!.id
    // A legacy plan, with a free-text dismissal and no structured dispositions,
    // adjudicates no finding: it is retried, and the last attempt saves it with
    // the finding left open and disclosed.
    const legacy = JSON.stringify({ shared_decisions: [], groups: [], dismissed: ["Defer the research proposition to advice."], questions: [] })
    expect(acceptChangePlan(directory, "d1", legacy).issues.join(" ")).toContain(`These review findings are not adjudicated: ${id}.`)
    expect(acceptChangePlan(directory, "d1", legacy, { final: true }).issues).toEqual([])
    expect(deferredFindingItems(directory)).toContainEqual(expect.objectContaining({ stage: "change-plan", kind: "omission",
      detail: "Further authority is needed before relying on enforceability.", ref: expect.objectContaining({ document: "d1", finding: id }) }))
    const context = reviewDispositionText(directory, "d1")
    expect(context).toContain("legacy_decisions")
    expect(context).toContain("unadjudicated_review_history")
    expect(context).toContain("Further authority is needed")
    expect(context).not.toContain('"disposition": "rejected"')
    expect(completedDocumentsText(directory, "d2")).not.toContain("p111111111111")
  })

  test("prior-document free-text pointers are navigation while saved records and canonical quotes stay intact", () => {
    const directory = setup()
    const referenced = plan()
    referenced.findings![0].detail += " Review [p111111111111]."
    referenced.findings![0].reason += " Final location p222222222222."
    referenced.findings![0].limitations.push("Unverified for [p111111111111].")
    referenced.dismissed = ["Earlier idea at [p111111111111] was out of scope."]
    expect(acceptChangePlan(directory, "d1", JSON.stringify(referenced)).issues).toEqual([])
    const earlier = completedDocumentsText(directory, "d2")
    expect(earlier).not.toContain("p111111111111")
    expect(earlier).not.toContain("p222222222222")
    expect(earlier).toContain("earlier paragraph")
    expect(earlier).toContain(SUPPORT[0].quote)
    expect(readChangePlan(directory, "d1")).toEqual(referenced)
    expect(packageReviewPrompt(directory)).toContain("p111111111111")
  })

  test("coordinator applies accepted amendments while deferred findings survive apply and durable resume", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.changes]: changesFirst, [AGENT.review]: reviewFirst,
      [AGENT["change-plan"]]: prompt => {
        const own = /<draft_paragraphs>\n([\s\S]*?)\n<\/draft_paragraphs>/.exec(prompt)?.[1] ?? ""
        if (!own.includes("The supplier shall deliver the goods.")) return JSON.stringify(plan([]))
        return JSON.stringify({ ...plan(), groups: [{ id: "delivery", pids: [pidOf(own, "The supplier shall deliver the goods.")], goal: "Retain the supplied delivery period.", sources: SUPPORT, limitations: [], resolves: listedFindingIDs(prompt) }] })
      },
    })
    const result = await runLean(directory, { sessions: model.port })
    expect(result.stop_reason).toBeNull()
    expect(result.status).toBe("completed")
    expect(readParagraphs(directory, "d1", "paragraphs-applied.json").map(row => row.text)).toContain(APPLIED)
    expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [{ mode: "code" }] })
    expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft" })
    expect(readOpenItems(directory).some(row => row.detail === findings[0].detail)).toBe(false)
    // The delivery group adjudicates the review's finding, so no finding is left open.
    expect(readChangePlan(directory, "d1")?.groups[0]?.resolves).toHaveLength(1)
    expect(deferredFindingItems(directory)).toHaveLength(2)
    const resumed = scriptedModel({})
    expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
    expect(resumed.prompts).toEqual([])
    findings.forEach(row => expect(stagePrompt(directory, "d2.draft")).toContain(row.detail))
    expect(stagePrompt(directory, "finalize")).toContain(deferredFindingItems(directory)[0].id)
  }, 120_000)
})
