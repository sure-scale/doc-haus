import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { afterEach, expect, test } from "bun:test"
import { renderMarkdownToDocx } from "../../lib/render"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import { readFileSync } from "node:fs"
import { buildRepairAdjudicationInput, mergeRepairAdjudications, repairAdjudicationIssues, runFocusedRepairAdjudication, type RepairAdjudication } from "./checklist-repair-adjudication"
import { acceptChangePlan, changePreparationPrompt, type ChangePlan } from "./change-pipeline"
import { writeLeanPlan, type StepID } from "./plan"
import { packageCompositionBlock } from "./prompts"
import { handedBackBlock, hashStageInputs, updateLeanStage, writeQuestions, leanFile, readLeanState, readLeanText, writeLeanBytes, writeLeanText } from "./state"
import { startRepairRound } from "./repairs"
import { openItem } from "./quotes"
import { cleanupMatters, matter, requestMemo, scriptedModel, SOURCE, SUPPORT } from "./test-support"
import type { ChecklistItem, ChecklistFinding } from "./checklist-contract"

afterEach(cleanupMatters)
const pid = "p111111111111", other = "p222222222222"
const item: ChecklistItem = { check: "Identify timing uncertainty as advice, without adopting a new term.", basis: "directive", qualifications: ["Approval remains unknown."], required_action: "flag", source_refs: [{ document: SOURCE, locator: "source", quote: SUPPORT[0]!.quote }], target: { document: "d1", role: "Advice", treatment: "advice" } }
const plan: ChangePlan = { shared_decisions: ["Generated scope: omit timing advice.", "Preserve unrelated source conditions."], groups: [{ id: "old", pids: [pid], goal: "Preserve supplied timing.", sources: SUPPORT, limitations: [] }, { id: "other", pids: [other], goal: "Preserve payment.", sources: SUPPORT, limitations: [] }], findings: [{ disposition: "rejected", kind: "suggestion", detail: "Prior timing assessment.", reason: "Prior generated decision.", documents: ["d1"], sources: SUPPORT, limitations: [] }, { disposition: "recommendation", kind: "suggestion", detail: "Unrelated advice.", reason: "Keep visible.", documents: ["d1"], sources: [], limitations: ["Unrelated uncertainty."] }], dismissed: [], questions: [] }
const result: RepairAdjudication = { shared_decisions: [], groups: [{ id: "timing", pids: [pid], goal: "Identify uncertainty without adopting a new deadline.", sources: SUPPORT, limitations: ["No approval inferred."] }], dismissed: [], findings: [], questions: [], superseded_decisions: [{ index: 0, reason: "Generated scope cannot waive original instructions to identify uncertainty.", sources: SUPPORT }], superseded_findings: [{ index: 0, reason: "The original source calls for advice, not adoption.", sources: SUPPORT }] }
async function setup() {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Advice", purpose: "Explain timing", file: "advice.docx", parts: [{ heading: "Timing", scope: "Timing" }] }])
  await writeLeanBytes(directory, "d1/body-applied.docx", (await renderMarkdownToDocx("Timing is missing.\n\nPayment remains unchanged.")).bytes)
  const current = readDocxCommentParagraphs(readFileSync(leanFile(directory, "d1/body-applied.docx")))
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid, sha256: "initial", text: "Timing is missing." }, { pid: other, sha256: "other", text: "Payment remains unchanged." }]))
  writeLeanText(directory, "d1/apply.json", JSON.stringify({ applied: [{ id: "prior", pid, result_pids: current[0]!.pids }], unapplied: [] }))
  expect(acceptChangePlan(directory, "d1", JSON.stringify(plan)).issues).toEqual([])
  const finding: ChecklistFinding = { item_id: "timing", status: "unsatisfied", reason: "Timing uncertainty missing.", evidence: [{ document: "d1", locator: "paragraph0", kind: "body", text: "Timing is missing." }], source_refs: item.source_refs, inspection_scope: "Current target.", limitation: "No approval assumed.", action_status: "not_fulfilled", open_decisions: ["Approval"] }
  const assessment = { item_id: "timing", item, finding }
  startRepairRound(directory, [openItem("checklist-check", "omission", `Reassess source.\n${JSON.stringify(assessment)}`, { document: "d1", repair_target: "body", pid })], true)
  writeLeanText(directory, "d1/review.md", "UNRELATED_REVIEW_SENTINEL ".repeat(1000))
  writeLeanText(directory, "research.json", JSON.stringify({ limitations: ["GENERATED_SCOPE_IS_NOT_WAIVER"] }))
  updateLeanStage(directory, readLeanState(directory)!, "d1.apply", { status: "done", input_sha: hashStageInputs("d1.apply", LEAN_PROMPT_VERSION, stageInputs(directory, "d1.apply")) })
  return { directory, assessment }
}
test("one assessment gets original instructions and exact current/initial passages without review/research dump", async () => {
  const { directory, assessment } = await setup()
  const input = buildRepairAdjudicationInput(directory, "d1", assessment)
  expect(input.prompt).toContain(JSON.stringify(assessment))
  expect(input.prompt).toContain("Draft a short supply note")
  expect(input.prompt).toContain("Generated research limits")
  expect(input.prompt).toContain("flagging or proposing")
  expect(input.prompt).not.toContain("UNRELATED_REVIEW_SENTINEL")
  expect(input.prompt).not.toContain("GENERATED_SCOPE_IS_NOT_WAIVER")
  expect(input.prompt).toContain(`[${pid}] Timing is missing.`)
  expect(input.finding_indices).toEqual([0])
  expect(input.prompt).toContain("Generated scope: omit timing advice.")
  expect(input.decision_indices).toEqual([0, 1])
  expect(JSON.stringify(input.artifacts)).toContain("Timing is missing.")
  expect(readLeanText(directory, "checklist-context/d1.json")).toBeUndefined()
})
test("supersession only names supplied history; completed groups cannot be superseded wholesale", async () => {
  const { directory, assessment } = await setup()
  const input = buildRepairAdjudicationInput(directory, "d1", assessment)
  expect(repairAdjudicationIssues(directory, "d1", input, result)).toEqual([])
  expect(repairAdjudicationIssues(directory, "d1", input, { ...result, superseded_findings: [{ index: 1, reason: "Drop" }] }).join(" ")).toContain("only supplied indices")
  expect(repairAdjudicationIssues(directory, "d1", input, { ...result, groups: [], superseded_groups: [{ index: 0, reason: "Remove" }] }).join(" ")).toContain("not superseded or replayed")
  expect(repairAdjudicationIssues(directory, "d1", input, { ...result, groups: [{ ...result.groups[0], pids: ["p999999999999"] }] }).join(" ")).toContain("not an initial draft paragraph")
})
test("shared PID goals merge transitively and unrelated amendments/dispositions remain", () => {
  const merged = mergeRepairAdjudications(plan, [result, { ...result, groups: [{ ...result.groups[0]!, id: "dependent", pids: [pid, other], goal: "Preserve dependent effects." }], superseded_findings: [] }])
  expect(merged.groups).toHaveLength(1)
  expect(merged.groups[0]!.pids.sort()).toEqual([pid, other])
  expect(merged.groups[0]!.goal).toContain("Preserve dependent effects.")
  expect(merged.groups[0]!.goal).not.toContain("Preserve payment.")
  expect(merged.shared_decisions).toEqual(["Preserve unrelated source conditions."])
  expect(merged.findings).toEqual([plan.findings![1]!])
})
test("native product dispatch accepts cumulative plan with one focused job and no generic read", async () => {
  const { directory } = await setup()
  const model = scriptedModel({ "lean-change-plan": () => JSON.stringify(result) })
  const run = await runFocusedRepairAdjudication(directory, readLeanState(directory)!, { sessions: model.port }, "d1", () => {})
  expect(run.status).toBe("done")
  expect(model.prompts).toHaveLength(1)
  expect(model.sessions[0]!.permissions).not.toContainEqual({ action: "read", resource: "*", effect: "allow" })
  const accepted = JSON.parse(readLeanText(directory, "d1/change-plan.json")!) as ChangePlan
  expect(accepted.groups).toHaveLength(1)
  expect(accepted.groups[0]!.goal).toBe(result.groups[0]!.goal)
  expect(accepted.groups.some(row => row.id === "other")).toBe(false)
  expect(accepted.findings).toEqual([plan.findings![1]!])
})

test("six independent repairs are bounded to four native sessions and transport errors are not retried", async () => {
  const { directory, assessment } = await setup()
  const assessments = Array.from({ length: 6 }, (_, index) => ({ ...assessment, item_id: `finding-${index}`, finding: { ...assessment.finding, item_id: `finding-${index}` } }))
  startRepairRound(directory, assessments.map(value => openItem("checklist-check", "omission", `Reassess.\n${JSON.stringify(value)}`, { document: "d1", repair_target: "body", pid })), true)
  let active = 0, maximum = 0, created = 0, prompts = 0
  const sessions = {
    async createSession() { return `ses_${++created}` },
    async prompt(input: { text: string }) {
      prompts++; active++; maximum = Math.max(maximum, active)
      const packet = JSON.parse(input.text.split("<assessment>\n")[1]!.split("\n</assessment>")[0]!)
      expect(packet.finding.item_id).toBe(packet.item_id)
      expect(input.text.match(/<assessment>/g)).toHaveLength(1)
      await new Promise(resolve => setTimeout(resolve, 5))
      active--
      if (packet.item_id === "finding-5") throw new Error("transport unavailable")
      return JSON.stringify(result)
    },
  }
  const before = readLeanText(directory, "d1/change-plan.json")
  const run = await runFocusedRepairAdjudication(directory, readLeanState(directory)!, { sessions }, "d1", () => {})
  expect(maximum).toBe(4)
  expect(created).toBe(6)
  expect(prompts).toBe(6)
  expect(run.status).toBe("failed")
  expect(readLeanText(directory, "d1/change-plan.json")).toBe(before)
})


test("a local goal never expands to an old omnibus paragraph group", () => {
  const omnibus = { ...plan, groups: [{ ...plan.groups[0]!, pids: Array.from({ length: 70 }, (_, index) => `old-${index}`), goal: "Historical whole-document work already applied." }] }
  const merged = mergeRepairAdjudications(omnibus, [result])
  expect(merged.groups).toHaveLength(1)
  expect(merged.groups[0]!.pids).toEqual([pid])
  expect(merged.groups[0]!.goal).not.toContain("whole-document")
})


test("rejection needs source support while unsupported uncertainty remains deferred", async () => {
  const { directory, assessment } = await setup()
  const input = buildRepairAdjudicationInput(directory, "d1", assessment)
  const noChange = { ...result, groups: [], superseded_findings: [], superseded_decisions: [], findings: [{
    ...plan.findings![0]!, detail: "Do not adopt a different delivery deadline.",
    reason: "The source supplies thirty days; no different deadline is instructed.", sources: SUPPORT,
  }] }
  expect(repairAdjudicationIssues(directory, "d1", input, noChange)).toEqual([])
  const unsupported = { ...noChange, findings: [{ ...noChange.findings[0]!, sources: [], limitations: ["No authority established."] }] }
  expect(repairAdjudicationIssues(directory, "d1", input, unsupported).join(" ")).toContain("rejection requires supporting source")
  expect(repairAdjudicationIssues(directory, "d1", input, { ...unsupported, findings: [{ ...unsupported.findings[0]!, disposition: "deferred" }] })).toEqual([])
  expect(repairAdjudicationIssues(directory, "d1", input, { ...noChange, findings: [{ ...noChange.findings[0]!, sources: [{ source_id: SOURCE, quote: "Invented permission to omit advice." }] }] }).join(" ")).toContain("but those words")
})

test("each supersession needs its own canonical exact support, not a sourced unrelated group", async () => {
  const { directory, assessment } = await setup()
  const input = buildRepairAdjudicationInput(directory, "d1", assessment)
  expect(repairAdjudicationIssues(directory, "d1", input, result)).toEqual([])
  for (const field of ["superseded_findings", "superseded_decisions"] as const) {
    const row = result[field][0]!
    for (const sources of [undefined, []])
      expect(repairAdjudicationIssues(directory, "d1", input, { ...result, [field]: [{ ...row, sources }] }).join(" ")).toContain("supersession requires nonempty")
    expect(repairAdjudicationIssues(directory, "d1", input, { ...result, [field]: [{ ...row, sources: [{ source_id: "generated-review", quote: SUPPORT[0]!.quote }] }] }).join(" ")).toContain("canonical matter sources")
    expect(repairAdjudicationIssues(directory, "d1", input, { ...result, [field]: [{ ...row, sources: [{ source_id: SOURCE, quote: "Permission to abandon prior duties." }] }] }).join(" ")).toContain("but those words")
  }
})

test("unsupported rejection and supersession cannot replace the accepted cumulative plan", async () => {
  const { directory } = await setup()
  const before = readLeanText(directory, "d1/change-plan.json")
  const unsupported = { ...result, groups: [], findings: [{ ...plan.findings![0]!, sources: [], limitations: ["No source inspected."] }],
    superseded_findings: [{ index: 0, reason: "Unnecessary." }], superseded_decisions: [{ index: 0, reason: "Ignore dependency." }] }
  const model = scriptedModel({ "lean-change-plan": () => JSON.stringify(unsupported) })
  const run = await runFocusedRepairAdjudication(directory, readLeanState(directory)!, { sessions: model.port }, "d1", () => {})
  expect(run.status).toBe("failed")
  expect(readLeanText(directory, "d1/change-plan.json")).toBe(before)
  expect(readLeanText(directory, "d1/repair-adjudications.json")).toBeUndefined()
})


test("actual request and human answers can support supersession; generated question text cannot", async () => {
  const { directory, assessment } = await setup()
  writeQuestions(directory, [{ id: "scope", stage: "change-plan", document: "d1", pid,
    question: "The generated question is not a client instruction.", context: "Confirm scope.",
    status: "answered", answer: "Retain the delivery timing as supplied." }])
  const input = buildRepairAdjudicationInput(directory, "d1", assessment)
  const withQuote = (quote: string) => ({ ...result, superseded_decisions: [{ ...result.superseded_decisions[0]!, sources: [{ source_id: "request", quote }] }] })
  expect(repairAdjudicationIssues(directory, "d1", input, withQuote("Draft a short supply note from the supply terms."))).toEqual([])
  expect(repairAdjudicationIssues(directory, "d1", input, withQuote("Retain the delivery timing as supplied."))).toEqual([])
  expect(repairAdjudicationIssues(directory, "d1", input, withQuote("The generated question is not a client instruction.")).join(" ")).toContain("but those words")
  expect(repairAdjudicationIssues(directory, "d1", input, withQuote("Retain the delivery timing as supplied. Approval granted.")).join(" ")).toContain("but those words")
})

test("a question left to the drafting reaches the focused adjudication beside the instructions, never as citable words; without one the prompt is unchanged", async () => {
  const { directory, assessment } = await setup()
  const without = buildRepairAdjudicationInput(directory, "d1", assessment).prompt
  writeQuestions(directory, [{ id: "q-left", stage: "intake", question: "Which timing should the supply note give?", context: "Two dates conflict.", status: "answered", answer_kind: "decide_for_me" }])
  // The hand-back re-keys the run as an answer does; the applied draft stays current here.
  updateLeanStage(directory, readLeanState(directory)!, "d1.apply", { status: "done", input_sha: hashStageInputs("d1.apply", LEAN_PROMPT_VERSION, stageInputs(directory, "d1.apply")) })
  const block = `<left_to_drafting>\n${handedBackBlock(directory)}\n</left_to_drafting>`
  const input = buildRepairAdjudicationInput(directory, "d1", assessment)
  expect(input.prompt.split(block).length - 1).toBe(1)
  expect(input.prompt.replace(`\n\n${block}`, "")).toBe(without)
  const withQuote = (quote: string) => ({ ...result, superseded_decisions: [{ ...result.superseded_decisions[0]!, sources: [{ source_id: "request", quote }] }] })
  expect(repairAdjudicationIssues(directory, "d1", input, withQuote("Which timing should the supply note give?")).join(" ")).toContain("but those words")
})

test("a requested memo reaches the focused adjudication and checklist preparation prompts; without one they are unchanged", async () => {
  const { directory, assessment } = await setup()
  const prompts = () => ({ adjudication: buildRepairAdjudicationInput(directory, "d1", assessment).prompt, preparation: changePreparationPrompt(directory, "d1", plan, plan.groups[0]!) })
  const without = prompts()
  expect(without.preparation).toContain("Prepare ONE accepted repair group")
  requestMemo(directory)
  const block = packageCompositionBlock(directory, "d1")
  expect(block).toStartWith("<package_composition>\n")
  const memo = prompts()
  for (const name of ["adjudication", "preparation"] as const) {
    expect([name, memo[name].split(block).length - 1]).toEqual([name, 1])
    expect([name, memo[name].replace(`\n\n${block}`, "")]).toEqual([name, without[name]])
    expect([name, without[name].includes("package_composition")]).toEqual([name, false])
  }
  // The checklist-only stages are keyed by the block, so resume sees it.
  for (const stage of ["change-plan", "change-prepare", "changes", "change-review"])
    expect([stage, stageInputs(directory, `d1.${stage}` as StepID).filter(row => row.name === "package-composition")]).toEqual([stage, [{ name: "package-composition", bytes: block }]])
})
