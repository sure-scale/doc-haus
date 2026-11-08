import { afterEach, expect, test } from "bun:test"
import { readFileSync, readdirSync } from "node:fs"
import path from "node:path"
import { readDocxCommentParagraphs } from "../../lib/docx-comments"
import type { ChecklistCheckerInput } from "./checklist-contract"
import type { ChecklistDocumentRecord } from "./checklist-document-context"
import { runChecklistCheckFlow, type ChecklistCheckResult } from "./checklist-flow"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { readLeanDocuments } from "./deliverables"
import { currentFactDisposition, withFactOptions } from "./fact-answers"
import { enqueueLeanRequest } from "./requests"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, startRepairRound } from "./repairs"
import { runLean } from "./runner"
import { LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanState, readLeanText, readQuestions, updateLeanStage, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, drafter, happyPath, matter, NO_FINDINGS, noChanges, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const priorContract = "question-specific-provisional-markers-2-product-verification"
const questionText = "What pressure limit applies, and has the workshop checklist been signed off?"
const permission = "Prepare the review packet with both the pressure limit and checklist sign-off left unknown. Disclose the outstanding checks; issue and signature are not authorised."
const conditional = "Workshop release remains conditional on [open item: obtain the pressure limit] and [open item: verify checklist sign-off]."
const inventory = "The packet includes a conditional inspection sheet and a checklist for completion before workshop release."
const summaries = {
  inventory,
  "single component": "The inspection pressure setting awaits [open item: obtain the pressure limit before setting the gauge].",
  assertion: "The workshop checklist has been signed off.",
  "unconditional effect": "Workshop release is effective now.",
  "wrong marker": "The workshop checklist has been signed off [open item: confirm the equipment colour].",
  "compound marker": "The pressure limit and checklist sign-off now authorise workshop release [open item: obtain the pressure limit].",
  uninspected: inventory,
  "consumed adverse": inventory,
  wait: inventory,
}

for (const variant of Object.keys(summaries) as (keyof typeof summaries)[]) {
  test(`marker dependency contract refreshes only unconsumed late checks (${variant})`, async () => {
    const directory = matter("The packet contains an inspection sheet and a workshop checklist. Release requires the pressure limit and confirmed checklist sign-off; neither fact is supplied.",
      "Prepare a workshop review packet containing the conditional inspection record and an inventory of its included documents.")
    const offered = withFactOptions({ id: "q-workshop-release", stage: "intake", status: "open", kind: "fact", question: questionText, context: "Neither component is supplied." })
    const question: Question = { ...offered, status: "answered", answer: permission, answer_class: "answered", fact_status: "provisional" }
    writeQuestions(directory, [question])
    for (let i = 0; i < LEAN_MAX_SUBSTANTIVE_REPAIRS; i++) startRepairRound(directory, [], true)
    let recovery = false
    const first = new Map<string, ChecklistCheckerInput>()
    const packets: ChecklistCheckerInput[] = []
    const registration = registering()
    const model = scriptedModel({ ...happyPath,
      [AGENT.draft]: drafter(`# Workshop review packet\n\n## Conditional inspection record\n\n${conditional}\n\n## Included documents\n\n${summaries[variant]}\n`),
      [AGENT.review]: () => NO_FINDINGS,
      [AGENT.changes]: noChanges,
      [AGENT.finalize]: () => "The pressure limit and checklist sign-off remain unknown. Obtain both before relying on workshop release. Permission covers review preparation only; issue and signature remain unapproved.",
      [CHECKLIST_JOB_AGENTS.check]: (text, _index, sessionID) => {
        const json = text.split("CHECK_INPUT:\n\n")[1]
        if (json) first.set(sessionID!, JSON.parse(json))
        const input = first.get(sessionID!)!
        packets.push(input)
        const record = readLeanJson(directory, `checklist-context/${input.item.target.document}.json`) as ChecklistDocumentRecord
        const current = record.paragraphs.map(row => row.paragraph_text)
        expect(current).toContain(conditional)
        expect(current).toContain(summaries[variant])
        // These are explicit semantic result fixtures, not an LLM or a claim
        // that matching prompt wording demonstrates semantic understanding.
        // The test exercises saved-output refresh, boundaries and delivery.
        const positive = recovery && ["inventory", "single component"].includes(variant)
        const gap = recovery && variant === "uninspected"
        return JSON.stringify({ item_id: input.item_id, status: positive ? "satisfied" : gap ? "indeterminate" : "unsatisfied",
          action_status: positive ? "fulfilled" : gap ? "indeterminate" : "not_fulfilled", open_decisions: [],
          ...(gap ? { indeterminate_cause: "missing_inspection" } : {}),
          reason: !recovery ? "The earlier topic-based assessment treated the inventory reference as an unmarked dependency."
            : positive ? variant === "inventory" ? "The inventory accurately identifies included documents without asserting either missing fact or unconditional release. The actual release condition visibly preserves both unknown components."
              : "The pressure-setting statement relies only on the pressure-limit component and marks it explicitly; it neither asserts checklist sign-off nor grants release. The separate release condition preserves both unknown components."
            : gap ? "The inventory target has not been sufficiently inspected to establish its assertions or dependency scope."
            : variant === "assertion" || variant === "wrong marker" ? "The summary asserts completed checklist sign-off, an unknown fact; the colour marker cannot qualify that assertion."
            : variant === "unconditional effect" ? "The summary states release is effective now even though the pressure limit and sign-off conditions remain unknown."
            : "The summary relies on both release components but marks only the pressure limit; checklist sign-off remains unmarked.",
          // Positive body evidence supports the actual marked dependencies.
          // The inspected non-dependent inventory is explained in the reason
          // and scope, rather than mislabelled as a marked dependent passage.
          evidence: gap ? [] : record.paragraphs.filter(row => (positive && variant === "inventory" ? [conditional] : [conditional, summaries[variant]]).includes(row.paragraph_text)).map(row => ({
            document: input.item.target.document, locator: `native PIDs ${row.pids.join(", ")}; accepted paragraph_index ${row.paragraph_index}`, text: row.paragraph_text, kind: "body",
          })),
          source_refs: input.item.source_refs, inspection_scope: gap ? "The target dependency is not fully inspected." : "The conditional record and complete included-documents paragraph in the current accepted body.",
          limitation: "Neither unknown fact is answered or approved. This checks its handling only; review preparation permits no issue or signature.",
        })
      },
    })

    expect((await runLean(directory, { sessions: model.port, ...registration })).status).toBe("blocked")
    // Settle first-gate factual-item bookkeeping without Continue, so the
    // accepted notes/prefix are genuinely current before the upgrade snapshot.
    let blocked = await runLean(directory, { sessions: model.port, ...registration })
    if (variant === "consumed adverse") {
      const inspected = await runChecklistCheckFlow(directory, { sessions: model.port, model: null }, { reinspectAdverse: true })
      expect(inspected.checks[0]!.reinspection?.status).toBe("succeeded")
      blocked = await runLean(directory, { sessions: model.port, ...registration })
    }
    expect(blocked.status).toBe("blocked")
    expect(blocked.stages.gate?.status).toBe("blocked")
    expect(blocked.stages["checklist-check"]?.status).toBe("done")
    expect(blocked.stages.finalize?.input_sha).toBe(hashStageInputs("finalize", LEAN_PROMPT_VERSION, stageInputs(directory, "finalize")))
    const checks = readLeanJson(directory, "checklist/checks.json") as { checks: ChecklistCheckResult[] }
    expect(checks.checks).toHaveLength(1)
    const consumed = structuredClone(checks.checks[0]!.reinspection)
    expect(!!consumed).toBe(variant === "consumed adverse")
    // This is the persisted pre-upgrade job/key, with unchanged Word/source
    // evidence and an existing adverse finding. No factual result is invented.
    checks.checks[0]!.input_sha = "prior-topic-dependency-check"
    writeLeanText(directory, "checklist/checks.json", JSON.stringify(checks))
    for (const step of ["checklist-check", "gate"] as const) {
      const inputs = stageInputs(directory, step).map(part => part.name === "provisional-fact-marker-contract" ? { ...part, bytes: priorContract } : part)
      updateLeanStage(directory, readLeanState(directory)!, step, { input_sha: hashStageInputs(step, LEAN_PROMPT_VERSION, inputs) })
    }
    const saved = readLeanState(directory)!
    const prefix = Object.fromEntries(Object.entries(saved.stages).filter(([step, stage]) => stage?.status === "done" && !["checklist-check", "package-review"].includes(step)).map(([step, stage]) => [step, structuredClone(stage)]))
    const repairs = readLeanText(directory, "repairs.json")
    const answers = readLeanText(directory, "questions.json")
    const body = readFileSync(leanFile(directory, "d1/body-applied.docx"))
    const paragraphs = readFileSync(leanFile(directory, "d1/paragraphs-applied.json"))
    const notes = readFileSync(leanFile(directory, "drafting-notes.docx"))
    const sections = new Map(readdirSync(leanFile(directory, "notes-sections")).map(file => [file, readFileSync(leanFile(directory, `notes-sections/${file}`))]))
    recovery = true
    if (variant === "wait") enqueueLeanRequest(directory, { kind: "answer", question_id: question.id, option_id: "o2" })
    const before = model.prompts.length
    enqueueLeanRequest(directory, { kind: "continue" })
    const resumed = await runLean(directory, { sessions: model.port, ...registration })
    const later = model.prompts.slice(before)
    expect(readLeanText(directory, "repairs.json")).toBe(repairs)
    expect(readRepairRound(directory)).toMatchObject({ round: 5, substantive_attempts: 5 })
    expect(later.some(row => row.agent === AGENT["change-plan"])).toBe(false)
    expect(readFileSync(leanFile(directory, "d1/body-applied.docx"))).toEqual(body)
    expect(readFileSync(leanFile(directory, "d1/paragraphs-applied.json"))).toEqual(paragraphs)
    if (variant === "wait") {
      expect(resumed.status).toBe("blocked")
      expect(currentFactDisposition(directory, readQuestions(directory)[0]!)).toBe("wait")
      expect(registration.names).toEqual([])
      return
    }
    expect(readLeanText(directory, "questions.json")).toBe(answers)
    expect(currentFactDisposition(directory, readQuestions(directory)[0]!)).toBe("provisional")
    for (const [step, value] of Object.entries(prefix)) expect(resumed.stages[step as keyof typeof resumed.stages]).toEqual(value)
    expect(readFileSync(leanFile(directory, "drafting-notes.docx"))).toEqual(notes)
    for (const [file, bytes] of sections) expect(readFileSync(leanFile(directory, `notes-sections/${file}`))).toEqual(bytes)
    expect(later.every(row => new Set<string>([CHECKLIST_JOB_AGENTS.check, AGENT["package-review"]]).has(row.agent))).toBe(true)
    const finalChecks = readLeanJson(directory, "checklist/checks.json") as { checks: ChecklistCheckResult[] }
    if (variant === "consumed adverse") {
      expect(later.filter(row => row.agent === CHECKLIST_JOB_AGENTS.check)).toHaveLength(0)
      expect(finalChecks.checks[0]!.reinspection).toEqual(consumed)
      expect(finalChecks.checks[0]!.finding).toEqual(checks.checks[0]!.finding)
    } else expect(later.some(row => row.agent === CHECKLIST_JOB_AGENTS.check)).toBe(true)
    if (!["inventory", "single component"].includes(variant)) {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      return
    }
    expect(resumed.status).toBe("completed")
    expect(finalChecks.checks[0]!.finding).toMatchObject({ status: "satisfied", action_status: "fulfilled", open_decisions: [] })
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    const documents = readLeanDocuments(directory)!
    expect(documents.status).toBe("completed")
    const deliveredBody = documents.documents.find(row => row.kind === "body")!.deliverable!
    expect(registration.names).toContain(deliveredBody)
    expect(readDocxCommentParagraphs(readFileSync(path.join(directory, deliveredBody))).map(row => row.paragraph_text)).toEqual(readDocxCommentParagraphs(body).map(row => row.paragraph_text))
    const deliveredNotes = documents.documents.find(row => row.kind === "drafting_notes")!.deliverable!
    const text = readDocxCommentParagraphs(readFileSync(path.join(directory, deliveredNotes))).map(row => row.paragraph_text).join("\n")
    expect(text).toContain("pressure limit and checklist sign-off remain unknown")
    expect(text).toContain("issue and signature remain unapproved")
    expect(packets.every(input => input.verification === "provisional_fact_marker" && input.item.source_refs.some(ref => ref.question === questionText))).toBe(true)
  }, 120_000)
}
