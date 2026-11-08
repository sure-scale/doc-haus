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
import { finalGate } from "./gate"
import { enqueueLeanRequest } from "./requests"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, readRepairRound, startRepairRound } from "./repairs"
import { runLean } from "./runner"
import { LEAN_MAX_RETRIES, LEAN_PROMPT_VERSION, stageInputs } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readLeanState, readLeanText, readQuestions, updateLeanStage, writeLeanText, writeQuestions, type Question } from "./state"
import { AGENT, cleanupMatters, drafter, happyPath, matter, NO_FINDINGS, noChanges, registering, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const questionText = "What inspection code applies to the training rig?"
const heading = "Inspection code"
const marked = "The training-rig inspection code remains unknown [open item: obtain the inspection code before relying on this record]."
const assertion = "The training rig has passed the inspection required by its code."
const inventory = "This packet includes a conditional inspection record for the training rig."
type Variant = "corrected" | "partial paragraph" | "merged paragraphs" | "unmarked dependent" | "stubborn invalid" | "consumed review" | "not applicable peer"

for (const variant of ["corrected", "partial paragraph", "merged paragraphs", "unmarked dependent", "stubborn invalid", "consumed review", "not applicable peer"] satisfies Variant[]) {
  test(`saved gate requires bounded same-session marker evidence correction (${variant})`, async () => {
    const directory = matter("The training-rig inspection code is not supplied. The packet includes a conditional inspection record.",
      "Prepare a training-rig review record and describe its included documents; do not invent the missing inspection code.")
    const offered = withFactOptions({ id: "q-rig-code", stage: "intake", status: "open", kind: "fact", question: questionText, context: "The code has not been supplied." })
    const question: Question = { ...offered, status: "answered", answer: "Prepare a review draft while leaving the inspection code unknown and disclosing its outstanding check. Issue and signature are not authorised.", answer_class: "answered", fact_status: "provisional" }
    writeQuestions(directory, [question])
    for (let i = 0; i < LEAN_MAX_SUBSTANTIVE_REPAIRS; i++) startRepairRound(directory, [], true)
    const draft = `# Training-rig review record\n\n## ${heading}\n\n${marked}\n\n## Included documents\n\n${variant === "unmarked dependent" ? assertion : inventory}\n`
    const unrelated = "The storage register lists a green rack and a blue rack."
    const documents = variant === "not applicable peer" ? [draft, `# Storage register\n\n## Racks\n\n${unrelated}\n`] : [draft]
    let recovery = false
    const packets = new Map<string, ChecklistCheckerInput>()
    const recoveryCalls: { sessionID: string; input: ChecklistCheckerInput; text: string }[] = []
    const registration = registering()
    const model = scriptedModel({ ...happyPath,
      [AGENT.draft]: drafter(...documents),
      [AGENT.review]: () => NO_FINDINGS,
      [AGENT.changes]: noChanges,
      [AGENT.finalize]: () => "The inspection code remains unknown. Obtain it before relying on the training-rig inspection record. Permission covers review preparation only; issue and signature remain unapproved.",
      [CHECKLIST_JOB_AGENTS.check]: (text, _index, sessionID) => {
        const json = text.split("CHECK_INPUT:\n\n")[1]
        if (json) packets.set(sessionID!, JSON.parse(json))
        const input = packets.get(sessionID!)!
        const record = readLeanJson(directory, `checklist-context/${input.item.target.document}.json`) as ChecklistDocumentRecord
        const evidence = (texts: string[]) => record.paragraphs.filter(row => texts.includes(row.paragraph_text)).map(row => ({
          document: input.item.target.document, locator: `native PIDs ${row.pids.join(", ")}; accepted paragraph_index ${row.paragraph_index}`, text: row.paragraph_text, kind: "body",
        }))
        const invalidEvidence = variant === "partial paragraph" ? [{ ...evidence([marked])[0]!, text: marked.slice(marked.indexOf("unknown")) }]
          : variant === "merged paragraphs" ? [{ ...evidence([heading])[0]!, text: `${heading}\n${marked}` }] : evidence([heading, marked])
        const base = { item_id: input.item_id, open_decisions: [], source_refs: input.item.source_refs,
          inspection_scope: "The complete current target, including the heading, inspection paragraph and inventory.",
          limitation: "The inspection code remains unknown; this assessment supplies neither its answer nor issue/signature approval.",
        }
        if (recovery) recoveryCalls.push({ sessionID: sessionID!, input, text })
        if (input.item.target.document === "d2") return JSON.stringify({ ...base, status: "not_applicable", action_status: "not_applicable",
          reason: "The inspected rack register makes no assertion or operative effect that relies on the training-rig inspection code.", evidence: evidence([unrelated]) })
        expect(record.paragraphs.some(row => row.paragraph_text === heading)).toBe(true)
        expect(record.paragraphs.some(row => row.paragraph_text === marked)).toBe(true)
        if (!recovery) return JSON.stringify({ ...base, status: "unsatisfied", action_status: "not_fulfilled",
          reason: "The synthetic prior assessment did not establish satisfactory marker coverage.", evidence: evidence([heading, marked]) })
        const attempt = recoveryCalls.filter(row => row.sessionID === sessionID).length
        // Scripted semantic outcomes test orchestration, never live model
        // understanding. Code must reject the first output unchanged, then ask
        // this SAME checker to reassess its actual current evidence selection.
        if (attempt === 1 || variant === "stubborn invalid") return JSON.stringify({ ...base, status: "satisfied", action_status: "fulfilled",
          reason: "The dependent inspection paragraph is marked; the heading is nondependent context.", evidence: invalidEvidence })
        expect(text).toContain("Your reply did not pass the checklist contract")
        if (variant === "unmarked dependent") return JSON.stringify({ ...base, status: "unsatisfied", action_status: "not_fulfilled",
          reason: "The separate summary asserts completed inspection under an unknown code without a marker; omitting that dependent assertion from evidence cannot establish success.", evidence: evidence([assertion]) })
        return JSON.stringify({ ...base, status: "satisfied", action_status: "fulfilled",
          reason: "The inspected heading and inventory make no dependent assertion. The following inspection paragraph is the actual dependent passage and explicitly preserves the unknown code.", evidence: evidence([marked]) })
      },
    })
    expect((await runLean(directory, { sessions: model.port, ...registration })).status).toBe("blocked")
    let blocked = await runLean(directory, { sessions: model.port, ...registration })
    if (variant === "consumed review") {
      const reviewed = await runChecklistCheckFlow(directory, { sessions: model.port, model: null }, { reinspectAdverse: true })
      expect(reviewed.checks[0]!.reinspection?.status).toBe("succeeded")
      blocked = await runLean(directory, { sessions: model.port, ...registration })
    }
    expect(blocked.stages.gate?.status).toBe("blocked")
    expect(blocked.stages["checklist-check"]?.status).toBe("done")
    expect(blocked.stages.finalize?.input_sha).toBe(hashStageInputs("finalize", LEAN_PROMPT_VERSION, stageInputs(directory, "finalize")))
    const checks = readLeanJson(directory, "checklist/checks.json") as { checks: ChecklistCheckResult[] }
    const target = checks.checks.find(row => row.item_id.endsWith(":d1"))!
    const consumed = structuredClone(target.reinspection)
    expect(!!consumed).toBe(variant === "consumed review")
    target.finding = { ...target.finding!, status: "satisfied", action_status: "fulfilled", open_decisions: [],
      reason: "The prior output included a nondependent heading beside the marked dependent paragraph." }
    if (variant === "partial paragraph") target.finding.evidence = [{ ...target.finding.evidence.find(row => row.text === marked)!, text: marked.slice(marked.indexOf("unknown")) }]
    if (variant === "merged paragraphs") target.finding.evidence = [{ ...target.finding.evidence.find(row => row.text === heading)!, text: `${heading}\n${marked}` }]
    for (const row of checks.checks) row.input_sha = `prior-per-row-marker-evidence:${row.item_id}`
    writeLeanText(directory, "checklist/checks.json", JSON.stringify(checks))
    // Obtain the genuine saved gate result for this exact invalid positive:
    // the unchanged gate rejects the unmarked heading, despite satisfied.
    const savedGate = finalGate(directory)
    expect(savedGate.status).toBe("blocked")
    writeLeanText(directory, "gate.json", JSON.stringify(savedGate))
    // The old checker and final-gate contracts are obsolete. Their existing
    // content-keyed resume must refresh checks before admitting any repair.
    for (const step of ["checklist-check", "gate"] as const) updateLeanStage(directory, readLeanState(directory)!, step, { input_sha: `prior-per-row-marker-evidence:${step}` })
    const saved = readLeanState(directory)!
    const prefix = Object.fromEntries(Object.entries(saved.stages).filter(([step, stage]) => stage?.status === "done" && !["checklist-check", "package-review"].includes(step)).map(([step, stage]) => [step, structuredClone(stage)]))
    const repairs = readLeanText(directory, "repairs.json")
    const answers = readLeanText(directory, "questions.json")
    const bodyFiles = documents.flatMap((_, index) => [`d${index + 1}/body-applied.docx`, `d${index + 1}/paragraphs-applied.json`])
    const bytes = new Map([...bodyFiles, "drafting-notes.docx", ...readdirSync(leanFile(directory, "notes-sections")).map(file => `notes-sections/${file}`)]
      .map(file => [file, readFileSync(leanFile(directory, file))]))
    const before = model.prompts.length
    recovery = true
    enqueueLeanRequest(directory, { kind: "continue" })
    const resumed = await runLean(directory, { sessions: model.port, ...registration })
    expect(readLeanText(directory, "repairs.json")).toBe(repairs)
    expect(readRepairRound(directory)).toMatchObject({ round: 5, substantive_attempts: 5 })
    expect(readLeanText(directory, "questions.json")).toBe(answers)
    expect(currentFactDisposition(directory, readQuestions(directory)[0]!)).toBe("provisional")
    for (const [step, value] of Object.entries(prefix)) expect(resumed.stages[step as keyof typeof resumed.stages]).toEqual(value)
    for (const [file, value] of bytes) expect(readFileSync(leanFile(directory, file))).toEqual(value)
    const later = model.prompts.slice(before)
    expect(later.every(row => new Set<string>([CHECKLIST_JOB_AGENTS.check, AGENT["package-review"]]).has(row.agent))).toBe(true)
    const targetCalls = recoveryCalls.filter(row => row.input.item.target.document === "d1")
    expect(new Set(targetCalls.map(row => row.sessionID)).size).toBe(1)
    expect(targetCalls).toHaveLength(variant === "stubborn invalid" ? LEAN_MAX_RETRIES + 1 : 2)
    expect(targetCalls[0]!.text).toContain("CHECK_INPUT:")
    expect(targetCalls[1]!.text).toContain("Your reply did not pass the checklist contract")
    const finalChecks = readLeanJson(directory, "checklist/checks.json") as { checks: ChecklistCheckResult[] }
    const finalTarget = finalChecks.checks.find(row => row.item_id.endsWith(":d1"))!
    expect(finalTarget.reinspection).toEqual(consumed)
    if (variant === "stubborn invalid" || variant === "unmarked dependent") {
      expect(resumed.status).toBe("blocked")
      expect(registration.names).toEqual([])
      if (variant === "stubborn invalid") expect(finalTarget).toMatchObject({ status: "failed", attempts: LEAN_MAX_RETRIES + 1 })
      else expect(finalTarget.finding).toMatchObject({ status: "unsatisfied", action_status: "not_fulfilled" })
      return
    }
    expect(finalTarget.finding).toMatchObject({ status: "satisfied", action_status: "fulfilled", evidence: [{ text: marked }] })
    if (variant === "not applicable peer") {
      const peer = finalChecks.checks.find(row => row.item_id.endsWith(":d2"))!
      expect(peer.finding).toMatchObject({ status: "not_applicable", action_status: "not_applicable", evidence: [{ text: unrelated }] })
      expect(recoveryCalls.filter(row => row.input.item.target.document === "d2")).toHaveLength(1)
    }
    expect(resumed.status).toBe("completed")
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    const delivered = readLeanDocuments(directory)!
    expect(delivered.status).toBe("completed")
    for (const row of delivered.documents.filter(row => row.kind === "body")) {
      expect(registration.names).toContain(row.deliverable!)
      expect(readDocxCommentParagraphs(readFileSync(path.join(directory, row.deliverable!))).map(paragraph => paragraph.paragraph_text))
        .toEqual(readDocxCommentParagraphs(readFileSync(leanFile(directory, `${row.document}/body-applied.docx`))).map(paragraph => paragraph.paragraph_text))
    }
    const notes = delivered.documents.find(row => row.kind === "drafting_notes")!.deliverable!
    const text = readDocxCommentParagraphs(readFileSync(path.join(directory, notes))).map(row => row.paragraph_text).join("\n")
    expect(text).toContain("inspection code remains unknown")
    expect(text).toContain("issue and signature remain unapproved")
  }, 120_000)
}
