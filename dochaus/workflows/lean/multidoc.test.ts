import { afterEach, describe, expect, test } from "bun:test"
import { existsSync, readFileSync, writeFileSync } from "node:fs"
import { readParagraphs } from "./checks"
import { readLeanDocuments } from "./deliverables"
import type { GateResult } from "./gate"
import { readLeanPlan, type StepID } from "./plan"
import { enqueueLeanRequest } from "./requests"
import { restoreLeanStage } from "./restore-stage"
import { stagePrompt } from "./prompts"
import { runLean } from "./runner"
import { LEAN_PROMPT_VERSION, LEAN_STAGE_AGENTS, stageInputs } from "./stages"
import { hashStageInputs, leanFile, readLeanJson, readOpenItems, readQuestions, writeLeanText } from "./state"
import {
  AGENT, APPLIED, DRAFT, DRAFTED, LETTER, QUESTION, SUPPORT, changes, changesFirst, cleanupMatters, drafter, draftingNotes, happyPath, listed, matter, pidIn,
  registering, reviewFirst, sameSession, scriptedModel, type Reply,
} from "./test-support"

// Runs that plan more than one document. Real coordinator, checkers, renderer,
// applier and gate; only the model is scripted, as a function of each prompt.

const REVIEW_SCOPE = "Where a problem concerns another planned document rather than this one, report it with that document's id as its document; the run carries it to that document. Content the request assigns to another planned document is not missing from this one."
const CHANGES_SCOPE = "Amend only this document. Content the request assigns to another planned document belongs in that document, and findings about it reach it separately; do not draft it here."
const SCHEDULE = "# Delivery schedule\n\n## Timetable\n\nDelivery follows the supply note and the covering letter.\n"

afterEach(cleanupMatters)

function promptsOf(model: ReturnType<typeof scriptedModel>, agent: string, title: string) {
  const sessions = new Set(model.sessions.flatMap((session, index) => session.agent === agent && session.title.includes(`: ${title}: `) ? [`ses_child${index + 1}`] : []))
  return model.prompts.filter(prompt => sessions.has(prompt.sessionID)).map(prompt => prompt.text)
}

describe("lean drafting with several documents", () => {
  test("a one-document request plans one document and delivers one body, one redline and the notes", async () => {
    const directory = matter()
    const deliverables = registering()
    const state = await runLean(directory, { sessions: scriptedModel(happyPath).port, registerDeliverable: deliverables.registerDeliverable })

    expect(state.status).toBe("completed")
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title])).toEqual([["d1", "Supply note"]])
    expect(Object.keys(state.stages)).toEqual(["discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "finalize", "checklist-check", "package-review", "gate"])
    expect(existsSync(leanFile(directory, "d2"))).toBe(false)
    expect(readLeanDocuments(directory)!.documents.map(row => `${row.kind}:${row.document}:${row.deliverable}`)).toEqual([
      "body:d1:Supply note.docx", "redline:d1:Supply note - redline.docx", "drafting_notes:null:Supply note - drafting notes.docx", "drafting_notes_markdown:null:null",
    ])
    expect(readLeanJson(directory, "gate.json")).toMatchObject({ status: "pass", delivery: "execution_ready", disclosed: [], blocking: [] })
  }, 120_000)

  test("later document writers remain scoped while semantic review can identify an earlier owned body", async () => {
    const directory = matter()
    const deliverables = registering()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
    const state = await runLean(directory, { sessions: model.port, registerDeliverable: deliverables.registerDeliverable })

    expect(state.status).toBe("completed")
    expect(Object.keys(state.stages)).toEqual([
      "discrepancy", "intake", "research", "research-review", "spine", "d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "d2.draft", "d2.review", "d2.review-independent", "d2.review-wording", "d2.review-client-disadvantage", "d2.change-plan", "d2.change-prepare", "d2.changes", "d2.change-review", "d2.apply", "finalize", "checklist-check", "package-review", "gate",
    ])
    expect(readLeanPlan(directory)?.documents.map(row => [row.id, row.title, row.parts.length])).toEqual([["d1", "Supply note", 2], ["d2", "Covering letter", 2]])
    const firstPids = [...readParagraphs(directory, "d1"), ...readParagraphs(directory, "d1", "paragraphs-applied.json")].map(row => row.pid)
    const later = [promptsOf(model, AGENT.draft, "Covering letter")[0]!, ...promptsOf(model, AGENT.review, "Covering letter"), ...promptsOf(model, AGENT["review-independent"], "Covering letter"), ...promptsOf(model, AGENT["review-wording"], "Covering letter"), ...promptsOf(model, AGENT["review-client-disadvantage"], "Covering letter"), ...promptsOf(model, AGENT["change-plan"], "Covering letter"), ...promptsOf(model, AGENT.changes, "Covering letter"), ...promptsOf(model, AGENT["change-review"], "Covering letter")]
    expect(later).toHaveLength(8)
    later.forEach(prompt => {
      expect(prompt).toContain("<completed_documents>")
      expect(prompt).toContain(APPLIED)
      expect(prompt).not.toContain(`${DRAFTED}\n`)
      if (prompt.includes("<artifact_phase>before-apply</artifact_phase>")) {
        expect(prompt).toContain('<earlier_body document="d1">')
        firstPids.forEach(pid => expect(prompt).toContain(pid))
      } else firstPids.forEach(pid => expect(prompt).not.toContain(pid))
    })
    expect(promptsOf(model, AGENT.draft, "Covering letter")[0]).toContain('Draft part 1 of 2 of "Covering letter": Enclosure')
    expect(model.prompts.filter(prompt => prompt.agent !== AGENT.finalize && !later.includes(prompt.text)).some(prompt => prompt.text.includes("<completed_documents>"))).toBe(false)
    expect(model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text).toMatch(/<document id="d1" title="Supply note">[\s\S]*<document id="d2" title="Covering letter">/)
    ;["d2/body-applied.docx", "d2/redline.docx", "d1/redline.docx"].forEach(name => expect(existsSync(leanFile(directory, name))).toBe(true))
    expect(readLeanDocuments(directory)!.documents.map(row => `${row.kind}:${row.document}:${row.title}:${row.deliverable}`)).toEqual([
      "body:d1:Supply note:Supply note.docx",
      "redline:d1:Supply note:Supply note - redline.docx",
      "body:d2:Covering letter:Covering letter.docx",
      "redline:d2:Covering letter:Covering letter - redline.docx",
      "drafting_notes:null:null:Supply note - drafting notes.docx",
      "drafting_notes_markdown:null:null:null",
    ])
  }, 120_000)

  test("reviews and amendment stages scope other documents by id, title and purpose while retaining their own outline", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    const plan = readLeanPlan(directory)!
    expect(plan.documents.map(row => row.purpose)).toEqual(["", ""])
    const line = (index: number, marked: boolean) => `- ${plan.documents[index]!.id} "${plan.documents[index]!.title}"${marked ? " (this document)\n" : "\n"}`
    ;[["Supply note", 0], ["Covering letter", 1]].forEach(([title, index]) => {
      const review = promptsOf(model, AGENT.review, title as string)[0]!
      const change = promptsOf(model, AGENT.changes, title as string)[0]!
      ;[review, change].forEach(prompt => {
        expect(prompt).toContain("<planned_documents>")
        expect(prompt).toContain(line(0, index === 0))
        expect(prompt).toContain(line(1, index === 1))
      })
      expect(review).toContain(REVIEW_SCOPE)
      expect(change).toContain(CHANGES_SCOPE)
    })

    const steps: StepID[] = ["d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review"]
    const keys = () => steps.map(step => hashStageInputs(step, LEAN_PROMPT_VERSION, stageInputs(directory, step)))
    const before = keys()
    const edit = (documents: typeof plan.documents) => writeLeanText(directory, "plan.json", `${JSON.stringify({ ...readLeanPlan(directory)!, documents }, null, 2)}\n`)
    edit(plan.documents.map(row => row.id === "d2" ? { ...row, parts: [...row.parts].reverse() } : row))
    expect(keys()).toEqual(before)
    edit(plan.documents.map(row => row.id === "d1" ? { ...row, parts: [...row.parts].reverse() } : row))
    expect(keys().map((key, index) => key !== before[index])).toEqual([false, false, false, false, true, true, true, true])
    edit(plan.documents.map(row => row.id === "d2" ? { ...row, purpose: "Covers the delivery." } : row))
    expect(keys().map((key, index) => key !== before[index])).toEqual(steps.map(() => true))
    expect(stagePrompt(directory, "d1.review")).toContain('- d2 "Covering letter": Covers the delivery.\n')
  }, 120_000)

  test("a one-document plan shows no planned-documents list or scope line", async () => {
    const directory = matter()
    const model = scriptedModel(happyPath)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    // The coverage check reads the plan itself; drafting and review prompts show no list.
    model.prompts.filter(prompt => prompt.agent !== LEAN_STAGE_AGENTS.coverage).forEach(prompt => {
      expect(prompt.text).not.toContain("<planned_documents>")
      expect(prompt.text).not.toContain(REVIEW_SCOPE)
      expect(prompt.text).not.toContain(CHANGES_SCOPE)
    })
    expect(stageInputs(directory, "d1.review").find(input => input.name === "planned-documents")?.bytes).toBe("")
  }, 120_000)

  test("one placeholder wording in two documents is two open items; the notes must name both and the gate settles each in its document", async () => {
    const address = "\n\n## Address\n\nDeliver to [open item: confirm the delivery address].\n"
    const directory = matter()
    const first = { prompt: "" }
    const notes: Reply = (prompt, index) => {
      first.prompt = index === 0 ? prompt : first.prompt
      return draftingNotes(index === 0 ? /<document id="d1"[\s\S]*?<\/document>/.exec(prompt)![0] : first.prompt, index)
    }
    const model = scriptedModel({
      ...happyPath, [AGENT.draft]: drafter(`${DRAFT.trimEnd()}${address}`, `${LETTER.trimEnd()}${address}`),
      [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst, [AGENT.finalize]: notes,
    })
    const state = await runLean(directory, { sessions: model.port })

    const placeholders = readOpenItems(directory).filter(item => item.stage === "apply" && item.kind === "placeholder")
    expect(placeholders.map(item => item.ref.document)).toEqual(["d1", "d2"])
    expect(new Set(placeholders.map(item => item.id)).size).toBe(2)
    expect(new Set(placeholders.map(item => item.detail)).size).toBe(1)
    const retries = model.prompts.filter(prompt => prompt.agent === AGENT.finalize).slice(1)
    expect(retries).toHaveLength(1)
    expect(retries[0]!.text).toContain("did not pass the finalize checks")
    expect(retries[0]!.text).toContain(placeholders[1]!.id)
    expect(retries[0]!.text).not.toContain(placeholders[0]!.id)
    expect(state.status).toBe("completed")
    const gate = readLeanJson(directory, "gate.json") as GateResult
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.disclosed.map(item => [item.id, item.ref.document])).toEqual(placeholders.map(item => [item.id, item.ref.document]))
  }, 120_000)

  test("a question in the second document stops application and the third document", async () => {
    const directory = matter()
    const marked = "The goods will be delivered as the supply note states [open item: whether late delivery carries a remedy]."
    const asking: Reply = prompt => {
      if (listed(prompt, DRAFTED)) return changes(prompt, 0)
      const letter = listed(prompt, "The goods will be delivered as the supply note states.")
      const pid = letter ? pidIn(prompt, "The goods will be delivered") : pidIn(prompt, "Delivery follows")
      const id = /^- (\S+)(?: \[p[0-9a-f]{12}\])?: Should late delivery/m.exec(prompt)?.[1]
      return JSON.stringify({
        amendments: [{ pid, updated: letter ? marked : "Delivery follows the supply note and the covering letter [open item: whether late delivery carries a remedy].",
          reason: "The remedy for late delivery is the lawyer's decision.", sources: SUPPORT }],
        questions: [{ ...(id ? { id } : {}), pid, question: QUESTION, kind: "judgment", context: "The terms are silent." }],
      })
    }
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER, SCHEDULE), [AGENT.review]: reviewFirst, [AGENT.changes]: asking })
    const state = await runLean(directory, { sessions: model.port })

    expect(Object.keys(state.stages).filter(step => step.startsWith("d3."))).toEqual(["d3.draft", "d3.review", "d3.review-independent", "d3.review-wording", "d3.review-client-disadvantage", "d3.change-plan", "d3.change-prepare", "d3.changes", "d3.change-review", "d3.apply"])
    expect(state.status).toBe("blocked")
    expect(promptsOf(model, AGENT.draft, "Delivery schedule")).toEqual([])
    expect(state.stages["d2.apply"]?.status).toBe("pending")
    expect(state.stages["d3.draft"]?.status).toBe("pending")
    expect(readQuestions(directory)).toEqual([expect.objectContaining({ question: QUESTION, document: "d2", status: "open" })])
    expect(model.count(AGENT.finalize)).toBe(0)
  }, 120_000)

  test("an answer to a first-document question repeats earlier work but its accepted draft before starting the later document", async () => {
    const directory = matter()
    const firstChanges: Reply = prompt => {
      const answered = prompt.includes("No remedy; say so.")
      return JSON.stringify({
        amendments: [{ pid: pidIn(prompt, "The supplier shall deliver the goods\\."),
          updated: answered ? `${APPLIED.slice(0, -1)}, and late delivery carries no remedy.` : APPLIED, reason: "Delivery period.", sources: SUPPORT }],
        questions: answered ? [] : [{ question: QUESTION, kind: "judgment", context: "The terms are silent." }],
      })
    }
    // The letter restates the supply note's delivery sentence, so its change list follows that wording.
    const letterChanges: Reply = prompt => JSON.stringify({
      amendments: [{ pid: pidIn(prompt, "The goods will be delivered"), updated: `The goods will be delivered as the supply note states: ${/^The supplier shall deliver.*$/m.exec(prompt)![0]}`,
        reason: "Agrees with the supply note.", sources: SUPPORT }],
      questions: [],
    })
    const replies = (): Record<string, Reply> => ({
      ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst,
      [AGENT.changes]: sameSession((prompt, index) => listed(prompt, DRAFTED) ? firstChanges(prompt, index) : letterChanges(prompt, index)),
    })
    const blocked = await runLean(directory, { sessions: scriptedModel(replies()).port })
    expect(blocked.status).toBe("blocked")
    const [question] = readQuestions(directory)
    expect(question).toMatchObject({ document: "d1", status: "open" })

    enqueueLeanRequest(directory, { kind: "answer", question_id: question!.id, answer: "No remedy; say so." })
    const answered = scriptedModel(replies())
    const resumed = await runLean(directory, { sessions: answered.port })
    expect(resumed.status).toBe("completed")
    // Only the later document is drafted; the first document's accepted draft is kept.
    expect(answered.counts()).toEqual([2, 2, 2, 1])
    const rerun = Object.keys(resumed.stages).filter(step => resumed.stages[step as keyof typeof resumed.stages]?.finished_at !== blocked.stages[step as keyof typeof blocked.stages]?.finished_at)
    expect(rerun).toEqual(["discrepancy", "intake", "research", "research-review", "spine", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply", "d2.draft", "d2.review", "d2.review-independent", "d2.review-wording", "d2.review-client-disadvantage", "d2.change-plan", "d2.change-prepare", "d2.changes", "d2.change-review", "d2.apply", "finalize", "checklist-check", "package-review", "gate"])
    expect(readParagraphs(directory, "d2", "paragraphs-applied.json").map(row => row.text)).toContain(
      "The goods will be delivered as the supply note states: The supplier shall deliver the goods within thirty days of the order date, and late delivery carries no remedy.")

    // Changing the accepted explanation revisits later documents even when the final wording is unchanged.
    const list = JSON.parse(readFileSync(leanFile(directory, "d1/changes.json"), "utf8"))
    list.amendments[0].reason = "Edited by the lawyer."
    writeFileSync(leanFile(directory, "d1/changes.json"), JSON.stringify(list, null, 2))
    const edited = scriptedModel(replies())
    const reapplied = await runLean(directory, { sessions: edited.port })
    expect(reapplied.status).toBe("completed")
    // An edited rationale reaches the dependent covering letter through its review and change plan, not a redraft.
    expect(edited.counts()).toEqual([0, 1, 1, 1])
    expect(edited.count(AGENT["change-review"])).toBe(2)
    expect(edited.count(AGENT["change-plan"])).toBe(1)
    expect(edited.count(AGENT["change-prepare"])).toBe(0)
    expect(reapplied.stages["d1.apply"]?.finished_at).not.toBe(resumed.stages["d1.apply"]?.finished_at)
    // The letter's accepted draft is kept, and with its change list unchanged its applied body stands.
    ;(["d2.draft", "d2.apply"] as const).forEach(step => expect(reapplied.stages[step]?.finished_at).toBe(resumed.stages[step]?.finished_at))
    ;(["d2.review", "d2.review-independent", "d2.review-wording", "d2.review-client-disadvantage", "d2.change-plan", "d2.change-prepare", "d2.changes", "d2.change-review"] as const).forEach(step => expect(reapplied.stages[step]?.finished_at).not.toBe(resumed.stages[step]?.finished_at))
  }, 180_000)

  test("restoring to just after the first document's apply re-runs only the later document, the notes and the gate", async () => {
    const directory = matter()
    const replies = { ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }
    const first = await runLean(directory, { sessions: scriptedModel(replies).port })
    expect(first.status).toBe("completed")

    const restored = restoreLeanStage(directory, "d1.apply")
    expect(Object.entries(restored.stages).map(([step, row]) => `${step}=${row?.status}`)).toEqual([
      "discrepancy=done", "intake=done", "research=done", "research-review=done", "spine=done", "d1.draft=done", "d1.review=done", "d1.review-independent=done", "d1.review-wording=done", "d1.review-client-disadvantage=done", "d1.change-plan=done", "d1.change-prepare=done", "d1.changes=done", "d1.change-review=done", "d1.apply=done",
      "d2.draft=pending", "d2.review=pending", "d2.review-independent=pending", "d2.review-wording=pending", "d2.review-client-disadvantage=pending", "d2.change-plan=pending", "d2.change-prepare=pending", "d2.changes=pending", "d2.change-review=pending", "d2.apply=pending", "finalize=pending", "checklist-check=pending", "package-review=pending", "gate=pending",
    ])
    expect(existsSync(leanFile(directory, "d1/redline.docx"))).toBe(true)
    expect(existsSync(leanFile(directory, "d2/body-applied.docx"))).toBe(false)

    const resume = scriptedModel(replies)
    const resumed = await runLean(directory, { sessions: resume.port })
    expect(resumed.status).toBe("completed")
    expect(resume.counts()).toEqual([2, 1, 1, 1])
    expect(resume.sessions.every(session => !session.title.includes("Supply note"))).toBe(true)
    ;(["d1.draft", "d1.review", "d1.review-independent", "d1.review-wording", "d1.review-client-disadvantage", "d1.change-plan", "d1.change-prepare", "d1.changes", "d1.change-review", "d1.apply"] as const).forEach(step => expect(resumed.stages[step]?.finished_at).toBe(first.stages[step]?.finished_at))
  }, 120_000)

  test("an outline that repeats a document title is retried; the fourth failure stops the run", async () => {
    const directory = matter()
    const twice = "# Supply note\n## Delivery\n# Supply note\n## Payment"
    const model = scriptedModel({ [AGENT.draft]: () => twice })
    const state = await runLean(directory, { sessions: model.port })

    expect(model.count(AGENT.draft)).toBe(4)
    expect(state.status).toBe("failed")
    expect(state.stop_reason).toStartWith('d1.draft: outline: checker failed after 3 retries: These document titles appear more than once: "Supply note"')
    expect(existsSync(leanFile(directory, "plan.json"))).toBe(false)
  }, 120_000)
})
