import { afterEach, describe, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { finalBodyParagraphs, readParagraphs } from "./checks"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { leanFile, readLeanJson, readLeanState, writeLeanText, writeQuestions } from "./state"
import { AGENT, DRAFT, LETTER, QUESTION, SOURCE, changesFirst, cleanupMatters, drafter, happyPath, matter, pidOf, reviewFirst, scriptedModel } from "./test-support"

// A produced docx is never redrafted. Once a document's body.docx and
// paragraphs.json are accepted, what changes afterwards (an answer, a lawyer's
// update, research, an amendment to an earlier document) reaches it through its
// review and change-plan steps as amendments, never through its draft step.
// Real coordinator, checkers, renderer and applier; only the model is scripted.

afterEach(cleanupMatters)

const replies = { ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }
const PASS = JSON.stringify({ status: "pass", findings: [], limitations: [], delivery: [] })

/** Each document's accepted draft: its Word bytes and its draft step's state. */
function accepted(directory: string) {
  const state = readLeanState(directory)!
  return (["d1", "d2"] as const).map(document => ({ body: readFileSync(leanFile(directory, `${document}/body.docx`)), step: state.stages[`${document}.draft`] }))
}

/** Prompts sent in a document's sessions of one stage. */
function stagePrompts(model: ReturnType<typeof scriptedModel>, title: string, stage: string) {
  const sessions = new Set(model.sessions.flatMap((session, index) => session.title === `Lean drafting: ${title}: ${stage}` ? [`ses_child${index + 1}`] : []))
  return model.prompts.filter(prompt => sessions.has(prompt.sessionID))
}

const later: [string, (directory: string) => void][] = [
  ["a queued answer", directory => {
    writeQuestions(directory, [{ id: "q-late", stage: "change-plan", document: "d2", pid: readParagraphs(directory, "d2")[0]!.pid, question: QUESTION, context: "The terms are silent.", status: "open" }])
    enqueueLeanRequest(directory, { kind: "answer", question_id: "q-late", answer: "No remedy." })
  }],
  ["a lawyer's update", directory => { enqueueLeanRequest(directory, { kind: "update", text: "Use plain English throughout." }) }],
  ["a research change", directory => {
    const research = readLeanJson(directory, "research.json") as { limitations: string[] }
    writeLeanText(directory, "research.json", `${JSON.stringify({ ...research, limitations: [...research.limitations, "Local practice on delivery periods was not researched."] }, null, 2)}\n`)
  }],
]

describe("an accepted draft is final", () => {
  for (const [name, change] of later)
    test(`after both documents apply, ${name} re-runs the change plans and leaves each accepted draft as it is`, async () => {
      const directory = matter()
      expect((await runLean(directory, { sessions: scriptedModel(replies).port })).status).toBe("completed")
      const before = accepted(directory)
      change(directory)

      const model = scriptedModel(replies)
      expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
      expect(model.count(AGENT.draft)).toBe(0)
      expect(accepted(directory)).toEqual(before)
      expect(stagePrompts(model, "Supply note", "change-plan").length).toBeGreaterThan(0)
      expect(stagePrompts(model, "Covering letter", "change-plan").length).toBeGreaterThan(0)
    }, 180_000)

  test("after both documents apply, an amendment to the first reaches the second as its change-plan input, not a redraft", async () => {
    const directory = matter()
    const DETAIL = "State expressly in the supply note that the buyer pays on delivery."
    const amended = "The buyer shall pay on delivery."
    let before: ReturnType<typeof accepted> | undefined
    const model = scriptedModel({ ...replies,
      // Package review runs once both documents have applied; its one finding amends the first.
      [AGENT["package-review"]]: () => {
        if (before) return PASS
        before = accepted(directory)
        return JSON.stringify({ status: "changes-required", findings: [{ target: "body", document: "d1", detail: DETAIL }], limitations: [], delivery: [] })
      },
      [AGENT.changes]: (prompt, index) => {
        const first = JSON.parse(changesFirst(prompt, index))
        if (!prompt.includes(DETAIL)) return JSON.stringify(first)
        return JSON.stringify({ ...first, amendments: [...first.amendments, { pid: pidOf(prompt, "Payment is due on delivery."), updated: amended,
          reason: "Expressly identifies the supplied payment timing.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }] })
      },
    })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(before).toBeDefined()
    expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).toContain(amended)
    expect(accepted(directory)).toEqual(before!)
    // One draft session per document: the first document's outline and parts, and the second's parts.
    expect(model.sessions.filter(session => session.agent === AGENT.draft)).toHaveLength(2)
    expect(stagePrompts(model, "Covering letter", "draft")).toHaveLength(2)
    // The second document's change plan ran before the amendment and again after it.
    expect(stagePrompts(model, "Covering letter", "change-plan")).toHaveLength(2)
  }, 180_000)
})
