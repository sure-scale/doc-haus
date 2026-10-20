import { afterEach, describe, expect, test } from "bun:test"
import type { GateResult } from "./gate"
import { runLean } from "./runner"
import { leanFile, readLeanJson, readLeanText, writeLeanText } from "./state"
import { AGENT, DRAFT, DRAFTED, LETTER, SUPPORT, changesFirst, cleanupMatters, drafter, happyPath, listed, matter, pidOf, scriptedModel, type Reply } from "./test-support"
import { existsSync } from "node:fs"

// Review findings that concern another document of the package. A finding
// names the document it concerns; it reaches that document when it is still to
// be drafted, is filed and disclosed under it when it is already final, and is
// carried to the notes when it belongs in the accompanying memo.

afterEach(cleanupMatters)

const FORWARD = "The covering letter should tell the buyer that delivery is due within thirty days of the order date."
const BACKWARD = "The supply note does not say where the goods are delivered."
const MEMO = "The memo should explain that the delivery period runs from the order date, not from acceptance."

function promptsOf(model: ReturnType<typeof scriptedModel>, agent: string, title: string) {
  const sessions = new Set(model.sessions.flatMap((session, index) => session.agent === agent && session.title.includes(`: ${title}: `) ? [`ses_child${index + 1}`] : []))
  return model.prompts.filter(prompt => sessions.has(prompt.sessionID)).map(prompt => prompt.text)
}

/** The supply note's review finds its own delivery period and a point for the
 * covering letter; the covering letter's review finds a point for the supply
 * note, which is final by then. */
const reviews = (memo = false): Reply => prompt => JSON.stringify({
  findings: listed(prompt, DRAFTED) ? [
    { pid: pidOf(prompt, DRAFTED), origin: "source_conflict", finding: "The delivery period is missing.", sources: SUPPORT },
    { document: "d2", origin: "request", finding: FORWARD, sources: SUPPORT },
    ...(memo ? [{ document: "memo", origin: "request", finding: MEMO, sources: [] }] : []),
  ] : [{ document: "d1", origin: "drafting", finding: BACKWARD, sources: [] }],
})

describe("cross-document review findings", () => {
  test("a finding for a later document reaches its change plan and draft, not the reviewing document's plan", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviews(), [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    promptsOf(model, AGENT.review, "Supply note").forEach(prompt => {
      expect(prompt).not.toContain("do not report it here")
      expect(prompt).toContain("report it with that document's id as its document")
    })
    expect(promptsOf(model, AGENT["change-plan"], "Supply note")[0]).not.toContain(FORWARD)
    expect(promptsOf(model, AGENT["change-plan"], "Covering letter")[0]).toContain(FORWARD)
    expect(promptsOf(model, AGENT.draft, "Covering letter").some(prompt => prompt.includes(FORWARD))).toBe(true)
  }, 120_000)

  test("a finding for an earlier, final document is filed under it and disclosed", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviews(), [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    const finalize = model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text
    expect(/<document id="d1"[\s\S]*?<\/document>/.exec(finalize)![0]).toContain(BACKWARD)
    expect((readLeanJson(directory, "gate.json") as GateResult).items.some(item => item.ref.document === "d1" && item.status === "disclosed" && item.detail.includes(BACKWARD))).toBe(true)
  }, 120_000)

  test("a finding for the accompanying memo reaches the notes", async () => {
    const directory = matter()
    writeLeanText(directory, "request.md", "Draft a short supply note from the supply terms, a covering letter and a cover memo.\n")
    const draft = drafter(DRAFT, LETTER)
    const outline: Reply = (prompt, index) => /^Draft part \d+ of \d+ of /m.test(prompt) ? draft(prompt, index) : `Memo: cover memo\n${draft(prompt, index)}`
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: outline, [AGENT.review]: reviews(true), [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(model.prompts.find(prompt => prompt.agent === AGENT.finalize)!.text).toContain(MEMO)
    expect((readLeanJson(directory, "gate.json") as GateResult).items.some(item => item.status === "disclosed" && item.detail.includes(MEMO))).toBe(true)
  }, 120_000)

  test("the register lists every review finding with its document, origin, basis and disposition", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviews(), [AGENT.changes]: changesFirst })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")

    expect(existsSync(leanFile(directory, "register.json"))).toBe(true)
    const register = readLeanJson(directory, "register.json") as { entries: { id: string; document: string; raised_in: string; origin: string; basis: string; disposition: string; finding: string }[] }
    expect(register.entries.map(row => [row.document, row.raised_in, row.origin, row.basis, row.finding]).toSorted()).toEqual([
      ["d1", "d1", "source_conflict", "source", "The delivery period is missing."],
      ["d1", "d2", "drafting", "practice", BACKWARD],
      ["d2", "d1", "request", "source", FORWARD],
    ].toSorted())
    register.entries.forEach(row => expect(row.id).toMatch(/^f-[0-9a-f]{10}$/))
    expect(register.entries.find(row => row.finding === BACKWARD)!.disposition).toBe("earlier_document")
    const notes = readLeanText(directory, "drafting-notes.md")!
    register.entries.forEach(row => expect(notes).not.toContain(row.id))
  }, 120_000)
})
