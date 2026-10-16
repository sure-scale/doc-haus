import { afterEach, expect, test } from "bun:test"
import { readFileSync } from "node:fs"
import { checkChanges } from "./change-checks"
import { runApply } from "./apply-stage"
import { changePlanPrompt, projectedChangeParagraphs, repairBodyText } from "./change-pipeline"
import { finalBodyParagraphs, readParagraphs } from "./checks"
import { stagePrompt } from "./prompts"
import { readRepairRound } from "./repairs"
import { runLean } from "./runner"
import { acceptStageReply, stageOutputValid } from "./stages"
import { leanFile, readLeanJson, readLeanState, writeLeanText } from "./state"
import { AGENT, APPLIED, DRAFT, LETTER, SOURCE, SUPPORT, changesFirst, cleanupMatters, drafter, happyPath, matter, pidOf, reviewFirst, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

test("an earlier-body repair regenerates dependent drafting without projecting its old amendments onto new paragraphs", async () => {
  const directory = matter()
  const original = "The goods will be delivered as the supply note states."
  const corrected = "The goods will be delivered within thirty days of the order date."
  const earlierDetail = "State expressly that payment is due on delivery in the operative note."
  const laterDetail = "Reassess the delivery summary after the operative note is corrected."
  let challenged = false
  let oldLaterPids: string[] = []
  let oldLaterDraft = ""
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst,
    [AGENT.changes]: (prompt, index) => {
      if (prompt.includes('"title":"Covering letter"')) return JSON.stringify({ amendments: [{ pid: pidOf(prompt, original), updated: corrected,
        reason: "Restates the supplied delivery period accurately.", sources: SUPPORT }], questions: [] })
      const first = JSON.parse(changesFirst(prompt, index))
      if (!prompt.includes(earlierDetail)) return JSON.stringify(first)
      return JSON.stringify({ ...first, amendments: [...first.amendments, { pid: pidOf(prompt, "Payment is due on delivery."),
        updated: "The buyer shall pay on delivery.", reason: "Expressly identifies the supplied payment timing.",
        sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] }] })
    },
    [AGENT["change-review"]]: prompt => {
      if (!challenged && prompt.includes('"title":"Covering letter"')) {
        challenged = true
        oldLaterPids = readParagraphs(directory, "d2").map(row => row.pid)
        oldLaterDraft = readFileSync(leanFile(directory, "d2/draft.md"), "utf8")
        return JSON.stringify({ status: "changes-required", findings: [
          { document: "d1", detail: earlierDetail }, { document: "d2", pid: readParagraphs(directory, "d2").find(row => row.text === original)!.pid, detail: laterDetail },
        ], limitations: [] })
      }
      return JSON.stringify({ status: "pass", findings: [], limitations: [] })
    },
  })
  const result = await runLean(directory, { sessions: model.port })
  expect(result.status).toBe("completed")
  expect(challenged).toBe(true)
  expect(readFileSync(leanFile(directory, "d2/draft.md"), "utf8")).toBe(oldLaterDraft)
  expect(readParagraphs(directory, "d2").every(row => !oldLaterPids.includes(row.pid))).toBe(true)
  expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).toContain(APPLIED)
  expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).toContain("The buyer shall pay on delivery.")
  expect(finalBodyParagraphs(directory, "d2").map(row => row.text)).toContain(corrected)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
  const rebuiltPlan = model.prompts.filter(row => row.agent === AGENT["change-plan"]).at(-1)!.text
  expect(rebuiltPlan).toContain(laterDetail)
  expect(rebuiltPlan).toContain("earlier draft")
  for (const pid of oldLaterPids) {
    expect(rebuiltPlan).not.toContain(pid)
    expect(model.prompts.filter(row => row.agent === AGENT.changes).at(-1)?.text).not.toContain(pid)
  }
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.prompts).toEqual([])
}, 120_000)

test("empty saved amendments cannot present an older applied body as the final body of a regenerated draft", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const oldFinal = finalBodyParagraphs(directory, "d2")
  expect(JSON.parse(readFileSync(leanFile(directory, "d2/changes.json"), "utf8")).amendments).toEqual([])
  expect(repairBodyText(directory, "d2", false)).toContain(`[${oldFinal[0]!.pid}]`)
  expect(repairBodyText(directory, "d2", false)).not.toContain("earlier draft")
  expect((await acceptStageReply(directory, "d2.draft", LETTER)).issues).toEqual([])
  expect(finalBodyParagraphs(directory, "d2")).toEqual(oldFinal)
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 2, substantive_attempts: 2, requested_at: new Date().toISOString(),
    items: [{ id: "saved-package-finding", stage: "package-review", kind: "omission", status: "open", detail: "Reassess this summary against the corrected earlier note.",
      ref: { document: "d2", repair_target: "body", pid: oldFinal[0]!.pid } }] }))
  for (const text of [repairBodyText(directory, "d2", false), changePlanPrompt(directory, "d2"), stagePrompt(directory, "d2.changes")]) {
    expect(text).toContain("earlier draft")
    expect(text).toContain(`[${readParagraphs(directory, "d2")[0]!.pid}]`)
    for (const row of oldFinal) expect(text).not.toContain(`[${row.pid}]`)
  }
  expect(readRepairRound(directory)?.substantive_attempts).toBe(2)
}, 120_000)

test("stale amendment context is advisory only; strict projection and accepted output still reject it", async () => {
  const directory = matter()
  expect((await acceptStageReply(directory, "d1.draft", DRAFT)).issues).toEqual([])
  const prompt = stagePrompt(directory, "d1.changes")
  const reply = changesFirst(prompt, 0)
  expect((await acceptStageReply(directory, "d1.changes", reply)).issues).toEqual([])
  expect(projectedChangeParagraphs(directory, "d1").map(row => row.text)).toContain(APPLIED)
  const stale = readFileSync(leanFile(directory, "d1/changes.json"), "utf8")
  expect((await acceptStageReply(directory, "d1.draft", DRAFT)).issues).toEqual([])
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 3, substantive_attempts: 3, requested_at: new Date().toISOString(),
    items: [{ id: "historical-finding", stage: "change-review", kind: "omission", status: "open", detail: "Check the delivery period in the regenerated note.",
      ref: { document: "d1", repair_target: "body", pid: JSON.parse(stale).amendments[0].pid } }] }))
  expect(() => stagePrompt(directory, "d1.review")).not.toThrow()
  expect(changePlanPrompt(directory, "d1")).toContain("Check the delivery period in the regenerated note.")
  expect(changePlanPrompt(directory, "d1")).toContain("earlier draft")
  expect(stagePrompt(directory, "d1.changes")).toContain("earlier draft")
  expect(() => projectedChangeParagraphs(directory, "d1")).toThrow("unknown initial paragraph")
  expect(checkChanges(directory, stale, readParagraphs(directory, "d1"), { document: "d1" }).issues.length).toBeGreaterThan(0)
  expect(stageOutputValid(directory, "d1.changes")).toBe(false)
  const apply = await runApply(directory, readLeanState(directory)!, { sessions: scriptedModel({}).port }, "d1")
  expect(apply.status).toBe("done")
  expect(readLeanJson(directory, "d1/apply.json")).toMatchObject({ applied: [], unapplied: [expect.objectContaining({ pid: JSON.parse(stale).amendments[0].pid })] })
  expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).not.toContain(APPLIED)
  expect(readFileSync(leanFile(directory, "d1/changes.json"), "utf8")).toBe(stale)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(3)
}, 120_000)
