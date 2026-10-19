import { afterEach, expect, test } from "bun:test"
import { existsSync } from "node:fs"
import { checkChangeReview } from "./change-pipeline"
import { finalBodyParagraphs, readParagraphs } from "./checks"
import { readRepairRound } from "./repairs"
import { runLean } from "./runner"
import { leanFile, readLeanJson, writeLeanText } from "./state"
import { AGENT, APPLIED, DRAFT, LETTER, SOURCE, SUPPORT, changes, changesFirst, cleanupMatters, drafter, happyPath, matter, pidOf, reviewFirst, sameSession, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const PAYMENT = "Payment shall be due on delivery."
const MISSING = "Payment timing remains to be agreed."

test("a later-body finding repairs the earlier owned body and retains accepted amendments", async () => {
  const directory = matter()
  const first = DRAFT.replace("Payment is due on delivery.", MISSING)
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(first, LETTER), [AGENT.review]: reviewFirst,
    [AGENT.changes]: (prompt, index) => {
      if (!prompt.includes("Restore source payment timing.")) return changesFirst(prompt, index)
      return JSON.stringify({ amendments: [
        JSON.parse(changes(prompt, index)).amendments[0],
        { pid: pidOf(prompt, MISSING), updated: PAYMENT, reason: "Restores the supplied payment timing.", sources: [{ source_id: SOURCE, quote: "Payment is due on delivery." }] },
      ], questions: [] })
    },
    [AGENT["change-review"]]: sameSession(prompt => {
      if (!prompt.includes("<projected_full_body>\n") || !prompt.includes("We enclose the supply note for signature."))
        return JSON.stringify({ status: "pass", findings: [], limitations: [] })
      const missing = finalBodyParagraphs(directory, "d1").some(row => row.text === MISSING)
      return JSON.stringify(missing ? { status: "changes-required", findings: [{ document: "d1", pid: readParagraphs(directory, "d1").find(row => row.text === MISSING)!.pid, detail: "Restore source payment timing." }], limitations: [] }
        : { status: "pass", findings: [], limitations: [] })
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const body = finalBodyParagraphs(directory, "d1").map(row => row.text)
  expect(body).toContain(PAYMENT)
  expect(body).toContain(APPLIED)
  expect(finalBodyParagraphs(directory, "d1").map(row => row.pid)).toEqual(readParagraphs(directory, "d1").map(row => row.pid))
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
  expect(readRepairRound(directory)?.items[0]?.ref.document).toBe("d1")
  // The covering letter is drafted once; the earlier body's repair reaches it through its change review.
  expect(model.prompts.filter(row => row.agent === AGENT.draft && /Draft part \d+ of \d+ of "Covering letter"/.test(row.text))).toHaveLength(2)
  expect(model.prompts.filter(row => row.agent === AGENT["change-review"] && row.text.includes("We enclose the supply note for signature.")).length).toBeGreaterThanOrEqual(2)
}, 120_000)

test("review ownership admits current and earlier bodies but rejects future, external and mismatched pointers", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst, [AGENT.changes]: changesFirst }).port })).status).toBe("completed")
  const pid = readParagraphs(directory, "d1")[0]!.pid
  const finding = { status: "changes-required", findings: [{ document: "d1", pid, detail: "Restore the supplied wording in its owning body." }], limitations: [] }
  expect(checkChangeReview(directory, "d2", finding)).toEqual([])
  expect(checkChangeReview(directory, "d1", finding)).toEqual([])
  expect(checkChangeReview(directory, "d2", { ...finding, findings: [{ ...finding.findings[0], document: "../../outside" }] }).length).toBeGreaterThan(0)
  expect(checkChangeReview(directory, "d1", { ...finding, findings: [{ ...finding.findings[0], document: "d2" }] }).length).toBeGreaterThan(0)
  expect(checkChangeReview(directory, "d2", { ...finding, findings: [{ ...finding.findings[0], pid: readParagraphs(directory, "d2")[0]!.pid }] }).length).toBeGreaterThan(0)
}, 120_000)

test("pre-apply review exposes planned explanations while actual comments remain post-apply evidence", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT["change-review"]]: prompt => {
    expect(existsSync(leanFile(directory, "d1/redline.docx"))).toBe(false)
    expect(prompt).toContain('<artifact_phase>before-apply</artifact_phase>')
    expect(prompt).toContain("States the delivery period from the supply terms.")
    expect(prompt).toContain(SUPPORT[0]!.quote)
    return JSON.stringify({ status: "pass", findings: [], limitations: [] })
  }, [AGENT["package-review"]]: prompt => {
    expect(existsSync(leanFile(directory, "d1/redline.docx"))).toBe(true)
    expect(prompt).toContain('"native_comment"')
    expect(prompt).toContain('"status": "attached"')
    return JSON.stringify({ status: "pass", delivery: [], findings: [], limitations: [] })
  } })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
}, 120_000)

test("a resumed substantive allowance preserves three previous attempts and permits the fourth repair", async () => {
  const directory = matter()
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 3, requested_at: new Date().toISOString(), items: [], substantive_attempts: 3 }))
  const model = scriptedModel({ ...happyPath, [AGENT["package-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] }
    : { status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Describe the supplied wording accurately." }], limitations: [] }) })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(4)
  expect(readLeanJson(directory, "repairs.json")).toMatchObject({ round: 3 })
}, 120_000)
