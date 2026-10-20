import { afterEach, expect, test } from "bun:test"
import { finalGate } from "./gate"
import { finalBodyParagraphs } from "./checks"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, runLean } from "./runner"
import { readLeanState, writeLeanState, writeLeanText } from "./state"
import { readRepairRound } from "./repairs"
import { enqueueLeanRequest } from "./requests"
import { AGENT, APPLIED, DRAFT, LETTER, NO_FINDINGS, SUPPORT, changesFirst, cleanupMatters, drafter, happyPath, matter, pidOf, reviewFirst, sameSession, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

test("semantic body correction routes through changes while independent reviews remain separate", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT["review-independent"]]: () => NO_FINDINGS,
    [AGENT["package-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] } : {
      status: "changes-required", delivery: [], findings: [{ target: "body", document: "d1", detail: "State the operative payment timing precisely from supplied terms." }], limitations: [],
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(2)
  expect(model.count(AGENT.review)).toBe(1)
  expect(model.count(AGENT["review-independent"])).toBe(1)
  expect(model.prompts.filter(row => row.agent === AGENT.changes).at(-1)?.text).toContain("State the operative payment timing")
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
})

test("memo-only correction re-finalizes without body rework", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath,
    [AGENT["package-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] } : {
      status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Memo must describe the final operative wording accurately." }], limitations: [],
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(1)
  expect(model.count(AGENT.finalize)).toBe(2)
})

test("a memo finding with a document owner repairs that deliverable, not the run-wide notes", async () => {
  const directory = matter()
  const original = "The goods will be delivered as the supply note states."
  const corrected = "The goods will be delivered within thirty days of the order date."
  const detail = "Correct the covering letter's delivery summary from the supplied terms."
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst,
    [AGENT.changes]: sameSession((prompt, index) => prompt.includes(detail)
      ? JSON.stringify({ amendments: [{ pid: pidOf(prompt, original), updated: corrected,
        reason: "States the supplied delivery period in the covering letter.", sources: SUPPORT }], questions: [] })
      : changesFirst(prompt, index)),
    [AGENT["package-review"]]: () => {
      const row = finalBodyParagraphs(directory, "d2").find(row => row.text === original)
      return JSON.stringify(row ? { status: "changes-required", delivery: [],
        findings: [{ target: "memo", document: "d2", pid: row.pid, detail }], limitations: [] }
        : { status: "pass", delivery: [], findings: [], limitations: [] })
    },
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(finalBodyParagraphs(directory, "d2").map(row => row.text)).toContain(corrected)
  expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).toContain(APPLIED)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(1)
  expect(readRepairRound(directory)?.items[0]?.ref).toMatchObject({ document: "d2", repair_target: "body" })
}, 120_000)

test("repeated unresolved semantic correction stays blocked and cannot be disclosed away", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath,
    [AGENT["package-review"]]: () => JSON.stringify({ status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Memo contradicts the final body." }], limitations: [] }),
  })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("blocked")
  expect(state.stop_reason).toContain(`${LEAN_MAX_SUBSTANTIVE_REPAIRS} substantive repair rounds`)
  expect(model.count(AGENT["package-review"])).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS + 1)
  expect(finalGate(directory).status).toBe("blocked")
})

test("a previously charged document-owned memo repair resumes without resetting the allowance", async () => {
  const directory = matter()
  const original = "The goods will be delivered as the supply note states."
  const corrected = "The goods will be delivered within thirty days of the order date."
  const detail = "Correct the saved covering letter's delivery summary from the supplied terms."
  let continuing = false
  const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst,
    [AGENT.changes]: sameSession((prompt, index) => prompt.includes(detail)
      ? JSON.stringify({ amendments: [{ pid: pidOf(prompt, original), updated: corrected,
        reason: "States the supplied delivery period in the covering letter.", sources: SUPPORT }], questions: [] })
      : changesFirst(prompt, index)),
    [AGENT["package-review"]]: () => {
      const row = finalBodyParagraphs(directory, "d2").find(row => row.text === original)
      return JSON.stringify(continuing && row ? { status: "changes-required", delivery: [],
        findings: [{ target: "memo", document: "d2", pid: row.pid, detail }], limitations: [] }
        : { status: "pass", delivery: [], findings: [], limitations: [] })
    },
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  const pid = finalBodyParagraphs(directory, "d2").find(row => row.text === original)!.pid
  writeLeanText(directory, "repairs.json", JSON.stringify({ round: 4, requested_at: new Date().toISOString(),
    substantive_attempts: LEAN_MAX_SUBSTANTIVE_REPAIRS, items: [{ id: "saved-document-correction", stage: "package-review", kind: "omission",
      detail, ref: { repair_target: "memo", document: "d2", pid }, status: "open" }] }))
  const saved = readLeanState(directory)!
  writeLeanState(directory, { ...saved, status: "blocked", stop_reason: "Saved memo correction was not dispatched to its owning document." })
  const draftCalls = model.count(AGENT.draft)
  enqueueLeanRequest(directory, { kind: "continue" })
  continuing = true
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(finalBodyParagraphs(directory, "d2").map(row => row.text)).toContain(corrected)
  expect(finalBodyParagraphs(directory, "d1").map(row => row.text)).toContain(APPLIED)
  expect(model.count(AGENT.draft)).toBe(draftCalls)
  expect(readRepairRound(directory)?.substantive_attempts).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS)
  expect(readRepairRound(directory)?.round).toBe(4)
}, 120_000)

test("new authority issue revisits research before rebuilding dependent work", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath,
    [AGENT["package-review"]]: (_, index) => JSON.stringify(index ? { status: "pass", delivery: [], findings: [], limitations: [] } : {
      status: "changes-required", delivery: [], findings: [{ target: "research", detail: "Resolve the newly identified authority applicability issue from actual sources." }], limitations: [],
    }),
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(model.count(AGENT.research)).toBe(2)
  expect(model.prompts.filter(row => row.agent === AGENT.research).at(-1)?.text).toContain("newly identified authority applicability")
})


test("independent reviews run concurrently while acceptance stays in stage order", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  let release: (() => void) | undefined
  let started = false
  const independent = new Promise<void>(resolve => { release = resolve })
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT["review-independent"]) {
      started = true
      await independent
      return NO_FINDINGS
    }
    if (input.agent === AGENT.review) {
      expect(started).toBe(true)
      expect(readLeanState(directory)?.stages["d1.review-independent"]).toMatchObject({ status: "running", session_id: expect.any(String) })
      release!()
    }
    return model.port.prompt(input)
  } }
  expect((await runLean(directory, { sessions: port })).status).toBe("completed")
})
