import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { checkChangeReview } from "./change-pipeline"
import { finalBodyParagraphs, readParagraphs } from "./checks"
import { finalGate } from "./gate"
import { openItem } from "./quotes"
import { documentRepairs, readRepairRound, startRepairRound } from "./repairs"
import { enqueueLeanRequest } from "./requests"
import { LEAN_MAX_SUBSTANTIVE_REPAIRS, runLean, stageInputSha } from "./runner"
import { leanFile, readLeanJson, readLeanState, readLeanText, writeLeanState, writeLeanText, writeQuestions } from "./state"
import { AGENT, APPLIED, DRAFT, LETTER, NO_FINDINGS, SOURCE, SUPPORT, changes, changesFirst, cleanupMatters, drafter, happyPath, matter, pidOf, reviewFirst, sameSession, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

describe("repair lifecycle", () => {
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
})

describe("document repair context", () => {
  function finding(directory: string, document: "d1" | "d2", detail: string, target: "body" | "research" = "body") {
    return openItem("change-review", "omission", detail, { document, repair_target: target, pid: readParagraphs(directory, document)[0]!.pid })
  }

  async function completedRepair() {
    const directory = matter()
    const model = scriptedModel({ ...happyPath, [AGENT.draft]: drafter(DRAFT, LETTER), [AGENT.review]: reviewFirst,
      [AGENT.changes]: (prompt, index) => {
        const value = JSON.parse(changesFirst(prompt, index))
        if (prompt.includes("Add the supplied payment timing to delivery.")) value.amendments[0].updated = `${APPLIED}\n\nPayment is due on delivery.`
        return JSON.stringify(value)
      } })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    // Legacy repair records have no per-document context; round3 must migrate
    // without changing the input key accepted by its existing reader.
    writeLeanText(directory, "repairs.json", JSON.stringify({ round: 3, requested_at: "2026-01-01T00:00:00Z", substantive_attempts: 2,
      items: [finding(directory, "d1", "Reassess the supplied delivery timing.")] }))
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    return { directory, model }
  }

  test("a memo-only round preserves completed upstream stages and draft, but repairs that memo", async () => {
    const { directory, model } = await completedRepair()
    const before = readLeanState(directory)!
    const outputs = Object.fromEntries(Object.entries(before.stages).filter(([step]) => step.startsWith("d1.") || step === "d2.draft")
      .flatMap(([, stage]) => (stage?.outputs ?? []).map(file => [file, readFileSync(leanFile(directory, file))])))
    const memoPlan = stageInputSha(directory, "d2.change-plan")
    const memoDraft = stageInputSha(directory, "d2.draft")
    startRepairRound(directory, [finding(directory, "d2", "Reassess the covering explanation.")], true)
    expect(readRepairRound(directory)).toMatchObject({ round: 4, substantive_attempts: 3, items: [expect.objectContaining({ ref: expect.objectContaining({ document: "d2" }) })] })
    expect(documentRepairs(directory, "d1")).toMatchObject({ round: 3 })
    for (const [step, row] of Object.entries(before.stages)) {
      if (step.startsWith("d1.")) expect(stageInputSha(directory, step as "d1.draft")).toBe(row!.input_sha!)
    }
    expect(stageInputSha(directory, "d2.draft")).toBe(memoDraft)
    expect(stageInputSha(directory, "d2.change-plan")).not.toBe(memoPlan)
    const calls = model.prompts.length
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    const after = readLeanState(directory)!
    for (const [step, row] of Object.entries(before.stages)) {
      if (step.startsWith("d1.") || step === "d2.draft") expect(after.stages[step as "d1.draft"]).toEqual(row)
    }
    for (const [file, bytes] of Object.entries(outputs)) expect(readFileSync(leanFile(directory, file))).toEqual(bytes)
    expect(model.prompts.slice(calls).some(row => row.agent === AGENT.draft)).toBe(false)
    expect(after.stages["d2.change-plan"]?.input_sha).not.toBe(before.stages["d2.change-plan"]?.input_sha)

    // A real subsequent upstream correction replaces, rather than unions, its
    // retired input and propagates its changed final text to the memo's change plan.
    const upstreamKey = stageInputSha(directory, "d1.change-plan")
    const memoKey = stageInputSha(directory, "d2.change-plan")
    startRepairRound(directory, [finding(directory, "d1", "Add the supplied payment timing to delivery.")], true)
    expect(documentRepairs(directory, "d1")?.items.map(row => row.detail)).toEqual(["Add the supplied payment timing to delivery."])
    expect(readRepairRound(directory)?.document_contexts?.d1).toBeUndefined()
    expect(stageInputSha(directory, "d1.change-plan")).not.toBe(upstreamKey)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    // The changed final text reaches the memo through its change plan; its accepted draft is kept.
    expect(stageInputSha(directory, "d2.change-plan")).not.toBe(memoKey)
    expect(stageInputSha(directory, "d2.draft")).toBe(memoDraft)
  }, 120_000)

  test("retained repair conditions expire on changed source or answered instructions", async () => {
    for (const change of ["source", "answer", "draft"]) {
      const { directory } = await completedRepair()
      startRepairRound(directory, [finding(directory, "d2", "Reassess advice.")])
      expect(documentRepairs(directory, "d1")?.round).toBe(3)
      if (change === "source") {
        const db = new Database(path.join(directory, ".dochaus", "legal.db"))
        try { db.run("UPDATE chunks SET text = text || ' Revised timing applies.'") } finally { db.close() }
      }
      if (change === "answer") writeQuestions(directory, [{ id: "timing-answer", question: "Which timing applies?", context: "Timing instruction", status: "answered", answer: "Use the newly supplied terms.", stage: "change-plan", document: "d1" }])
      if (change === "draft") writeLeanText(directory, "d1/paragraphs.json", JSON.stringify(readParagraphs(directory, "d1").map((row, index) => ({ ...row, pid: `p${String(index).padStart(12, "0")}` }))))
      expect(documentRepairs(directory, "d1")).toBeUndefined()
      expect(stageInputSha(directory, "d1.change-plan")).not.toBe(readLeanState(directory)!.stages["d1.change-plan"]!.input_sha!)
    }
  }, 120_000)

  test("addressing a document retires its old body repair even when the new repair is research-only", async () => {
    const { directory } = await completedRepair()
    startRepairRound(directory, [finding(directory, "d2", "Reassess advice.")])
    expect(documentRepairs(directory, "d1")?.round).toBe(3)
    startRepairRound(directory, [finding(directory, "d1", "Research the new proposition.", "research")])
    expect(documentRepairs(directory, "d1")).toBeUndefined()
    expect(readRepairRound(directory)?.document_contexts?.d1).toBeUndefined()
  }, 120_000)

  test("an already stale former admission is not retained under a new basis", async () => {
    const { directory } = await completedRepair()
    writeLeanText(directory, "updated-notes.md", "A new instruction changes the earlier assessment.")
    startRepairRound(directory, [finding(directory, "d2", "Reassess advice.")])
    expect(documentRepairs(directory, "d1")).toBeUndefined()
    expect(readRepairRound(directory)?.document_contexts?.d1).toBeUndefined()
  }, 120_000)

  test("run-wide notes repair does not retire the first document's completed body repair", async () => {
    const { directory } = await completedRepair()
    const key = stageInputSha(directory, "d1.change-plan")
    startRepairRound(directory, [openItem("package-review", "omission", "Clarify the run-wide drafting notes.", { repair_target: "memo" })])
    expect(documentRepairs(directory, "d1")?.round).toBe(3)
    expect(stageInputSha(directory, "d1.change-plan")).toBe(key)
  }, 120_000)
})

describe("semantic repair", () => {
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
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.finalize && prompt.text.includes("<section_catalog>"))).toHaveLength(1)
    expect(model.prompts.filter(prompt => prompt.agent === AGENT.finalize && prompt.text.includes("<required_repairs>"))).toHaveLength(1)
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

  test("repeated unresolved semantic correction is disclosed with a review draft once the repair allowance is used up, never passed as ready", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath,
      [AGENT["package-review"]]: () => JSON.stringify({ status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Memo contradicts the final body." }], limitations: [] }),
    })
    const state = await runLean(directory, { sessions: model.port })
    expect(state.status).toBe("completed")
    expect(model.count(AGENT["package-review"])).toBe(LEAN_MAX_SUBSTANTIVE_REPAIRS + 1)
    const gate = finalGate(directory)
    expect(gate).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(gate.disclosed.filter(row => row.stage === "package-review")).toEqual([expect.objectContaining({ detail: "Memo contradicts the final body." })])
    expect(readLeanText(directory, "delivered-notes.md")).toContain("- Memo contradicts the final body.")
    // With repair allowance left the same finding is open: it starts a repair round.
    writeLeanText(directory, "repairs.json", `${JSON.stringify({ ...readRepairRound(directory), substantive_attempts: LEAN_MAX_SUBSTANTIVE_REPAIRS - 1 }, null, 2)}\n`)
    expect(finalGate(directory).status).toBe("blocked")
  })

  test("with the repair allowance used up, a package review still asking for changes is not run again on unchanged inputs", async () => {
    const directory = matter()
    const model = scriptedModel({ ...happyPath,
      [AGENT["package-review"]]: () => JSON.stringify({ status: "changes-required", delivery: [], findings: [{ target: "memo", detail: "Memo contradicts the final body." }], limitations: [] }),
    })
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    const reviews = model.count(AGENT["package-review"])
    const disclosed = finalGate(directory).disclosed.filter(row => row.stage === "package-review").map(row => row.id)
    expect(disclosed).toHaveLength(1)
    expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
    expect(model.count(AGENT["package-review"])).toBe(reviews)
    expect(finalGate(directory)).toMatchObject({ status: "pass", delivery: "review_draft", blocking: [] })
    expect(finalGate(directory).disclosed.filter(row => row.stage === "package-review").map(row => row.id)).toEqual(disclosed)
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
})
