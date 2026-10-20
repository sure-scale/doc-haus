import { afterEach, expect, test } from "bun:test"
import { existsSync, rmSync } from "node:fs"
import { readLeanEvents } from "./events"
import { stagePrompt } from "./prompts"
import { enqueueLeanRequest } from "./requests"
import { runLean, stageInputSha } from "./runner"
import { LEAN_CHILD_PERMISSIONS, stepOutputs } from "./stages"
import { REVIEW_STAGES, currentOpenItems, leanFile, readLeanJson, readLeanState, readOpenItems, writeLeanText } from "./state"
import { AGENT, NO_FINDINGS, SOURCE, cleanupMatters, happyPath, matter, noChanges, scriptedModel } from "./test-support"

afterEach(cleanupMatters)
const reviewers: readonly string[] = REVIEW_STAGES.map(stage => AGENT[stage])

/** A review reply with one finding in these words. */
const finding = (text: string, sources: { source_id: string; quote: string }[] = []) =>
  JSON.stringify({ findings: [{ origin: "drafting", finding: text, sources }] })

/** The review findings a change plan prompt lists, as [review stage, finding]. */
function listedFindings(prompt: string) {
  const block = /<review_findings>\n([\s\S]*?)\n<\/review_findings>/.exec(prompt)?.[1] ?? "[]"
  return (JSON.parse(block) as { review: string; finding: string }[]).map(row => [row.review, row.finding])
}

function barrier(count: number) {
  let release: (() => void) | undefined
  const ready = new Promise<void>(resolve => { release = resolve })
  const started = new Set<string>()
  return { started, async arrive(agent: string) {
    started.add(agent)
    if (started.size === count) release!()
    await ready
  } }
}

test("all four independent reviews run together and fan into adjudication before global reconciliation", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  const calls = barrier(reviewers.length)
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (reviewers.includes(input.agent)) {
      await calls.arrive(input.agent)
      REVIEW_STAGES.forEach(stage => expect(readLeanState(directory)?.stages[`d1.${stage}`]?.status).toBe("running"))
      return finding(`Independent finding from ${input.agent}.`)
    }
    if (input.agent === AGENT["change-plan"]) {
      REVIEW_STAGES.forEach(stage => expect(readLeanState(directory)?.stages[`d1.${stage}`]?.status).toBe("done"))
      expect(listedFindings(input.text)).toEqual(REVIEW_STAGES.map(stage => [stage, `Independent finding from ${AGENT[stage]}.`]))
      expect(input.text).toContain("Consolidate duplicates")
      expect(input.text).toContain("Reviewer agreement and exact quotation matching alone do not establish")
    }
    if (input.agent === AGENT.changes) {
      expect(readLeanState(directory)?.stages["d1.change-plan"]?.status).toBe("done")
      expect(readLeanState(directory)?.stages["d1.change-prepare"]?.status).toBe("done")
      expect(input.text).toContain("<change_plan>")
      expect(input.text).toContain("<prepared_changes>")
      expect(input.text).not.toContain("Independent finding from")
      expect(input.text).toContain("Prepared amendments are proposals, not authority.")
    }
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions: port })
  expect(state.status).toBe("completed")
  expect(calls.started.size).toBe(4)
  expect(model.count(AGENT["change-plan"])).toBe(1)
  expect(model.count(AGENT.changes)).toBe(1)
  expect(model.count(AGENT["change-review"])).toBe(1)
  REVIEW_STAGES.forEach(stage => {
    expect(state.stages[`d1.${stage}`]).toMatchObject({ status: "done", attempts: 1, calls: 1, outputs: [`d1/${stage}.json`] })
    expect(model.sessions.find(session => session.agent === AGENT[stage])).toMatchObject({ parentID: "ses_owner", permissions: LEAN_CHILD_PERMISSIONS })
  })
  expect(readLeanEvents(directory).filter(event => event.event === "done" && reviewers.includes(AGENT[event.stage as typeof REVIEW_STAGES[number]])).map(event => event.stage)).toEqual([...REVIEW_STAGES])
}, 120_000)

test("resume preserves completed reviews; specialist edits rekey consumers and missing specialists run together", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: scriptedModel(happyPath).port })).status).toBe("completed")
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.sessions).toHaveLength(0)
  const key = stageInputSha(directory, "d1.changes")
  writeLeanText(directory, "d1/review-wording.json", `${JSON.stringify({ findings: [{ id: "f-0123456789", document: "d1", origin: "drafting", finding: "Wording rechecked by the lawyer.", sources: [] }] }, null, 2)}\n`)
  expect(stageInputSha(directory, "d1.changes")).not.toBe(key)
  const edited = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: edited.port })).status).toBe("completed")
  expect(edited.count(AGENT.changes)).toBe(1)
  reviewers.forEach(agent => expect(edited.count(agent)).toBe(0))
  rmSync(leanFile(directory, "d1/review-wording.json"))
  rmSync(leanFile(directory, "d1/review-client-disadvantage.json"))
  const resumed = scriptedModel(happyPath)
  const calls = barrier(2)
  const added: readonly string[] = [AGENT["review-wording"], AGENT["review-client-disadvantage"]]
  const port = { ...resumed.port, async prompt(input: Parameters<typeof resumed.port.prompt>[0]) {
    if (added.includes(input.agent)) {
      await calls.arrive(input.agent)
      expect(input.text).not.toContain("Wording rechecked by the lawyer.")
      return NO_FINDINGS
    }
    return resumed.port.prompt(input)
  } }
  expect((await runLean(directory, { sessions: port })).status).toBe("completed")
  expect(calls.started.size).toBe(2)
  expect(resumed.count(AGENT.review)).toBe(0)
  expect(resumed.count(AGENT["review-independent"])).toBe(0)
}, 120_000)

test("a review-boundary update discards prefetched old replies and re-reviews before changes", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  // The update arrives once, while the first reviews are in flight. A retry
  // prompt does not repeat the instructions, so it must not enqueue another.
  const update = { sent: false }
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (reviewers.includes(input.agent)) {
      const current = input.text.includes("Use the updated instructions.")
      if (input.agent === AGENT.review && !current && !update.sent) {
        update.sent = true
        enqueueLeanRequest(directory, { kind: "update", text: "Use the updated instructions." })
      }
      return finding(`${current ? "Current" : "Stale"} independent review from ${input.agent}.`)
    }
    if (input.agent === AGENT["change-plan"] || input.agent === AGENT.changes) {
      expect(input.text).not.toContain("Stale independent review")
      expect(input.text).toContain("Use the updated instructions.")
      if (input.agent === AGENT["change-plan"])
        reviewers.forEach(agent => expect(input.text).toContain(`Current independent review from ${agent}.`))
      else expect(input.text).not.toContain("Current independent review from")
    }
    return model.port.prompt(input)
  } }
  expect((await runLean(directory, { sessions: port })).status).toBe("completed")
  expect(model.count(AGENT.changes)).toBe(1)
  REVIEW_STAGES.forEach(stage => expect((readLeanJson(directory, `d1/${stage}.json`) as { findings: { finding: string }[] }).findings.map(row => row.finding))
    .toEqual([`Current independent review from ${AGENT[stage]}.`]))
  reviewers.forEach(agent => expect(model.sessions.filter(session => session.agent === agent)).toHaveLength(2))
}, 120_000)

test("a failed reviewer settles siblings without saving them or applying; resume reuses accepted reviews", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  const calls = barrier(4)
  const settled: string[] = []
  const port = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (reviewers.includes(input.agent)) {
      await calls.arrive(input.agent)
      settled.push(input.agent)
      if (input.agent === AGENT["review-wording"]) throw new Error("review transport unavailable")
      return NO_FINDINGS
    }
    return model.port.prompt(input)
  } }
  const failed = await runLean(directory, { sessions: port })
  expect(failed.status).toBe("failed")
  expect(failed.stop_reason).toContain("d1.review-wording: stopped by a runtime error")
  expect(settled).toHaveLength(4)
  expect(failed.stages["d1.review"]?.status).toBe("done")
  expect(failed.stages["d1.review-independent"]?.status).toBe("done")
  expect(failed.stages["d1.review-wording"]?.status).toBe("failed")
  expect(failed.stages["d1.review-client-disadvantage"]?.status).toBe("pending")
  expect(existsSync(leanFile(directory, "d1/review-client-disadvantage.json"))).toBe(false)
  expect(model.count(AGENT.changes)).toBe(0)
  const resumed = scriptedModel(happyPath)
  expect((await runLean(directory, { sessions: resumed.port })).status).toBe("completed")
  expect(resumed.count(AGENT.review)).toBe(0)
  expect(resumed.count(AGENT["review-independent"])).toBe(0)
  expect(resumed.count(AGENT["review-wording"])).toBe(1)
  expect(resumed.count(AGENT["review-client-disadvantage"])).toBe(1)
}, 120_000)

test("specialist source misses reach reconciliation and are superseded by the checked body", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath,
    [AGENT["review-wording"]]: () => finding("Unsupported reviewer passage.", [{ source_id: SOURCE, quote: "words absent from the supplied terms" }]),
    [AGENT["review-client-disadvantage"]]: () => finding("Optional negotiating point, preserving the agreed terms."),
    [AGENT.changes]: noChanges,
  })
  expect((await runLean(directory, { sessions: model.port })).status).toBe("completed")
  expect(readOpenItems(directory)).toContainEqual(expect.objectContaining({ stage: "review-wording", kind: "quote_miss", ref: expect.objectContaining({ document: "d1" }) }))
  const prompt = model.prompts.find(row => row.agent === AGENT["change-plan"])!.text
  expect(prompt).toContain("<unverified_findings>")
  expect(prompt).toContain("words absent from the supplied terms")
  expect(currentOpenItems(directory).some(item => item.stage === "review-wording")).toBe(false)
  REVIEW_STAGES.forEach(stage => {
    expect(stepOutputs(`d1.${stage}`)).toEqual([`d1/${stage}.json`])
    const review = stagePrompt(directory, `d1.${stage}`)
    expect(review).not.toContain("Optional negotiating point")
    expect(review).not.toContain("words absent from the supplied terms")
    expect(review).toContain("scope_and_decisions")
  })
}, 120_000)


test("a specialist checker retries only that reviewer in the same session", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath,
    [AGENT["review-wording"]]: (_, index) => index ? NO_FINDINGS : JSON.stringify({ findings: [{ pid: "pffffffffffff", origin: "drafting", finding: "Invalid draft pointer." }] }),
  })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("completed")
  const prompts = model.prompts.filter(row => row.agent === AGENT["review-wording"])
  expect(prompts).toHaveLength(2)
  expect(prompts[1]!.sessionID).toBe(prompts[0]!.sessionID)
  expect(prompts[1]!.text).toContain("findings[0]: [pffffffffffff] is not a paragraph of the draft.")
  expect(state.stages["d1.review-wording"]).toMatchObject({ attempts: 2, calls: 2 })
  reviewers.filter(agent => agent !== AGENT["review-wording"]).forEach(agent => expect(model.sessions.filter(session => session.agent === agent)).toHaveLength(1))
}, 120_000)


test("session creation failure leaves unstarted reviews pending and waits for already started siblings", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  let settled = false
  const port = { ...model.port,
    async createSession(input: Parameters<typeof model.port.createSession>[0]) {
      if (input.agent === AGENT["review-wording"]) throw new Error("session creation unavailable")
      return model.port.createSession(input)
    },
    async prompt(input: Parameters<typeof model.port.prompt>[0]) {
      if (input.agent === AGENT["review-independent"]) {
        await new Promise(resolve => setTimeout(resolve, 1))
        settled = true
        return NO_FINDINGS
      }
      return model.port.prompt(input)
    },
  }
  const state = await runLean(directory, { sessions: port })
  expect(state.status).toBe("failed")
  expect(state.stop_reason).toContain("session creation unavailable")
  expect(settled).toBe(true)
  expect(state.stages["d1.review"]?.status).toBe("pending")
  expect(state.stages["d1.review-independent"]?.status).toBe("pending")
  expect(state.stages["d1.review-wording"]?.status).toBe("failed")
  expect(model.count(AGENT.changes)).toBe(0)
  expect(existsSync(leanFile(directory, "d1/review-independent.json"))).toBe(false)
}, 120_000)
