import { afterEach, beforeEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { DEFAULT_DRAFTING } from "../../lib/drafting-preferences"
import { readLeanEvents } from "./events"
import { leanSteps, writeLeanPlan } from "./plan"
import { runLean, stageInputSha } from "./runner"
import { readLeanJson, readLeanState } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

const previous = process.env.WORKSPACE_ROOT
let workspace: string
beforeEach(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "lean-drafting-preferences-"))
  process.env.WORKSPACE_ROOT = workspace
})
afterEach(() => {
  cleanupMatters()
  rmSync(workspace, { recursive: true, force: true })
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
})

function save(input: Record<string, unknown>) {
  mkdirSync(path.join(workspace, ".preferences"), { recursive: true })
  writeFileSync(path.join(workspace, ".preferences", "preferences.json"), JSON.stringify(input))
}

test("unchanged, equivalent defaults and operational-only saves skip accepted stages", async () => {
  const directory = matter()
  const first = await runLean(directory, { sessions: scriptedModel(happyPath).port })
  expect(first.status).toBe("completed")
  for (const preferences of [undefined, { ...DEFAULT_DRAFTING, houseStyle: " \n " }, {
    process: "plan-first", searchApiKey: "synthetic-search", caseLawApiKey: "synthetic-case-law",
    telemetry: { retention_days: 21, rate_overrides: { synthetic: { input_per_mtok: 5 } } },
  }]) {
    if (preferences) save(preferences)
    const idle = scriptedModel({})
    const resumed = await runLean(directory, { sessions: idle.port })
    expect(resumed.status).toBe("completed")
    expect(idle.sessions).toHaveLength(0)
    expect(resumed.stages["d1.apply"]?.input_sha).toBe(first.stages["d1.apply"]?.input_sha)
  }
}, 120_000)

test("house style and posture changes revisit completed model work but the accepted draft, and code outputs", async () => {
  const directory = matter()
  let prior = await runLean(directory, { sessions: scriptedModel(happyPath).port })
  expect(prior.status).toBe("completed")
  for (const preferences of [{ houseStyle: "Use active voice." }, { houseStyle: "Use active voice.", posture: "client-favorable" }]) {
    const keys = new Map(leanSteps(directory).map(step => [step, stageInputSha(directory, step)]))
    save(preferences)
    // The accepted draft is keyed on the outline contract alone; the change reaches it through its review.
    for (const step of leanSteps(directory)) {
      if (step.endsWith(".draft")) expect(stageInputSha(directory, step)).toBe(keys.get(step)!)
      else expect(stageInputSha(directory, step)).not.toBe(keys.get(step)!)
    }
    const revised = scriptedModel(happyPath)
    const state = await runLean(directory, { sessions: revised.port })
    expect(state.status).toBe("completed")
    for (const agent of [AGENT.discrepancy, AGENT.intake, AGENT.research, AGENT["research-review"], AGENT.spine,
      AGENT.review, AGENT["review-independent"], AGENT["review-wording"], AGENT["review-client-disadvantage"],
      AGENT["change-plan"], AGENT.changes, AGENT["change-review"], AGENT.finalize, AGENT["package-review"]])
      expect(revised.prompts.map(row => row.agent)).toContain(agent)
    expect(revised.count(AGENT.draft)).toBe(0)
    expect(state.stages["d1.apply"]?.input_sha).not.toBe(prior.stages["d1.apply"]?.input_sha)
    expect(state.stages.finalize?.input_sha).not.toBe(prior.stages.finalize?.input_sha)
    prior = state
  }
}, 120_000)

test("a save during a sectioned draft rejects that attempt before accepting mixed work", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  let changed = false
  let rejected = false
  const sessions = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT.draft && /^Draft part 1/m.test(input.text) && !changed) {
      changed = true
      save({ houseStyle: "Use active voice." })
    }
    if (input.agent === AGENT.discrepancy && changed) {
      rejected = true
      expect(readLeanState(directory)?.stages["d1.draft"]?.status).toBe("pending")
      expect(readLeanEvents(directory).filter(event => event.stage === "draft" && event.event === "done")).toHaveLength(0)
    }
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  expect(changed).toBe(true)
  expect(rejected).toBe(true)
  expect(model.count(AGENT.discrepancy)).toBe(2)
  expect(readLeanEvents(directory).filter(event => event.stage === "draft" && event.event === "done")).toHaveLength(1)
  for (const step of leanSteps(directory)) expect(state.stages[step]?.input_sha).toBe(stageInputSha(directory, step))
}, 120_000)

test("a save during parallel review discards prefetched old replies before adjudication", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  let changed = false
  const reviewers: string[] = [AGENT.review, AGENT["review-independent"], AGENT["review-wording"], AGENT["review-client-disadvantage"]]
  const sessions = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT.review && !changed) {
      changed = true
      save({ posture: "conservative" })
    }
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  for (const agent of reviewers) expect(model.count(agent)).toBe(2)
  expect(model.count(AGENT["change-plan"])).toBe(1)
  for (const step of leanSteps(directory)) expect(state.stages[step]?.input_sha).toBe(stageInputSha(directory, step))
}, 120_000)

test("an observed A to B to A change aborts a sectioned attempt before its next part and recovers without a loop", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  let switched = false
  let restored = false
  let firstDraftSession: string | undefined
  const staleSecondParts: string[] = []
  const sessions = { ...model.port,
    async createSession(input: Parameters<typeof model.port.createSession>[0]) {
      // Restore A during replay admission. The invalidated attempt must have
      // stopped before its next section, without requiring a settings ledger.
      if (switched && !restored && input.agent === AGENT.discrepancy) {
        restored = true
        save({})
      }
      return model.port.createSession(input)
    },
    async prompt(input: Parameters<typeof model.port.prompt>[0]) {
      if (input.agent === AGENT.draft && !firstDraftSession) firstDraftSession = input.sessionID
      if (input.agent === AGENT.draft && /^Draft part 1/m.test(input.text) && !switched) {
        switched = true
        save({ houseStyle: "Use active voice." })
      }
      if (input.sessionID === firstDraftSession && /^Draft part 2/m.test(input.text)) staleSecondParts.push(input.text)
      return model.port.prompt(input)
    },
  }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  expect(switched && restored).toBe(true)
  expect(staleSecondParts).toEqual([])
  expect(model.count(AGENT.discrepancy)).toBe(2)
  expect(readLeanEvents(directory).filter(event => event.stage === "draft" && event.event === "done")).toHaveLength(1)
  for (const step of leanSteps(directory)) expect(state.stages[step]?.input_sha).toBe(stageInputSha(directory, step))
}, 120_000)

test("parallel review retains an observed preference change after A is restored before outer acceptance", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  let switched = false
  let restored = false
  let release: () => void = () => {}
  const independent = new Promise<void>(resolve => { release = resolve })
  const sessions = { ...model.port,
    async createSession(input: Parameters<typeof model.port.createSession>[0]) {
      if (switched && !restored && input.agent === AGENT["review-client-disadvantage"]) {
        // Allow the running wording exchange to observe B, then restore A
        // before the outer stage accepts anything or drains prior replies.
        await new Promise(resolve => setTimeout(resolve, 0))
        save({})
        restored = true
        release()
      }
      return model.port.createSession(input)
    },
    async prompt(input: Parameters<typeof model.port.prompt>[0]) {
      if (input.agent === AGENT["review-wording"] && !switched) {
        switched = true
        save({ posture: "conservative" })
      }
      const reply = await model.port.prompt(input)
      if (input.agent === AGENT["review-independent"] && !restored) await independent
      return reply
    },
  }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  expect(switched && restored).toBe(true)
  expect(model.count(AGENT["review-wording"])).toBe(2)
  expect(model.count(AGENT["review-independent"])).toBe(2)
  expect(model.count(AGENT["change-plan"])).toBe(1)
  expect(readLeanEvents(directory).filter(event => event.stage === "review-wording" && event.event === "done")).toHaveLength(1)
  for (const step of leanSteps(directory)) expect(state.stages[step]?.input_sha).toBe(stageInputSha(directory, step))
}, 120_000)


test("helper-swallowed exchange invalidation restarts the walk without charging a substantive repair", async () => {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }, { heading: "Payment", scope: "Payment" }] }])
  const model = scriptedModel(happyPath)
  let switched = false
  let restored = false
  const sessions = { ...model.port,
    async createSession(input: Parameters<typeof model.port.createSession>[0]) {
      if (switched && !restored && input.agent === AGENT.discrepancy) {
        save({})
        restored = true
      }
      return model.port.createSession(input)
    },
    async prompt(input: Parameters<typeof model.port.prompt>[0]) {
      if (input.agent === "lean-checklist-extract" && !switched) {
        switched = true
        save({ houseStyle: "Use active voice." })
      }
      return model.port.prompt(input)
    },
  }
  const state = await runLean(directory, { sessions })
  expect(state.status).toBe("completed")
  expect(switched && restored).toBe(true)
  expect(model.count("lean-checklist-extract")).toBeGreaterThan(1)
  expect(model.count("lean-checklist-extract")).toBeLessThan(8)
  expect(model.count(AGENT.discrepancy)).toBe(2)
  expect(readLeanJson(directory, "repairs.json")).toBeUndefined()
  for (const step of leanSteps(directory)) expect(state.stages[step]?.input_sha).toBe(stageInputSha(directory, step))
}, 120_000)
