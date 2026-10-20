import { afterEach, beforeEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readFirmDraftingPreferences } from "../../lib/drafting-preferences"
import { matterPosture } from "../../lib/matter-client"
import { leanSteps, readLeanPlan, runPosture, stepStage, writeLeanPlan } from "./plan"
import { stagePrompt } from "./prompts"
import { runLean, stageInputSha } from "./runner"
import { stageInputs } from "./stages"
import { isReviewStage } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

// How far the run's drafts lean toward the party the lawyer acts for: the
// matter's Posture when it records one, else the firm's, frozen into the plan by
// the outline. Real coordinator; only the model is scripted.

const previous = process.env.WORKSPACE_ROOT
let workspace: string
beforeEach(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "lean-posture-"))
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
// The ingest service writes matter.json; a synthetic one stands in for it.
const recordMatter = (directory: string, fields: Record<string, unknown>) =>
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "matter-fixture", name: "Supply note", ...fields }))
const rendered = (directory: string, step: Parameters<typeof stageInputs>[1]) =>
  stageInputs(directory, step).find(part => part.name === "drafting-preferences")?.bytes

test("before the outline, the run takes the matter's Posture when it records one and the firm's otherwise", () => {
  const directory = matter()
  expect(runPosture(directory)).toBe("balanced")
  save({ posture: "client-favorable" })
  expect(runPosture(directory)).toBe("client-favorable")
  for (const posture of [undefined, "", "aggressive", 42, null]) {
    recordMatter(directory, { posture })
    expect(matterPosture(directory)).toBeUndefined()
    expect(runPosture(directory)).toBe("client-favorable")
  }
  // A matter.json that cannot be read records no Posture: the firm's applies.
  for (const text of ["{", "null", "[]", '{"posture":']) {
    writeFileSync(path.join(directory, "matter.json"), text)
    expect(matterPosture(directory)).toBeUndefined()
    expect(runPosture(directory)).toBe("client-favorable")
  }
  recordMatter(directory, { posture: "conservative" })
  expect(matterPosture(directory)).toBe("conservative")
  expect(runPosture(directory)).toBe("conservative")
  // The stages read the firm's preferences with the matter's Posture in place of the firm's.
  expect(rendered(directory, "intake")).toBe(readFirmDraftingPreferences("full", "conservative"))
  expect(rendered(directory, "intake")).toContain("Draft conservatively")
  expect(rendered(directory, "intake")).not.toContain("Favor the represented person's")
})

test("the outline freezes the run's Posture into the plan, and a settings or matter change mid-run does not alter it", async () => {
  const directory = matter()
  save({ posture: "client-favorable" })
  recordMatter(directory, { posture: "conservative" })
  const model = scriptedModel(happyPath)
  let changed = false
  let frozen: unknown
  const sessions = { ...model.port, async prompt(input: Parameters<typeof model.port.prompt>[0]) {
    if (input.agent === AGENT.review && !changed) {
      changed = true
      frozen = readLeanPlan(directory)?.posture
      save({ posture: "balanced" })
      recordMatter(directory, { posture: "client-favorable" })
    }
    return model.port.prompt(input)
  } }
  const state = await runLean(directory, { sessions })
  expect(changed).toBe(true)
  expect(frozen).toBe("conservative")
  expect(state.status).toBe("completed")
  expect(readLeanPlan(directory)?.posture).toBe("conservative")
  expect(runPosture(directory)).toBe("conservative")
  // No stage was revisited: the change reached neither a running stage nor any key.
  for (const agent of [AGENT.discrepancy, AGENT.review, AGENT["review-independent"], AGENT["change-plan"], AGENT.finalize]) expect(model.count(agent)).toBe(1)
  for (const step of leanSteps(directory)) expect(state.stages[step]?.input_sha).toBe(stageInputSha(directory, step))
  expect(rendered(directory, "d1.review")).toBe(readFirmDraftingPreferences("full", "conservative"))
}, 120_000)

test("a completed run on a matter without a Posture keeps the firm's Posture it was drafted under", async () => {
  const directory = matter()
  const first = await runLean(directory, { sessions: scriptedModel(happyPath).port })
  expect(first.status).toBe("completed")
  expect(readLeanPlan(directory)?.posture).toBe("balanced")
  const keys = new Map(leanSteps(directory).map(step => [step, stageInputSha(directory, step)]))
  save({ posture: "client-favorable" })
  recordMatter(directory, { posture: "conservative" })
  for (const step of leanSteps(directory)) expect(stageInputSha(directory, step)).toBe(keys.get(step)!)
  const idle = scriptedModel({})
  expect((await runLean(directory, { sessions: idle.port })).status).toBe("completed")
  expect(idle.sessions).toHaveLength(0)
}, 120_000)

test("a matter without a Posture keeps its stage keys, and a plan saved without one follows the matter and the firm as before", () => {
  const directory = matter()
  // A plan saved before the Posture was frozen into it.
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }])
  const keys = () => leanSteps(directory).map(step => stageInputSha(directory, step))
  const before = keys()
  for (const step of leanSteps(directory)) {
    const bytes = rendered(directory, step)
    if (bytes !== undefined) expect(bytes).toBe(readFirmDraftingPreferences())
  }
  // No Posture, an unreadable one, or the firm's own: the same rendering and keys.
  for (const fields of [{}, { posture: "" }, { posture: 42 }, { posture: "aggressive" }, { posture: "balanced" }]) {
    recordMatter(directory, fields)
    expect(keys()).toEqual(before)
  }
  // A Posture of its own reaches every step but the accepted draft.
  recordMatter(directory, { posture: "conservative" })
  expect(leanSteps(directory).filter((step, index) => keys()[index] === before[index])).toEqual(leanSteps(directory).filter(step => step.endsWith(".draft")))
})

test("every review takes the frozen Posture's standard for the side the lawyer acts for; with no side stated it reviews even-handedly", () => {
  const directory = matter()
  const documents = [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }]
  const reviews = () => leanSteps(directory).filter(step => isReviewStage(stepStage(step)))
  const PROTECTS = "whether it protects that party's position, and whether a protection a careful drafter would include for that party is missing."
  const MARKET = "whether it lacks what a conventional, market-standard document for that side would include. Do not press for protections beyond market practice"
  const NEUTRAL = "Assess each provision as a reviewing lawyer who acts for neither party"
  // The live settings say Client-favorable; the plan's frozen Posture governs.
  save({ posture: "client-favorable" })
  recordMatter(directory, { posture: "client-favorable" })
  for (const posture of ["client-favorable", "balanced", "conservative"] as const) {
    writeLeanPlan(directory, documents, undefined, { party: "Northwind Traders Ltd", stated: true }, undefined, posture)
    expect(reviews().length).toBeGreaterThan(1)
    for (const step of reviews()) {
      const prompt = stagePrompt(directory, step)
      expect(prompt).toContain("Assess each provision as the reviewing lawyer for the party the lawyer acts for")
      if (posture === "client-favorable") {
        expect(prompt).toContain(PROTECTS)
        expect(prompt).not.toContain("market practice")
      } else {
        expect(prompt).toContain(MARKET)
        expect(prompt).not.toContain("whether it protects that party's position")
      }
    }
    // No side stated: the same neutral review under every Posture.
    writeLeanPlan(directory, documents, undefined, undefined, undefined, posture)
    for (const step of reviews()) {
      const prompt = stagePrompt(directory, step)
      expect(prompt).toContain(NEUTRAL)
      expect(prompt).not.toContain("market practice")
      expect(prompt).not.toContain("whether it protects that party's position")
    }
  }
})

test("the review agents defer to the prompt's standard for the party rather than setting their own", () => {
  for (const name of ["lean-review", "lean-review-independent", "lean-review-client-disadvantage"]) {
    const text = readFileSync(path.join(import.meta.dir, "..", "..", "agent", `${name}.md`), "utf8")
    expect(text).toContain("by the standard")
    expect(text).not.toContain("whether it protects that party's position")
    expect(text).not.toContain("a protection a careful drafter would include for that party")
  }
})
