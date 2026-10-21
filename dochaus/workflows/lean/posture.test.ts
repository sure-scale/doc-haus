import { afterEach, beforeEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readFirmDraftingPreferences } from "../../lib/drafting-preferences"
import { matterPosture } from "../../lib/matter-client"
import { acceptChangePlan, changePlanPrompt, deferredFindingItems } from "./change-pipeline"
import { packageReviewPrompt } from "./package-review"
import { leanSteps, readLeanPlan, runPosture, stepStage, writeLeanPlan, type ActingFor } from "./plan"
import { stagePrompt } from "./prompts"
import { issueRegister } from "./register"
import { findingsFor, routedFindingItems } from "./review-findings"
import { runLean, stageInputSha } from "./runner"
import { acceptStageReply, stageInputs } from "./stages"
import { isReviewStage, writeLeanText } from "./state"
import { AGENT, SOURCE, SUPPORT, cleanupMatters, happyPath, matter, pidOf, scriptedModel } from "./test-support"

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

test("every review takes the frozen Posture's standard for the side the lawyer acts for, and measures balance against neutral market practice whoever it acts for; with no side stated it reviews even-handedly", () => {
  const directory = matter()
  const documents = [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }] }]
  const reviews = () => leanSteps(directory).filter(step => isReviewStage(stepStage(step)))
  const PROTECTS = "whether it protects that party's position, and whether a protection a careful drafter would include for that party is missing."
  const MARKET = "whether it lacks a protection for that party that a neutral, mid-market document of this kind, prepared for neither party, would contain."
  // The balance label's reference is neutral market practice under every Posture and side,
  // never a market document for the client's side (itself a push under Balanced).
  const REFERENCE = "measured against what a neutral, mid-market document of this kind, prepared for neither party, would contain, whichever party the lawyer acts for"
  const SIDED_FRAME = ["for that side", "market-standard document"]
  const NEUTRAL = "Assess each provision as a reviewing lawyer who acts for neither party"
  const BALANCE = '"balance":"neutral|market|beyond_market"'
  // The live settings say Client-favorable; the plan's frozen Posture governs.
  save({ posture: "client-favorable" })
  recordMatter(directory, { posture: "client-favorable" })
  for (const posture of ["client-favorable", "balanced", "conservative"] as const) {
    writeLeanPlan(directory, documents, undefined, { party: "Northwind Traders Ltd", stated: true }, undefined, posture)
    expect(reviews().length).toBeGreaterThan(1)
    for (const step of reviews()) {
      const prompt = stagePrompt(directory, step)
      expect(prompt).toContain("Assess each provision as the reviewing lawyer for the party the lawyer acts for")
      // Every finding states its balance; code, not a plea, holds back what goes beyond market.
      expect(prompt).toContain(BALANCE)
      expect(prompt).toContain(REFERENCE)
      for (const frame of SIDED_FRAME) expect(prompt).not.toContain(frame)
      expect(prompt).not.toContain("Do not press")
      if (posture === "client-favorable") {
        expect(prompt).toContain(PROTECTS)
        expect(prompt).not.toContain(MARKET)
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
      expect(prompt).toContain(BALANCE)
      expect(prompt).toContain(REFERENCE)
      for (const frame of SIDED_FRAME) expect(prompt).not.toContain(frame)
      expect(prompt).not.toContain(MARKET)
      expect(prompt).not.toContain("whether it protects that party's position")
    }
  }
})

// A review finding that gives one party more than market practice is held back
// unless the run acts for a side under Client-favorable: it reaches no change
// plan and is disclosed as a recommendation naming the side it favours.
const NORTHWIND: ActingFor = { party: "Northwind Traders Ltd", stated: true }
const PID = "p111111111111"
const BEYOND = "Delivery gives the buyer no right to cancel the whole contract for any late delivery."
const MARKET_GAP = "Delivery gives the buyer no remedy for late delivery."
const INSTRUCTED = "The draft leaves out the delivery period the request asks for."
const GUIDANCE = "The firm's guidance asks for a buyer's right to cancel for late delivery."
const raised = (finding: string, origin: string, balance: string) =>
  ({ pids: [PID], origin, balance, ...(balance === "neutral" ? {} : { favours: "Northwind Traders Ltd" }), finding, sources: [] })

async function reviewed(actingFor: ActingFor | undefined, posture: "client-favorable" | "balanced" | "conservative") {
  const directory = matter()
  writeLeanPlan(directory, [{ title: "Supply note", purpose: "Record supplied terms", parts: [{ heading: "Delivery", scope: "Delivery" }],
    playbook_review: { sources: [{ source_id: SOURCE, quote: SUPPORT[0]!.quote }], reason: "The firm's supply guidance applies." } }], undefined, actingFor, undefined, posture)
  writeLeanText(directory, "d1/paragraphs.json", JSON.stringify([{ pid: PID, sha256: "initial", text: "The supplier shall deliver the goods." }]))
  const replies = {
    "d1.review-client-disadvantage": { findings: [raised(BEYOND, "missing_protection", "beyond_market"), raised(MARKET_GAP, "missing_protection", "market")] },
    // An instruction and the firm's chosen guidance are never held back, whatever their balance.
    "d1.review": { findings: [raised(INSTRUCTED, "request", "beyond_market")] },
    "d1.review-playbook": { coverage: "- Late delivery: correction required.", findings: [raised(GUIDANCE, "drafting", "beyond_market")] },
  } as const
  for (const [step, reply] of Object.entries(replies)) expect((await acceptStageReply(directory, step as keyof typeof replies, JSON.stringify(reply))).issues).toEqual([])
  return directory
}

test("a beyond-market finding is held back unless the run acts for a side under Client-favorable, and reaches the notes as a recommendation naming that side", async () => {
  const cases = [[NORTHWIND, "client-favorable", ""], [NORTHWIND, "balanced", "the run's Posture is Balanced"], [NORTHWIND, "conservative", "the run's Posture is Conservative"],
    [undefined, "client-favorable", "the run acts for neither party"], [undefined, "balanced", "the run acts for neither party"]] as const
  for (const [actingFor, posture, because] of cases) {
    const directory = await reviewed(actingFor, posture)
    const held = Boolean(because)
    const planned = findingsFor(directory, "d1").map(row => row.finding)
    expect(planned).toEqual(expect.arrayContaining([MARKET_GAP, INSTRUCTED, GUIDANCE]))
    expect(planned.includes(BEYOND)).toBe(!held)
    expect(changePlanPrompt(directory, "d1").includes(BEYOND)).toBe(!held)
    // Held back is a recommendation, never an open decision.
    expect(routedFindingItems(directory)).toEqual([])
    const adjudicated = { shared_decisions: [], groups: [], dismissed: [], questions: [], findings: [{ disposition: "rejected", kind: "suggestion",
      detail: "The fixture needs no change.", reason: "Scripted adjudication.", documents: ["d1"], sources: [], limitations: ["Fixture."], resolves: findingsFor(directory, "d1").map(row => row.id) }] }
    expect(acceptChangePlan(directory, "d1", JSON.stringify(adjudicated)).issues).toEqual([])
    expect(deferredFindingItems(directory)).toEqual([])
    // A re-plan is not handed it back as a prior disposition.
    expect(changePlanPrompt(directory, "d1").includes(BEYOND)).toBe(!held)
    const disclosure = `Not drafted: it favours Northwind Traders Ltd beyond market practice, and ${because}.`
    for (const prompt of [stagePrompt(directory, "finalize"), packageReviewPrompt(directory)]) {
      expect(prompt.includes(BEYOND)).toBe(held)
      expect(prompt.includes(disclosure)).toBe(held)
    }
    expect(issueRegister(directory).entries.find(row => row.finding === BEYOND)).toMatchObject(held ? { disposition: "held_back", reason: disclosure } : { disposition: "rejected" })
  }
}, 120_000)

test("on a matter whose client is recorded under a Balanced Posture, the coordinator drafts no beyond-market finding and the notes disclose it", async () => {
  const directory = matter()
  recordMatter(directory, { client: "Northwind Traders Ltd", posture: "balanced" })
  const model = scriptedModel({ ...happyPath,
    [AGENT["review-client-disadvantage"]]: prompt => JSON.stringify({ findings: [{ ...raised(BEYOND, "missing_protection", "beyond_market"), pids: [pidOf(prompt, "The supplier shall deliver the goods.")] }] }) })
  const state = await runLean(directory, { sessions: model.port })
  expect(state.status).toBe("completed")
  expect(readLeanPlan(directory)).toMatchObject({ acting_for: NORTHWIND, posture: "balanced" })
  const shown = (agent: string) => model.prompts.filter(row => row.agent === agent).map(row => row.text).join("\n")
  expect(shown(AGENT["review-client-disadvantage"])).toContain("beyond_market")
  for (const agent of [AGENT["change-plan"], AGENT.changes]) expect(shown(agent)).not.toContain(BEYOND)
  expect(shown(AGENT.finalize)).toContain(BEYOND)
  expect(shown(AGENT.finalize)).toContain("Not drafted: it favours Northwind Traders Ltd beyond market practice, and the run's Posture is Balanced.")
  expect(readFileSync(path.join(directory, ".dochaus", "lean", "register.json"), "utf8")).toContain('"disposition": "held_back"')
}, 120_000)

test("the review agents defer to the prompt's standard for the party rather than setting their own", () => {
  for (const name of ["lean-review", "lean-review-independent", "lean-review-client-disadvantage"]) {
    const text = readFileSync(path.join(import.meta.dir, "..", "..", "agent", `${name}.md`), "utf8")
    expect(text).toContain("by the standard")
    expect(text).not.toContain("whether it protects that party's position")
    expect(text).not.toContain("a protection a careful drafter would include for that party")
  }
})
