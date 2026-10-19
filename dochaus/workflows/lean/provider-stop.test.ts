import { afterEach, expect, test } from "bun:test"
import { readdirSync, readFileSync, statSync } from "node:fs"
import path from "node:path"
import { PROVIDER_STOP_MESSAGES } from "../../lib/provider-stop"
import { CHECKLIST_JOB_AGENTS } from "./checklist-job"
import { enqueueLeanRequest } from "./requests"
import { runLean } from "./runner"
import { engineSessionPort } from "./session-port"
import type { StageSessionPort } from "./stages"
import { readLeanJson, readLeanState, writeLeanState } from "./state"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel, type Reply } from "./test-support"

// Real coordinator; the scripted session port fails the way the engine port
// does: the native turn error or the engine's refusal body is the error's cause.

afterEach(cleanupMatters)

const failing = (cause: Record<string, unknown>): Reply => () => {
  throw new Error("The engine could not complete the turn.", { cause })
}

const EXPIRED = { type: "provider.auth", message: "This API key has expired", status: 401 }

test.each([
  ["auth", EXPIRED],
  ["setup", { error: "Finish installation setup before generating.", code: "setup_required" }],
  ["model", { type: "provider.no-route", message: "Model unavailable: test/scripted" }],
] as const)("a %s failure stops the run with its own reason, and Continue resumes once it is fixed", async (kind, cause) => {
  const directory = matter()
  const broken = scriptedModel({ ...happyPath, [AGENT.discrepancy]: failing(cause) })
  const stopped = await runLean(directory, { sessions: broken.port })

  expect(stopped.status).toBe("failed")
  expect(stopped.provider_stop).toBe(kind)
  // The lawyer's fix only, never the provider's own payload.
  expect(stopped.stop_reason).toBe(`discrepancy: ${PROVIDER_STOP_MESSAGES[kind]}`)
  expect(stopped.stages.discrepancy?.status).toBe("failed")
  expect(broken.prompts.map(row => row.agent)).toEqual([AGENT.discrepancy])

  const fixed = scriptedModel(happyPath)
  const resumed = await runLean(directory, { sessions: fixed.port })
  expect(resumed.status).toBe("completed")
  expect(resumed.provider_stop).toBeUndefined()
}, 120_000)

test("a provider failure a helper records as an issue still stops the step as that failure", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, "lean-checklist-extract": failing(EXPIRED) })
  const state = await runLean(directory, { sessions: model.port })

  expect(state.status).toBe("failed")
  expect(state.provider_stop).toBe("auth")
  expect(state.stop_reason).toStartWith("d1.draft: ")
  expect(state.stages["d1.draft"]?.status).toBe("failed")
}, 120_000)

test("any other engine failure stops as a runtime error after the usual checker retry", async () => {
  const directory = matter()
  const transport = failing({ type: "provider.transport", message: "Decode error", status: 200 })
  const model = scriptedModel({ ...happyPath, [AGENT.discrepancy]: (prompt, index, session) => index ? transport(prompt, index, session) : "not a JSON reply" })
  const state = await runLean(directory, { sessions: model.port })

  expect(model.count(AGENT.discrepancy)).toBe(2)
  expect(state.status).toBe("failed")
  expect(state.provider_stop).toBeUndefined()
  expect(state.stop_reason).toBe("discrepancy: stopped by a runtime error: The engine could not complete the turn.")
}, 120_000)

test("a sibling review's provider failure stops the run at that review, not at the one that succeeded", async () => {
  const directory = matter()
  const model = scriptedModel({ ...happyPath, [AGENT["review-independent"]]: failing(EXPIRED) })
  const state = await runLean(directory, { sessions: model.port })

  expect(state.provider_stop).toBe("auth")
  expect(state.stop_reason).toStartWith("d1.review-independent: ")
  expect(state.stages["d1.review"]?.status).toBe("done")
  expect(state.stages["d1.review-independent"]?.status).toBe("failed")
}, 120_000)

/** An engine on which every turn fails with this native error. */
function failingEngine(error: Record<string, unknown>) {
  const transcripts = new Map<string, Record<string, unknown>[]>()
  return (async (target: string | URL | Request, init?: RequestInit) => {
    const url = new URL(typeof target === "string" || target instanceof URL ? target : target.url)
    const [, id = "", route] = /^\/api\/session\/([^/]+)(?:\/(\w+))?$/.exec(url.pathname) ?? []
    const transcript = transcripts.get(id) ?? []
    transcripts.set(id, transcript)
    if (!route) return Response.json({ data: { id, location: { directory: new Headers(init?.headers).get("x-opencode-directory") } } })
    if (route === "prompt") {
      const body = JSON.parse(String(init?.body)) as { id: string; text: string }
      transcript.push({ id: body.id, type: "user", text: body.text, time: { created: 1 } },
        { id: `msg_${String(transcript.length).padStart(12, "0")}${"a".repeat(14)}`, type: "assistant", time: { created: 2, completed: 3 }, finish: "error", error, content: [] },
        { id: `idle_${transcript.length}`, type: "idle", outcome: "failed", time: { created: 4 } })
      return Response.json({ data: { id: body.id, sessionID: id, type: "user", payload: { text: body.text }, delivery: "queue", time: { created: 1 } } })
    }
    return Response.json({ data: transcript.toReversed().slice(0, Number(url.searchParams.get("limit"))), cursor: {} })
  }) as typeof fetch
}

test("no record keeps the provider's own words, only the failure", async () => {
  const directory = matter()
  const model = scriptedModel(happyPath)
  const engine = engineSessionPort({ engineURL: "http://engine.test", fetch: failingEngine(EXPIRED), pollMs: 1 })
  const sessions: StageSessionPort = { ...model.port, prompt: input => input.agent === "lean-checklist-extract" ? engine.prompt(input) : model.port.prompt(input) }
  const state = await runLean(directory, { sessions })

  expect(state.provider_stop).toBe("auth")
  expect(JSON.stringify(readLeanJson(directory, "checklist/extraction.json"))).toContain(PROVIDER_STOP_MESSAGES.auth)
  const files = (readdirSync(directory, { recursive: true }) as string[]).map(name => path.join(directory, name)).filter(file => statSync(file).isFile())
  expect(files.filter(file => readFileSync(file, "utf8").includes(EXPIRED.message))).toEqual([])
}, 120_000)

const deferChecklistRepair: Reply = text => JSON.stringify({ shared_decisions: [], groups: [], dismissed: [], questions: [],
  ...(text.includes("ONE supplied checklist assessment") ? { superseded_findings: [], superseded_decisions: [], findings: [{ disposition: "deferred", kind: "information",
    detail: "The assessed requirement still needs a supported correction.", reason: "No supported correction is identified for this assessment.", documents: ["d1"], sources: [], limitations: ["No supported correction identified."] }] } : {}) })

/** One requested item, checked with this result or failing its check. */
function checklistModel(check: Reply | Record<string, unknown>) {
  return scriptedModel({
    ...happyPath,
    [AGENT["change-plan"]]: deferChecklistRepair,
    [CHECKLIST_JOB_AGENTS.extract]: text => {
      const input = JSON.parse(text.split("SOURCE_INPUT:\n\n")[1]!)
      const item = { check: "The requested supply note is drafted.", basis: "directive", source_refs: input.passages, qualifications: [], target: { ...input.targets[0], treatment: "operative" } }
      return JSON.stringify({ items: input.source.document === "request" ? [item] : [] })
    },
    [CHECKLIST_JOB_AGENTS.check]: typeof check === "function" ? check : text => {
      const input = JSON.parse(text.split("CHECK_INPUT:\n\n")[1]!)
      return JSON.stringify({ item_id: input.item_id, reason: `Focused result ${check.status}.`, evidence: input.current_evidence,
        source_refs: input.item.source_refs, inspection_scope: "Current body paragraphs only.", limitation: "Comments not inspected.", ...check })
    },
  })
}

test("a provider failure during the one checklist reinspection leaves it unspent, so Continue reinspects", async () => {
  const directory = matter()
  expect((await runLean(directory, { sessions: checklistModel({ status: "indeterminate", indeterminate_cause: "not_inspectable" }).port })).status).toBe("completed")
  const settled = readLeanState(directory)!
  writeLeanState(directory, { ...settled, status: "blocked", stop_reason: "Checklist handoff blocked. See checklist/report.md.",
    stages: { ...settled.stages, "checklist-check": { ...settled.stages["checklist-check"]!, status: "blocked" } } })
  enqueueLeanRequest(directory, { kind: "continue" })
  const stopped = await runLean(directory, { sessions: checklistModel(failing(EXPIRED)).port })
  expect(stopped.provider_stop).toBe("auth")

  enqueueLeanRequest(directory, { kind: "continue" })
  const fixed = checklistModel({ status: "satisfied" })
  const resumed = await runLean(directory, { sessions: fixed.port })
  expect(fixed.count(CHECKLIST_JOB_AGENTS.check)).toBe(1)
  expect(resumed.status).toBe("completed")
}, 120_000)
