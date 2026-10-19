import { afterEach, expect, test } from "bun:test"
import { PROVIDER_STOP_MESSAGES } from "../../lib/provider-stop"
import { runLean } from "./runner"
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
