import { afterEach, expect, test } from "bun:test"
import { nativeStageResult, OutputLimitError } from "./native-stage-result"
import { runLean } from "./runner"
import { AGENT, cleanupMatters, happyPath, matter, scriptedModel } from "./test-support"

afterEach(cleanupMatters)

const USER = { info: { id: "msg_user", role: "user", time: { created: 1 } }, parts: [{ type: "text", text: "Draft it." }] }

/** The turn's reply, as the native transcript projects it. */
function reply(finish: string, parts: Record<string, unknown>[]) {
  return { info: { id: "msg_reply", role: "assistant", parentID: "msg_user", agent: "expander", finish, time: { created: 2, completed: 3 },
    tokens: { input: 9_000, output: 131_072, reasoning: 131_072, cache: { read: 0, write: 0 } } }, parts }
}

const REASONING = { type: "reasoning", text: "Thinking it over again..." }

test("a reply cut off at its output limit before any text stops the step with the reasoning it spent", () => {
  expect(() => nativeStageResult([USER, reply("length", [REASONING])], "msg_user", "expander"))
    .toThrow("The model's reply reached its output limit before giving any answer (131,072 reasoning tokens).")
  expect(() => nativeStageResult([USER, reply("length", [REASONING, { type: "text", text: "  " }])], "msg_user", "expander"))
    .toThrow("reached its output limit before giving any answer")
})

test("a reply cut off at its output limit after some text or a tool call is still the reply the checker reads", () => {
  expect(nativeStageResult([USER, reply("length", [REASONING, { type: "text", text: "{\"partial\":" }])], "msg_user", "expander")).toBe("{\"partial\":")
  expect(nativeStageResult([USER, reply("length", [REASONING, { type: "tool", tool: "write", state: { status: "error", error: "Tool call was cut off." } }])], "msg_user", "expander")).toBe("")
  expect(nativeStageResult([USER, reply("stop", [REASONING])], "msg_user", "expander")).toBe("")
})

test("without a reasoning count the stop names the output it spent, or no count at all", () => {
  const counted = (tokens: Record<string, unknown> | undefined) => {
    const message = reply("length", [])
    return [USER, { ...message, info: { ...message.info, tokens } }]
  }
  expect(() => nativeStageResult(counted({ input: 9_000, output: 4_096, reasoning: 0 }), "msg_user", "expander"))
    .toThrow("The model's reply reached its output limit before giving any answer (4,096 output tokens).")
  for (const tokens of [undefined, { input: 9_000 }]) {
    let thrown: unknown
    try { nativeStageResult(counted(tokens), "msg_user", "expander") } catch (error) { thrown = error }
    expect(thrown).toBeInstanceOf(OutputLimitError)
    expect((thrown as Error).message).toBe("The model's reply reached its output limit before giving any answer.")
  }
})

test("the output-limit stop is a runtime stop, not a checker retry, and Continue re-runs the step", async () => {
  const directory = matter()
  const cutOff = scriptedModel({ ...happyPath, [AGENT.discrepancy]: () => nativeStageResult([USER, reply("length", [REASONING])], "msg_user", "expander") })
  const stopped = await runLean(directory, { sessions: cutOff.port })

  // The differences and gaps jobs are asked together; neither is retried.
  expect(cutOff.count(AGENT.discrepancy)).toBe(2)
  expect(stopped.status).toBe("failed")
  expect(stopped.provider_stop).toBeUndefined()
  expect(stopped.stop_reason).toBe("discrepancy: stopped by a runtime error: The model's reply reached its output limit before giving any answer (131,072 reasoning tokens).")

  const resumed = await runLean(directory, { sessions: scriptedModel(happyPath).port })
  expect(resumed.status).toBe("completed")
}, 120_000)
