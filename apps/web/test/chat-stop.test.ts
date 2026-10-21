import { expect, test } from "bun:test"
import type { SessionMessageInfo } from "@opencode/client"
import { projectMessages } from "../src/api/session-view"
import { toTurns } from "../src/components/chat/turn-model"

// A stopped chat answer keeps its partial output and reads as stopped, in the
// live view after Stop and in the stored history after a reload.

const model = { providerID: "test", id: "model" }
const transcript = [
  { type: "user", id: "msg_u1", text: "Who are the parties?", time: { created: 1 } },
  { type: "assistant", id: "msg_a1", agent: "qa", model, finish: "stop", time: { created: 2, completed: 3 }, content: [{ type: "text", text: "Acme Ltd and Beta LLP." }] },
  { type: "idle", id: "msg_i1", outcome: "succeeded", time: { created: 4 } },
  // Stopped mid-step: the engine fails the running step as interrupted.
  { type: "user", id: "msg_u2", text: "Summarise the letter.", time: { created: 5 } },
  { type: "assistant", id: "msg_a2", agent: "qa", model, finish: "tool-calls", time: { created: 6, completed: 7 }, content: [{ type: "text", text: "I'll read the letter." }] },
  { type: "assistant", id: "msg_a3", agent: "qa", model, finish: "error", error: { type: "aborted", message: "Step interrupted" }, time: { created: 8, completed: 9 }, content: [{ type: "text", text: "The letter demands" }] },
  { type: "idle", id: "msg_i2", outcome: "interrupted", time: { created: 10 } },
  // Stopped between steps: the last step had completed.
  { type: "user", id: "msg_u3", text: "When is payment due?", time: { created: 11 } },
  { type: "assistant", id: "msg_a4", agent: "qa", model, finish: "tool-calls", time: { created: 12, completed: 13 }, content: [{ type: "text", text: "Checking the terms." }] },
  { type: "idle", id: "msg_i3", outcome: "interrupted", time: { created: 14 } },
  // Stopped before any output: the step failed with nothing streamed.
  { type: "user", id: "msg_u4", text: "Who signed it?", time: { created: 15 } },
  { type: "assistant", id: "msg_a5", agent: "qa", model, finish: "error", error: { type: "aborted", message: "Step interrupted" }, time: { created: 16, completed: 17 }, content: [] },
  { type: "idle", id: "msg_i4", outcome: "interrupted", time: { created: 18 } },
  // A provider failure is not a stop.
  { type: "user", id: "msg_u5", text: "Which invoice is unpaid?", time: { created: 19 } },
  { type: "assistant", id: "msg_a6", agent: "qa", model, finish: "error", error: { type: "provider.unknown", message: "The provider is unavailable." }, time: { created: 20, completed: 21 }, content: [] },
  { type: "idle", id: "msg_i5", outcome: "failed", time: { created: 22 } },
] as SessionMessageInfo[]

test("an interrupted turn's last answer step is stored as stopped; other turns are not", () => {
  const stopped = projectMessages("ses", transcript)
    .filter((message) => message.info.stopped)
    .map((message) => message.info.id)
  expect(stopped).toEqual(["msg_a3", "msg_a4", "msg_a5"])
})

test("a stopped answer keeps its partial output, is marked stopped, and does not read as an error", () => {
  const turns = toTurns(projectMessages("ses", transcript), "matter").filter((turn) => turn.role === "assistant")
  expect(turns.map(({ text, stopped, error }) => ({ text, stopped: Boolean(stopped), error }))).toEqual([
    { text: "Acme Ltd and Beta LLP.", stopped: false, error: undefined },
    { text: "I'll read the letter.\n\nThe letter demands", stopped: true, error: undefined },
    { text: "Checking the terms.", stopped: true, error: undefined },
    // Nothing was streamed before Stop, but the turn still says it was stopped.
    { text: "", stopped: true, error: undefined },
    { text: "", stopped: false, error: "The provider is unavailable." },
  ])
})
