import { expect, test } from "bun:test"
import type { SessionContext } from "@opencode/plugin/promise/session"
import { applyLegalContextChanges } from "./v2-context"
import type { Message, Part } from "./policy-types"
import { MAX_STEPS_PROMPT } from "@opencode/core/session/runner/max-steps"

type History = Array<{ info: Message; parts: Part[] }>
const history = (parts: unknown[]): History => [{ info: { id: "msg_user", role: "user", sessionID: "ses_test" }, parts }] as History
const context = (messages: unknown[]): SessionContext => ({ messages }) as SessionContext

test("native final-step control preserves the lean stage output contract", () => {
  const request = { id: "request", role: "user", content: [{ type: "text", text: 'Return {"amendments":[],"questions":[]} using the supplied evidence.' }] }
  const reasoning = { id: "reasoning", role: "assistant", content: [{ type: "reasoning", text: "verified work", providerMetadata: { vertex: { thoughtSignature: "opaque" } } }] }
  const final = { role: "user", content: [{ type: "text", text: MAX_STEPS_PROMPT }] }
  const event = { ...context([request, reasoning, final]), agent: "lean-change-prepare", tools: { read: { description: "read", input: {} } }, options: { maxTokens: 1000 }, system: [{ type: "text", text: "stage rules" }] } as SessionContext
  const before = structuredClone(event)
  applyLegalContextChanges(event, [], [])
  expect(event.messages.slice(0, 2)).toEqual(before.messages.slice(0, 2))
  expect(event.tools).toEqual(before.tools)
  expect(event.options).toEqual(before.options)
  expect(event.system).toEqual(before.system)
  expect(event.messages.at(-1)?.role).toBe("user")
  const part = event.messages.at(-1)?.content[0]
  expect(part?.type).toBe("text")
  expect(part?.type === "text" && part.text).toContain("required output format")
  expect(part?.type === "text" && part.text).toContain("Tools remain disabled")
  expect(part?.type === "text" && part.text).not.toContain("MUST provide a text response summarizing")
})

test("final-step adaptation leaves durable instructions, other agents and unrelated controls intact", () => {
  const durable = { id: "durable", role: "user", content: [{ type: "text", text: MAX_STEPS_PROMPT }] }
  const mixed = { role: "user", content: [{ type: "text", text: MAX_STEPS_PROMPT }, { type: "text", text: "additional content" }] }
  const assistant = { role: "assistant", content: [{ type: "text", text: MAX_STEPS_PROMPT }] }
  const unrelated = { role: "user", content: [{ type: "text", text: "CRITICAL - MAXIMUM STEPS REACHED\nDifferent user instruction." }] }
  const event = { ...context([durable, mixed, assistant, unrelated]), agent: "lean-change-prepare" } as SessionContext
  const before = structuredClone(event.messages)
  applyLegalContextChanges(event, [], [])
  expect(event.messages).toEqual(before)
  const other = { ...context([{ role: "user", content: [{ type: "text", text: MAX_STEPS_PROMPT }] }]), agent: "draft" } as SessionContext
  const original = structuredClone(other.messages)
  applyLegalContextChanges(other, [], [])
  expect(other.messages).toEqual(original)
})

test("prepared evidence preserves native attachments, skill text and signed reasoning order", () => {
  const before = history([{ id: "p1", type: "text", text: "request" }])
  const after = structuredClone(before)
  after[0]!.parts.push({ id: "p2", type: "text", text: "sealed evidence" } as Part)
  const reasoning = { type: "reasoning", text: "reason", providerMetadata: { vertex: { thoughtSignature: "opaque" } } }
  const media = { type: "media", media: { source: "url", url: "https://example.invalid/file" } }
  const event = context([
    { id: "checkpoint", role: "user", content: [{ type: "text", text: "native checkpoint" }], native: { state: "opaque" } },
    { id: "msg_user", role: "user", content: [{ type: "text", text: "skill" }, { type: "text", text: "request" }, media] },
    { id: "msg_assistant", role: "assistant", content: [reasoning] },
  ])
  const checkpoint = event.messages[0]
  applyLegalContextChanges(event, before, after)
  expect(event.messages[0]).toEqual(checkpoint)
  expect(event.messages[1]!.content).toEqual([{ type: "text", text: "skill" }, { type: "text", text: "request" }, media, { type: "text", text: "sealed evidence" }])
  expect(event.messages[2]!.content).toEqual([reasoning])
})

test("policy removal removes both tool call and result without removing unrelated checkpoint", () => {
  const before = history([{ id: "p1", type: "tool", callID: "call_1", tool: "task", state: { status: "completed", output: "old" } }])
  const event = context([
    { id: "msg_user", role: "assistant", content: [{ type: "tool-call", id: "call_1", name: "task", input: {} }] },
    { role: "tool", content: [{ type: "tool-result", id: "call_1", name: "task", result: { type: "text", value: "old" } }] },
    { id: "checkpoint", role: "user", content: [{ type: "text", text: "checkpoint" }] },
  ])
  applyLegalContextChanges(event, before, [])
  expect(event.messages).toHaveLength(1)
  expect(event.messages[0]!.id).toBe("checkpoint")
})

test("unsigned local task replacement has text and no dangling result", () => {
  const before = history([{ id: "p1", type: "tool", callID: "call_1", tool: "task", state: { status: "completed", output: "old" } }])
  const after = history([{ id: "p1", type: "text", text: "historical outcome" }])
  const event = context([
    { id: "msg_user", role: "assistant", content: [{ type: "tool-call", id: "call_1", name: "task", input: {} }] },
    { role: "tool", content: [{ type: "tool-result", id: "call_1", name: "task", result: { type: "text", value: "old" } }] },
  ])
  applyLegalContextChanges(event, before, after)
  expect(event.messages).toHaveLength(1)
  expect(event.messages[0]!.content).toEqual([{ type: "text", text: "historical outcome" }])
})

test("updated error guidance keeps the tool error channel", () => {
  const before = history([{ id: "p1", type: "tool", callID: "call_1", tool: "read", state: { status: "error", error: "stopped" } }])
  const after = history([{ id: "p1", type: "tool", callID: "call_1", tool: "read", state: { status: "error", error: "stopped; publish current result" } }])
  const event = context([{ role: "tool", content: [{ type: "tool-result", id: "call_1", name: "read", result: { type: "error", value: "stopped" } }] }])
  applyLegalContextChanges(event, before, after)
  expect(event.messages[0]!.content[0]).toMatchObject({ result: { type: "error", value: "stopped; publish current result" } })
})


test("compacted requests rehydrate only new sealed evidence without replaying the original input", () => {
  const before = history([{ id: "p1", type: "text", text: "original admitted task" }])
  const after = structuredClone(before)
  after[0]!.parts.push({ id: "p2", type: "text", text: "current sealed task and evidence" } as Part)
  const event = context([{ id: "compacted", role: "user", content: [{ type: "text", text: "native checkpoint" }] }])
  applyLegalContextChanges(event, before, after)
  expect(event.messages).toHaveLength(2)
  expect(event.messages[0]!.content).toEqual([{ type: "text", text: "native checkpoint" }])
  expect(event.messages[1]!.content).toEqual([{ type: "text", text: "current sealed task and evidence" }])
  expect(JSON.stringify(event.messages)).not.toContain("original admitted task")
})
