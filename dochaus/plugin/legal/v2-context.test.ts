import { expect, test } from "bun:test"
import type { SessionContext } from "@opencode/plugin/promise/session"
import { applyLegalContextChanges } from "./v2-context"
import type { Message, Part } from "./policy-types"
import { MAX_STEPS_PROMPT } from "@opencode/core/session/runner/max-steps"
import { projectSessionMessages } from "../../lib/engine-v2-projection"
import { nativeStageResult } from "../../workflows/lean/native-stage-result"

type History = Array<{ info: Message; parts: Part[] }>
const history = (parts: unknown[]): History => [{ info: { id: "msg_user", role: "user", sessionID: "ses_test" }, parts }] as History
const context = (messages: unknown[]): SessionContext => ({ messages }) as SessionContext
const interruptedControl = "The previous response was interrupted. Continue from where you left off without repeating completed content."

function interruptedTurn(agent = "lean-drafting-notes") {
  const native = [
    { id: "admitted", type: "user", text: "Return this stage's complete output with the supplied evidence and unknowns.", time: { created: 1 } },
    { id: "interrupted", type: "assistant", agent, finish: "error", error: { type: "provider.rate-limit" }, time: { created: 2, completed: 3 },
      content: [{ type: "text", text: "DISCARDED_ERRORED_PREFIX: The substantive draft was interrupted." }] },
    { id: "native-retry", type: "synthetic", text: interruptedControl, time: { created: 4 } },
  ]
  const before = projectSessionMessages(native, { id: "native-session", agent }) as unknown as History
  const event = { ...context([
    { id: "admitted", role: "user", content: [{ type: "text", text: native[0]!.text }] },
    { id: "interrupted", role: "assistant", content: native[1]!.content },
    { id: "native-retry", role: "user", content: [{ type: "text", text: interruptedControl }] },
  ]), agent, tools: { read: { description: "read", input: {} } }, options: { maxTokens: 1000 }, system: [{ type: "text", text: "Keep unknowns and approval requirements." }] } as SessionContext
  return { native, before, event, agent }
}

for (const output of [
  { format: "JSON", complete: '{"observations":[{"text":"The seal design remains unknown; the custodian must supply it."}],"questions":[]}', suffix: '],"questions":[]}' },
  { format: "prose", complete: "The seal design remains unknown; the custodian must supply it before relying on the handover.", suffix: "Complete that check before relying on the handover." },
]) test(`an authenticated interrupted lean turn requests complete replacement ${output.format} for latest-reply acceptance`, () => {
  const { native, before, event, agent } = interruptedTurn()
  const durable = structuredClone(before)
  const original = structuredClone(event)
  // Deterministic continuation double: the native append instruction supplies
  // only a suffix; the replacement contract supplies the full checked artifact.
  const continueReply = (text: string) => text.includes("complete replacement") && !text.includes("without repeating completed content") ? output.complete : output.suffix
  expect(continueReply(interruptedControl)).toBe(output.suffix)
  applyLegalContextChanges(event, before, structuredClone(before))
  const control = event.messages.find(message => message.id === "native-retry")!.content[0]!
  const reply = continueReply(control.type === "text" ? control.text : "")
  const projected = projectSessionMessages([...native,
    { id: "completed", type: "assistant", agent, finish: "stop", time: { created: 5, completed: 6 }, content: [{ type: "text", text: reply }] },
  ], { id: "native-session", agent })
  const result = nativeStageResult(projected, "admitted", agent)
  expect(result).toBe(output.complete)
  if (output.format === "JSON") expect(JSON.parse(result)).toEqual(JSON.parse(output.complete))
  expect(result).not.toContain("DISCARDED_ERRORED_PREFIX")
  expect(before).toEqual(durable)
  expect(event.messages.slice(0, 2)).toEqual(original.messages.slice(0, 2))
  expect(event.tools).toEqual(original.tools)
  expect(event.options).toEqual(original.options)
  expect(event.system).toEqual(original.system)
  expect(control.type === "text" && control.text).toContain("latest successful reply alone")
  expect(control.type === "text" && control.text).toContain("Do not repeat completed tool actions")
  expect(control.type === "text" && control.text).toContain("unknowns")
  expect(control.type === "text" && control.text).toContain("approval")
})

test("interruption adaptation authenticates native identity and the current admitted lean turn", () => {
  const cases: Array<(value: ReturnType<typeof interruptedTurn>) => void> = [
    value => { value.event.agent = "draft" },
    value => { value.before = [] },
    value => { value.event.messages[2]!.id = "unmatched-control" },
    value => { (value.before[2]!.info as unknown as { type: string }).type = "user" },
    value => { delete (value.before[2]!.info as unknown as { type?: string }).type },
    value => { value.before[2]!.info.role = "assistant" },
    value => { value.event.messages[2]!.role = "assistant" },
    value => { value.event.messages[2]!.content.push({ type: "text", text: "source attachment" }) },
    value => { value.before.push(...history([{ id: "later-part", type: "text", text: "A later real user request." }])) },
    value => { value.before = value.before.slice(1) },
    value => {
      const text = "Your last response hit the output token limit (stop reason: length). Do not apologize, recap, or repeat yourself. Break the remaining work into smaller pieces."
      const part = value.before[2]!.parts[0]!
      if (part.type === "text") part.text = text
      value.event.messages[2]!.content = [{ type: "text", text }]
    },
    value => {
      const part = value.before[2]!.parts[0]!
      if (part.type === "text") part.text = "An unrelated native synthetic instruction."
    },
  ]
  for (const alter of cases) {
    const value = interruptedTurn()
    alter(value)
    const original = structuredClone(value.event)
    applyLegalContextChanges(value.event, value.before, structuredClone(value.before))
    expect(value.event).toEqual(original)
  }
})

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
