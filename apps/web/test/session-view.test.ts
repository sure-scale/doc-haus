import { expect, test } from "bun:test"
import { projectEvent, projectMessages } from "../src/api/session-view"
import type { SessionMessageInfo, V2Event } from "@opencode/client"

test("native session messages retain the chat's agent, text, tools, and citations", () => {
  const messages = [
    { type: "agent-switched", id: "msg_agent", agent: "extract", time: { created: 1 } },
    { type: "user", id: "msg_user", text: "Summarise", time: { created: 2 } },
    {
      type: "assistant", id: "msg_answer", agent: "extract", model: { providerID: "test", id: "model" },
      time: { created: 3, completed: 5 },
      content: [
        { type: "tool", id: "call_search", name: "search-document", time: { created: 3, completed: 4 }, state: {
          status: "completed", input: { query: "Summarise" }, content: [{ type: "text", text: "Found passage" }],
          metadata: { citations: [{ section: "A", excerpt: "Passage" }] },
        } },
        { type: "text", text: "Answer" },
      ],
    },
  ] as SessionMessageInfo[]

  const projected = projectMessages("ses_matter", messages)
  expect(projected).toHaveLength(2)
  expect(projected[0]?.info.agent).toBe("extract")
  expect(projected[1]?.info.parentID).toBe("msg_user")
  expect(projected[1]?.parts[0]).toMatchObject({
    type: "tool", tool: "search-document", state: { status: "completed", output: "Found passage", metadata: { citations: [{ section: "A" }] } },
  })
  expect(projected[1]?.parts[1]).toMatchObject({ type: "text", text: "Answer" })
})

test("native permission events retain the decision target and document resource", () => {
  const event = {
    type: "permission.asked", id: "evt_permission", created: 1,
    data: { id: "per_1", sessionID: "ses_matter", action: "redline", resources: ["contract.docx"], source: { type: "tool", messageID: "msg_answer", id: "call_edit" } },
  } as V2Event
  expect(projectEvent(event)).toEqual([{
    type: "permission.asked",
    properties: {
      id: "per_1", sessionID: "ses_matter", permission: "redline", patterns: ["contract.docx"],
      metadata: {}, always: [], tool: { messageID: "msg_answer", callID: "call_edit" },
    },
  }])
})

test("native execution success projects to session.idle; a step end projects nothing", () => {
  const native = (type: string, data: Record<string, unknown>) =>
    ({ type, id: `evt_${type}`, created: 1, durable: { aggregateID: "ses_matter", seq: 1, version: 1 }, data: { sessionID: "ses_matter", ...data } }) as V2Event
  // An answer built over several steps completes an assistant message per step.
  expect(projectEvent(native("session.step.ended", { assistantMessageID: "msg_step", finish: "tool-calls" }))).toEqual([])
  expect(projectEvent(native("session.execution.succeeded", {}))).toEqual([
    { type: "session.idle", properties: { sessionID: "ses_matter" } },
  ])
})
