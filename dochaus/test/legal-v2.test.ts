import { expect, test } from "bun:test"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import legal from "../plugin/legal-v2"
import { recordSessionAgent, sessionAgent } from "../plugin/legal/session"

test("native tool hooks restore trusted agent identity independently of prior context", async () => {
  const directory = await mkdtemp(path.join(os.tmpdir(), "native-tool-identity-"))
  const hooks = new Map<string, (event: Record<string, unknown>) => Promise<void>>()
  await legal.setup({
    location: { directory, project: { directory } },
    tool: { transform: async () => {}, hook: async (name: string, callback: (event: Record<string, unknown>) => Promise<void>) => { hooks.set(name, callback) } },
    session: { hook: async () => {} },
  } as never)
  try {
    const before = hooks.get("execute.before")!
    const sessionID = "ses_native_identity"
    const event = { tool: "read", sessionID, messageID: "msg_identity", id: "call_identity", agent: "qa", input: { filePath: "note.md" } }
    // No model-context callback has populated the identity cache.
    await before(event)
    expect(sessionAgent(sessionID)).toBe("qa")
    recordSessionAgent(sessionID, "research")
    // Model-supplied fields must never change the trusted native identity.
    await before({ ...event, id: "call_identity_input", input: { filePath: "note.md", agent: "research" } })
    expect(sessionAgent(sessionID)).toBe("qa")
    // Runtime-private fields are never accepted from model tool input.
    await expect(before({ ...event, id: "call_private", input: { filePath: "note.md", __prepared: true } }))
      .rejects.toThrow("Tool input cannot supply runtime-private preparation fields")
    await hooks.get("execute.after")!({ tool: "read", sessionID, messageID: "msg_identity", id: "call_read", agent: "extract", input: { path: "note.md" }, status: "completed", result: { content: "Read file note.md, lines 1-1\n1: Text", metadata: { truncated: false } } })
    expect(sessionAgent(sessionID)).toBe("extract")
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})
