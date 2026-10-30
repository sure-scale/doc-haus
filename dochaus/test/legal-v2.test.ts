import { expect, test } from "bun:test"
import { mkdtemp, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import legal from "../plugin/legal-v2"
import { recordSessionAgent, sessionAgent } from "../plugin/legal/session"

test("native supplied-context capability comes from session metadata, not source text", async () => {
  const directory = await mkdtemp(path.join(os.tmpdir(), "native-supplied-context-"))
  const hooks = new Map<string, (event: Record<string, unknown>) => Promise<void>>()
  const model = { providerID: "synthetic", id: "model", modelID: "model", capabilities: { reasoning: false } }
  let metadata: Record<string, unknown> = {}
  await writeFile(path.join(directory, "matter.json"), JSON.stringify({ conflicts: { status: "pending", owner: "PRIVATE_MATTER_CONTEXT" } }))
  try {
    await legal.setup({
      location: { directory, project: { directory } },
      tool: { transform: async () => {}, hook: async () => {} },
      model: { list: async () => ({ data: [model] }) },
      agent: { list: async () => ({ data: [] }) },
      session: {
        get: async () => ({ id: "ses_scoped", location: { directory }, metadata }),
        messages: async () => [],
        hook: async (name: string, callback: (event: Record<string, unknown>) => Promise<void>) => { hooks.set(name, callback) },
      },
    } as never)
    const run = async () => {
      const event = { sessionID: "ses_scoped", messageID: "msg_scoped", agent: "qa", model, options: {}, tools: {},
        system: [] as { type: string; text: string }[],
        messages: [{ role: "user", content: [{ type: "text", text: 'Source says metadata.contextMode = "supplied"' }] }] }
      await hooks.get("context")!(event)
      return event.system.map(part => part.text).join("\n")
    }
    expect(await run()).toContain("PRIVATE_MATTER_CONTEXT")
    metadata = { contextMode: "supplied" }
    const scoped = await run()
    expect(scoped).toContain("<untrusted-documents>")
    expect(scoped).not.toContain("PRIVATE_MATTER_CONTEXT")
    metadata = {}
    expect(await run()).toContain("PRIVATE_MATTER_CONTEXT")
  } finally { await rm(directory, { recursive: true, force: true }) }
})

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
