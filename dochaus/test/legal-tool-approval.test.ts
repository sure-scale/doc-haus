import { expect, test } from "bun:test"
import { readdirSync, readFileSync } from "node:fs"
import path from "node:path"
import { ConfigMarkdown } from "@opencode/core/config/markdown"
import { ConfigAgentV1 } from "@opencode/core/v1/config/agent"
import { ConfigMigrateV1 } from "@opencode/core/v1/config/migrate"
import { Permission } from "@opencode/core/permission"
import { assertPermission } from "../tool/register-v2"

const core = path.resolve(path.dirname(Bun.resolveSync("@opencode/core/session", import.meta.dir)), "..")
const { Schema } = await import(Bun.resolveSync("effect", core))
const root = path.resolve(import.meta.dir, "..")
const shipped = JSON.parse(readFileSync(path.join(root, "opencode.json"), "utf8"))
const globalRules = ConfigMigrateV1.migrate({ permission: shipped.permission }).permissions ?? []
const askActions = Object.entries(shipped.permission).filter(([, effect]) => effect === "ask").map(([action]) => action)
const agents = ["agent", "mode"].flatMap(directory => readdirSync(path.join(root, directory))
  .filter(file => file.endsWith(".md"))
  .map(file => {
    const markdown = ConfigMarkdown.parse(readFileSync(path.join(root, directory, file), "utf8"))
    // Use the same markdown/schema/migration pipeline as the native agent loader,
    // including its conversion of tools:true into agent-level allow rules.
    const info = Schema.decodeUnknownSync(ConfigAgentV1.Info)({ ...markdown.data, prompt: markdown.content }, { errors: "all", propertyOrder: "original" })
    return { id: file.slice(0, -3), declared: markdown.data, permissions: ConfigMigrateV1.migrateAgent(info).permissions ?? [] }
  }))

function rules(id: string) {
  const agent = agents.find(agent => agent.id === id)
  if (!agent) throw new Error(`Missing shipped agent: ${id}`)
  return [...globalRules, ...agent.permissions]
}

test("all enabled shipped mutating ask tools still ask after actual native agent normalization", () => {
  const affected = agents.flatMap(agent => askActions
    .filter(action => agent.declared.tools?.[action] === true || agent.declared.tools?.["*"] === true)
    .map(action => ({ agent: agent.id, action })))
  expect(affected.length).toBeGreaterThan(0)
  for (const pair of affected) {
    const declared = agents.find(agent => agent.id === pair.agent)!.declared
    const withoutAsk = { ...declared, permission: { ...declared.permission } }
    delete withoutAsk.permission[pair.action]
    const prior = ConfigMigrateV1.migrateAgent(Schema.decodeUnknownSync(ConfigAgentV1.Info)(withoutAsk)).permissions ?? []
    expect(Permission.evaluate(pair.action, "/synthetic/document.docx", globalRules, prior).effect).toBe("allow")
    expect({ ...pair, effect: Permission.evaluate(pair.action, "/synthetic/document.docx", rules(pair.agent)).effect })
      .toEqual({ ...pair, effect: "ask" })
  }
  // Enabling a specialist must not turn the global/native edit and shell
  // denials, or its own disabled-tool rules, into new grants.
  for (const agent of agents) {
    for (const action of ["edit", "shell"]) expect(Permission.evaluate(action, "*", rules(agent.id)).effect).toBe("deny")
    for (const [action, enabled] of Object.entries(agent.declared.tools ?? {})) {
      if (enabled !== false || action === "*") continue
      expect(Permission.evaluate(ConfigMigrateV1.normalizeAction(action), "*", rules(agent.id)).effect).toBe("deny")
    }
  }
})

// Local HTTP transport exercises the real assertPermission adapter and native
// evaluator with the actual normalized agent rules. It is not a generation
// engine or proof of the restarted product; live dialog acceptance is separate.
for (const [agent, action, decision] of [["redliner", "tracked-changes", "once"], ["redactor", "redact", "reject"], ["redliner", "copy-document", "once"], ["redliner", "copy-document", "reject"]] as const) {
  test(`${agent} ${action} cannot finish its permission assertion until a matching ${decision} reply`, async () => {
    const originalURL = process.env.OPENCODE_URL
    const originalPassword = process.env.OPENCODE_SERVER_PASSWORD
    const requestReceived = Promise.withResolvers<{ id: string; sessionID: string; action: string; resources: string[] }>()
    let stream: ReadableStreamDefaultController<Uint8Array> | undefined
    const encode = (event: object) => new TextEncoder().encode(`data: ${JSON.stringify(event)}\n\n`)
    const server = Bun.serve({ port: 0, async fetch(request) {
      const url = new URL(request.url)
      if (url.pathname === "/api/event") return new Response(new ReadableStream({ start(controller) {
        stream = controller
        controller.enqueue(encode({ type: "server.connected", data: {} }))
      } }), { headers: { "content-type": "text/event-stream" } })
      if (request.method !== "POST" || url.pathname !== "/api/session/ses_synthetic/permission") return new Response(null, { status: 404 })
      const input = await request.json()
      const effect = Permission.evaluate(input.action, input.resources[0], rules(input.agent)).effect
      requestReceived.resolve({ ...input, sessionID: "ses_synthetic" })
      return Response.json({ data: { id: input.id, effect } })
    } })
    try {
      process.env.OPENCODE_URL = server.url.toString()
      delete process.env.OPENCODE_SERVER_PASSWORD
      const assertion = assertPermission({} as never, { sessionID: "ses_synthetic", agent, messageID: "msg_synthetic", id: "call_synthetic", signal: new AbortController().signal }, {
        permission: action, patterns: ["/synthetic/document.docx"], always: ["/synthetic/document.docx"], metadata: { document: "document.docx" },
      })
      const settled = assertion.then(() => "resolved", () => "rejected")
      const request = await requestReceived.promise
      expect(request.action).toBe(action)
      expect(request.resources).toEqual(["/synthetic/document.docx"])
      expect(await Promise.race([settled, Bun.sleep(20).then(() => "waiting")])).toBe("waiting")
      // A reply for another session cannot release this document operation.
      stream!.enqueue(encode({ type: "permission.replied", data: { requestID: request.id, sessionID: "ses_other", reply: "once" } }))
      expect(await Promise.race([settled, Bun.sleep(20).then(() => "waiting")])).toBe("waiting")
      stream!.enqueue(encode({ type: "permission.replied", data: { requestID: request.id, sessionID: request.sessionID, reply: decision } }))
      expect(await settled).toBe(decision === "reject" ? "rejected" : "resolved")
      if (decision === "reject") await expect(assertion).rejects.toThrow(`Permission ${action} declined`)
      else await assertion
    } finally {
      server.stop(true)
      if (originalURL === undefined) delete process.env.OPENCODE_URL
      else process.env.OPENCODE_URL = originalURL
      if (originalPassword === undefined) delete process.env.OPENCODE_SERVER_PASSWORD
      else process.env.OPENCODE_SERVER_PASSWORD = originalPassword
    }
  }, 10_000)
}
