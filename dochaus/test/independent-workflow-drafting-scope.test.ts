import { expect, test } from "bun:test"
import { mkdtempSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { LLM, Message } from "../../packages/ai/src/index"
import { OpenAIChat } from "../../packages/ai/src/protocols/openai-chat"
import { AI } from "../../packages/ai/src/promise"
import { ConfigMarkdown } from "@opencode/core/config/markdown"
import { ConfigAgentV1 } from "@opencode/core/v1/config/agent"
import { ConfigMigrateV1 } from "@opencode/core/v1/config/migrate"
import { Permission } from "@opencode/core/permission"

const core = path.resolve(path.dirname(Bun.resolveSync("@opencode/core/session", import.meta.dir)), "..")
const { Schema } = await import(Bun.resolveSync("effect", core))

// Select the product checkout, not a substitute implementation.
const source = process.env.DOCHAUS_PROOF_SOURCE_ROOT ?? path.resolve(import.meta.dir, "../..")
const { default: legal } = await import(path.join(source, "dochaus/plugin/legal-v2.ts"))

test("independent ordinary workflows use the installed legal hooks", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "independent-workflow-scope-"))
  const originalFetch = globalThis.fetch
  const originalURL = process.env.OPENCODE_URL
  const lowering = AI.make()
  const prompts = new Map(["qa", "research", "extract", "legal-review", "document-reviewer"].map(agent => [
    `ordinary-${agent}`, { agent, id: `ordinary-${agent}-user`, text: "Read the recorded document and report its stated condition with a citation." },
  ]))
  const reads: string[] = []
  try {
    const readAgent = (id: string) => {
      const markdown = ConfigMarkdown.parse(readFileSync(path.join(source, "dochaus/agent", `${id}.md`), "utf8"))
      return { id, ...ConfigMigrateV1.migrateAgent(Schema.decodeUnknownSync(ConfigAgentV1.Info)({
        ...markdown.data, prompt: markdown.content,
      }, { errors: "all", propertyOrder: "original" })) }
    }
    const reviewAgent = readAgent("legal-review")
    const documentReviewer = readAgent("document-reviewer")
    expect(Permission.evaluate("subagent", "document-reviewer", reviewAgent.permissions ?? []).effect).toBe("allow")
    for (const action of ["edit", "task"]) {
      expect(Permission.evaluate(action, "*", documentReviewer.permissions ?? []).effect).toBe("deny")
    }
    expect(documentReviewer.mode).toBe("subagent")
    // Neutral session snapshots bind the native read methods directly.
    globalThis.fetch = Object.assign(async () => {
      throw new Error("The independent workflow fixture permits no network requests.")
    }, { preconnect: originalFetch.preconnect }) as typeof fetch
    const hooks = new Map<string, (event: never) => Promise<void>>()
    const tools = new Map<string, (event: never) => Promise<void>>()
    await legal.setup({ location: { directory, project: { directory } },
      model: { list: async () => ({ data: [{ providerID: "provider", id: "model", modelID: "model",
        package: "@opencode/ai/providers/openai-compatible", capabilities: { reasoning: true } }] }) },
      agent: { list: async () => ({ data: [...["qa", "research", "extract"].map(id => ({
        id, mode: "primary", permissions: [], request: { body: {} },
      })), reviewAgent, documentReviewer] }) },
      tool: { transform: async () => {}, hook: async (name: string, callback: (event: never) => Promise<void>) => { tools.set(name, callback) } },
      session: {
        hook: async (name: string, callback: (event: never) => Promise<void>) => { hooks.set(name, callback) },
        get: async ({ sessionID }: { sessionID: string }) => {
          reads.push(`get:${sessionID}`)
          const prompt = prompts.get(sessionID)
          if (!prompt) throw new Error(`Unexpected native session: ${sessionID}`)
          return { id: sessionID, agent: prompt.agent, location: { directory },
            title: "Independent document conversation", ...(prompt.agent === "document-reviewer" ? { parentID: "ordinary-legal-review" } : {}),
            model: { providerID: "provider", id: "model" } }
        },
        messages: async ({ sessionID }: { sessionID: string }) => {
          reads.push(`messages:${sessionID}`)
          const prompt = prompts.get(sessionID)
          if (!prompt) throw new Error(`Unexpected native history: ${sessionID}`)
          return [{ id: prompt.id, type: "user", text: prompt.text, time: { created: 1 } }]
        },
      },
    } as never)
    for (const agent of ["qa", "research", "extract", "document-reviewer"]) {
      const sessionID = `ordinary-${agent}`
      const prompt = prompts.get(sessionID)!
      const event = { sessionID, messageID: `${sessionID}-request`, agent,
        model: { providerID: "provider", id: "model" }, options: {},
        system: [{ type: "text" as const, text: "Read and report the document evidence." }],
        messages: [Message.make({ id: prompt.id, role: "user", content: prompt.text })], tools: {},
      }
      await hooks.get("context")!(event as never)
      const body = await lowering.run(OpenAIChat.fromRequest(LLM.request({
        model: OpenAIChat.route.with({ endpoint: { baseURL: "http://provider.invalid" } }).model({ id: "model" }),
        system: event.system, messages: event.messages, tools: [],
      })))
      const http = { sessionID, agent, kind: "primary", model: event.model,
        request: new Request("http://provider.invalid/chat/completions", { method: "POST", body: JSON.stringify(body) }) }
      await hooks.get("http.request")!(http as never)
      expect((await http.request.json()).messages).toEqual(body.messages)
    }
    const review = { sessionID: "ordinary-legal-review", messageID: "ordinary-review-request", agent: "legal-review",
      model: { providerID: "provider", id: "model" }, options: {},
      system: [{ type: "text" as const, text: "Delegate the requested independent document review." }],
      messages: [Message.make({ id: prompts.get("ordinary-legal-review")!.id, role: "user",
        content: prompts.get("ordinary-legal-review")!.text })], tools: {},
    }
    await hooks.get("context")!(review as never)
    const ordinaryTask = { sessionID: review.sessionID, messageID: review.messageID, agent: review.agent,
      id: "ordinary-document-review-task", tool: "task", input: { subagent_type: "document-reviewer",
        description: "Review the recorded document", prompt: "Read the recorded document and report supported findings with citations." } }
    await tools.get("execute.before")!(ordinaryTask as never)
    expect(ordinaryTask.input.subagent_type).toBe("document-reviewer")
    expect(reads.length).toBeGreaterThan(0)
    console.log(JSON.stringify({ source, ordinary_context_http: ["qa", "research", "extract", "document-reviewer"],
      review_parent_context: "accepted", ordinary_task_policy: "accepted with actual migrated native subagent permission",
      limitations: "installed hook and public lowering only; local read transport; no native core permission request/session creation/provider execution or expired-child proof; permission uses actual config parser/migration/evaluator" }))
  } finally {
    await lowering.dispose()
    globalThis.fetch = originalFetch
    if (originalURL === undefined) delete process.env.OPENCODE_URL
    else process.env.OPENCODE_URL = originalURL
    rmSync(directory, { recursive: true, force: true })
  }
})
