// Deterministic guarantees of the general legal-session policy hooks, exercised
// through the real LegalPlugin instance against temp matters — no engine boot
// needed: the hooks are plain async functions once the plugin factory runs.
// Run from dochaus/: bun test test/legal-policy.test.ts
import { afterAll, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { isOfficialSource } from "../lib/research"
import { resolveTierModel } from "../lib/model-tiers"
import { adaptOpenAIPlanParams, optimizeRouterParams } from "../plugin/legal/model-policy"
import { LegalPlugin } from "../plugin/legal/policy"
import { recordSessionAgent } from "../plugin/legal/session"
import { liveText } from "../lib/extract"
import type { DocumentCitation } from "../lib/citations"

// Two matters side by side in one workspace — the ethical-wall scenario.
const workspace = mkdtempSync(path.join(tmpdir(), "legal-policy-workspace-"))
const matterA = path.join(workspace, "matter-a")
const matterB = path.join(workspace, "matter-b")
mkdirSync(matterA)
mkdirSync(matterB)
writeFileSync(path.join(matterB, "secret.docx"), "other matter's document")

afterAll(() => {
  rmSync(workspace, { recursive: true, force: true })
})

async function pluginFor(directory: string) {
  return LegalPlugin({ directory } as Parameters<typeof LegalPlugin>[0])
}

function callArgs(tool: string, args: Record<string, unknown>) {
  return [{ tool, sessionID: "test", callID: "test" }, { args }] as const
}

test("citation position repair verifies quotes without claiming the source was amended", async () => {
  const directory = mkdtempSync(path.join(workspace, "citation-position-"))
  const docPath = path.join(directory, "instructions.eml")
  const bytes = "Subject: Approved instructions\r\nFrom: client@example.invalid\r\nTo: counsel@example.invalid\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nPayment is due after receipt of the invoice.\r\n"
  writeFileSync(docPath, bytes)
  const text = await liveText(docPath)
  const excerpt = "Payment is due after receipt of the invoice."
  const start = text.indexOf(excerpt)
  expect(start).toBeGreaterThan(0)
  const citation: DocumentCitation = { documentName: "instructions.eml", docPath, section: "Instructions", excerpt, charStart: start, charEnd: start + excerpt.length }
  const hooks = await pluginFor(directory)
  const exact = { title: "passage", output: "unchanged tool result", metadata: { citations: [citation] } }
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "citation-position", callID: "exact", args: {} }, exact)
  expect(exact.output).toBe("unchanged tool result")
  expect(exact.metadata.citations[0]?.verified).toBe(true)

  const shifted = { title: "passage", output: "indexed passage", metadata: { citations: [{ ...citation, charStart: 0, charEnd: excerpt.length }] } }
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "citation-position", callID: "shifted", args: {} }, shifted)
  expect(shifted.metadata.citations[0]?.charStart).toBe(start)
  expect(shifted.metadata.citations[0]?.verified).toBe(true)
  expect(shifted.metadata.citations[0]?.reanchored).toBe(true)
  expect(shifted.output).toContain(excerpt)
  expect(shifted.output).not.toContain("changed after indexing")
  expect(shifted.output).not.toContain("Recommend re-uploading")

  const unsupported = { title: "passage", output: "unverified", metadata: { citations: [{ ...citation, excerpt: "Payment is due before delivery." }] } }
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "citation-position", callID: "unsupported", args: {} }, unsupported)
  expect(unsupported.metadata.citations).toEqual([])
  expect(unsupported.output).toContain("Do not quote or rely on the rejected passages.")
  expect(await Bun.file(docPath).text()).toBe(bytes)
})

test("rebuilt search output keeps the notes search-document recorded after its passages", async () => {
  const directory = mkdtempSync(path.join(workspace, "citation-notes-"))
  const docPath = path.join(directory, "instructions.eml")
  writeFileSync(docPath, "Subject: Approved instructions\r\nFrom: client@example.invalid\r\nTo: counsel@example.invalid\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nPayment is due after receipt of the invoice.\r\n")
  const excerpt = "Payment is due after receipt of the invoice."
  const citation: DocumentCitation = { documentName: "instructions.eml", docPath, section: "Instructions", excerpt, charStart: 0, charEnd: excerpt.length }
  const notes =
    '\n\nNote: amended base(s) not matched to an indexed document: "Instructions" (amended by Update.eml). ' +
    "Passages from the documents these names refer to may be superseded and are not flagged above."
  const hooks = await pluginFor(directory)
  const superseded = { amenderName: "Update.eml", operation: "restate" as const, charStart: 0, charEnd: 10 }
  const shifted = { title: "passage", output: `indexed passage${notes}`, metadata: { citations: [{ ...citation, superseded_by: superseded }], notes } }
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "citation-notes", callID: "shifted", args: {} }, shifted)
  expect(shifted.metadata.citations[0]?.reanchored).toBe(true)
  // Per-passage warnings are re-rendered from the citation itself.
  expect(shifted.output).toContain("[⚠ SUPERSEDED]")
  expect(shifted.output).toContain(`${excerpt}${notes}\n\n[citation-verification]`)

  const rejected = { title: "passage", output: `indexed passage${notes}`, metadata: { citations: [{ ...citation, excerpt: "Payment is due before delivery." }], notes } }
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "citation-notes", callID: "rejected", args: {} }, rejected)
  expect(rejected.metadata.citations).toEqual([])
  // With no verified passage left, notes about the passages have nothing to qualify.
  expect(rejected.output).toContain("No passages survived verification.\n\n[citation-verification]")
  expect(rejected.output).not.toContain("not matched to an indexed document")

  const mixed = { title: "passage", output: `indexed passages${notes}`, metadata: { citations: [citation, { ...citation, excerpt: "Payment is due before delivery." }], notes } }
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "citation-notes", callID: "mixed", args: {} }, mixed)
  expect(mixed.metadata.citations).toHaveLength(1)
  expect(mixed.output).toContain(`${excerpt}${notes}\n\n[citation-verification]`)
})

function chatParams(providerID: string, maxOutputTokens: number | undefined) {
  return {
    input: {
      sessionID: "test",
      agent: "test-agent",
      model: { providerID },
      provider: {},
      message: {},
    } as never,
    output: { temperature: 0.1, topP: 1, topK: 0, maxOutputTokens, options: {} },
  }
}

function chatParamsForAgent(sessionID: string, agent: string) {
  return {
    input: {
      sessionID,
      agent,
      model: { providerID: "google-vertex", capabilities: { reasoning: false } },
      provider: {},
      message: {},
    } as never,
    output: { temperature: 0.1, topP: 1, topK: 0, maxOutputTokens: undefined, options: {} },
  }
}

describe("matter isolation", () => {
  test("read reaching into a sibling matter is denied", async () => {
    const hooks = await pluginFor(matterA)
    await expect(
      hooks["tool.execute.before"]!(...callArgs("read", { filePath: path.join(matterB, "secret.docx") })),
    ).rejects.toThrow("matter-scoped")
  })

  test("relative ../ escape into a sibling matter is denied", async () => {
    const hooks = await pluginFor(matterA)
    await expect(
      hooks["tool.execute.before"]!(...callArgs("read", { filePath: "../matter-b/secret.docx" })),
    ).rejects.toThrow("matter-scoped")
    await expect(hooks["tool.execute.before"]!(...callArgs("glob", { path: "../matter-b" }))).rejects.toThrow(
      "matter-scoped",
    )
    await expect(hooks["tool.execute.before"]!(...callArgs("grep", { path: ".." }))).rejects.toThrow("matter-scoped")
  })

  test("paths inside the matter pass", async () => {
    const hooks = await pluginFor(matterA)
    await hooks["tool.execute.before"]!(...callArgs("read", { filePath: "contract.docx" }))
    await hooks["tool.execute.before"]!(...callArgs("read", { filePath: path.join(matterA, "contract.docx") }))
    await hooks["tool.execute.before"]!(...callArgs("glob", { pattern: "**/*.docx" }))
  })

  test("paths outside the workspace fall through to the engine's permission layer", async () => {
    const hooks = await pluginFor(matterA)
    await hooks["tool.execute.before"]!(...callArgs("read", { filePath: "/etc/hosts" }))
  })
})

describe("conflicts gate", () => {
  function matterWith(conflicts?: Record<string, unknown>) {
    const dir = mkdtempSync(path.join(workspace, "matter-conflicts-"))
    writeFileSync(path.join(dir, "matter.json"), JSON.stringify({ id: path.basename(dir), title: "t", conflicts }))
    return dir
  }

  test("pending blocks every editing and drafting tool", async () => {
    const hooks = await pluginFor(matterWith({ status: "pending", owner: "K. Counsel" }))
    for (const tool of ["redline", "tracked-changes", "word-integration", "draft-document"]) {
      await expect(hooks["tool.execute.before"]!(...callArgs(tool, {}))).rejects.toThrow("K. Counsel")
    }
  })

  test("pending never blocks reading or searching", async () => {
    const hooks = await pluginFor(matterWith({ status: "pending", owner: "K. Counsel" }))
    await hooks["tool.execute.before"]!(...callArgs("read", { filePath: "contract.docx" }))
    await hooks["tool.execute.before"]!(...callArgs("search-document", { query: "indemnity" }))
    await hooks["tool.execute.before"]!(...callArgs("cite", {}))
  })

  test("cleared and bypassed do not block", async () => {
    for (const conflicts of [
      { status: "cleared", by: "K. Counsel" },
      { status: "bypassed", by: "K. Counsel", reason: "sole counsel in jurisdiction" },
    ]) {
      const hooks = await pluginFor(matterWith(conflicts))
      await hooks["tool.execute.before"]!(...callArgs("redline", {}))
    }
  })

  test("an absent record never blocks — legacy matters predate the field", async () => {
    const hooks = await pluginFor(matterWith(undefined))
    await hooks["tool.execute.before"]!(...callArgs("redline", {}))
    const noMatterJson = await pluginFor(mkdtempSync(path.join(workspace, "matter-legacy-")))
    await noMatterJson["tool.execute.before"]!(...callArgs("draft-document", {}))
  })
})

describe("conflicts system fragment", () => {
  async function systemFor(dir: string, agent?: string) {
    const hooks = await pluginFor(dir)
    const output = { system: [] as string[] }
    const sessionID = `conflicts-policy-${path.basename(dir)}`
    if (agent) recordSessionAgent(sessionID, agent)
    await hooks["experimental.chat.system.transform"]!({ sessionID } as never, output)
    return output.system.join("\n")
  }

  function matterWith(conflicts?: Record<string, unknown>) {
    const dir = mkdtempSync(path.join(workspace, "matter-fragment-"))
    writeFileSync(path.join(dir, "matter.json"), JSON.stringify({ id: path.basename(dir), title: "t", conflicts }))
    return dir
  }

  test("absent record injects the none-recorded fragment", async () => {
    const system = await systemFor(matterWith(undefined))
    expect(system).toContain("<conflicts-status>")
    expect(system).toContain("No conflicts check is recorded")
  })

  test("pending injects owner and due", async () => {
    const system = await systemFor(matterWith({ status: "pending", owner: "K. Counsel", due: "2026-07-01" }))
    expect(system).toContain("pending, owned by K. Counsel")
    expect(system).toContain("2026-07-01")
  })

  test("bypassed injects who and why", async () => {
    const system = await systemFor(matterWith({ status: "bypassed", by: "K. Counsel", reason: "sole counsel" }))
    expect(system).toContain("bypassed by K. Counsel: sole counsel")
  })

  test("cleared is silent — fragment absence means cleared", async () => {
    const system = await systemFor(matterWith({ status: "cleared", by: "K. Counsel" }))
    expect(system).not.toContain("<conflicts-status>")
    expect(system).toContain("<untrusted-documents>")
  })

  test("internal stage disclosure preserves its output contract and the human reviewer note", async () => {
    for (const conflicts of [
      undefined,
      { status: "pending", owner: "K. Counsel" },
      { status: "bypassed", by: "K. Counsel", reason: "sole counsel" },
    ]) {
      const system = await systemFor(matterWith(conflicts), "lean-package-review")
      expect(system).toContain("<conflicts-status>")
      expect(system).toContain("<structured-stage-conflicts>")
      expect(system).toContain("no prose outside it")
      expect(system).toContain("do not add fields")
      expect(system).toContain("human-facing reviewer note must still disclose")
    }
    expect(await systemFor(matterWith(undefined), "reviewer")).not.toContain("<structured-stage-conflicts>")
    expect(await systemFor(matterWith({ status: "cleared", by: "K. Counsel" }), "lean-package-review")).not.toContain("<structured-stage-conflicts>")
  })
})

describe("low-throughput model budget", () => {
  test("caps configured low-throughput providers before the LLM call", async () => {
    const hooks = await pluginFor(matterA)
    const params = chatParams("azure-foundry", 32_000)
    await hooks["chat.params"]!(params.input, params.output)
    expect(params.output.maxOutputTokens).toBe(4096)
  })

  test("leaves other providers unchanged", async () => {
    const hooks = await pluginFor(matterA)
    const params = chatParams("google-vertex", 32_000)
    await hooks["chat.params"]!(params.input, params.output)
    expect(params.output.maxOutputTokens).toBe(32_000)
  })
})

describe("router model policy", () => {
  test("preserves the engine response allowance for reasoning models across the live policy hook", async () => {
    const dir = mkdtempSync(path.join(tmpdir(), "reasoning-response-budget-"))
    try {
      const plugin = await pluginFor(dir)
      for (const agent of ["router", "draft"]) {
        const params = chatParamsForAgent(`reasoning-${agent}`, agent)
        Object.assign(params.input, {
          model: { id: "reasoning-model", providerID: "test-provider", capabilities: { reasoning: true } },
        })
        Object.assign(params.output, { maxOutputTokens: 32_000 })
        await plugin["chat.params"]!(params.input, params.output)
        expect(params.output.maxOutputTokens).toBe(32_000)
      }
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test("omits the unsupported output-token parameter for OpenAI plan models after legal role caps", () => {
    const openai = { maxOutputTokens: 8192 as number | undefined }
    adaptOpenAIPlanParams("openai", openai)
    expect(openai.maxOutputTokens).toBeUndefined()

    const other = { maxOutputTokens: 8192 as number | undefined }
    adaptOpenAIPlanParams("opencode-go", other)
    expect(other.maxOutputTokens).toBe(8192)
  })

  test("disables Gemma thinking and caps the classification response", () => {
    const output = {
      temperature: 0.8,
      maxOutputTokens: 128_000,
      options: { chat_template_kwargs: { enable_thinking: true, other: "kept" } },
    }
    optimizeRouterParams("router", "gemma-4-26b-a4b-it-maas", output, false)
    expect(output).toEqual({
      temperature: 0,
      maxOutputTokens: 16,
      options: {
        chat_template_kwargs: { enable_thinking: false, other: "kept" },
        thinkingConfig: { includeThoughts: false, thinkingLevel: "minimal" },
      },
    })
  })

  test("does not change substantive agents", () => {
    const output = { temperature: 0.2, maxOutputTokens: 4096, options: {} }
    optimizeRouterParams("draft", "gemma-4-26b-a4b-it-maas", output, false)
    expect(output).toEqual({ temperature: 0.2, maxOutputTokens: 4096, options: {} })
  })
})

describe("model tiering", () => {
  test("resolveTierModel follows the live config selection", () => {
    const cfg = { model: "google-vertex/pro", small_model: "google-vertex/flash" }
    expect(resolveTierModel("smart", cfg)).toBe("google-vertex/pro")
    expect(resolveTierModel("fast", cfg)).toBe("google-vertex/flash")
    expect(resolveTierModel("fastest", cfg)).toBe("google-vertex/flash")
  })

  test("throughput tiers fall back to the primary model when small_model is unset", () => {
    const cfg = { model: "google-vertex/only" }
    expect(resolveTierModel("smart", cfg)).toBe("google-vertex/only")
    expect(resolveTierModel("fast", cfg)).toBe("google-vertex/only")
  })

  test("an empty config leaves generation unconfigured", () => {
    expect(resolveTierModel("smart", {})).toBeUndefined()
    expect(resolveTierModel("fast", {})).toBeUndefined()
    expect(resolveTierModel("fastest", {})).toBeUndefined()
  })
})

describe("Gemini continuation history", () => {
  test("replays an unsigned local task outcome as text while preserving signed provider calls", async () => {
    const dir = mkdtempSync(path.join(workspace, "matter-gemini-task-history-"))
    const hooks = await pluginFor(dir)
    const messages = [{
      info: {
        id: "gemini-history-user",
        sessionID: "gemini-history",
        role: "user",
        agent: "draft",
        model: { providerID: "google-vertex", modelID: "gemini-3.1-pro-preview" },
      },
      parts: [{ type: "text", text: "Continue from the local task outcome." }],
    }, {
      info: {
        id: "gemini-history-assistant",
        sessionID: "gemini-history",
        role: "assistant",
        agent: "draft",
        providerID: "google-vertex",
        modelID: "gemini-3.1-pro-preview",
      },
      parts: [{
        type: "tool",
        tool: "task",
        callID: "local-task",
        state: {
          status: "error",
          input: { description: "Repair the artifact" },
          error: "The local task stopped before publication.",
          metadata: { parentSessionId: "gemini-history", sessionId: "gemini-child" },
        },
      }, {
        type: "tool",
        tool: "task",
        callID: "signed-task",
        metadata: { vertex: { thoughtSignature: "provider-signature" } },
        state: {
          status: "completed",
          input: { description: "Complete the next task" },
          output: "Task completed.",
        },
      }],
    }]

    await hooks["experimental.chat.messages.transform"]!({} as never, { messages } as never)

    expect(messages[1]!.parts[0]).toMatchObject({
      type: "text",
      synthetic: true,
      text: expect.stringContaining("The local task stopped before publication."),
    })
    expect(messages[1]!.parts[1]).toMatchObject({
      type: "tool",
      metadata: { vertex: { thoughtSignature: "provider-signature" } },
    })
  })
})

describe("context budget telemetry", () => {
  test("records a matter-local SQLite turn before model conversion", async () => {
    const dir = mkdtempSync(path.join(workspace, "matter-context-"))
    const hooks = await pluginFor(dir)
    const sessionID = "context-budget-session"
    const params = chatParamsForAgent(sessionID, "draft")
    await hooks["chat.params"]!(params.input, params.output)
    const text = "Please draft the agreement."
    await hooks["experimental.chat.messages.transform"]!({} as never, {
      messages: [
        {
          info: { id: "msg-context-1", sessionID, type: "user" },
          parts: [{ type: "text", text }],
        },
      ],
    } as never)

    const db = new Database(path.join(dir, ".dochaus", "telemetry.db"), { readonly: true })
    const row = db
      .query("SELECT session_id, estimated_input_tokens, model, agent FROM turns WHERE session_id = ?")
      .get(sessionID) as { session_id: string; estimated_input_tokens: number; model: string; agent: string } | null
    db.close()
    expect(row).toEqual({
      session_id: sessionID,
      estimated_input_tokens: Math.ceil(text.length / 4),
      model: "google-vertex/unknown",
      agent: "draft",
    })
    expect(existsSync(path.join(dir, ".dochaus", "context-budget.jsonl"))).toBe(false)
  })

  test("attributes overlapping matter sessions from durable message identity", async () => {
    const dir = mkdtempSync(path.join(workspace, "matter-context-overlap-"))
    const hooks = await pluginFor(dir)
    await hooks["chat.params"]!(
      chatParamsForAgent("context-first", "draft").input,
      chatParamsForAgent("context-first", "draft").output,
    )
    await hooks["chat.params"]!(
      chatParamsForAgent("context-second", "review").input,
      chatParamsForAgent("context-second", "review").output,
    )
    await hooks["experimental.chat.messages.transform"]!({} as never, {
      messages: [
        {
          info: { id: "msg-context-second", sessionID: "context-second", type: "user" },
          parts: [{ type: "text", text: "Review the document." }],
        },
      ],
    } as never)

    const db = new Database(path.join(dir, ".dochaus", "telemetry.db"), { readonly: true })
    const rows = db.query("SELECT session_id, agent FROM turns ORDER BY id").all()
    db.close()
    expect(rows).toEqual([{ session_id: "context-second", agent: "review" }])
  })

  test("skips tool telemetry without a session id", async () => {
    const dir = mkdtempSync(path.join(workspace, "matter-tool-telemetry-"))
    const hooks = await pluginFor(dir)
    const errors: unknown[] = []
    const originalError = console.error
    console.error = (...args: unknown[]) => {
      errors.push(args)
    }
    try {
      await hooks["tool.execute.after"]!({ tool: "glob", callID: "missing-session", args: {} } as never, {
        output: "done",
      } as never)
    } finally {
      console.error = originalError
    }

    expect(errors).toHaveLength(0)
    if (!existsSync(path.join(dir, ".dochaus", "telemetry.db"))) return
    const db = new Database(path.join(dir, ".dochaus", "telemetry.db"), { readonly: true })
    const row = db.query("SELECT COUNT(*) AS count FROM tool_calls").get() as { count: number }
    db.close()
    expect(row.count).toBe(0)
  })
})

describe("tool output budget", () => {
  test("clips oversized generic tool output", async () => {
    const hooks = await pluginFor(matterA)
    const output = { output: "A".repeat(13_000) }
    await hooks["tool.execute.after"]!({ tool: "glob" } as never, output as never)
    expect(output.output).toContain("tool-output-preview")
    expect(output.output).not.toContain("A".repeat(6_500))
  })

  test("preserves exact read continuation and EOF boundaries when clipping", async () => {
    const hooks = await pluginFor(matterA)
    const lines = Array.from({ length: 667 }, (_, index) => `${index + 1}: ${"A".repeat(80)}`).join("\n")
    const continuation = {
      output:
        `<content>\n${lines}\n\n` +
        `(Output capped at 50 KB. Showing lines 1-489. Use offset=490 to continue.)\n</content>`,
    }
    await hooks["tool.execute.after"]!({ tool: "read" } as never, continuation as never)
    expect(continuation.output).toContain("tool-output-preview")
    const preview = /Tool preview showing through line (\d+)\. Use offset=(\d+) to continue\./.exec(
      continuation.output,
    )
    expect(preview).not.toBeNull()
    expect(Number(preview![2])).toBe(Number(preview![1]) + 1)
    expect(Number(preview![2])).toBeLessThan(490)
    expect(continuation.output).toContain("Changing only the limit repeats discarded prefix text")

    const eof = {
      output:
        `<content>\n1: complete\n\n(End of file - total 1 lines)\n</content>` +
        `<system-reminder>${"B".repeat(13_000)}</system-reminder>`,
    }
    await hooks["tool.execute.after"]!({ tool: "read" } as never, eof as never)
    expect(eof.output).toContain("tool-output-preview")
    expect(eof.output).toEndWith("(End of file - total 1 lines)")
  })
})

describe("official source classification", () => {
  test("treats legal-information mirrors as non-official by default", () => {
    expect(isOfficialSource("https://www.law.cornell.edu/uscode/text/1/1")).toBe(false)
    expect(isOfficialSource("https://www.courtlistener.com/opinion/123/example/")).toBe(false)
    expect(isOfficialSource("https://uscode.house.gov/view.xhtml?req=1+USC+1")).toBe(true)
  })
})
