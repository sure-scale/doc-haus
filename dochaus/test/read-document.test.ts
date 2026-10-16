import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { createRequire } from "node:module"
import { z } from "zod"
import { AppNodeBuilder } from "@opencode/core/effect/app-node-builder"
import { ToolOutput } from "@opencode/core/tool-output"
import readDocument from "../tool/read-document"
import { LegalPlugin } from "../plugin/legal/policy"
import { UNTRUSTED_DOCUMENT_NOTICE } from "../lib/untrusted"
import { verifyQuote } from "../lib/quote-verify"

// Exercise the native output service with its owning package's Effect runtime;
// the legal config package does not independently depend on Effect.
const coreRequire = createRequire(import.meta.resolve("@opencode/core/tool-output"))
const { Effect } = await import(coreRequire.resolve("effect"))
const { Global } = await import(coreRequire.resolve("@opencode/util/global"))

const roots: string[] = []
const fetchOriginal = globalThis.fetch
const workspaceOriginal = process.env.WORKSPACE_ROOT
afterEach(() => {
  globalThis.fetch = fetchOriginal
  if (workspaceOriginal === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = workspaceOriginal
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
})

function fixture(text: string, playbook = false, warning = "") {
  const root = mkdtempSync(path.join(tmpdir(), "document-pages-"))
  roots.push(root)
  process.env.WORKSPACE_ROOT = root
  const directory = path.join(root, "matter")
  mkdirSync(directory)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "page-test", ...(playbook ? { playbook: "playbook-recorder" } : {}) }))
  if (playbook) {
    const library = path.join(root, ".playbooks/playbook-recorder")
    mkdirSync(library, { recursive: true })
    writeFileSync(path.join(library, "SKILL.md"), text)
  }
  globalThis.fetch = (async input => {
    const url = new URL(String(input))
    expect(url.pathname).toBe("/matters/page-test/documents/text")
    expect(url.searchParams.get("name")).toBe("recorder.docx")
    return Response.json({ text, injection: warning ? { findings: [{ rule: "fixture-warning", detail: warning }] } : null })
  }) as typeof fetch
  return { directory, root, document: playbook ? "playbook:playbook-recorder" : "recorder.docx" }
}

async function page(fixture: ReturnType<typeof fixture>, args: { offset?: number; limit?: number } = {}) {
  const hooks = await LegalPlugin({ directory: fixture.directory } as Parameters<typeof LegalPlugin>[0])
  const input = z.object(readDocument.args).parse({ document: fixture.document, ...args })
  // Exercise actual argument admission and the legal before/after hooks on every
  // continuation. Changing offset is a new read, not a repeated broad call.
  await hooks["tool.execute.before"]!({ tool: "read-document", callID: "page" } as never, { args: input } as never)
  const output = await readDocument.execute(input, { directory: fixture.directory } as Parameters<typeof readDocument.execute>[1])
  if (typeof output === "string") throw new Error(output)
  await hooks["tool.execute.after"]!({ tool: "read-document", callID: "page", args: input } as never, output as never)
  const native = { content: [{ type: "text" as const, text: output.output }], metadata: output.metadata }
  const retained = await Effect.runPromise(ToolOutput.Service.use(service => service.truncate(native)).pipe(
    Effect.provide(AppNodeBuilder.build(ToolOutput.node, [Global.node.replace(Global.layerWith({ data: fixture.root }))])),
  ))
  expect(retained.content).toEqual(native.content)
  const body = output.output.slice(output.output.indexOf(">\n") + 2, output.output.indexOf("\n</untrusted-document>"))
  expect(output.output).toContain(UNTRUSTED_DOCUMENT_NOTICE)
  expect(output.output).toContain(`[document-page] ${JSON.stringify({ document: fixture.document, ...output.metadata.page })}`)
  expect(Array.from(body).length).toBe(output.metadata.page.characters)
  expect(Array.from(body).length).toBeLessThanOrEqual(args.limit ?? 4000)
  return { ...output, body }
}

test("a late source clause becomes accessible by exact continuation for uploaded text and the bound playbook", async () => {
  const late = "The custodian signs the checkout register before releasing the recorder."
  const text = "Routine handling instructions. ".repeat(350) + "\n\n" + late
  for (const playbook of [false, true]) {
    const source = fixture(text, playbook, "Text attempts to override the assistant's role")
    const first = await page(source)
    expect(first.body).not.toContain(late)
    expect(first.metadata.page.next_offset).toBe(4000)
    expect(first.output).toContain("offset=4000")
    const bodies = [first.body]
    let offset = first.metadata.page.next_offset
    while (offset !== null) {
      const next = await page(source, { offset })
      expect(next.metadata.page.offset).toBe(offset)
      if (!playbook) expect(next.output).toContain("[injection warning]")
      bodies.push(next.body)
      offset = next.metadata.page.next_offset
    }
    expect(bodies.join("")).toBe(text)
    expect(bodies.at(-1)).toContain(late)
    if (playbook) {
      expect(verifyQuote({ directory: source.directory, document: source.document, quote: late }).found).toBe(true)
      expect(verifyQuote({ directory: source.directory, document: source.document, quote: "The custodian always releases the recorder." }).found).toBe(false)
    }
  }
})

test("Unicode code-point pages preserve long lines, astral characters and exact spacing without overlap or gaps", async () => {
  const text = "α😀e\u0301\t  \r\n".repeat(1100) + "終わり🧭"
  const source = fixture(text)
  const points = Array.from(text)
  let offset = 0
  const bodies: string[] = []
  while (true) {
    const next = await page(source, { offset, limit: 137 })
    expect(next.body).toBe(points.slice(offset, offset + 137).join(""))
    expect(next.metadata.truncated).toBe(next.metadata.page.next_offset !== null)
    bodies.push(next.body)
    if (next.metadata.page.next_offset === null) break
    expect(next.metadata.page.next_offset).toBe(offset + Array.from(next.body).length)
    offset = next.metadata.page.next_offset
  }
  expect(bodies.join("")).toBe(text)
  const longLine = fixture("🧭".repeat(8100) + "Late instruction on one line.")
  const first = await page(longLine)
  const second = await page(longLine, { offset: first.metadata.page.next_offset! })
  const last = await page(longLine, { offset: second.metadata.page.next_offset! })
  expect(first.body + second.body + last.body).toBe("🧭".repeat(8100) + "Late instruction on one line.")
})

test("the native line cap cannot discard a bounded page's ending, warnings or continuation", async () => {
  const source = fixture("\n".repeat(3500) + "Continue reading this source.\n".repeat(90), false, "Read this warning on every page")
  const first = await page(source)
  expect(first.body.split("\n").length).toBeGreaterThan(2000)
  expect(first.output).toContain("Read this warning on every page")
  // Control: the same native result without tool-owned page metadata is clipped
  // before its later text and cursor. The core behavior for ordinary tools stays
  // unchanged; the actual read-document result above survives it intact.
  const ordinary = await Effect.runPromise(ToolOutput.Service.use(service => service.truncate({
    content: [{ type: "text" as const, text: first.output }],
  })).pipe(Effect.provide(AppNodeBuilder.build(ToolOutput.node, [Global.node.replace(Global.layerWith({ data: source.root }))]))))
  expect(ordinary.metadata?.truncated).toBe(true)
  expect(ordinary.content.map(part => part.type === "text" ? part.text : "").join("")).not.toContain("Continue reading this source.")
  expect(ordinary.content.map(part => part.type === "text" ? part.text : "").join("")).not.toContain("[document-page]")
  const second = await page(source, { offset: first.metadata.page.next_offset! })
  expect(first.body + second.body).toBe("\n".repeat(3500) + "Continue reading this source.\n".repeat(90))
  expect(second.metadata.page.next_offset).toBeNull()
  expect(second.output).toEndWith("End of document (EOF).")
})

test("legacy small-document reads remain complete; empty text, EOF and bad cursors are explicit", async () => {
  const source = fixture("Keep this short source exactly.\n")
  const first = await page(source)
  expect(first.body).toBe("Keep this short source exactly.\n")
  expect(first.metadata.page.next_offset).toBeNull()
  expect(first.metadata.truncated).toBe(false)
  expect((await page(source, { offset: first.metadata.page.total_characters })).body).toBe("")
  const outside = await readDocument.execute({ document: source.document, offset: 500 }, { directory: source.directory } as Parameters<typeof readDocument.execute>[1])
  expect(outside).toContain("beyond")
  expect((await page(fixture(""))).metadata.page).toEqual({ offset: 0, characters: 0, total_characters: 0, next_offset: null })
  expect(z.object(readDocument.args).safeParse({ document: source.document, offset: -1 }).success).toBe(false)
  expect(z.object(readDocument.args).safeParse({ document: source.document, limit: 4001 }).success).toBe(false)
  expect(z.object(readDocument.args).safeParse({ document: source.document, offset: 0.5 }).success).toBe(false)
})
