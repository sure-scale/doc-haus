import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { createRequire } from "node:module"
import { z } from "zod"
import { AppNodeBuilder } from "@opencode/core/effect/app-node-builder"
import { ToolOutput } from "@opencode/core/tool-output"
import readRecord from "../tool/read-record"
import { registerLegalTools } from "../tool/register-v2"
import { LegalPlugin } from "../plugin/legal/policy"
import { UNTRUSTED_RECORD_NOTICE } from "../lib/untrusted"
import { boundLargeToolOutput } from "../plugin/legal/tool-output"

const coreRequire = createRequire(import.meta.resolve("@opencode/core/tool-output"))
const { Effect } = await import(coreRequire.resolve("effect"))
const { Global } = await import(coreRequire.resolve("@opencode/util/global"))
const roots: string[] = []
const budgetOriginal = process.env.DOCHAUS_TOOL_OUTPUT_CHARS
afterEach(() => {
  roots.splice(0).forEach(root => rmSync(root, { recursive: true, force: true }))
  if (budgetOriginal === undefined) delete process.env.DOCHAUS_TOOL_OUTPUT_CHARS
  else process.env.DOCHAUS_TOOL_OUTPUT_CHARS = budgetOriginal
})

function fixture(value: unknown) {
  const root = mkdtempSync(path.join(tmpdir(), "record-pages-"))
  roots.push(root)
  const directory = path.join(root, "matter")
  mkdirSync(path.join(directory, ".dochaus"), { recursive: true })
  const file = path.join(directory, ".dochaus/record.json")
  writeFileSync(file, JSON.stringify(value))
  return { root, directory, file }
}

async function page(source: ReturnType<typeof fixture>, args: { pointer?: string; offset?: number; limit?: number; path?: string } = {}) {
  const input = z.object(readRecord.args).parse({ path: ".dochaus/record.json", ...args })
  const output = await readRecord.execute(input, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])
  if (typeof output === "string") throw new Error(output)
  expect(output.output.length).toBeLessThanOrEqual(6000)
  const body = JSON.parse(output.output)
  expect(body.notice).toBe(UNTRUSTED_RECORD_NOTICE)
  return { ...output, body }
}

test("a complete selected small record retains every qualifier and is read from current canonical bytes", async () => {
  const qualified = { id: "custody-release", statement: "The custodian may release the recorder only after approval.", condition: { all: ["written approval", "signed register"], exception: "Do not release while a hold is active." }, status: "proposal", accepted: false, count: 0, evidence: null }
  const source = fixture({ records: [qualified], unrelated: "x".repeat(25000) })
  const before = readFileSync(source.file)
  const selected = await page(source, { pointer: "/records/0" })
  expect(selected.body.kind).toBe("value")
  expect(selected.body.value).toEqual(qualified)
  expect(selected.metadata.truncated).toBe(false)
  expect(readFileSync(source.file).equals(before)).toBe(true)
  expect(readdirSync(path.join(source.directory, ".dochaus"))).toEqual(["record.json"])
  writeFileSync(source.file, JSON.stringify({ records: [{ ...qualified, status: "withdrawn" }] }))
  expect((await page(source, { pointer: "/records/0" })).body.value.status).toBe("withdrawn")
  // Reading creates neither a rewritten source nor paging sidecars.
  expect(before.toString()).toContain('"status":"proposal"')
  expect(readFileSync(source.file).toString()).toContain('"status":"withdrawn"')
})

test("exact string continuation exposes late evidence on one long line and reconstructs Unicode and whitespace", async () => {
  const text = 'α😀e\u0301\t  \r\n\\"'.repeat(10000) + "The original approval is still pending."
  const source = fixture({ content: { text } })
  const before = readFileSync(source.file)
  const segments: string[] = []
  let offset = 0
  while (true) {
    const next = await page(source, { pointer: "/content/text", offset, limit: 4000 })
    expect(next.body.kind).toBe("text")
    expect(next.body.offset).toBe(offset)
    expect(next.body.text).toBe(Array.from(text).slice(offset, offset + next.body.returned).join(""))
    expect(next.body.total).toBe(Array.from(text).length)
    expect(next.metadata.truncated).toBe(next.body.next_offset !== null)
    segments.push(next.body.text)
    if (next.body.next_offset === null) break
    expect(next.body.next_offset).toBe(offset + next.body.returned)
    offset = next.body.next_offset
  }
  expect(segments[0]).not.toContain("The original approval is still pending.")
  expect(segments.join("")).toBe(text)
  expect(readFileSync(source.file).equals(before)).toBe(true)
  expect((await page(source, { pointer: "/content/text", offset: Array.from(text).length })).body.text).toBe("")
})

test("large arrays and objects expose every direct child with resolvable pointers and original identity hints", async () => {
  const records = Array.from({ length: 180 }, (_, index) => ({ id: `record-${index}`, frame: { condition: "Only after authorization.", alternative: "Return to storage." }, detail: "evidence ".repeat(200) }))
  for (const value of [records, Object.fromEntries(records.map((record, index) => [`item/${index}~`, record]))]) {
    const source = fixture({ value })
    const pointers: string[] = []
    let offset = 0
    while (true) {
      const next = await page(source, { pointer: "/value", offset, limit: 7 })
      expect(next.body.kind).toBe("children")
      expect(next.body.offset).toBe(offset)
      expect(next.body.children.length).toBeGreaterThan(0)
      expect(next.body.children.length).toBeLessThanOrEqual(7)
      for (const child of next.body.children) {
        expect(child.type).toBe("object")
        expect(child.identity.value).toBe(`record-${pointers.length}`)
        pointers.push(child.pointer)
      }
      if (next.body.next_offset === null) break
      offset = next.body.next_offset
    }
    expect(new Set(pointers).size).toBe(180)
    expect((await page(source, { pointer: pointers[179] + "/frame" })).body.value).toEqual(records[179].frame)
  }
})

test("JSON Pointer keeps escaped keys, own properties, root primitives and array coordinates distinct", async () => {
  const source = fixture({ "": { "a/b": { "~key": false } }, "01": 0, ["__proto__"]: null, rows: [null] })
  expect((await page(source, { pointer: "//a~1b/~0key" })).body.value).toBe(false)
  expect((await page(source, { pointer: "/01" })).body.value).toBe(0)
  expect((await page(source, { pointer: "/rows/0" })).body.value).toBeNull()
  expect((await page(source, { pointer: "/__proto__" })).body.value).toBeNull()
  for (const value of [null, false, 0, [], {}]) expect((await page(fixture(value))).body.value).toEqual(value)
  for (const pointer of ["rows/0", "/bad~2key", "/rows/01", "/rows/-", "/rows/2", "/toString", "/rows/0/x"]) {
    expect(await readRecord.execute({ path: ".dochaus/record.json", pointer }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])).toContain("Could not read")
  }
})

test("a readable matter record passes, then path traversal and symlink escape are rejected without modifying either file", async () => {
  const source = fixture({ status: "available" })
  expect((await page(source)).body.value.status).toBe("available")
  const external = path.join(source.root, "outside.json")
  writeFileSync(external, '{"private":"outside"}')
  symlinkSync(source.root, path.join(source.directory, "escape"), "dir")
  symlinkSync(external, path.join(source.directory, "outside.json"))
  for (const candidate of ["../outside.json", external, "escape/outside.json", "outside.json", ".dochaus"]) {
    const result = await readRecord.execute({ path: candidate }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])
    expect(typeof result).toBe("string")
    expect(String(result)).not.toContain('"private":"outside"')
    expect(result).toContain("Could not read")
  }
  expect(readFileSync(external, "utf8")).toBe('{"private":"outside"}')
})

test("numeric values retain their original JSON tokens rather than rounding identifiers or inventing null", async () => {
  const source = fixture(null)
  const original = '{"large_id":9007199254740993,"large_exponent":1e999,"decimal":1.2300,"negative_zero":-0,"rawJSON":{"rawJSON":"ordinary data"}}'
  writeFileSync(source.file, original)
  const selected = await page(source)
  expect(selected.output).toContain('"value":' + original)
  const numeric = await page(source, { pointer: "/large_id" })
  expect(numeric.body.type).toBe("number")
  expect(numeric.output).toContain('"value":9007199254740993')
  expect(await readRecord.execute({ path: ".dochaus/record.json", pointer: "/large_id/rawJSON" }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])).toContain("traverses a scalar")
  expect((await page(source, { pointer: "/rawJSON/rawJSON" })).body.text).toBe("ordinary data")
  expect(readFileSync(source.file, "utf8")).toBe(original)
})

test("malformed JSON, missing files, bad offsets and invalid argument admission are explicit", async () => {
  const source = fixture({ text: "short" })
  expect(await readRecord.execute({ path: ".dochaus/missing.json" }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])).toContain("not found")
  expect(await readRecord.execute({ path: ".dochaus/record.json", pointer: "/text", offset: 6 }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])).toContain("beyond")
  writeFileSync(source.file, "{bad JSON}")
  expect(await readRecord.execute({ path: ".dochaus/record.json" }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])).toContain("valid JSON")
  writeFileSync(source.file, Buffer.from([123, 34, 120, 34, 58, 34, 255, 34, 125]))
  expect(await readRecord.execute({ path: ".dochaus/record.json" }, { directory: source.directory } as Parameters<typeof readRecord.execute>[1])).toContain("UTF-8")
  for (const args of [{ offset: -1 }, { offset: 0.5 }, { limit: 0 }, { limit: 4001 }]) expect(z.object(readRecord.args).safeParse({ path: ".dochaus/record.json", ...args }).success).toBe(false)
})

test("the additive registered tool retains its bounded page and exact cursor through legal and native output hooks", async () => {
  const source = fixture({ text: "\n".repeat(9000) + "Late current evidence." })
  const definitions = new Map<string, { input: z.ZodType; execute: (input: unknown, runtime: unknown) => Promise<{ content: string; metadata: Record<string, unknown> }> }>()
  await registerLegalTools({
    location: { directory: source.directory, project: { directory: source.directory } },
    tool: { transform: async (transform: (editor: unknown) => void) => {
      transform({ get: () => ({ description: "Native child", output: undefined, execute: () => {} }), remove: () => {}, add: (definition: never) => { const entry = definition as unknown as { name: string }; definitions.set(entry.name, definition) } })
      return { dispose: () => {} }
    } },
  } as unknown as Parameters<typeof registerLegalTools>[0])
  const registered = definitions.get("read-record")!
  expect(registered).toBeDefined()
  const args = registered.input.parse({ path: ".dochaus/record.json", pointer: "/text" })
  const hooks = await LegalPlugin({ directory: source.directory } as Parameters<typeof LegalPlugin>[0])
  await hooks["tool.execute.before"]!({ tool: "read-record", callID: "record-page" } as never, { args } as never)
  const result = await registered.execute(args, { agent: "review", sessionID: "record-session", messageID: "record-message", id: "record-page", signal: new AbortController().signal, progress: () => {} })
  const before = result.content
  const projected = { output: result.content, metadata: result.metadata }
  process.env.DOCHAUS_TOOL_OUTPUT_CHARS = "1000"
  const ordinary = { output: before }
  boundLargeToolOutput({ tool: "ordinary-record-preview" }, ordinary)
  expect(ordinary.output).not.toBe(before)
  expect(ordinary.output).toContain("[tool-output-preview]")
  expect(() => JSON.parse(ordinary.output)).toThrow()
  await hooks["tool.execute.after"]!({ tool: "read-record", callID: "record-page", args } as never, projected as never)
  expect(projected.output).toBe(before)
  const retained = await Effect.runPromise(ToolOutput.Service.use(service => service.truncate({ content: [{ type: "text" as const, text: projected.output }], metadata: projected.metadata })).pipe(Effect.provide(AppNodeBuilder.build(ToolOutput.node, [Global.node.replace(Global.layerWith({ data: source.root }))]))))
  expect(retained.content).toEqual([{ type: "text", text: before }])
  const body = JSON.parse(before)
  expect(body.next_offset).toBe(body.returned)
  expect(body.text.length).toBeGreaterThan(1000)
  expect(body.notice).toBe(UNTRUSTED_RECORD_NOTICE)
})
