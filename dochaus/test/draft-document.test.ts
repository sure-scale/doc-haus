import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { DOMParser } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { openRevision } from "../lib/docx-amendments"
import { readDocxCommentParagraphs } from "../lib/docx-comments"
import { docxodus } from "../lib/docxodus"
import { withParagraphPointers } from "../lib/drafting-amendment"
import { renderMarkdownToDocx } from "../lib/render"
import createTemplate from "../tool/create-template"
import draftDocument from "../tool/draft-document"
import updateTemplate from "../tool/update-template"

async function runDraftTool(input: {
  prefix: string
  matterId: string
  preferences?: Record<string, unknown>
  template: Uint8Array
  sessionID: string
  calls: Record<string, unknown>[]
}) {
  const root = mkdtempSync(path.join(tmpdir(), input.prefix))
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify(input.preferences ?? {}))
  writeFileSync(path.join(root, "matter.json"), JSON.stringify({ id: input.matterId }))
  const uploads: { actor: string | null; run: string | null }[] = []
  const server = Bun.serve({
    port: 0,
    async fetch(request) {
      if (new URL(request.url).pathname === "/templates/content") return new Response(new Uint8Array(input.template))
      uploads.push({ actor: request.headers.get("x-dochaus-actor"), run: request.headers.get("x-dochaus-run") })
      const file = (await request.formData()).get("file") as File
      writeFileSync(path.join(root, file.name), new Uint8Array(await file.arrayBuffer()))
      return Response.json({})
    },
  })
  try {
    // Isolate import-time URLs so this test cannot change another test's tool
    // transport or contact a configured stack.
    const child = Bun.spawn(
      [
        process.execPath,
        "-e",
        `
          import draft from ${JSON.stringify(path.resolve(import.meta.dir, "../tool/draft-document.ts"))};
          const directory = process.env.WORKSPACE_ROOT;
          const context = {
            directory, worktree: directory, sessionID: ${JSON.stringify(input.sessionID)}, messageID: ${JSON.stringify(input.sessionID)},
            agent: "draft", abort: new AbortController().signal, metadata() {}, async ask() {}
          };
          const results = [];
          for (const args of ${JSON.stringify(input.calls)}) {
            const result = await draft.execute(args, context);
            if (typeof result === "string") throw new Error(result);
            results.push(result);
          }
          console.log(JSON.stringify(results));
        `,
      ],
      { env: { ...process.env, WORKSPACE_ROOT: root, INGEST_URL: server.url.href.replace(/\/$/, "") }, stdout: "pipe", stderr: "pipe" },
    )
    const [output, error] = await Promise.all([new Response(child.stdout).text(), new Response(child.stderr).text()])
    expect(await child.exited, error).toBe(0)
    const results = JSON.parse(output.trim().split("\n").at(-1)!) as { output: string; metadata: Record<string, unknown> }[]
    const files = new Map<string, Uint8Array>()
    for (const call of input.calls) {
      const name = String(call.name)
      files.set(name, new Uint8Array(readFileSync(path.join(root, name))))
    }
    return { results, files, uploads }
  } finally {
    server.stop(true)
    rmSync(root, { recursive: true, force: true })
  }
}

test("draft-document knows no document form, so neither a from-scratch draft nor a template draft gets new cross-reference fields", async () => {
  const advice = [
    "# Advice on termination",
    "## 1. Background",
    "1.1 You asked whether the Supplier may terminate early.",
    "1.2 We reviewed the Agreement.",
    "## 2. Analysis",
    "2.1 The Agreement permits termination under clause 1.2.",
    "2.2 Clause 2.1 requires notice in writing.",
  ].join("\n\n")
  const templateText = [
    "# Services Agreement",
    "## 1. Definitions",
    "1.1 In this Agreement the terms in clause 2.1 apply.",
    "## 2. Services",
    "2.1 The Supplier shall provide the Services.",
    "2.2 The Customer shall pay as set out in clause 2.1.",
  ].join("\n\n")
  const template = (await renderMarkdownToDocx(templateText)).bytes
  const { files } = await runDraftTool({
    prefix: "dh-draft-crossrefs-",
    matterId: "draft-crossrefs",
    template,
    sessionID: "draft-crossrefs",
    calls: [
      { name: "advice.docx", content: advice },
      { name: "services.docx", template: "services.docx" },
    ],
  })
  for (const name of ["advice.docx", "services.docx"]) {
    const body = strFromU8(unzipSync(files.get(name)!)["word/document.xml"]!)
    expect(body).toContain("<w:numPr>")
    expect(body).not.toContain("<w:instrText")
  }
}, 120_000)

test("template drafting stamps the new document and leaves template bytes untouched", async () => {
  const template = await Bun.file(path.resolve(import.meta.dir, "../templates/_base.docx")).bytes()
  const snapshot = new Uint8Array(template)
  const { files, uploads } = await runDraftTool({
    prefix: "dh-template-attribution-",
    matterId: "template-test",
    preferences: { attorney: "Template Reviewing Lawyer", firm: "Template Reviewing Firm" },
    template,
    sessionID: "template-identity",
    calls: [{ name: "prepared.docx", template: "review.docx" }],
  })
  const parts = unzipSync(files.get("prepared.docx")!)
  const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
  expect(parse(parts["docProps/core.xml"]).getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "creator")[0]?.textContent).toBe("Template Reviewing Lawyer")
  expect(parse(parts["docProps/app.xml"]).getElementsByTagNameNS("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "Company")[0]?.textContent).toBe("Template Reviewing Firm")
  expect(template).toEqual(snapshot)
  expect(uploads).toEqual([{ actor: "agent:draft", run: "template-identity" }])
}, 120_000)

test("a template reference to an omitted component becomes an open item, reported in the result", async () => {
  const templateText = [
    "# Services Agreement",
    "## 1. Definitions",
    "1.1 In this Agreement the terms in clause 2.1 and Schedule 1 apply.",
    "## 2. Services",
    "2.1 [optional: supplier services] The Supplier shall provide the Services.",
    "2.2 The Customer shall pay as set out in clause 2.1.",
    "## Schedule 1 – Services",
    "The services listed here.",
  ].join("\n\n")
  const template = (await renderMarkdownToDocx(templateText, undefined, undefined, { form: "agreement" })).bytes
  expect(strFromU8(unzipSync(template)["word/document.xml"]!).match(/ REF _Ref\d+ \\r \\h /g)).toHaveLength(2)
  const { results, files } = await runDraftTool({
    prefix: "dh-template-crossrefs-",
    matterId: "template-crossrefs",
    template,
    sessionID: "template-crossrefs",
    calls: [{ name: "prepared.docx", template: "services.docx", omit: ["supplier services"] }],
  })
  const result = results[0] as { output: string; metadata: { removedReferences: string[] } }
  const marker = "[open item: refers to 2.1, which was removed]"
  expect(result.output).toContain(`2 cross-references to parts no longer in the draft are marked as open items: ${marker}.`)
  expect(result.metadata.removedReferences).toEqual([marker, marker])
  const opened = await openRevision(files.get("prepared.docx")!)
  try {
    const shown = opened.paragraphs.map((row) => `${row.metadata.list?.generatedLabel?.trim() ?? ""} ${row.match.text}`.trim())
    expect(shown.join("\n")).not.toContain("The Supplier shall provide")
    expect(shown).toEqual(expect.arrayContaining([
      `1.1 In this Agreement the terms in clause ${marker} and Schedule 1 apply.`,
      `2.1 The Customer shall pay as set out in clause ${marker}.`,
    ]))
  } finally {
    opened.session.close()
  }
}, 120_000)

test("actual scratch, template draft and template-builder tools emit pointed Word before upload", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "dh-generation-tools-"))
  await writeFile(path.join(directory, "matter.json"), JSON.stringify({ id: "generation-tools" }))
  const baseParts = unzipSync(await Bun.file(path.resolve(import.meta.dir, "../templates/_base.docx")).bytes())
  const baseXml = strFromU8(baseParts["word/document.xml"]!)
  const sectPr = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(baseXml)?.[0] ?? ""
  const body = [
    "<w:p><w:r><w:t>Recipient: [insert recipient]</w:t></w:r></w:p>",
    "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Notice</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>[insert notice date]</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
  ].join("")
  baseParts["word/document.xml"] = strToU8(baseXml.replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${body}${sectPr}</w:body>`))
  const imported = zipSync(baseParts)
  const importedBefore = new Uint8Array(imported)
  const templates = new Map<string, Uint8Array<ArrayBuffer>>([["imported-template.docx", new Uint8Array(imported)]])
  const emitted: Array<{ name: string; bytes: Uint8Array<ArrayBuffer> }> = []
  const server = Bun.serve({ port: 0, async fetch(request) {
    const url = new URL(request.url)
    if (url.pathname === "/templates/content") return new Response(templates.get(url.searchParams.get("name")!) ?? null)
    if (request.method === "GET" && url.pathname === "/templates")
      return Response.json({ templates: Array.from(templates.keys()).map(name => ({ name, description: "Generic record template" })) })
    if (request.method === "POST") {
      const form = await request.formData()
      const file = form.get("file") as File
      const bytes = new Uint8Array(await file.arrayBuffer())
      emitted.push({ name: file.name, bytes })
      if (url.pathname === "/templates") templates.set(file.name, bytes)
      if (url.pathname === "/matters/generation-tools/deliverables") await writeFile(path.join(directory, file.name), bytes)
      return Response.json({ ok: true })
    }
    return new Response("Unexpected route", { status: 404 })
  } })
  const originalFetch = globalThis.fetch
  globalThis.fetch = ((input, init) => {
    const original = new URL(input instanceof Request ? input.url : String(input))
    return originalFetch(new URL(original.pathname + original.search, server.url), init)
  }) as typeof fetch
  const context = { directory, worktree: directory, agent: "test", sessionID: "test", messageID: "test", callID: "test",
    abort: new AbortController().signal, metadata() {}, async ask() {} }
  try {
    expect(readDocxCommentParagraphs(imported).every(row => row.pids.length === 0)).toBe(true)
    expect(await draftDocument.execute({ name: "scratch.docx", content: "# Record note\n\nRetain the supplied copy." }, context)).toMatchObject({ title: "Drafted scratch.docx" })
    expect(await draftDocument.execute({ name: "filled.docx", template: "imported-template.docx", fills: [
      { placeholder: "[insert recipient]", value: "Customer" }, { placeholder: "[insert notice date]", value: "To be confirmed" },
    ] }, context)).toMatchObject({ title: "Drafted filled.docx" })
    expect(await createTemplate.execute({ name: "record-template.docx", content: "# Record template\n\n[insert recipient] shall receive a copy.", description: "Generic copy template" }, context)).toMatchObject({ title: "Created template record-template.docx" })
    expect(await updateTemplate.execute({ name: "record-template.docx", content: "# Record template\n\n[insert recipient] shall receive a copy and confirmation." }, context)).toMatchObject({ title: "Updated template record-template.docx" })
    expect(emitted.map(row => row.name)).toEqual(["scratch.docx", "filled.docx", "record-template.docx", "record-template.docx"])
    for (const row of emitted) {
      const paragraphs = readDocxCommentParagraphs(row.bytes)
      expect(paragraphs.length).toBeGreaterThan(0)
      expect(paragraphs.every(paragraph => paragraph.pids.length === 1)).toBe(true)
      expect(new Set(paragraphs.flatMap(paragraph => paragraph.pids)).size).toBe(paragraphs.length)
      expect(withParagraphPointers(row.bytes)).toBe(row.bytes)
      const session = (await docxodus()).openDocxSession(row.bytes, {})
      try { expect(session.project().markdown.length).toBeGreaterThan(0) } finally { session.close() }
    }
    expect(readDocxCommentParagraphs(emitted[1]!.bytes).map(row => row.paragraph_text)).toEqual(["Recipient: Customer", "Notice", "To be confirmed"])
    expect(new Uint8Array(imported)).toEqual(importedBefore)
    expect(templates.get("imported-template.docx")).toEqual(importedBefore)
    expect(new Uint8Array(await readFile(path.join(directory, "filled.docx")))).toEqual(emitted[1]!.bytes)
  } finally {
    globalThis.fetch = originalFetch
    server.stop(true)
    await rm(directory, { recursive: true, force: true })
  }
}, 120_000)
