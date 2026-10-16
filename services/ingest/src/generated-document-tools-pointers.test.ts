import { expect, test } from "bun:test"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph, Table, TableCell, TableRow } from "docx"
import draftDocument from "../../../dochaus/tool/draft-document"
import createTemplate from "../../../dochaus/tool/create-template"
import updateTemplate from "../../../dochaus/tool/update-template"
import { readDocxCommentParagraphs } from "../../../dochaus/lib/docx-comments"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { docxodus } from "../../../dochaus/lib/docxodus"

test("actual scratch, template draft and template-builder tools emit pointed Word before upload", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "dh-generation-tools-"))
  await writeFile(path.join(directory, "matter.json"), JSON.stringify({ id: "generation-tools" }))
  const imported = await Packer.toBuffer(new Document({ sections: [{ children: [
    new Paragraph("Recipient: [insert recipient]"),
    new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("Notice")] }), new TableCell({ children: [new Paragraph("[insert notice date]")] })] })] }),
  ] }] }))
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
}, 120000)
