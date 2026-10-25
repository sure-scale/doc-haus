// Run from dochaus/: bun test test/draft-template-crossrefs.test.ts
import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { openRevision } from "../lib/docx-amendments"
import { renderMarkdownToDocx } from "../lib/render"

// A neutral numbered template whose optional clause two others cite.
const TEMPLATE = ["# Services Agreement", "## 1. Definitions", "1.1 In this Agreement the terms in clause 2.1 and Schedule 1 apply.",
  "## 2. Services", "2.1 [optional: supplier services] The Supplier shall provide the Services.",
  "2.2 The Customer shall pay as set out in clause 2.1.", "## Schedule 1 – Services", "The services listed here."].join("\n\n")

test("a template reference to an omitted component becomes an open item, reported in the result", async () => {
  const template = (await renderMarkdownToDocx(TEMPLATE, undefined, undefined, { form: "agreement" })).bytes
  expect(strFromU8(unzipSync(template)["word/document.xml"]!).match(/ REF _Ref\d+ \\r \\h /g)).toHaveLength(2)
  const root = mkdtempSync(path.join(tmpdir(), "dh-template-crossrefs-"))
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify({}))
  writeFileSync(path.join(root, "matter.json"), JSON.stringify({ id: "template-crossrefs" }))
  const server = Bun.serve({ port: 0, async fetch(request) {
    if (new URL(request.url).pathname === "/templates/content") return new Response(new Uint8Array(template))
    const file = (await request.formData()).get("file") as File
    writeFileSync(path.join(root, file.name), new Uint8Array(await file.arrayBuffer()))
    return Response.json({})
  } })
  try {
    // Isolate import-time URLs so this test cannot change another test's tool
    // transport or contact a configured stack.
    const child = Bun.spawn([process.execPath, "-e", `
      import draft from ${JSON.stringify(path.resolve(import.meta.dir, "../tool/draft-document.ts"))};
      const directory = process.env.WORKSPACE_ROOT;
      const result = await draft.execute({name:"prepared.docx",template:"services.docx",omit:["supplier services"]}, {
        directory,worktree:directory,sessionID:"template-crossrefs",messageID:"template-crossrefs",agent:"draft",
        abort:new AbortController().signal,metadata() {},async ask() {}
      });
      if (typeof result === "string") throw new Error(result);
      console.log(JSON.stringify(result));
    `], { env: { ...process.env, WORKSPACE_ROOT: root, INGEST_URL: server.url.href.replace(/\/$/, "") }, stdout: "pipe", stderr: "pipe" })
    const [output, error] = await Promise.all([new Response(child.stdout).text(), new Response(child.stderr).text()])
    expect(await child.exited, error).toBe(0)
    const result = JSON.parse(output.trim().split("\n").at(-1)!) as { output: string; metadata: { removedReferences: string[] } }
    const marker = "[open item: refers to 2.1, which was removed]"
    expect(result.output).toContain(`2 cross-references to parts no longer in the draft are marked as open items: ${marker}.`)
    expect(result.metadata.removedReferences).toEqual([marker, marker])
    // The clause now numbered 2.1 does not cite itself.
    const opened = await openRevision(new Uint8Array(readFileSync(path.join(root, "prepared.docx"))))
    try {
      const shown = opened.paragraphs.map(row => `${row.metadata.list?.generatedLabel?.trim() ?? ""} ${row.match.text}`.trim())
      expect(shown.join("\n")).not.toContain("The Supplier shall provide")
      expect(shown).toEqual(expect.arrayContaining([
        `1.1 In this Agreement the terms in clause ${marker} and Schedule 1 apply.`,
        `2.1 The Customer shall pay as set out in clause ${marker}.`,
      ]))
    } finally {
      opened.session.close()
    }
  } finally {
    server.stop(true)
    rmSync(root, { recursive: true, force: true })
  }
}, 120_000)
