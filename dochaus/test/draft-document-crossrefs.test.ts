// Run from dochaus/: bun test test/draft-document-crossrefs.test.ts
import { expect, test } from "bun:test"
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { renderMarkdownToDocx } from "../lib/render"

// Neutral numbered text citing its own clause numbers: advice discussing an
// agreement's clauses, and a template whose references are typed.
const ADVICE = ["# Advice on termination", "## 1. Background", "1.1 You asked whether the Supplier may terminate early.", "1.2 We reviewed the Agreement.",
  "## 2. Analysis", "2.1 The Agreement permits termination under clause 1.2.", "2.2 Clause 2.1 requires notice in writing."].join("\n\n")
const TEMPLATE = ["# Services Agreement", "## 1. Definitions", "1.1 In this Agreement the terms in clause 2.1 apply.",
  "## 2. Services", "2.1 The Supplier shall provide the Services.", "2.2 The Customer shall pay as set out in clause 2.1."].join("\n\n")

test("draft-document knows no document form, so neither a from-scratch draft nor a template draft gets new cross-reference fields", async () => {
  const template = (await renderMarkdownToDocx(TEMPLATE)).bytes
  const root = mkdtempSync(path.join(tmpdir(), "dh-draft-crossrefs-"))
  mkdirSync(path.join(root, ".preferences"))
  writeFileSync(path.join(root, ".preferences/preferences.json"), JSON.stringify({}))
  writeFileSync(path.join(root, "matter.json"), JSON.stringify({ id: "draft-crossrefs" }))
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
      const context = { directory, worktree: directory, sessionID: "draft-crossrefs", messageID: "draft-crossrefs", agent: "draft",
        abort: new AbortController().signal, metadata() {}, async ask() {} };
      for (const args of [{ name: "advice.docx", content: ${JSON.stringify(ADVICE)} }, { name: "services.docx", template: "services.docx" }]) {
        const result = await draft.execute(args, context);
        if (typeof result === "string") throw new Error(result);
      }
    `], { env: { ...process.env, WORKSPACE_ROOT: root, INGEST_URL: server.url.href.replace(/\/$/, "") }, stdout: "pipe", stderr: "pipe" })
    const error = await new Response(child.stderr).text()
    expect(await child.exited, error).toBe(0)
    for (const name of ["advice.docx", "services.docx"]) {
      const body = strFromU8(unzipSync(readFileSync(path.join(root, name)))["word/document.xml"]!)
      // Premise: the clauses are Word list numbered, so references could be linked.
      expect(body).toContain("<w:numPr>")
      expect(body).not.toContain("<w:instrText")
    }
  } finally {
    server.stop(true)
    rmSync(root, { recursive: true, force: true })
  }
}, 120_000)
