import { expect, test } from "bun:test"
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { CommentRangeStart, CommentRangeEnd, CommentReference, InsertedTextRun, Document, Packer, Paragraph, TextRun } from "docx"
import { documentCopyRoutes } from "./document-copy"
import { listDocuments, listSourceDocuments, openDb, upsertDocument, pendingRedlineCounts } from "./db"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { legalTools } from "../../../dochaus/tool/registry"

async function fixture() {
  const directory = realpathSync(mkdtempSync(path.join(tmpdir(), "dochaus-copy-")))
  const bytes = await Packer.toBuffer(new Document({
    comments: { children: [{ id: 0, author: "Reviewer", children: [new Paragraph("Existing note")] }] },
    sections: [{ children: [new Paragraph({ children: [
      new CommentRangeStart(0), new TextRun({ text: "Keep the original wording.", bold: true }),
      new CommentRangeEnd(0), new TextRun({ children: [new CommentReference(0)] }),
      new InsertedTextRun({ id: 7, author: "Reviewer", date: "2020-01-01T00:00:00Z", text: "Existing revision" }),
    ] })] }],
  }))
  const original = path.join(directory, "original.docx")
  writeFileSync(original, bytes)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "copy-matter" }))
  const db = openDb(directory)
  upsertDocument(db, original, "original.docx", Date.now(), null)
  db.close()
  const app = documentCopyRoutes(id => { expect(id).toBe("copy-matter"); return directory })
  const request = (source: string, name: string) => app.request("/matters/copy-matter/documents/copy", {
    method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ source, name }),
  })
  return { directory, bytes, original, app, request }
}

test("exact canonical copy registers a deliverable; source proposals stay on the unchanged original", async () => {
  const f = await fixture()
  try {
    const proposal = recordRedlineProposal(f.directory, { docPath: f.original, docName: "original.docx", scope: "phrase", anchorId: "p:body:original", findText: "original", oldText: "original", newText: "proposed", author: "Reviewer", reviewComment: "Original proposal stays here." })
    const response = await f.request("original.docx", "negotiation.docx")
    expect(response.status).toBe(200)
    expect(await response.json()).toMatchObject({ name: "negotiation.docx", role: "deliverable", source: "original.docx", pendingProposalsExcluded: 1 })
    expect(readFileSync(path.join(f.directory, "negotiation.docx"))).toEqual(Buffer.from(f.bytes))
    expect(readFileSync(f.original)).toEqual(Buffer.from(f.bytes))
    const db = openDb(f.directory)
    expect(listDocuments(db)).toHaveLength(2)
    expect(listSourceDocuments(db)).toEqual([expect.objectContaining({ name: "original.docx" })])
    expect(db.query("SELECT COUNT(*) AS n FROM chunks").get()).toEqual({ n: 0 })
    expect(db.query("SELECT id, doc_path, status FROM redlines").all()).toEqual([{ id: proposal.id, doc_path: f.original, status: "pending" }])
    db.close()
    const tracked = (await import("../../../dochaus/tool/tracked-changes")).default
    const result = await tracked.execute({ document: "negotiation.docx", find: "original", replace: "revised", comment: "Requested separate-copy edit.", author: "Reviewer" }, {
      callID: "copy-call", sessionID: "copy-test", messageID: "copy-test", agent: "redliner", directory: f.directory, worktree: f.directory,
      abort: new AbortController().signal, metadata() {}, async ask(request) { expect(request.patterns).toEqual([path.join(f.directory, "negotiation.docx")]) },
    })
    expect(result).toBeObject()
    const after = openDb(f.directory)
    expect(Object.values(pendingRedlineCounts(after)).sort()).toEqual([1, 1])
    after.close()
    expect(readFileSync(f.original)).toEqual(Buffer.from(f.bytes))
  } finally { rmSync(f.directory, { recursive: true, force: true }) }
})

test("valid copy then unsafe paths, aliases, absent registration and existing destinations are rejected", async () => {
  const f = await fixture()
  try {
    expect((await f.request("original.docx", "valid.docx")).status).toBe(200)
    for (const name of ["../escape.docx", "/tmp/escape.docx", "nested/name.docx", "nested\\name.docx", "file.pdf", "bad\0.docx"])
      expect((await f.request("original.docx", name)).status).toBe(400)
    expect((await f.request("../original.docx", "other.docx")).status).toBe(400)
    expect((await f.request("original.docx", "original.docx")).status).toBe(409)
    expect((await f.request("original.docx", "valid.docx")).status).toBe(409)
    if (process.platform === "darwin" || process.platform === "win32")
      expect((await f.request("original.docx", "VALID.DOCX")).status).toBe(409)
    writeFileSync(path.join(f.directory, "unregistered.docx"), f.bytes)
    expect((await f.request("unregistered.docx", "other.docx")).status).toBe(409)
    symlinkSync(f.original, path.join(f.directory, "alias.docx"))
    expect((await f.request("alias.docx", "other.docx")).status).toBe(409)
    expect((await f.request("original.docx", "alias.docx")).status).toBe(409)
    symlinkSync(path.join(tmpdir(), "outside-copy.docx"), path.join(f.directory, "outside.docx"))
    expect((await f.request("original.docx", "outside.docx")).status).toBe(409)
    const db = openDb(f.directory)
    upsertDocument(db, path.join(f.directory, "missing.docx"), "missing.docx", Date.now(), null)
    db.close()
    expect((await f.request("original.docx", "missing.docx")).status).toBe(409)
    expect((await f.request("missing.docx", "other.docx")).status).toBe(409)
    expect(existsSync(path.join(f.directory, "other.docx"))).toBe(false)
    expect(readFileSync(f.original)).toEqual(Buffer.from(f.bytes))
  } finally { rmSync(f.directory, { recursive: true, force: true }) }
})

test("registry failure removes new file; competing same-name creates never overwrite", async () => {
  const f = await fixture()
  try {
    expect((await f.request("original.docx", "valid.docx")).status).toBe(200)
    const db = openDb(f.directory)
    db.run("CREATE TRIGGER refuse_copy BEFORE INSERT ON documents WHEN NEW.name = 'blocked.docx' BEGIN SELECT RAISE(ABORT, 'isolated registry failure'); END")
    f.app.onError(() => new Response("Registry write failed", { status: 500 }))
    expect((await f.request("original.docx", "blocked.docx")).status).toBe(500)
    expect(existsSync(path.join(f.directory, "blocked.docx"))).toBe(false)
    expect(listDocuments(db)).toHaveLength(2)
    db.close()
    const results = await Promise.all([f.request("original.docx", "concurrent.docx"), f.request("original.docx", "concurrent.docx")])
    expect(results.map(response => response.status).sort()).toEqual([200, 409])
    expect(readFileSync(path.join(f.directory, "concurrent.docx"))).toEqual(Buffer.from(f.bytes))
    expect(readFileSync(f.original)).toEqual(Buffer.from(f.bytes))
  } finally { rmSync(f.directory, { recursive: true, force: true }) }
})

test("native copy tool asks before HTTP and returns the registered new document name", async () => {
  const f = await fixture()
  let calls = 0
  const server = Bun.serve({ port: 0, fetch(request) { calls++; return f.app.fetch(request) } })
  const prior = process.env.INGEST_URL
  process.env.INGEST_URL = `http://127.0.0.1:${server.port}`
  try {
    const copy = (await import("../../../dochaus/tool/copy-document")).default
    expect(legalTools.some(([name]) => name === "copy-document")).toBe(true)
    const agent = readFileSync(path.resolve(import.meta.dir, "../../../dochaus/agent/redliner.md"), "utf8")
    expect(agent).toContain("  copy-document: true")
    expect(agent).toContain("  copy-document: ask")
    const configuration = JSON.parse(readFileSync(path.resolve(import.meta.dir, "../../../dochaus/opencode.json"), "utf8"))
    expect(configuration.permission["copy-document"]).toBe("ask")
    const context = { callID: "copy-call", sessionID: "copy-test", messageID: "copy-test", agent: "redliner", directory: f.directory, worktree: f.directory, abort: new AbortController().signal, metadata() {},
      async ask(request: { permission: string; patterns: string[] }) { expect(request.permission).toBe("copy-document"); expect(request.patterns).toEqual([path.join(f.directory, "new.docx")]); throw new Error("Permission denied") },
    }
    await expect(copy.execute({ document: "original.docx", name: "new.docx" }, context)).rejects.toThrow("Permission denied")
    expect(calls).toBe(0)
    expect(existsSync(path.join(f.directory, "new.docx"))).toBe(false)
    const result = await copy.execute({ document: "original.docx", name: "new.docx" }, { ...context, async ask(request) { expect(request.permission).toBe("copy-document"); expect(request.always).toBeUndefined() } })
    expect(result).toMatchObject({ title: "Created new.docx", metadata: { name: "new.docx", document: path.join(f.directory, "new.docx") } })
    expect(calls).toBe(1)
    expect(readFileSync(f.original)).toEqual(Buffer.from(f.bytes))
    expect(await copy.execute({ document: "../original.docx", name: "other.docx" }, context)).toBeString()
    expect(calls).toBe(1)
  } finally {
    if (prior === undefined) delete process.env.INGEST_URL; else process.env.INGEST_URL = prior
    server.stop(true)
    rmSync(f.directory, { recursive: true, force: true })
  }
})
