import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import mammoth from "mammoth"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { getRedline, listSourceDocuments, openDb } from "./db"
import { docxodus } from "./docxodus"
import { readDocxCommentParagraphs, readDocxComments } from "../../../dochaus/lib/docx-comments"

test("HTTP redline decisions preserve rejected bytes and publish an accepted, reindexed source", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-redline-http-"))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  // No generation engine is needed. Any accidental engine request is observable.
  const engineRequests: string[] = []
  const engine = Bun.serve({ port: 0, fetch(request) {
    engineRequests.push(request.url)
    return new Response("No engine in this document-control test", { status: 503 })
  } })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const serviceLog = path.join(temporary, "ingest.log")
  const service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}` },
    stdout: Bun.file(serviceLog), stderr: Bun.file(serviceLog),
  })
  const base = `http://127.0.0.1:${port}`
  const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
  const request = (route: string, options?: RequestInit) => fetch(base + route, {
    ...options, signal: AbortSignal.timeout(30_000),
  })
  try {
    const deadline = Date.now() + 20_000
    while (true) {
      if (await request("/matters").then(response => response.ok).catch(() => false)) break
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
      await Bun.sleep(50)
    }
    const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ title: "Isolated document decisions" }) })
    expect(created.status).toBe(200)
    const matter = await created.json() as { id: string; dir: string }
    expect(path.dirname(matter.dir)).toBe(workspace)
    const name = "source.docx"
    const documentPath = path.join(matter.dir, name)
    const original = await Packer.toBuffer(new Document({ sections: [{ children: [
      new Paragraph("1. Document wording"), new Paragraph("Original marker describes the document."),
      new Paragraph("2. Review note"), new Paragraph("Second marker remains unchanged."),
    ] }] }))
    const form = new FormData()
    form.append("file", new File([new Uint8Array(original)], name))
    const uploaded = await request(`/matters/${matter.id}/documents`, { method: "POST", body: form })
    expect(uploaded.status).toBe(200)
    const before = openDb(matter.dir)
    const originalRegistry = before.query("SELECT doc_path, name, role FROM documents WHERE role = 'source'").all()
    before.close()
    const proposal = (findText: string, newText: string, anchorId: string) => recordRedlineProposal(matter.dir, {
      docPath: documentPath, docName: name, scope: "phrase", findText, oldText: findText, newText,
      author: "Test reviewer", reviewComment: "Review this proposed wording.", anchorId,
    }).id
    const acceptedID = proposal("Original marker", "Accepted marker", "first-paragraph")
    const rejectedID = proposal("Second marker", "Unwanted marker", "second-paragraph")
    const redlines = `/matters/${matter.id}/redlines`
    const pending = await (await request(`${redlines}?name=${name}`)).json() as Array<{ id: number; status: string }>
    expect(pending.map(row => ({ id: row.id, status: row.status }))).toEqual([
      { id: acceptedID, status: "pending" }, { id: rejectedID, status: "pending" },
    ])
    const dx = await docxodus()
    const comparison = await request(`/matters/${matter.id}/documents/redlined?name=${name}`)
    expect(comparison.status).toBe(200)
    const exported = new Uint8Array(await comparison.arrayBuffer())
    expect((await dx.getRevisions(exported)).length).toBeGreaterThan(0)
    expect(readDocxComments(exported).map(row => ({ text: row.text, paragraph: row.markers[0]?.accepted_paragraph_text }))).toEqual([
      { text: "Review this proposed wording.", paragraph: "Accepted marker describes the document." },
      { text: "Review this proposed wording.", paragraph: "Unwanted marker remains unchanged." },
    ])
    expect(hash(await Bun.file(documentPath).bytes())).toBe(hash(original))
    const previewDb = openDb(matter.dir)
    try {
      expect([acceptedID, rejectedID].map(id => getRedline(previewDb, id)?.status)).toEqual(["pending", "pending"])
      previewDb.run("UPDATE redlines SET find_text = ? WHERE id = ?", ["Missing exact proposal target", rejectedID])
      const conflict = await request(`/matters/${matter.id}/documents/redlined?name=${name}`)
      expect(conflict.status).toBe(409)
      expect((await conflict.json() as { error: string }).error).toContain("missing or ambiguous")
      previewDb.run("UPDATE redlines SET find_text = ? WHERE id = ?", ["Second marker", rejectedID])
    } finally { previewDb.close() }
    expect(hash(await Bun.file(documentPath).bytes())).toBe(hash(original))
    expect((await request(`${redlines}/${rejectedID}/reject`, { method: "POST" })).status).toBe(200)
    expect(hash(await Bun.file(documentPath).bytes())).toBe(hash(original))
    expect((await request(`${redlines}/${acceptedID}/accept`, { method: "POST" })).status).toBe(200)
    const accepted = await Bun.file(documentPath).bytes()
    expect(hash(accepted)).not.toBe(hash(original))
    expect((await dx.getRevisions(accepted)).length).toBe(0)
    const expectedText = "1. Document wording Accepted marker describes the document. 2. Review note Second marker remains unchanged."
    expect((await mammoth.extractRawText({ buffer: Buffer.from(accepted) })).value.trim().replace(/\s+/g, " ")).toBe(expectedText)
    const after = openDb(matter.dir)
    try {
      expect(getRedline(after, acceptedID)?.status).toBe("accepted")
      expect(getRedline(after, rejectedID)?.status).toBe("rejected")
      // Reindexing replaces version rows; the registered document identity and role persist.
      expect(listSourceDocuments(after)).toHaveLength(1)
      expect(listSourceDocuments(after)).toMatchObject(originalRegistry)
      expect(after.query("SELECT text FROM chunks_fts WHERE chunks_fts MATCH 'Accepted'").all().length).toBeGreaterThan(0)
      expect(after.query("SELECT text FROM chunks_fts WHERE chunks_fts MATCH 'Original OR Unwanted'").all()).toEqual([])
      expect((after.query("SELECT text FROM chunks ORDER BY chunk_index").all() as Array<{ text: string }>).map(row => row.text).join(" ")).toContain("Second marker remains unchanged.")
    } finally {
      after.close()
    }
    const sourceIndex = await Bun.file(path.join(matter.dir, "source-index.json")).json() as {
      sources: Array<{ document: string; extracted_text: string }>
    }
    expect(sourceIndex.sources.map(source => source.document)).toEqual([name])
    expect((await Bun.file(path.join(matter.dir, sourceIndex.sources[0].extracted_text)).text()).trim().replace(/\s+/g, " ")).toBe(expectedText)
    const download = await request(`/matters/${matter.id}/documents/content?name=${name}`)
    expect(download.status).toBe(200)
    expect(download.headers.get("content-disposition")).toContain('filename="source.docx"')
    expect(hash(new Uint8Array(await download.arrayBuffer()))).toBe(hash(accepted))
    // Generated titles can contain Unicode. Both the header and bytes must
    // survive the same download endpoint the browser uses.
    const unicodeName = "Révision — 商業.docx"
    const generated = new FormData()
    generated.append("file", new File([new Uint8Array(original)], unicodeName))
    expect((await request(`/matters/${matter.id}/deliverables`, { method: "POST", body: generated })).status).toBe(200)
    const unicodeDownload = await request(`/matters/${matter.id}/documents/content?name=${encodeURIComponent(unicodeName)}`)
    expect(unicodeDownload.status).toBe(200)
    expect(unicodeDownload.headers.get("content-disposition")).toContain(`filename*=UTF-8''${encodeURIComponent(unicodeName)}`)
    const generatedBytes = new Uint8Array(await unicodeDownload.arrayBuffer())
    expect(hash(generatedBytes)).toBe(hash(await Bun.file(path.join(matter.dir, unicodeName)).bytes()))
    expect(readDocxCommentParagraphs(generatedBytes).map(row => row.paragraph_text)).toEqual(readDocxCommentParagraphs(original).map(row => row.paragraph_text))
    expect(readDocxCommentParagraphs(generatedBytes).every(row => row.pids.length === 1)).toBe(true)
    expect(await (await request(`${redlines}?name=${name}`)).json()).toEqual([])
    for (const [id, action] of [[acceptedID, "accept"], [acceptedID, "reject"], [rejectedID, "reject"], [rejectedID, "accept"]])
      expect((await request(`${redlines}/${id}/${action}`, { method: "POST" })).status).toBe(404)
    expect(hash(await Bun.file(documentPath).bytes())).toBe(hash(accepted))
    expect(engineRequests).toEqual([])
    console.log("Isolated HTTP accept/reject verified with real local embeddings; no generation engine requests.")
  } finally {
    service.kill()
    await service.exited
    engine.stop(true)
    const evidence = process.env.DOCHAUS_REDLINE_TEST_EVIDENCE
    if (evidence) {
      await mkdir(evidence, { recursive: true })
      await writeFile(path.join(evidence, "isolated-ingest-redline.log"), await readFile(serviceLog))
    }
    await rm(temporary, { recursive: true, force: true })
  }
}, 90_000)
