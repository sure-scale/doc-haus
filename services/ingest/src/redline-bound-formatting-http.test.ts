import { expect, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdir, mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { createDocxRevision } from "../../../dochaus/lib/docx-amendments"
import { paragraphTargets, withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { getRedline, openDb, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"

type ListedProposal = RedlineRow & { stale?: true; stale_reason?: string }
const refusal = "This proposal has an invalid Word revision binding. Prepare a fresh change list."

test("HTTP bound paragraph proposals refuse unsupported formatting without accepting or publishing", async () => {
  const temporary = await realpath(await mkdtemp(path.join(tmpdir(), "dochaus-bound-format-http-")))
  const workspace = path.join(temporary, "workspace"), config = path.join(temporary, "config")
  await mkdir(workspace); await mkdir(config); await writeFile(path.join(config, "opencode.json"), "{}\n")
  const engineRequests: string[] = []
  const engine = Bun.serve({ port: 0, fetch(request) {
    engineRequests.push(request.url)
    return new Response("No generation engine in this document-control test", { status: 503 })
  } })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() }), port = probe.port; probe.stop(true)
  const log = path.join(temporary, "ingest.log")
  const service = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config,
      INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "" },
    stdout: Bun.file(log), stderr: Bun.file(log),
  })
  const request = (route: string, options?: RequestInit) => fetch(`http://127.0.0.1:${port}${route}`, { ...options, signal: AbortSignal.timeout(30_000) })
  try {
    const deadline = Date.now() + 20_000
    while (!await request("/matters").then(response => response.ok).catch(() => false)) {
      if (service.exitCode !== null || Date.now() >= deadline) throw new Error(`Isolated ingest startup failed:\n${await readFile(log, "utf8")}`)
      await Bun.sleep(50)
    }
    const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Synthetic bound formatting review" }) })
    expect(created.status).toBe(200)
    const matter = await created.json() as { id: string; dir: string }, route = `/matters/${matter.id}`
    const source = withParagraphPointers((await renderMarkdownToDocx("First paragraph.\n\nSecond paragraph.")).bytes)
    const versions = (name: string) => {
      const db = openDb(matter.dir)
      try {
        return {
          current: db.query("SELECT current_version FROM versioned_documents WHERE name=?").get(name),
          kept: db.query("SELECT v.* FROM document_versions v JOIN versioned_documents d ON d.id=v.document_id WHERE d.name=? ORDER BY v.number").all(name),
          events: db.query("SELECT e.* FROM document_version_events e JOIN versioned_documents d ON d.id=e.document_id WHERE d.name=? ORDER BY e.id").all(name),
        }
      } finally { db.close() }
    }
    for (const mode of ["single", "all", "unsupported"] as const) {
      const name = `${mode}.docx`, form = new FormData(); form.append("file", new File([new Uint8Array(source).buffer], name))
      expect((await request(`${route}/documents`, { method: "POST", body: form })).status).toBe(200)
      const file = resolveMatterDocumentTarget(matter.dir, name).docPath, original = await Bun.file(file).bytes()
      const target = (await paragraphTargets(original))[1]!
      expect(target.pid).not.toBeNull()
      const id = recordRedlineProposal(matter.dir, { docPath: file, docName: name, scope: "clause",
        findText: "Second paragraph.", oldText: "Second paragraph.", newText: "Updated paragraph.", author: "Synthetic Reviewer",
        reviewComment: "Review the requested wording.", anchorId: target.pid! }).id
      const binding = JSON.stringify({ version: 2, pid: target.pid, paragraph_sha256: target.sha256, occurrence: 0,
        base_sha256: createHash("sha256").update(original).digest("hex"), amendment_id: String(id) })
      const db = openDb(matter.dir)
      try { db.run("UPDATE redlines SET target_json=? WHERE id=?", [binding, id]) } finally { db.close() }
      const list = async () => {
        const response = await request(`${route}/redlines?name=${name}`)
        expect(response.status).toBe(200)
        return await response.json() as ListedProposal[]
      }
      expect((await list()).map(row => [row.id, row.status, row.stale, row.stale_reason])).toEqual([[id, "pending", undefined, undefined]])
      const preview = await request(`${route}/documents/redlined?name=${name}`)
      expect(preview.status).toBe(200)
      expect((await createDocxRevision(new Uint8Array(await preview.arrayBuffer()), { pendingRevisions: "layer" })).paragraphs.filter(row => row.scope === "body").map(row => row.text)).toEqual(["First paragraph.", "Updated paragraph."])
      const before = versions(name)
      expect(before.kept).toHaveLength(1)
      if (mode !== "unsupported") {
        const accepted = await request(mode === "single" ? `${route}/redlines/${id}/accept` : `${route}/redlines/accept-all?name=${name}`, { method: "POST" })
        expect(accepted.status).toBe(200)
        expect((await createDocxRevision(await Bun.file(file).bytes())).paragraphs.filter(row => row.scope === "body").map(row => row.text)).toEqual(["First paragraph.", "Updated paragraph."])
        expect(versions(name).kept).toHaveLength(2)
        expect(await list()).toEqual([])
        const check = openDb(matter.dir)
        try { expect(getRedline(check, id)!.status).toBe("accepted") } finally { check.close() }
        continue
      }
      // The exact same bound target and source remain valid. Only the stored
      // formatting column changes, including malformed and empty non-null values.
      for (const formatting of [JSON.stringify({ spacingAfter: 6 }), "", "null", "{invalid"]) {
        const update = openDb(matter.dir)
        try { update.run("UPDATE redlines SET formatting_json=? WHERE id=?", [formatting, id]) } finally { update.close() }
        const listed = await list()
        expect(listed.map(row => [row.id, row.status, row.stale, row.stale_reason])).toEqual([[id, "pending", true, refusal]])
        expect(listed[0]!.target_json).toBe(binding)
        // The preview route intentionally omits stale proposals; it must show
        // the unchanged saved file, rather than a wording-only partial preview.
        const savedOnly = await request(`${route}/documents/redlined?name=${name}`)
        expect(savedOnly.status).toBe(200)
        expect(new Uint8Array(await savedOnly.arrayBuffer())).toEqual(original)
        for (const endpoint of [`${route}/redlines/${id}/accept`, `${route}/redlines/accept-all?name=${name}`]) {
          const response = await request(endpoint, { method: "POST" })
          expect(response.status).toBe(409)
          expect((await response.json() as { error: string }).error).toContain(refusal)
          expect(await Bun.file(file).bytes()).toEqual(original)
          expect(versions(name)).toEqual(before)
          const check = openDb(matter.dir)
          try { expect(getRedline(check, id)!).toMatchObject({ status: "pending", target_json: binding, formatting_json: formatting }) } finally { check.close() }
        }
      }
      const reset = openDb(matter.dir)
      try { reset.run("UPDATE redlines SET formatting_json=NULL WHERE id=?", [id]) } finally { reset.close() }
      expect((await list()).map(row => [row.id, row.stale, row.stale_reason])).toEqual([[id, undefined, undefined]])
      expect(await Bun.file(file).bytes()).toEqual(original)
      expect(versions(name)).toEqual(before)
    }
    expect(engineRequests).toEqual([])
  } finally {
    service.kill(); await service.exited; engine.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 120_000)
