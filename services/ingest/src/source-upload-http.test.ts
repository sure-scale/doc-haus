import { expect, test } from "bun:test"
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { getDocumentRole, listSourceDocuments, openDb } from "./db"

// A same-name upload never replaces an uploaded source silently: the lawyer's
// own Agreement.docx survives the counterparty's returned Agreement.docx unless
// the caller chooses replace, and keep-both stores the next free name.
test("a same-name source upload is refused unless the caller chooses replace or keep both", async () => {
  const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-source-upload-http-"))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
  // Uploads need no generation engine. Any accidental engine request is observable.
  const engineRequests: string[] = []
  const engine = Bun.serve({ port: 0, fetch(request) {
    engineRequests.push(request.url)
    return new Response("No engine in this upload test", { status: 503 })
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
  const request = (route: string, options?: RequestInit) => fetch(base + route, {
    ...options, signal: AbortSignal.timeout(30_000),
  })
  const docx = async (marker: string) => new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: [
    new Paragraph("1. Scope"), new Paragraph(`${marker} describes the invented scope.`),
  ] }] })))
  try {
    const deadline = Date.now() + 20_000
    while (true) {
      if (await request("/matters").then(response => response.ok).catch(() => false)) break
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
      await Bun.sleep(50)
    }
    const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ title: "Same-name uploads" }) })
    expect(created.status).toBe(200)
    const matter = await created.json() as { id: string; dir: string }
    const upload = (name: string, bytes: Uint8Array<ArrayBuffer>, existing?: string) => {
      const form = new FormData()
      form.append("file", new File([bytes], name))
      if (existing) form.append("existing", existing)
      return request(`/matters/${matter.id}/documents`, { method: "POST", body: form })
    }
    const stored = async (name: string) => new Uint8Array(await readFile(path.join(matter.dir, name)))
    const sources = () => {
      const db = openDb(matter.dir)
      try {
        return (listSourceDocuments(db) as { name: string }[]).map(row => row.name).toSorted()
      } finally { db.close() }
    }

    const ours = await docx("Our marker")
    const theirs = await docx("Their marker")
    const first = await upload("Agreement.docx", ours)
    expect(first.status).toBe(200)
    expect((await first.json() as { name: string }).name).toBe("Agreement.docx")

    // Without a choice the clash is refused with the existing name, and our file stays.
    const refused = await upload("Agreement.docx", theirs)
    expect(refused.status).toBe(409)
    expect(await refused.json()).toEqual({ error: expect.any(String), code: "source_exists", name: "Agreement.docx" })
    expect(await stored("Agreement.docx")).toEqual(ours)
    // Names match case-insensitively where the filesystem does.
    if (process.platform === "darwin" || process.platform === "win32") {
      const differentCase = await upload("AGREEMENT.DOCX", theirs)
      expect(differentCase.status).toBe(409)
      expect(await differentCase.json()).toMatchObject({ code: "source_exists", name: "Agreement.docx" })
      expect(await stored("Agreement.docx")).toEqual(ours)
    }
    expect((await upload("Agreement.docx", theirs, "overwrite")).status).toBe(400)

    // Keep both stores the next free name and returns it; the original is untouched.
    const kept = await upload("Agreement.docx", theirs, "keep-both")
    expect(kept.status).toBe(200)
    expect((await kept.json() as { name: string }).name).toBe("Agreement (2).docx")
    expect(await stored("Agreement (2).docx")).toEqual(theirs)
    expect(await stored("Agreement.docx")).toEqual(ours)
    const third = await upload("Agreement.docx", await docx("Third marker"), "keep-both")
    expect((await third.json() as { name: string }).name).toBe("Agreement (3).docx")
    // A free name is free case-insensitively on every platform.
    const lower = await upload("agreement.docx", await docx("Lower marker"), "keep-both")
    expect((await lower.json() as { name: string }).name).toBe("agreement (4).docx")
    expect(sources()).toEqual(["Agreement (2).docx", "Agreement (3).docx", "Agreement.docx", "agreement (4).docx"].toSorted())

    // Replace is the explicit choice to overwrite.
    const replaced = await upload("Agreement.docx", theirs, "replace")
    expect(replaced.status).toBe(200)
    expect((await replaced.json() as { name: string }).name).toBe("Agreement.docx")
    expect(await stored("Agreement.docx")).toEqual(theirs)

    // A generated deliverable is never replaced by a source, whatever the choice;
    // keep both stores the source beside it.
    const deliverable = await docx("Drafted marker")
    const form = new FormData()
    form.append("file", new File([deliverable], "Draft.docx"))
    expect((await request(`/matters/${matter.id}/deliverables`, { method: "POST", body: form })).status).toBe(200)
    const delivered = await stored("Draft.docx")
    for (const existing of [undefined, "replace"]) {
      const clash = await upload("Draft.docx", theirs, existing)
      expect(clash.status).toBe(409)
      expect((await clash.json() as { error: string; code?: string }))
        .toEqual({ error: "Cannot replace generated deliverable Draft.docx with a source document." })
    }
    const beside = await upload("Draft.docx", theirs, "keep-both")
    expect((await beside.json() as { name: string }).name).toBe("Draft (2).docx")
    expect(await stored("Draft.docx")).toEqual(delivered)
    const db = openDb(matter.dir)
    try {
      expect(getDocumentRole(db, path.join(matter.dir, "Draft.docx"))).toBe("deliverable")
      expect(getDocumentRole(db, path.join(matter.dir, "Draft (2).docx"))).toBe("source")
    } finally { db.close() }
    expect(engineRequests).toEqual([])
  } finally {
    service.kill()
    await service.exited
    engine.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}, 90_000)
