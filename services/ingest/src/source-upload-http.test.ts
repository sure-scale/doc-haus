import { expect, test } from "bun:test"
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises"
import { tmpdir } from "node:os"
import path from "node:path"
import { Document, Packer, Paragraph } from "docx"
import { strToU8, zipSync } from "fflate"
import { getDocumentRole, listSourceDocuments, openDb } from "./db"
import { MAX_ZIP_FILES } from "./zip-archive"

// A same-name upload never replaces an uploaded source silently: the lawyer's
// own Agreement.docx survives the counterparty's returned Agreement.docx unless
// the caller chooses replace, and keep-both stores the next free name.
test("a same-name source upload is refused unless the caller chooses replace or keep both", async () => {
  await withMatter("Same-name uploads", async ({ matter, upload, deliver, stored, role }) => {
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
    // the refusal names it so the caller can keep both, which stores the source beside it.
    const delivered = await deliver("Draft.docx", await docx("Drafted marker"))
    for (const existing of [undefined, "replace"]) {
      const clash = await upload("Draft.docx", theirs, existing)
      expect(clash.status).toBe(409)
      expect(await clash.json()).toEqual({
        error: "Cannot replace generated deliverable Draft.docx with a source document.",
        code: "deliverable_exists",
        name: "Draft.docx",
      })
    }
    const beside = await upload("Draft.docx", theirs, "keep-both")
    expect((await beside.json() as { name: string }).name).toBe("Draft (2).docx")
    expect(await stored("Draft.docx")).toEqual(delivered)
    expect(role("Draft.docx")).toBe("deliverable")
    expect(role("Draft (2).docx")).toBe("source")
  })
}, 90_000)

// A file already in the matter folder but never registered (copied in by hand,
// say) is refused like a source, so no existing file is overwritten silently.
test("a same-name upload is refused when the file is on disk but not registered", async () => {
  await withMatter("Unregistered files", async ({ matter, upload, stored, role }) => {
    const loose = await docx("Loose marker")
    const theirs = await docx("Their marker")
    await writeFile(path.join(matter.dir, "Agreement.docx"), loose)
    expect(role("Agreement.docx")).toBeNull()

    const refused = await upload("Agreement.docx", theirs)
    expect(refused.status).toBe(409)
    expect(await refused.json()).toEqual({ error: expect.any(String), code: "source_exists", name: "Agreement.docx" })
    if (process.platform === "darwin" || process.platform === "win32")
      expect(await (await upload("AGREEMENT.DOCX", theirs)).json()).toMatchObject({ code: "source_exists", name: "Agreement.docx" })
    expect(await stored("Agreement.docx")).toEqual(loose)

    const kept = await upload("Agreement.docx", theirs, "keep-both")
    expect((await kept.json() as { name: string }).name).toBe("Agreement (2).docx")
    expect(await stored("Agreement.docx")).toEqual(loose)
    expect((await upload("Agreement.docx", theirs, "replace")).status).toBe(200)
    expect(await stored("Agreement.docx")).toEqual(theirs)
    expect(role("Agreement.docx")).toBe("source")
  })
}, 90_000)

// macOS Finder can name a file in decomposed Unicode (NFD) while the matter holds
// the composed (NFC) spelling; APFS treats both as one file, so the upload is
// refused without a choice, and keep both stores it under a genuinely free name.
test.skipIf(process.platform !== "darwin")("a decomposed upload name clashes with its composed source", async () => {
  await withMatter("Unicode names", async ({ matter, upload, stored }) => {
    const composed = "Caf\u00e9.docx"
    const decomposed = "Cafe\u0301.docx"
    const ours = await docx("Our marker")
    const theirs = await docx("Their marker")
    expect((await upload(composed, ours)).status).toBe(200)

    const refused = await upload(decomposed, theirs)
    expect(refused.status).toBe(409)
    expect(await refused.json()).toMatchObject({ code: "source_exists", name: composed })

    const kept = await upload(decomposed, theirs, "keep-both")
    const keptName = (await kept.json() as { name: string }).name
    expect(keptName.normalize("NFC")).toBe("Caf\u00e9 (2).docx")
    expect(kept.status).toBe(200)
    expect(await stored(keptName)).toEqual(theirs)
    expect(await stored(composed)).toEqual(ours)
    const db = openDb(matter.dir)
    try {
      expect((listSourceDocuments(db) as { name: string }[]).map(row => row.name.normalize("NFC")).toSorted())
        .toEqual([composed, "Caf\u00e9 (2).docx"].toSorted())
    } finally { db.close() }
  })
}, 90_000)

// Converting a PDF never overwrites a document that already holds the .docx
// name: the conversion is stored under the next free name, which it returns.
test("a PDF conversion keeps both when its .docx name is already taken", async () => {
  await withMatter("PDF conversions", async ({ matter, request, upload, deliver, stored, role }) => {
    const convert = async (name: string) => {
      const response = await request(`/matters/${matter.id}/documents/convert?name=${encodeURIComponent(name)}`, { method: "POST" })
      expect(response.status).toBe(200)
      return (await response.json() as { name: string }).name
    }

    expect((await upload("Notice.pdf", pdf("The invented notice is given in writing."))).status).toBe(200)
    expect(await convert("Notice.pdf")).toBe("Notice.docx")

    const ours = await docx("Our marker")
    expect((await upload("Agreement.docx", ours)).status).toBe(200)
    expect((await upload("Agreement.pdf", pdf("The invented agreement returned as a PDF."))).status).toBe(200)
    expect(await convert("Agreement.pdf")).toBe("Agreement (2).docx")
    expect(await stored("Agreement.docx")).toEqual(ours)
    expect(role("Agreement (2).docx")).toBe("source")
    // Converting again adds another copy rather than replacing the first.
    expect(await convert("Agreement.pdf")).toBe("Agreement (3).docx")

    const delivered = await deliver("Memo.docx", await docx("Drafted marker"))
    expect((await upload("Memo.pdf", pdf("The invented memo returned as a PDF."))).status).toBe(200)
    expect(await convert("Memo.pdf")).toBe("Memo (2).docx")
    expect(await stored("Memo.docx")).toEqual(delivered)
    expect(role("Memo.docx")).toBe("deliverable")
  })
}, 120_000)

// The extraction quality gate's result (#239) goes out in the upload response,
// the document list and the document text route; a blank file is kept and
// listed, not indexed.
test("an upload's extraction quality is in the response, the document list and the text route", async () => {
  await withMatter("Extraction quality", async ({ matter, request, upload, stored }) => {
    const blank = new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: [] }] })))
    const empty = await upload("Blank.docx", blank)
    expect(empty.status).toBe(200)
    expect(await empty.json()).toMatchObject({ name: "Blank.docx", chunks: 0, injection: null, quality: { status: "no_readable_text" } })
    expect(await stored("Blank.docx")).toEqual(blank)
    const normal = await upload("Agreement.docx", await docx("Our marker"))
    expect(await normal.json()).toMatchObject({ quality: { status: "ok" } })

    const listed = await (await request(`/matters/${matter.id}`)).json() as { documents: { name: string; quality?: { status: string } }[] }
    expect(Object.fromEntries(listed.documents.map((document) => [document.name, document.quality?.status]))).toEqual({
      "Blank.docx": "no_readable_text",
      "Agreement.docx": "ok",
    })
    const text = await (await request(`/matters/${matter.id}/documents/text?name=Blank.docx`)).json()
    expect(text).toMatchObject({ text: "", injection: null, quality: { status: "no_readable_text" } })
  })
}, 120_000)

// A zip upload brings in the files inside it; one past the zip limits is
// refused with a message the upload row can show, and nothing of it is added.
test("a zip upload adds its files, and one over the limits is refused whole", async () => {
  await withMatter("Zip uploads", async ({ matter, request, upload }) => {
    const added = await upload("Bundle.zip", zipSync({ "Bundle/Agreement.docx": await docx("Zipped marker"), "Bundle/photo.png": strToU8("png") }))
    expect(added.status).toBe(200)
    expect(await added.json()).toMatchObject({
      name: "Bundle.zip",
      quality: { status: "ok" },
      attachments: [
        { name: "Bundle/Agreement.docx", document: "Agreement.docx", status: "indexed" },
        { name: "Bundle/photo.png", status: "unsupported" },
      ],
    })

    const many = zipSync(Object.fromEntries(Array.from({ length: MAX_ZIP_FILES + 1 }, (_, index) => [`f${index}.txt`, strToU8("x")])))
    const refused = await upload("Many.zip", many)
    expect(refused.status).toBe(413)
    expect(await refused.json()).toEqual({ error: expect.stringContaining("Many.zip was not added"), code: "zip_too_large" })
    const listed = await (await request(`/matters/${matter.id}`)).json() as { documents: { name: string; attachment_of?: { name: string } }[] }
    expect(listed.documents.map((document) => [document.name, document.attachment_of?.name]).toSorted()).toEqual([
      ["Agreement.docx", "Bundle.zip"],
      ["Bundle.zip", undefined],
    ])
  })
}, 120_000)

const docx = async (marker: string) => new Uint8Array(await Packer.toBuffer(new Document({ sections: [{ children: [
  new Paragraph("1. Scope"), new Paragraph(`${marker} describes the invented scope.`),
] }] })))

// A one-page PDF with one line of text.
function pdf(wording: string) {
  const stream = `BT /F1 12 Tf 40 700 Td (${wording}) Tj ET`
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ]
  let text = "%PDF-1.4\n"
  const offsets = objects.map((object, index) => {
    const offset = text.length
    text += `${index + 1} 0 obj\n${object}\nendobj\n`
    return offset
  })
  const xref = text.length
  text += `xref\n0 6\n0000000000 65535 f \n${offsets.map(offset => `${String(offset).padStart(10, "0")} 00000 n \n`).join("")}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return new Uint8Array(Buffer.from(text))
}

// An isolated ingest service with one new matter. Uploads need no generation
// engine; any accidental engine request fails the test.
async function withMatter(title: string, run: (matter: {
  matter: { id: string; dir: string }
  request: (route: string, options?: RequestInit) => Promise<Response>
  upload: (name: string, bytes: Uint8Array<ArrayBuffer>, existing?: string) => Promise<Response>
  deliver: (name: string, bytes: Uint8Array<ArrayBuffer>) => Promise<Uint8Array>
  stored: (name: string) => Promise<Uint8Array>
  role: (name: string) => ReturnType<typeof getDocumentRole>
}) => Promise<void>) {
  const temporary = await mkdtemp(path.join(tmpdir(), "dochaus-source-upload-http-"))
  const workspace = path.join(temporary, "workspace")
  const config = path.join(temporary, "config")
  await mkdir(workspace)
  await mkdir(config)
  await writeFile(path.join(config, "opencode.json"), "{}\n")
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
    ...options, signal: AbortSignal.timeout(60_000),
  })
  try {
    const deadline = Date.now() + 20_000
    while (true) {
      if (await request("/matters").then(response => response.ok).catch(() => false)) break
      if (Date.now() >= deadline) throw new Error(`Ingest startup failed:\n${await readFile(serviceLog, "utf8")}`)
      await Bun.sleep(50)
    }
    const created = await request("/matters", { method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ title }) })
    expect(created.status).toBe(200)
    const matter = await created.json() as { id: string; dir: string }
    const post = (route: string, name: string, bytes: Uint8Array<ArrayBuffer>, existing?: string) => {
      const form = new FormData()
      form.append("file", new File([bytes], name))
      if (existing) form.append("existing", existing)
      return request(`/matters/${matter.id}/${route}`, { method: "POST", body: form })
    }
    const stored = async (name: string) => new Uint8Array(await readFile(path.join(matter.dir, name)))
    await run({
      matter,
      request,
      upload: (name, bytes, existing) => post("documents", name, bytes, existing),
      // The deliverable writer adds paragraph pointers, so return the stored bytes.
      deliver: async (name, bytes) => {
        expect((await post("deliverables", name, bytes)).status).toBe(200)
        return stored(name)
      },
      stored,
      role: (name) => {
        const db = openDb(matter.dir)
        try {
          return getDocumentRole(db, path.join(matter.dir, name))
        } finally { db.close() }
      },
    })
    expect(engineRequests).toEqual([])
  } finally {
    service.kill()
    await service.exited
    engine.stop(true)
    await rm(temporary, { recursive: true, force: true })
  }
}
