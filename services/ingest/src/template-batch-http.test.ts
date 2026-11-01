// Run from services/ingest/: bun test src/template-batch-http.test.ts
import { expect, test } from "bun:test"
import { mkdtemp, mkdir, readdir, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { utils, write } from "xlsx"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"

const templateMarkdown = [
  "# Synthetic engagement letter",
  "## 1. Parties\n\nThis letter is addressed to [Client Name] of [Client Address].",
  "## 2. Fees\n\nFees are payable from [insert start date].",
  "## 3. Non-solicitation [optional: Non-solicit]\n\nNo solicitation for twelve months.",
].join("\n\n")

const documentText = (bytes: Uint8Array) =>
  [...strFromU8(unzipSync(bytes)["word/document.xml"]!).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("")

// An ingest on its own data root, with the synthetic template in its library.
async function startIngest() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-template-batch-http-"))
  const config = path.join(root, "config")
  await mkdir(config)
  await Bun.write(path.join(config, "opencode.json"), "{}")
  const probe = Bun.serve({ hostname: "127.0.0.1", port: 0, fetch: () => new Response() })
  const port = probe.port
  probe.stop(true)
  const child = Bun.spawn([process.execPath, "run", "src/server.ts"], {
    cwd: path.resolve(import.meta.dir, ".."),
    env: {
      ...process.env,
      DOCHAUS_DATA_ROOT: root,
      WORKSPACE_ROOT: undefined,
      DOCHAUS_RUNTIME_DIR: undefined,
      DOCHAUS_DIR: config,
      INGEST_HOST: "127.0.0.1",
      INGEST_PORT: String(port),
    },
    stdout: Bun.file(path.join(root, "ingest.log")),
    stderr: Bun.file(path.join(root, "ingest.log")),
  })
  const request = (route: string, options?: RequestInit) =>
    fetch(`http://127.0.0.1:${port}${route}`, { ...options, signal: AbortSignal.timeout(60000) })
  const post = (route: string, body: unknown) =>
    request(route, { method: "POST", headers: { "Content-Type": "application/json", "x-dochaus-actor": "user:example" }, body: JSON.stringify(body) })
  const readSheet = (file: File) => {
    const body = new FormData()
    body.set("name", "Engagement Letter.docx")
    body.set("file", file)
    return request("/templates/batch/read", { method: "POST", body })
  }
  const stop = async () => {
    child.kill("SIGTERM")
    await child.exited
    await rm(root, { recursive: true, force: true })
  }
  try {
    const deadline = Date.now() + 20000
    while (!(await request("/matters").catch(() => undefined))?.ok) {
      if (Date.now() > deadline || child.exitCode !== null) throw new Error("Isolated ingest did not start")
      await Bun.sleep(50)
    }
    const upload = new FormData()
    upload.set("file", new File([(await renderMarkdownToDocx(templateMarkdown)).bytes], "Engagement Letter.docx"))
    expect((await request("/templates", { method: "POST", body: upload })).status).toBe(200)
  } catch (error) {
    await stop()
    throw error
  }
  return { root, request, post, readSheet, stop }
}

test("spreadsheet batch: read, preview, sample, create, remember, refuse a changed template", async () => {
  const { root, request, post, readSheet, stop } = await startIngest()
  try {
    // The downloadable spreadsheet names every field the template can take.
    const sample = await request("/templates/batch/sample?name=Engagement%20Letter.docx")
    expect(sample.headers.get("content-type")).toContain("text/csv")
    expect(sample.headers.get("content-disposition")).toContain(encodeURIComponent("Engagement Letter – spreadsheet.csv"))
    expect((await sample.text()).replace(/^\uFEFF/, "")).toBe("Client Name,Client Address,Start date,Include: Non-solicit\r\n")

    // CSV: header repeated and blank rows dropped; values stay as typed.
    const csv = [
      "Client Name,Address,Start date,Include: Non-solicit",
      "Example Holdings Ltd,1 Example Street,01/02/2026,yes",
      ",,,",
      "Client Name,Address,Start date,Include: Non-solicit",
      "Sample Trading Ltd,,007,no",
    ].join("\n")
    const fromCsv = await (await readSheet(new File([csv], "clients.csv"))).json()
    expect(fromCsv).toMatchObject({
      template: { name: "Engagement Letter.docx" },
      unnamedBlanks: 0,
      columns: ["Client Name", "Address", "Start date", "Include: Non-solicit"],
      rows: [
        { row: 2, values: ["Example Holdings Ltd", "1 Example Street", "01/02/2026", "yes"] },
        { row: 5, values: ["Sample Trading Ltd", "", "007", "no"] },
      ],
      dropped: { blank: 1, header: 1 },
      mapping: { "[Client Name]": "Client Name", "[Client Address]": null, "[insert start date]": "Start date", "[optional: Non-solicit]": "Include: Non-solicit" },
      remembered: false,
      identifyingKey: "[Client Name]",
      source: "clients.csv",
    })
    expect(fromCsv.template.sha256).toMatch(/^[a-f0-9]{64}$/)

    // XLSX: cells read as displayed (a date stays a date, not a serial).
    const rows = Array.from({ length: 50 }, (_, index) => [`Example Client ${index + 1} Ltd`, `${index + 1} Example Road`, new Date(Date.UTC(2026, 0, index + 1)), index % 2 ? "no" : "yes"])
    const sheet = utils.aoa_to_sheet([["Client Name", "Client Address", "Start date", "Include: Non-solicit"], ...rows], { cellDates: true, dateNF: "d mmmm yyyy" })
    const book = utils.book_new()
    utils.book_append_sheet(book, sheet, "Clients")
    const xlsx = new Uint8Array(write(book, { type: "array", bookType: "xlsx" }) as ArrayBuffer)
    const read = await (await readSheet(new File([xlsx], "clients.xlsx"))).json()
    expect(read.rows).toHaveLength(50)
    expect(read.rows[0]).toEqual({ row: 2, values: ["Example Client 1 Ltd", "1 Example Road", "1 January 2026", "yes"] })
    expect(Object.values(read.mapping).every(Boolean)).toBe(true)

    // Preview: one row's document plus its missing labels.
    const preview = await post("/templates/batch/preview", { name: "Engagement Letter.docx", columns: read.columns, values: ["Example Client 1 Ltd", "", "1 January 2026", "yes"], mapping: read.mapping })
    expect(preview.status).toBe(200)
    expect(JSON.parse(decodeURIComponent(preview.headers.get("x-dochaus-batch-missing")!))).toEqual(["Client Address"])
    const previewText = documentText(new Uint8Array(await preview.arrayBuffer()))
    expect(previewText).toContain("addressed to Example Client 1 Ltd of [Client Address].")

    // Create fifty; row 4 skipped, so forty-nine documents.
    const matter = await (await post("/matters", { title: "Batch example" })).json() as { id: string }
    const created = await (await post(`/matters/${matter.id}/template-batches`, {
      name: read.template.name, sha256: read.template.sha256, source: read.source, columns: read.columns, rows: read.rows, skipped: [4], mapping: read.mapping,
    })).json() as { id: string; folder: string; documents: { name: string; row: number; missing: string[] }[]; skipped: number[]; failed: unknown[] }
    expect(created.failed).toEqual([])
    expect(created.skipped).toEqual([4])
    expect(created.documents).toHaveLength(49)
    expect(created.folder).toMatch(/^Engagement Letter – \d{4}-\d{2}-\d{2}$/)
    expect(created.documents[0]).toEqual({ name: "Engagement Letter – Example Client 1 Ltd – row 002.docx", row: 2, missing: [] })
    expect(created.documents.some((d) => d.row === 4)).toBe(false)

    const listed = await (await request(`/matters/${matter.id}`)).json() as { documents: { name: string; role: string; batch?: { id: string; folder: string; row: number; source: string; missing: string[] } }[] }
    const batchDocs = listed.documents.filter((d) => d.batch)
    expect(batchDocs).toHaveLength(49)
    expect(batchDocs.every((d) => d.role === "deliverable" && d.batch!.id === created.id && d.batch!.source === "clients.xlsx")).toBe(true)
    expect(listed.documents.find((d) => d.batch?.row === 10)?.batch).toEqual({ id: created.id, folder: created.folder, row: 10, source: "clients.xlsx", missing: [] })

    const tenth = created.documents.find((d) => d.row === 10)!
    const versions = await (await request(`/matters/${matter.id}/documents/versions?name=${encodeURIComponent(tenth.name)}`)).json() as { versions: { origin: string; actor: string; label: string }[] }
    expect(versions.versions).toMatchObject([{ origin: "generated", actor: "user:example", label: "From Engagement Letter · spreadsheet row 10 of clients.xlsx" }])
    const tenthText = documentText(new Uint8Array(await (await request(`/matters/${matter.id}/documents/content?name=${encodeURIComponent(tenth.name)}`)).arrayBuffer()))
    expect(tenthText).toContain("addressed to Example Client 9 Ltd of 9 Example Road.")
    expect(tenthText).not.toContain("Example Client 8 Ltd")

    const record = await Bun.file(path.join(root, "workspace", matter.id, ".dochaus", "template-batches", `${created.id}.json`)).json()
    expect(record).toMatchObject({ id: created.id, folder: created.folder, template: { name: "Engagement Letter.docx", sha256: read.template.sha256 }, source: "clients.xlsx", skipped: [4], failed: [] })
    expect(record.documents).toHaveLength(49)

    // The matching is remembered for the template: a renamed column is
    // matched again on the next read.
    await post(`/matters/${matter.id}/template-batches`, {
      name: read.template.name, sha256: read.template.sha256, source: "clients.csv", columns: ["Customer", "Address", "Start date", "Include: Non-solicit"],
      rows: [{ row: 2, values: ["Example Client 1 Ltd", "1 Example Road", "1 May 2026", "yes"] }], skipped: [],
      mapping: { "[Client Name]": "Customer", "[Client Address]": "Address", "[insert start date]": "Start date", "[optional: Non-solicit]": "Include: Non-solicit" },
    })
    const again = await (await readSheet(new File(["Customer,Address,Start date\nSample Trading Ltd,2 Example Road,1 June 2026\n"], "next.csv"))).json()
    expect(again.mapping).toEqual({ "[Client Name]": "Customer", "[Client Address]": "Address", "[insert start date]": "Start date", "[optional: Non-solicit]": null })
    expect(again.remembered).toBe(true)
    // The second batch's row-2 name was free only with a suffix.
    expect((await readdir(path.join(root, "workspace", matter.id))).filter((name) => name.includes("row 002"))).toEqual(
      expect.arrayContaining(["Engagement Letter – Example Client 1 Ltd – row 002.docx", "Engagement Letter – Example Client 1 Ltd – row 002 (2).docx"]),
    )

    // A template changed since the read is refused.
    const changed = new FormData()
    changed.set("file", new File([(await renderMarkdownToDocx(`${templateMarkdown}\n\n## 4. Changed\n\nNew text.`)).bytes], "Engagement Letter.docx"))
    await request("/templates", { method: "POST", body: changed })
    const refused = await post(`/matters/${matter.id}/template-batches`, {
      name: read.template.name, sha256: read.template.sha256, source: read.source, columns: read.columns, rows: read.rows.slice(0, 1), skipped: [], mapping: read.mapping,
    })
    expect(refused.status).toBe(409)
    expect(await refused.json()).toMatchObject({ code: "template_changed" })

    // Limits come back as plain 400s.
    const wide = await readSheet(new File([`${Array.from({ length: 201 }, (_, i) => `C${i}`).join(",")}\n${"x,".repeat(200)}x\n`], "wide.csv"))
    expect(wide.status).toBe(400)
    expect((await wide.json()).error).toContain("at most 200")
  } finally {
    await stop()
  }
}, 180000)

test("a batch document's missing values follow its content", async () => {
  const { root, request, post, readSheet, stop } = await startIngest()
  try {
    const read = await (await readSheet(new File([
      "Client Name,Client Address,Start date\nExample Holdings Ltd,,1 May 2026\nSample Trading Ltd,,1 June 2026\n",
    ], "clients.csv"))).json()
    const matter = await (await post("/matters", { title: "Batch content" })).json() as { id: string }
    const created = await (await post(`/matters/${matter.id}/template-batches`, {
      name: read.template.name, sha256: read.template.sha256, source: read.source, columns: read.columns, rows: read.rows, skipped: [], mapping: read.mapping,
    })).json() as { documents: { name: string; row: number; missing: string[] }[] }
    const missing = async () => Object.fromEntries(
      ((await (await request(`/matters/${matter.id}`)).json()) as { documents: { name: string; batch?: { missing: string[] } }[] }).documents
        .filter((d) => d.batch).map((d) => [d.name, d.batch!.missing]),
    )
    const [filled, edited] = created.documents.map((d) => d.name) as [string, string]
    expect(await missing()).toEqual({ [filled]: ["Client Address"], [edited]: ["Client Address"] })

    // Rewrite a document's text through the registered deliverable route.
    const rewrite = async (name: string, change: (xml: string) => string) => {
      const file = path.join(root, "workspace", matter.id, name)
      const current = new Uint8Array(await Bun.file(file).arrayBuffer())
      const parts = unzipSync(current)
      const xml = strFromU8(parts["word/document.xml"]!)
      const next = change(xml)
      expect(next).not.toBe(xml)
      parts["word/document.xml"] = strToU8(next)
      const body = new FormData()
      body.set("file", new File([zipSync(parts)], name))
      const response = await request(`/matters/${matter.id}/deliverables`, {
        method: "POST",
        headers: { "x-dochaus-deliverable-precondition": "replace", "x-dochaus-deliverable-prior-sha256": new Bun.CryptoHasher("sha256").update(current).digest("hex") },
        body,
      })
      expect(response.status).toBe(200)
    }
    // The lawyer fills the address: its badge goes.
    await rewrite(filled, (xml) => xml.replace("[Client Address]", "1 Example Street"))
    // Another document changes elsewhere but keeps the placeholder: still missing.
    await rewrite(edited, (xml) => xml.replace("No solicitation for twelve months.", "No solicitation for six months."))
    expect(await missing()).toEqual({ [filled]: [], [edited]: ["Client Address"] })
  } finally {
    await stop()
  }
}, 180000)
