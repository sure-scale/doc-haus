import { afterEach, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { tmpdir } from "node:os"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "../../../dochaus/lib/docxodus"
import { exportDocxComments, readDocxComments } from "../../../dochaus/lib/docx-comments"
import { inspectExactParagraphSpacing } from "../../../dochaus/lib/docx-exact-spacing"
import { frozenParagraphSpacingView, paragraphSpacingProposalMarker, paragraphSpacingProposalSchema, type ParagraphSpacingProposalBinding } from "../../../dochaus/lib/paragraph-spacing-proposal"
import { exclusiveRedlineProposalMarker, tableRowProposalMarker } from "../../../dochaus/lib/table-row-proposal"
import { recordRedline, recordRedlineProposal } from "../../../dochaus/lib/redlines"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import { getPreparedChanges, getRedline, listPendingRedlines, openDb, setPreparedChanges, upsertDocument, type RedlineRow } from "./db"
import { resolveMatterDocumentTarget } from "./document-identity"
import { checkedParagraphSpacingOperation, paragraphSpacingView, recordParagraphSpacingProposal } from "./paragraph-spacing-proposals"
import { recordTableRowProposal } from "./table-row-proposals"
import { bake, buildRedlined, redlineState } from "./redline"

const FILE = "working.docx", AUTHOR = "Synthetic Counsel", REASON = "Explains the expressly requested spacing after this paragraph."
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", BODY = "word/document.xml"
const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
async function source() {
  return exportDocxComments((await renderMarkdownToDocx("Before paragraph.\n\nComment paragraph.\n\n| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes,
    [{ id: "prior", author: "Prior Reviewer", text: "Preserve this note.", anchor: { paragraph_text: "Comment paragraph." } }]).bytes
}
function changed(bytes: Uint8Array, part: string, change: (doc: Document) => void) {
  const parts = unzipSync(bytes), doc = parse(parts[part]!); change(doc)
  parts[part] = strToU8(new XMLSerializer().serializeToString(doc)); return zipSync(parts)
}
async function operation(bytes: Uint8Array, text = "Same", index = 1): Promise<ParagraphSpacingProposalBinding> {
  const packet = await inspectExactParagraphSpacing(bytes, text), target = packet.matches[index]!
  expect(target.eligibility).toEqual({ eligible: true })
  expect(target.directAfter.kind).not.toBe("unavailable")
  if (target.directAfter.kind === "unavailable") throw new Error("Fixture has no direct spacing")
  return { version: 4, kind: "paragraph_spacing_after", prepared_role: "source", source_sha256: packet.sourceSha256,
    paragraph_handle: target.paragraphHandle, paragraph_text: target.text, table_context: target.tableContext,
    before_direct_after: target.directAfter, after_direct_after_twips: "120" }
}
const dirs: string[] = []
afterEach(() => dirs.splice(0).forEach(dir => rmSync(dir, { recursive: true, force: true })))
function matter(bytes: Uint8Array) {
  const directory = realpathSync.native(mkdtempSync(path.join(tmpdir(), "spacing-proposal-"))); dirs.push(directory)
  const file = path.join(directory, FILE); writeFileSync(file, bytes)
  const db = openDb(directory); try { upsertDocument(db, file, FILE, Date.now(), null) } finally { db.close() }
  return { directory, file, bytes }
}
const review = (file: string) => ({ author: AUTHOR, reason: REASON, openItems: ["Retain the review note."], preparedPath: file })
function pending(directory: string, file: string) { const db = openDb(directory); try { return listPendingRedlines(db, file) } finally { db.close() } }
const textProposal = (file: string) => ({ docPath: file, docName: FILE, scope: "phrase" as const, findText: "Before", oldText: "Before", newText: "Earlier", author: AUTHOR, reviewComment: "A separate wording change.", anchorId: "before" })

test("v4 is exclusive without becoming a table operation; frozen source properties cannot authorize another target", async () => {
  const m = matter(await source()), value = await operation(m.bytes), { id } = await recordParagraphSpacingProposal(m.directory, FILE, value, review(m.file)), row = pending(m.directory, m.file)[0]!
  expect(row.id).toBe(id); expect(row.old_text).toBe(row.new_text)
  expect(await paragraphSpacingView(m.bytes, row)).toEqual({ ...frozenParagraphSpacingView(value), table_context: value.table_context })
  const state = await redlineState(m.bytes, [row], {}, "source"), preview = await buildRedlined(m.bytes, [row], state)
  expect(readDocxComments(preview).map(comment => comment.text)).toEqual(["Preserve this note.", REASON])
  const accepted = await bake(m.bytes, [row], state)
  expect(readDocxComments(accepted).map(comment => comment.text)).toEqual(["Preserve this note."])
  expect(readFileSync(m.file)).toEqual(Buffer.from(m.bytes)); expect(row.status).toBe("pending")
  for (const disguised of [{ scope: "phrase", target_json: JSON.stringify({ version: 4, kind: "unknown" }) }, { scope: "clause", target_json: JSON.stringify({ kind: "paragraph_spacing_after" }) }] as const) {
    expect(paragraphSpacingProposalMarker(disguised)).toBe(true); expect(exclusiveRedlineProposalMarker(disguised)).toBe(true); expect(tableRowProposalMarker(disguised)).toBe(false)
    await expect(buildRedlined(m.bytes, [{ ...row, ...disguised }], state)).rejects.toThrow("invalid stored operation")
    await expect(bake(m.bytes, [{ ...row, ...disguised }], state)).rejects.toThrow("invalid stored operation")
  }
  expect(paragraphSpacingProposalSchema.safeParse({ ...value, before_direct_after: { kind: "twips", raw: "bad", normalized: "80" } }).success).toBe(false)
  expect(paragraphSpacingProposalSchema.safeParse({ ...value, before_direct_after: { kind: "twips", raw: "0080", normalized: "80" }, after_direct_after_twips: "80" }).success).toBe(false)
  expect(paragraphSpacingProposalSchema.safeParse({ ...value, position: "after" }).success).toBe(false)
  const empty = matter(m.bytes), sdk = await docxodus(), probe = sdk.openDocxSession(m.bytes, {}), proto = Object.getPrototypeOf(probe) as typeof probe; probe.close()
  const setter = spyOn(proto, "setParagraphFormat")
  try {
    const body = (await inspectExactParagraphSpacing(m.bytes, "Before paragraph.")).matches[0]!
    expect(body.eligibility).toEqual({ eligible: false, reason: "unsupported_empty_properties" })
    expect(body.directAfter).toEqual({ kind: "absent" })
    const unsupported: ParagraphSpacingProposalBinding = { ...value, paragraph_handle: body.paragraphHandle, paragraph_text: body.text, table_context: body.tableContext, before_direct_after: { kind: "absent" } }
    const bodyRow = { ...row, find_text: body.text, old_text: body.text, new_text: body.text, target_json: JSON.stringify(unsupported) }
    await expect(recordParagraphSpacingProposal(empty.directory, FILE, unsupported, review(empty.file))).rejects.toThrow("cannot be safely prepared")
    await expect(paragraphSpacingView(m.bytes, bodyRow)).rejects.toThrow("cannot be safely prepared")
    await expect(buildRedlined(m.bytes, [bodyRow], state)).rejects.toThrow("cannot be safely prepared")
    await expect(bake(m.bytes, [bodyRow], state)).rejects.toThrow("cannot be safely prepared")
    for (const bad of [
      { ...value, before_direct_after: { kind: "twips" as const, raw: "0080", normalized: "80" } },
      { ...value, paragraph_handle: value.paragraph_handle + "x" },
      { ...value, table_context: { ...value.table_context!, column: 1 } },
      { ...value, paragraph_text: "Different paragraph." },
    ]) {
      const forged = { ...row, find_text: bad.paragraph_text, old_text: bad.paragraph_text, new_text: bad.paragraph_text, target_json: JSON.stringify(bad) }
      await expect(recordParagraphSpacingProposal(empty.directory, FILE, bad, review(empty.file))).rejects.toThrow("does not match")
      await expect(paragraphSpacingView(m.bytes, forged)).rejects.toThrow("does not match")
      await expect(buildRedlined(m.bytes, [forged], state)).rejects.toThrow("does not match")
      await expect(bake(m.bytes, [forged], state)).rejects.toThrow("does not match")
    }
    await expect(checkedParagraphSpacingOperation(m.bytes, { ...row, new_text: "A hidden wording change." })).rejects.toThrow("invalid stored operation")
    await expect(checkedParagraphSpacingOperation(m.bytes, { ...row, formatting_json: "{}" })).rejects.toThrow("invalid stored operation")
    expect(setter).not.toHaveBeenCalled(); expect(pending(empty.directory, empty.file)).toEqual([])
  } finally { setter.mockRestore() }
})

test("every shared recorder refuses coexistence before supersession, including a disguised unknown v4", async () => {
  const m = matter(await source()), value = await operation(m.bytes)
  await recordParagraphSpacingProposal(m.directory, FILE, value, review(m.file))
  for (const recorder of [recordRedline, recordRedlineProposal]) expect(() => recorder(m.directory, textProposal(m.file))).toThrow("reviewed on its own")
  await expect(recordTableRowProposal(m.directory, FILE, { version: 3, kind: "table_row_insert", prepared_role: "source", source_sha256: value.source_sha256, target_handle: "not-reached", position: "after", cells: ["New", "Row"] }, review(m.file))).rejects.toThrow("reviewed on its own")
  const db = openDb(m.directory)
  try { db.run("UPDATE redlines SET scope='phrase',target_json=?", [JSON.stringify({ version: 4, kind: "unknown" })]) } finally { db.close() }
  for (const recorder of [recordRedline, recordRedlineProposal]) expect(() => recorder(m.directory, textProposal(m.file))).toThrow("reviewed on its own")
  expect(pending(m.directory, m.file).map(row => row.status)).toEqual(["pending"])
  const n = matter(m.bytes); recordRedline(n.directory, textProposal(n.file))
  await expect(recordParagraphSpacingProposal(n.directory, FILE, value, review(n.file))).rejects.toThrow("reviewed on its own")
  expect(pending(n.directory, n.file)).toHaveLength(1)
})

test("post-preflight source, physical path, role and pending races refuse recording without clobbering metadata", async () => {
  const bytes = await source(), value = await operation(bytes), sdk = await docxodus(), probe = sdk.openDocxSession(bytes, {}), proto = Object.getPrototypeOf(probe) as typeof probe; probe.close()
  const set = proto.setParagraphFormat
  for (const race of ["source", "path", "role", "pending"] as const) {
    const m = matter(bytes), db = openDb(m.directory), metadata = { redline_ids: [], open_items: ["Keep prior disclosure."] }
    try { setPreparedChanges(db, m.file, metadata) } finally { db.close() }
    let injected = false
    const spy = spyOn(proto, "setParagraphFormat").mockImplementation(function(this: typeof probe, ...args: Parameters<typeof set>) {
      const result = set.apply(this, args)
      if (!injected) {
        injected = true
        if (race === "source") writeFileSync(m.file, changed(bytes, "word/styles.xml", doc => doc.documentElement!.setAttribute("changed", "yes")))
        else if (race === "path") renameSync(m.file, path.join(m.directory, "Working.docx"))
        else if (race === "role") { const db = openDb(m.directory); try { db.run("UPDATE documents SET role='deliverable'") } finally { db.close() } }
        else recordRedline(m.directory, textProposal(m.file))
      }
      return result
    })
    try { await expect(recordParagraphSpacingProposal(m.directory, FILE, value, review(m.file))).rejects.toThrow(race === "pending" ? "reviewed on its own" : "earlier Word file") }
    finally { spy.mockRestore() }
    expect(injected).toBe(true); expect(pending(m.directory, m.file).filter(row => row.scope === "paragraph_spacing")).toEqual([])
    const check = openDb(m.directory); try { expect(getPreparedChanges(check, m.file)).toEqual(metadata) } finally { check.close() }
  }
})

test("real spacing list, native preview and single/all Accept publish only clean reviewed properties; stale and mixed requests write nothing", async () => {
  const temporary = realpathSync.native(mkdtempSync(path.join(tmpdir(), "spacing-http-"))); dirs.push(temporary)
  const workspace = path.join(temporary, "workspace"), config = path.join(temporary, "config"); mkdirSync(workspace); mkdirSync(config); writeFileSync(path.join(config, "opencode.json"), "{}")
  const engineRequests: string[] = [], engine = Bun.serve({ port: 0, fetch(request) { engineRequests.push(request.url); return new Response("No generation engine", { status: 503 }) } })
  const probe = Bun.serve({ port: 0, fetch: () => new Response() }), port = probe.port!; probe.stop(true)
  const log = path.join(temporary, "ingest.log"), service = Bun.spawn([process.execPath, "run", "src/server.ts"], { cwd: path.resolve(import.meta.dir, ".."), env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config, INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "" }, stdout: Bun.file(log), stderr: Bun.file(log) })
  const request = (route: string, options?: RequestInit) => fetch(`http://127.0.0.1:${port}${route}`, { ...options, signal: AbortSignal.timeout(30000) })
  try {
    const deadline = Date.now() + 20000
    while (!await request("/matters").then(response => response.ok).catch(() => false)) { if (service.exitCode !== null || Date.now() > deadline) throw new Error(await Bun.file(log).text()); await Bun.sleep(50) }
    for (const mode of ["single", "all", "body", "snapshot", "stale", "role", "mixed", "invalid"] as const) {
      const created = await (await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: `Synthetic spacing ${mode}` }) })).json() as { id: string; dir: string }, route = `/matters/${created.id}`
      const original = await source(), bytes = mode === "body" ? changed(original, BODY, doc => {
        const target = Array.from(doc.getElementsByTagNameNS(W, "p")).find(p => Array.from(p.getElementsByTagNameNS(W, "t")).some(t => t.textContent === "Before paragraph."))!
        let props = Array.from(target.childNodes).find((node): node is Element => node.nodeType === 1 && (node as Element).namespaceURI === W && (node as Element).localName === "pPr")
        if (!props) { props = doc.createElementNS(W, "w:pPr"); target.insertBefore(props, target.firstChild) }
        expect(props.getElementsByTagNameNS(W, "spacing").length).toBe(0)
        const spacing = doc.createElementNS(W, "w:spacing"); spacing.setAttributeNS(W, "w:after", "160"); props.appendChild(spacing)
      }) : original, form = new FormData(); form.append("file", new File([new Uint8Array(bytes).buffer], FILE))
      expect((await request(`${route}/documents`, { method: "POST", body: form })).status).toBe(200)
      const file = resolveMatterDocumentTarget(created.dir, FILE).docPath, value = mode === "body" ? await operation(bytes, "Before paragraph.", 0) : await operation(bytes), { id } = await recordParagraphSpacingProposal(created.dir, FILE, value, review(file)), rows = pending(created.dir, file)
      if (mode === "body") { expect(value.table_context).toBeNull(); expect(value.before_direct_after).toEqual({ kind: "twips", raw: "160", normalized: "160" }) }
      const expected = await bake(bytes, rows, await redlineState(bytes, rows, {}, "source"))
      const versions = () => { const db = openDb(created.dir); try { return { current: db.query("SELECT current_version FROM versioned_documents").all(), kept: db.query("SELECT * FROM document_versions ORDER BY number").all(), events: db.query("SELECT * FROM document_version_events ORDER BY id").all() } } finally { db.close() } }
      const beforeVersions = versions(); expect(beforeVersions.kept).toHaveLength(1)
      const preview = await request(`${route}/documents/redlined?name=${FILE}`); expect(preview.status).toBe(200)
      expect(readDocxComments(new Uint8Array(await preview.arrayBuffer())).map(comment => comment.text)).toEqual(["Preserve this note.", REASON])
      expect(readFileSync(file)).toEqual(Buffer.from(bytes)); expect(versions()).toEqual(beforeVersions)
      let textId: number | undefined
      if (mode === "snapshot") { const db = openDb(created.dir); try { db.run("UPDATE redlines SET target_json=? WHERE id=?", [JSON.stringify({ ...value, before_direct_after: { kind: "twips", raw: "0080", normalized: "80" } }), id]) } finally { db.close() } }
      if (mode === "stale") writeFileSync(file, changed(bytes, "word/styles.xml", doc => doc.documentElement!.setAttribute("changed", "yes")))
      if (mode === "role") { const db = openDb(created.dir); try { db.run("UPDATE documents SET role='deliverable'") } finally { db.close() } }
      if (mode === "mixed") { const db = openDb(created.dir); try { textId = Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,created_at) VALUES (?,?,?,?,?,?,?,?)", [file, FILE, "phrase", "Before", "Before", "Earlier", AUTHOR, Date.now()]).lastInsertRowid) } finally { db.close() } }
      if (mode === "invalid") { const db = openDb(created.dir); try { db.run("UPDATE redlines SET scope='phrase',target_json=? WHERE id=?", [JSON.stringify({ version: 4, kind: "unknown" }), id]) } finally { db.close() } }
      const before = readFileSync(file), listed = await (await request(`${route}/redlines?name=${FILE}`)).json() as Array<RedlineRow & { paragraph_spacing?: unknown; stale?: boolean }>, supported = ["single", "all", "body"].includes(mode)
      const display = mode === "snapshot" ? { ...value, before_direct_after: { kind: "twips" as const, raw: "0080", normalized: "80" } } : value
      expect(listed[0]!.paragraph_spacing).toEqual(supported ? { ...frozenParagraphSpacingView(value), table_context: value.table_context } : mode === "invalid" ? undefined : frozenParagraphSpacingView(display))
      expect(listed[0]!.stale).toBe(supported ? undefined : true)
      const response = await request(mode === "all" ? `${route}/redlines/accept-all?name=${FILE}` : `${route}/redlines/${id}/accept`, { method: "POST" })
      expect(response.status).toBe(supported ? 200 : 409)
      if (supported) {
        const actual = readFileSync(file), expectedParts = unzipSync(expected), actualParts = unzipSync(actual)
        // Independent authoring can have different ZIP timestamps. Every part
        // must be exact, and the actual published/kept/download bytes agree.
        expect(actualParts).toEqual(expectedParts); expect(readDocxComments(actual).map(comment => comment.text)).toEqual(["Preserve this note."])
        const kept = versions(); expect(kept.kept).toHaveLength(2)
        const latest = kept.kept[1] as { id: string; sha256: string }
        expect(latest.sha256).toBe(hash(actual))
        const download = await request(`${route}/documents/content?name=${FILE}`); expect(new Uint8Array(await download.arrayBuffer())).toEqual(new Uint8Array(actual))
      } else {
        expect(readFileSync(file)).toEqual(before); expect(versions()).toEqual(beforeVersions)
        expect((await request(`${route}/redlines/accept-all?name=${FILE}`, { method: "POST" })).status).toBe(409)
        if (mode === "mixed") { expect((await request(`${route}/redlines/${textId}/accept`, { method: "POST" })).status).toBe(409); expect((await request(`${route}/documents/redlined?name=${FILE}`)).status).toBe(409) }
      }
      const db = openDb(created.dir); try { expect(getRedline(db, id)!.status).toBe(supported ? "accepted" : "pending") } finally { db.close() }
    }
    expect(engineRequests).toEqual([])
  } finally { service.kill(); await service.exited; engine.stop(true) }
}, 120000)
