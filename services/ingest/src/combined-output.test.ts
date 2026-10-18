import { afterEach, describe, expect, setSystemTime, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { lstatSync, mkdtempSync, readdirSync, readFileSync, readlinkSync, rmSync, symlinkSync, unlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import * as combinedDocx from "../../../dochaus/lib/combined-docx"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import * as deliverables from "../../../dochaus/workflows/lean/deliverables"
import { memoNotesFile, readLeanDocuments, recordLeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
import { writeLeanPlan, type PlannedDocument } from "../../../dochaus/workflows/lean/plan"
import { initLeanState, leanFile, readLeanState, writeLeanBytes, writeLeanState, writeLeanText } from "../../../dochaus/workflows/lean/state"
import { leanFixtureMatter, scriptedLeanModel } from "../../../dochaus/workflows/lean/test-support"
import { combinedOutputRoutes } from "./combined-output"
import { attachmentDocumentDisposition } from "./content-disposition"
import { openDb } from "./db"
import { writeDeliverableRevision } from "./deliverable"
import { leanDraftingStatus } from "./lean-drafting"
import { registerLeanDeliverable } from "./lean-deliverables"

// The combined download: a document planned as "combined" with a requested memo
// is downloaded as one file built from its registered body and notes. Real
// registration through the matter registry; nothing is scripted but the model
// port, which registration never calls.

const directories: string[] = []
afterEach(() => directories.splice(0).forEach(directory => rmSync(directory, { recursive: true, force: true })))

type Document = Omit<PlannedDocument, "id">
const SUPPLY: Document = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const COMBINED: Document = { ...SUPPLY, file: "supply-pack.docx", deliverable: "combined" }
/** Another planned document, delivered as its own body. */
const SIDE: Document = { title: "Side letter", purpose: "Notice terms", parts: [{ heading: "Notice", scope: "Notice terms" }], file: "side-letter.docx" }
const MEMO = "# Cover memo\n\nMemo marker: the start date needs confirmation."
const BODY = "# Supply note\n\n## Delivery\n\nBody marker: the supplier delivers the goods."
const SIDE_BODY = "# Side letter\n\n## Notice\n\nSide marker: notices go by post."
/** The status a delivered COMBINED document reports, with its registered parts. */
const COMBINED_OUTPUT = { document: "d1", name: "supply-pack", file: "supply-pack.docx", body: "supply-pack - clean.docx", notes: "supply-pack - cover memo.docx" }

/** A matter whose lean run passed its gate: the plan, the run's outputs, and
 * their registration as matter deliverables. */
async function deliveredMatter(documents: Document[] | null, memo = true) {
  const directory = leanFixtureMatter("dochaus-combined-", dir => openDb(dir).close())
  directories.push(directory)
  if (documents) {
    writeLeanPlan(directory, documents, undefined, undefined, memo ? { name: "cover memo", file: memoNotesFile(documents, "cover memo") } : undefined)
    for (const [index, document] of documents.entries()) {
      const body = (await renderMarkdownToDocx(document.title === SIDE.title ? SIDE_BODY : BODY)).bytes
      writeLeanBytes(directory, `d${index + 1}/body-applied.docx`, body)
      writeLeanBytes(directory, `d${index + 1}/redline.docx`, body)
    }
    writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx(MEMO)).bytes)
    writeLeanText(directory, "drafting-notes.md", `${MEMO}\n`)
    await recordLeanDocuments(directory, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
  }
  const app = combinedOutputRoutes(id => { expect(id).toBe("combined-matter"); return directory })
  return { directory, download: () => app.request("/matters/combined-matter/documents/combined") }
}

function documentText(bytes: Uint8Array) {
  return strFromU8(unzipSync(bytes)["word/document.xml"]!)
}

/** Rewrites a registered part's document XML on disk, as an edit outside
 * doc.haus would. */
function rewriteDocumentXml(directory: string, name: string, change: (xml: string) => string) {
  const file = path.join(directory, name)
  const parts = unzipSync(readFileSync(file))
  writeFileSync(file, zipSync({ ...parts, "word/document.xml": strToU8(change(strFromU8(parts["word/document.xml"]!))) }))
}

/** Everything the route could write: every entry in the matter at any depth,
 * dot directories included (the run's files and the registry live under
 * .dochaus). A file is its size, sha256 and modification time, a link its
 * target, a directory its path. A directory's own time is left out: a download
 * takes and drops its parts' mutation locks, which touches only the lock
 * root's time. So is the content of SQLite's shared-memory index ("-shm"),
 * which every reader of the registry updates (macOS SQLite keeps it between
 * connections); the registry and its write-ahead log are compared in full. */
function tree(directory: string, relative = ""): Record<string, unknown> {
  const entries: Record<string, unknown> = {}
  for (const name of readdirSync(path.join(directory, relative)).toSorted()) {
    const entry = path.join(relative, name)
    const file = path.join(directory, entry)
    const stat = lstatSync(file, { bigint: true })
    if (stat.isSymbolicLink()) entries[entry] = { link: readlinkSync(file) }
    else if (stat.isDirectory()) Object.assign(entries, { [entry]: "directory" }, tree(directory, entry))
    else if (name.endsWith(".db-shm")) entries[entry] = "sqlite shared-memory index"
    else entries[entry] = { size: stat.size, sha256: createHash("sha256").update(readFileSync(file)).digest("hex"), mtime: stat.mtimeNs }
  }
  return entries
}

/** The response to `request`, which must leave the matter's tree as it was. */
async function readOnly(directory: string, request: () => Response | Promise<Response>) {
  const before = tree(directory)
  const response = await request()
  expect(tree(directory)).toEqual(before)
  return response
}

test("without a combined document there is no combined file", async () => {
  for (const documents of [null, [SUPPLY], [{ ...COMBINED, deliverable: "body" as const }], [COMBINED, { ...SIDE, deliverable: "combined" as const }]]) {
    const f = await deliveredMatter(documents)
    const response = await f.download()
    expect(response.status).toBe(404)
    expect(await response.json()).toEqual({ error: "no_combined_output" })
    expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
  }
  // A combined document without a requested memo is delivered as its body.
  const unrequested = await deliveredMatter([COMBINED], false)
  expect((await unrequested.download()).status).toBe(404)
  expect(leanDraftingStatus(unrequested.directory)?.combined_output).toBeNull()
}, 60_000)

test("a combined document downloads as one file, memo first, from the registered files, as the file saved in the matter; nothing is written", async () => {
  const f = await deliveredMatter([COMBINED])
  expect(leanDraftingStatus(f.directory)?.combined_output).toEqual(COMBINED_OUTPUT)
  expect(readLeanDocuments(f.directory)?.documents.map(row => row.deliverable))
    .toEqual(["supply-pack - clean.docx", "supply-pack - redline.docx", "supply-pack - cover memo.docx", null, "supply-pack.docx"])
  const before = tree(f.directory)
  // The snapshot reaches below the top level: the registry and the run's files.
  expect(Object.keys(before)).toEqual(expect.arrayContaining([
    path.join(".dochaus", "legal.db"),
    path.relative(f.directory, leanFile(f.directory, "documents.json")),
    path.relative(f.directory, leanFile(f.directory, "d1/body-applied.docx")),
  ]))

  const response = await f.download()
  expect(response.status).toBe(200)
  expect(response.headers.get("content-type")).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
  expect(response.headers.get("content-disposition")).toBe(`attachment; filename="supply-pack.docx"; filename*=UTF-8''supply-pack.docx`)
  const downloaded = new Uint8Array(await response.arrayBuffer())
  const xml = documentText(downloaded)
  expect(xml.indexOf("Memo marker")).toBeGreaterThan(-1)
  expect(xml.indexOf("Memo marker")).toBeLessThan(xml.indexOf("Body marker"))

  expect(tree(f.directory)).toEqual(before)
  // The download is the combined file the run saved in the matter, byte for byte.
  expect(Object.keys(before)).toContain("supply-pack.docx")
  expect(downloaded).toEqual(new Uint8Array(readFileSync(path.join(f.directory, "supply-pack.docx"))))

  // A change accepted into the registered body lives only in that file; the
  // combined file follows it, not the run's own copy.
  const registered = readFileSync(path.join(f.directory, "supply-pack - clean.docx"))
  const revised = strToU8(documentText(registered).replace("Body marker", "Revised marker"))
  await writeDeliverableRevision(f.directory, "supply-pack - clean.docx", Buffer.from(zipSync({ ...unzipSync(registered), "word/document.xml": revised })))
  const after = documentText(new Uint8Array(await (await f.download()).arrayBuffer()))
  expect(after).toContain("Revised marker")
  expect(after).not.toContain("Body marker")
  expect(after.indexOf("Memo marker")).toBeLessThan(after.indexOf("Revised marker"))
}, 60_000)

test("the checklist's combined row names, types and builds exactly the file the download delivers", async () => {
  // Delivered without a gate, so the registered parts are the run's own copies.
  const f = await deliveredMatter([COMBINED])
  const row = deliverables.deliveryArtifactEvidence(f.directory, { build: true }).at(-1)
  if (row?.kind !== "combined") throw new Error("no combined row")
  const response = await readOnly(f.directory, f.download)
  expect(response.status).toBe(200)
  const bytes = new Uint8Array(await response.arrayBuffer())
  expect(row.download_filename).toBe("supply-pack.docx")
  expect(response.headers.get("content-disposition")).toBe(attachmentDocumentDisposition(row.download_filename))
  expect(response.headers.get("content-type")).toBe(row.content_type)
  expect(row.files_delivered).toBe(1)
  expect(row.built_at_check).toMatchObject({ sha256: createHash("sha256").update(bytes).digest("hex"), paragraphs_equal_memo_then_body: true })
  // A memo both refuse: the row carries the download's own refusal code.
  const parts = unzipSync(readFileSync(leanFile(f.directory, "drafting-notes.docx")))
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:t></w:r>", `</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`))
  const noted = Buffer.from(zipSync(parts))
  writeLeanBytes(f.directory, "drafting-notes.docx", noted)
  await writeDeliverableRevision(f.directory, "supply-pack - cover memo.docx", noted)
  const refused = deliverables.deliveryArtifactEvidence(f.directory, { build: true }).at(-1)
  const download = await readOnly(f.directory, f.download)
  expect(download.status).toBe(409)
  expect(refused).toMatchObject({ built_at_check: { refused: (await download.json()).error } })
}, 60_000)

test("a combined document without its registered notes is refused", async () => {
  const f = await deliveredMatter([COMBINED])
  const summary = readLeanDocuments(f.directory)!
  writeLeanText(f.directory, "documents.json", `${JSON.stringify({ ...summary, documents: summary.documents.filter(row => row.kind !== "drafting_notes") }, null, 2)}\n`)
  const response = await readOnly(f.directory, f.download)
  expect(response.status).toBe(409)
  expect(await response.json()).toEqual({ error: "combined_parts_missing" })
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
}, 60_000)

test("with several planned documents the memo is paired with the combined document's own body", async () => {
  const f = await deliveredMatter([SIDE, COMBINED])
  expect(leanDraftingStatus(f.directory)?.combined_output).toEqual({ ...COMBINED_OUTPUT, document: "d2" })
  const response = await f.download()
  expect(response.status).toBe(200)
  const xml = documentText(new Uint8Array(await response.arrayBuffer()))
  expect(xml).toContain("Memo marker")
  expect(xml).toContain("Body marker")
  expect(xml).not.toContain("Side marker")
}, 60_000)

/** The run's next outputs, recorded as a run stopped at its gate records them. */
async function blockedRerun(directory: string, notes: string) {
  writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx(notes)).bytes)
  writeLeanText(directory, "drafting-notes.md", `${notes}\n`)
  await recordLeanDocuments(directory, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, false)
}

async function refused(f: { directory: string; download: () => Response | Promise<Response> }) {
  const response = await f.download()
  expect(response.status).toBe(409)
  expect(await response.json()).toEqual({ error: "combined_parts_missing" })
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
}

test("combined output is reported once both parts are delivered, not from the plan alone", async () => {
  const f = await deliveredMatter(null)
  writeLeanPlan(f.directory, [COMBINED], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
  expect((await f.download()).status).toBe(409)

  const delivered = await deliveredMatter([COMBINED])
  expect(leanDraftingStatus(delivered.directory)?.combined_output).toEqual(COMBINED_OUTPUT)
  expect((await delivered.download()).status).toBe(200)
}, 60_000)

test("an earlier delivery under another plan is not served as the combined file", async () => {
  // Run 1 delivered the document as its body, without a memo; a re-outline
  // then asks for it combined with a memo under the same title and file name.
  const f = await deliveredMatter([{ ...COMBINED, deliverable: "body" }], false)
  expect(readLeanDocuments(f.directory)?.documents.map(row => row.deliverable))
    .toEqual(["supply-pack.docx", "supply-pack - redline.docx", "supply-pack - drafting notes.docx", null])
  writeLeanPlan(f.directory, [COMBINED], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  await refused(f)
  // Still refused after run 2 stops at its gate with the same body.
  await blockedRerun(f.directory, MEMO)
  await refused(f)
}, 60_000)

test("a combined file renamed after delivery is not built from the parts delivered under the old name", async () => {
  const f = await deliveredMatter([COMBINED])
  writeLeanPlan(f.directory, [{ ...COMBINED, file: "supply-final.docx" }], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  await refused(f)
  const renamed = [{ ...COMBINED, file: "supply-final.docx" }]
  writeLeanPlan(f.directory, renamed, undefined, undefined, { name: "cover memo", file: memoNotesFile(renamed, "cover memo") })
  await refused(f)
}, 60_000)

test("parts delivered for a document the plan has since retitled or moved are not paired", async () => {
  // The same file and memo under another title: the body was drafted for the old one.
  const retitled = await deliveredMatter([COMBINED])
  writeLeanPlan(retitled.directory, [{ ...COMBINED, title: "Supply schedule" }], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  await refused(retitled)
  // The same document planned second: its parts were delivered as the first.
  const moved = await deliveredMatter([COMBINED])
  const plan = [SIDE, COMBINED]
  writeLeanPlan(moved.directory, plan, undefined, undefined, { name: "cover memo", file: memoNotesFile(plan, "cover memo") })
  await refused(moved)
}, 60_000)

test("a rerun stopped at its gate with changed notes has no combined file until it is delivered", async () => {
  const f = await deliveredMatter([COMBINED])
  await blockedRerun(f.directory, `${MEMO}\n\nChanged.`)
  expect(readLeanDocuments(f.directory)?.documents.map(row => [row.kind, row.deliverable])).toContainEqual(["drafting_notes", null])
  await refused(f)
}, 60_000)

test("registered notes the combined file cannot carry are refused", async () => {
  const f = await deliveredMatter([COMBINED])
  const notes = path.join(f.directory, "supply-pack - cover memo.docx")
  const parts = unzipSync(readFileSync(notes))
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:t></w:r>", `</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`))
  await writeDeliverableRevision(f.directory, "supply-pack - cover memo.docx", Buffer.from(zipSync(parts)))
  const response = await readOnly(f.directory, f.download)
  expect(response.status).toBe(409)
  expect(await response.json()).toEqual({
    error: "unsupported_memo_dependency",
    message: "The memo holds content the combined file cannot carry. Download the draft and the memo separately.",
  })
}, 60_000)

describe("combined download hardening: refusal codes, matter boundary, lock scope and contention", () => {
  const BODY_NAME = "supply-pack - clean.docx"
  const NOTES_NAME = "supply-pack - cover memo.docx"

  /** Holds `names`' mutation locks from an async context of its own, so a
   * request started by the test still waits for them, until `release`; then
   * runs `inside` under the held locks before letting them go. */
  async function holdLocks(directory: string, names: string[], inside: () => Promise<void> = async () => {}) {
    let release!: () => void
    const released = new Promise<void>(resolve => { release = resolve })
    let held!: () => void
    const holding = new Promise<void>(resolve => { held = resolve })
    const done = withArtifactMutationLocksAsync(directory, names, async () => { held(); await released; await inside() })
    await holding
    return { release: () => { release(); return done } }
  }

  const UNREADABLE = {
    memo: "The memo could not be read as a Word document. Open it in Word, save it again, and retry.",
    body: "The draft could not be read as a Word document. Open it in Word, save it again, and retry.",
  }

  test("each refusal reports its own code, and an unreadable part is named", async () => {
    const empty = await deliveredMatter([COMBINED])
    rewriteDocumentXml(empty.directory, NOTES_NAME, xml => xml.replace(/(<w:t(?:\s[^>]*)?>)[^<]*(<\/w:t>)/g, "$1$2"))
    const refusedEmpty = await readOnly(empty.directory, empty.download)
    expect(refusedEmpty.status).toBe(409)
    expect(await refusedEmpty.json()).toEqual({ error: "empty_memo", message: "The memo has no text to place before the draft." })

    const broken = await deliveredMatter([COMBINED])
    rewriteDocumentXml(broken.directory, NOTES_NAME, xml => xml.replace("<w:body>", "<w:body><w:p"))
    const unreadableMemo = await readOnly(broken.directory, broken.download)
    expect(unreadableMemo.status).toBe(409)
    expect(await unreadableMemo.json()).toEqual({ error: "invalid_docx_xml", part: "memo", message: UNREADABLE.memo })

    const brokenBody = await deliveredMatter([COMBINED])
    rewriteDocumentXml(brokenBody.directory, BODY_NAME, xml => xml.replace("<w:body>", "<w:body><w:p"))
    const unreadableBodyXml = await readOnly(brokenBody.directory, brokenBody.download)
    expect(unreadableBodyXml.status).toBe(409)
    expect(await unreadableBodyXml.json()).toEqual({ error: "invalid_docx_xml", part: "body", message: UNREADABLE.body })

    const notWord = await deliveredMatter([COMBINED])
    writeFileSync(path.join(notWord.directory, BODY_NAME), "not a Word package")
    const unreadableBody = await readOnly(notWord.directory, notWord.download)
    expect(unreadableBody.status).toBe(409)
    const refusal = await unreadableBody.json()
    expect(refusal).toEqual({ error: "invalid_docx_package", part: "body", message: UNREADABLE.body })
    expect(JSON.stringify(refusal)).not.toContain(notWord.directory)

    // A zip without a document part: the composer's own package code.
    const noDocument = await deliveredMatter([COMBINED])
    const package_ = unzipSync(readFileSync(path.join(noDocument.directory, NOTES_NAME)))
    delete package_["word/document.xml"]
    writeFileSync(path.join(noDocument.directory, NOTES_NAME), zipSync(package_))
    const packageMemo = await readOnly(noDocument.directory, noDocument.download)
    expect(packageMemo.status).toBe(409)
    expect(await packageMemo.json()).toEqual({ error: "invalid_docx_package", part: "memo", message: UNREADABLE.memo })
  }, 60_000)

  test("an unreadable memo is named even when the draft could not itself serve as a memo", async () => {
    // A note reference is fine in the draft but refused in a memo, so the
    // draft composed on its own fails too, for a reason other than reading it.
    const f = await deliveredMatter([COMBINED])
    rewriteDocumentXml(f.directory, BODY_NAME, xml => xml.replace("</w:t></w:r>", `</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`))
    const draft = new Uint8Array(readFileSync(path.join(f.directory, BODY_NAME)))
    expect(() => combinedDocx.combineMemoAndBody({ memo: draft, body: draft })).toThrow("unsupported_memo_dependency")
    rewriteDocumentXml(f.directory, NOTES_NAME, xml => xml.replace("<w:body>", "<w:body><w:p"))
    const response = await readOnly(f.directory, f.download)
    expect(response.status).toBe(409)
    expect(await response.json()).toEqual({ error: "invalid_docx_xml", part: "memo", message: UNREADABLE.memo })
  }, 60_000)

  test("memo markup the combined file cannot carry is refused with its own reason", async () => {
    // A memo list whose numbering links to a list style.
    const numbered = await deliveredMatter([COMBINED])
    const listed = unzipSync((await renderMarkdownToDocx("# Cover memo\n\n- Memo marker: the start date needs confirmation.")).bytes)
    listed["word/numbering.xml"] = strToU8(strFromU8(listed["word/numbering.xml"]!).replace(/<w:abstractNum\b[^>]*>/, `$&<w:numStyleLink w:val="MemoList"/>`))
    writeFileSync(path.join(numbered.directory, NOTES_NAME), zipSync(listed))
    const numbering = await readOnly(numbered.directory, numbered.download)
    expect(numbering.status).toBe(409)
    expect(await numbering.json()).toEqual({
      error: "unsupported_memo_numbering",
      message: "The memo uses list numbering the combined file cannot carry. Download the draft and the memo separately.",
    })

    // A memo prefix the draft binds to another namespace.
    const spaced = await deliveredMatter([COMBINED])
    rewriteDocumentXml(spaced.directory, NOTES_NAME, xml => xml.replace("<w:document ", `<w:document xmlns:zz="urn:example:memo" `).replace("<w:p>", `<w:p zz:mark="1">`))
    rewriteDocumentXml(spaced.directory, BODY_NAME, xml => xml.replace("<w:document ", `<w:document xmlns:zz="urn:example:draft" `))
    const namespace = await readOnly(spaced.directory, spaced.download)
    expect(namespace.status).toBe(409)
    expect(await namespace.json()).toEqual({
      error: "unsupported_memo_namespace",
      message: "The memo uses Word markup the combined file cannot carry. Download the draft and the memo separately.",
    })
  }, 60_000)

  test("a build that fails for any other reason is refused in plain words, not as a server error", async () => {
    const f = await deliveredMatter([COMBINED])
    const failing = spyOn(combinedDocx, "buildCombinedFile").mockReturnValue({ refused: "build_failed" })
    try {
      const response = await readOnly(f.directory, f.download)
      expect({ status: response.status, body: await response.json() }).toEqual({ status: 409, body: {
        error: "build_failed",
        message: "The combined file could not be built. Download the draft and the memo separately.",
      } })
    } finally {
      failing.mockRestore()
    }
  }, 60_000)

  test("part names other than plain matter file names are refused before any read", async () => {
    // combinedParts only returns names registration gives, which are plain;
    // the route holds its parts to plain names on its own. A relative or
    // absolute alias of a registered part is not one.
    const f = await deliveredMatter([COMBINED])
    const parts = deliverables.combinedParts(f.directory)!
    expect(parts).toEqual({ document: "d1", name: "supply-pack", file: "supply-pack.docx", body: BODY_NAME, notes: NOTES_NAME })
    for (const part of ["body", "notes"] as const) {
      for (const name of [`./${parts[part]}`, `sub/../${parts[part]}`, path.join(f.directory, parts[part]), "..", `../${path.basename(f.directory)}/${parts[part]}`]) {
        const stand = spyOn(deliverables, "combinedParts").mockReturnValue({ ...parts, [part]: name })
        try {
          const response = await readOnly(f.directory, f.download)
          expect({ name, status: response.status, body: await response.json() }).toEqual({ name, status: 409, body: { error: "combined_parts_missing" } })
        } finally {
          stand.mockRestore()
        }
      }
    }
    expect((await f.download()).status).toBe(200)
  }, 60_000)

  test("a part that is gone, that a source now holds, or whose registry is gone is refused without writing", async () => {
    const missing = async (f: { directory: string; download: () => Response | Promise<Response> }) => {
      const response = await readOnly(f.directory, f.download)
      expect(response.status).toBe(409)
      expect(await response.json()).toEqual({ error: "combined_parts_missing" })
    }
    const gone = await deliveredMatter([COMBINED])
    unlinkSync(path.join(gone.directory, NOTES_NAME))
    await missing(gone)

    const taken = await deliveredMatter([COMBINED])
    const db = openDb(taken.directory)
    try {
      expect(db.run("UPDATE documents SET role = 'source' WHERE name = ?", [BODY_NAME]).changes).toBe(1)
    } finally {
      db.close()
    }
    await missing(taken)

    // Without a registry nothing is registered, and none is created.
    const unregistered = await deliveredMatter([COMBINED])
    for (const file of ["legal.db", "legal.db-wal", "legal.db-shm"]) rmSync(path.join(unregistered.directory, ".dochaus", file), { force: true })
    await missing(unregistered)
  }, 60_000)

  test("a part named outside the matter, or linked out of it, is refused and never read", async () => {
    for (const name of ["..", `../${BODY_NAME}`, `sub/${BODY_NAME}`]) {
      const f = await deliveredMatter([COMBINED])
      const summary = readLeanDocuments(f.directory)!
      writeLeanText(f.directory, "documents.json", `${JSON.stringify({ ...summary, documents: summary.documents.map(row => row.kind === "body" ? { ...row, deliverable: name } : row) }, null, 2)}\n`)
      const response = await f.download()
      expect(response.status).toBe(409)
      expect(await response.json()).toEqual({ error: "combined_parts_missing" })
    }
    const f = await deliveredMatter([COMBINED])
    const outside = mkdtempSync(path.join(tmpdir(), "dochaus-combined-outside-"))
    directories.push(outside)
    const elsewhere = path.join(outside, "elsewhere.docx")
    writeFileSync(elsewhere, (await renderMarkdownToDocx("# Elsewhere\n\nOutside marker.")).bytes)
    unlinkSync(path.join(f.directory, BODY_NAME))
    symlinkSync(elsewhere, path.join(f.directory, BODY_NAME))
    const response = await readOnly(f.directory, f.download)
    expect(response.status).toBe(409)
    expect(await response.text()).toBe(JSON.stringify({ error: "combined_parts_missing" }))

    // A link to another registered file in the matter is no alias for the part.
    const alias = await deliveredMatter([COMBINED])
    unlinkSync(path.join(alias.directory, BODY_NAME))
    symlinkSync("supply-pack - redline.docx", path.join(alias.directory, BODY_NAME))
    const aliased = await readOnly(alias.directory, alias.download)
    expect(aliased.status).toBe(409)
    expect(await aliased.json()).toEqual({ error: "combined_parts_missing" })
  }, 60_000)

  test("a download waits for a write in progress on either part, then serves what it wrote", async () => {
    for (const [name, marker, revision] of [[BODY_NAME, "Body marker", "Revised marker"], [NOTES_NAME, "Memo marker", "Revised memo line"]] as const) {
      const f = await deliveredMatter([COMBINED])
      const registered = readFileSync(path.join(f.directory, name))
      const revised = Buffer.from(zipSync({ ...unzipSync(registered), "word/document.xml": strToU8(documentText(registered).replace(marker, revision)) }))
      const hold = await holdLocks(f.directory, [name], async () => { await writeDeliverableRevision(f.directory, name, revised) })
      let settled = false
      const download = Promise.resolve(f.download()).then(response => { settled = true; return response })
      await Bun.sleep(300)
      expect({ name, settled }).toEqual({ name, settled: false })
      await hold.release()
      const response = await download
      expect(response.status).toBe(200)
      const xml = documentText(new Uint8Array(await response.arrayBuffer()))
      expect(xml).toContain(revision)
      expect(xml).not.toContain(marker)
    }
  }, 60_000)

  test("a part held past the lock wait is reported busy and retryable", async () => {
    for (const name of [BODY_NAME, NOTES_NAME]) {
      const f = await deliveredMatter([COMBINED])
      const hold = await holdLocks(f.directory, [name])
      try {
        const before = tree(f.directory)
        const download = Promise.resolve(f.download())
        await Bun.sleep(200)
        // Past the directory lock's 30 s wait without sleeping through it.
        setSystemTime(new Date(Date.now() + 31_000))
        const response = await download
        expect({ name, status: response.status }).toEqual({ name, status: 503 })
        expect(response.headers.get("retry-after")).toBe("5")
        expect(await response.json()).toEqual({ error: "combined_parts_busy", message: "The draft or the memo is being updated. Try again in a moment." })
        expect(tree(f.directory)).toEqual(before)
      } finally {
        setSystemTime()
        await hold.release()
      }
      expect((await f.download()).status).toBe(200)
    }
  }, 60_000)

  test("the parts are composed after their locks are released", async () => {
    const f = await deliveredMatter([COMBINED])
    // A held lock shows as an entry under the lock root.
    const locks = path.join(f.directory, ".dochaus", "artifact-mutation-locks")
    const hold = await holdLocks(f.directory, [BODY_NAME, NOTES_NAME])
    expect(readdirSync(locks)).toHaveLength(2)
    await hold.release()
    expect(readdirSync(locks)).toEqual([])

    const compose = combinedDocx.buildCombinedFile
    const held: string[][] = []
    const watched = spyOn(combinedDocx, "buildCombinedFile").mockImplementation(input => {
      held.push(readdirSync(locks))
      return compose(input)
    })
    try {
      expect((await f.download()).status).toBe(200)
      expect(held).toEqual([[]])
      // A refusal that names its unreadable part is built once too, also unlocked.
      rewriteDocumentXml(f.directory, NOTES_NAME, xml => xml.replace("<w:body>", "<w:body><w:p"))
      held.length = 0
      expect((await f.download()).status).toBe(409)
      expect(held).toEqual([[]])
    } finally {
      watched.mockRestore()
    }
  }, 60_000)

  /** Runs `run` against an ingest server of its own, on a workspace of its own. */
  async function withIngest(run: (request: (route: string, init?: RequestInit) => Promise<Response>) => Promise<void>) {
    const temporary = mkdtempSync(path.join(tmpdir(), "dochaus-combined-accept-"))
    directories.push(temporary)
    const workspace = path.join(temporary, "workspace")
    const config = path.join(temporary, "config")
    await Bun.write(path.join(config, "opencode.json"), "{}")
    await Bun.write(path.join(workspace, ".keep"), "")
    const engine = Bun.serve({ port: 0, fetch: () => new Response(null, { status: 503 }) })
    const probe = Bun.serve({ port: 0, fetch: () => new Response() })
    const port = probe.port!
    probe.stop(true)
    const log = path.join(temporary, "ingest.log")
    const ingest = Bun.spawn([process.execPath, "run", "src/server.ts"], {
      cwd: path.resolve(import.meta.dir, ".."),
      env: { ...process.env, WORKSPACE_ROOT: workspace, DOCHAUS_DIR: config, INGEST_PORT: String(port), INGEST_HOST: "127.0.0.1", OPENCODE_URL: `http://127.0.0.1:${engine.port}`, OPENCODE_SERVER_PASSWORD: "" },
      stdout: Bun.file(log), stderr: Bun.file(log),
    })
    const request = (route: string, init?: RequestInit) => fetch(`http://127.0.0.1:${port}${route}`, { ...init, signal: AbortSignal.timeout(45_000) })
    try {
      const deadline = Date.now() + 20_000
      while (!await request("/matters").then(response => response.ok).catch(() => false)) {
        if (ingest.exitCode !== null || Date.now() > deadline) throw new Error(`ingest did not start:\n${await Bun.file(log).text()}`)
        await Bun.sleep(50)
      }
      await run(request)
    } finally {
      ingest.kill()
      await ingest.exited
      engine.stop(true)
    }
  }

  /** A new matter whose completed lean run delivered `documents` with the
   * requested memo, and the registered name of the first one's body. */
  async function completedMatter(request: (route: string, init?: RequestInit) => Promise<Response>, documents: Document[] = [COMBINED]) {
    const matter = await (await request("/matters", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ title: "Combined accept" }) })).json() as { id: string; dir: string }
    initLeanState(matter.dir, { request: "Draft a short supply note from the supply terms.", ownerSessionID: "ses_owner", model: null })
    writeLeanPlan(matter.dir, documents, undefined, undefined, { name: "cover memo", file: memoNotesFile(documents, "cover memo") })
    for (const index of documents.keys()) {
      const rendered = (await renderMarkdownToDocx(BODY)).bytes
      writeLeanBytes(matter.dir, `d${index + 1}/body-applied.docx`, rendered)
      writeLeanBytes(matter.dir, `d${index + 1}/redline.docx`, rendered)
    }
    writeLeanBytes(matter.dir, "drafting-notes.docx", (await renderMarkdownToDocx(MEMO)).bytes)
    writeLeanText(matter.dir, "drafting-notes.md", `${MEMO}\n`)
    await recordLeanDocuments(matter.dir, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
    writeLeanState(matter.dir, { ...readLeanState(matter.dir)!, status: "completed" })
    const body = readLeanDocuments(matter.dir)!.documents.find(row => row.kind === "body")!.deliverable!
    return { ...matter, route: `/matters/${matter.id}`, body }
  }

  /** A change proposed on `name` for the lawyer to accept. */
  function propose(directory: string, name: string, find = "delivers the goods", replacement = "delivers the goods within ten days") {
    const db = openDb(directory)
    try {
      return Number(db.run("INSERT INTO redlines (doc_path,doc_name,scope,find_text,old_text,new_text,author,review_comment,anchor_id,created_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
        [path.join(directory, name), name, "phrase", find, find, replacement, "Review", "Agreed delivery time.", null, Date.now()]).lastInsertRowid)
    } finally { db.close() }
  }

  const ADDENDUM = "Changes accepted after these notes were prepared"
  const SAVED = "supply-pack.docx"

  test("a Word accept on the body holds the notes lock across the body write and the notes addendum, and rebuilds the saved combined file", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const route = matter.route
      const bodyFile = path.join(matter.dir, BODY_NAME)
      const id = propose(matter.dir, BODY_NAME)
      const prepared = readFileSync(bodyFile)
      const savedBefore = readFileSync(path.join(matter.dir, SAVED))

      // While the notes are held, the accept writes nothing, not even the body.
      const hold = await holdLocks(matter.dir, [NOTES_NAME])
      const accepting = request(`${route}/redlines/${id}/accept`, { method: "POST" })
      const downloading = request(`${route}/documents/combined`)
      for (const until = Date.now() + 3_000; Date.now() < until;) {
        expect(readFileSync(bodyFile)).toEqual(prepared)
        expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(savedBefore)
        await Bun.sleep(50)
      }
      await hold.release()
      const accepted = await accepting
      // This run records no change list, so its redline cannot be rebuilt; the accept says so.
      expect({ status: accepted.status, body: await accepted.json() })
        .toEqual({ status: 200, body: { ok: true, notes_updated: true, redline_updated: false, redline_skip_reason: "rebuild_failed", combined_updated: true } })

      // The download that waited with it pairs the body and the notes from one side of the accept.
      const waited = await downloading
      expect(waited.status).toBe(200)
      const pair = documentText(new Uint8Array(await waited.arrayBuffer()))
      expect(pair.includes("within ten days")).toBe(pair.includes(ADDENDUM))

      const after = await request(`${route}/documents/combined`)
      expect(after.status).toBe(200)
      const bytes = new Uint8Array(await after.arrayBuffer())
      const xml = documentText(bytes)
      expect(xml).toContain(ADDENDUM)
      expect(xml.indexOf(ADDENDUM)).toBeLessThan(xml.indexOf("within ten days"))
      // The saved combined file is the download, byte for byte.
      expect(bytes).toEqual(new Uint8Array(readFileSync(path.join(matter.dir, SAVED))))
    })
  }, 120_000)

  test("accepting every change on the body rebuilds the saved combined file too", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      propose(matter.dir, BODY_NAME)
      const accepted = await request(`${matter.route}/redlines/accept-all?name=${encodeURIComponent(BODY_NAME)}`, { method: "POST" })
      expect({ status: accepted.status, body: await accepted.json() }).toEqual({ status: 200, body: {
        ok: true, accepted: 1, notes_updated: true, redline_updated: false, redline_skip_reason: "rebuild_failed", combined_updated: true,
      } })
      const download = await request(`${matter.route}/documents/combined`)
      expect(download.status).toBe(200)
      const bytes = new Uint8Array(await download.arrayBuffer())
      expect(documentText(bytes)).toContain("within ten days")
      expect(bytes).toEqual(new Uint8Array(readFileSync(path.join(matter.dir, SAVED))))
    })
  }, 120_000)

  test("a combined file saved under a numbered name is the one an accept rebuilds and the download names", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      await request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
      const accepted = readFileSync(path.join(matter.dir, SAVED))
      // A re-delivery keeps the accept's combined file and saves the new one under the next free name.
      await recordLeanDocuments(matter.dir, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
      const numbered = "supply-pack (2).docx"
      const status = leanDraftingStatus(matter.dir)?.combined_output
      expect(status?.file).toBe(numbered)
      const body = status!.body
      const response = await request(`${matter.route}/redlines/${propose(matter.dir, body)}/accept`, { method: "POST" })
      expect(await response.json()).toMatchObject({ ok: true, combined_updated: true })
      const download = await request(`${matter.route}/documents/combined`)
      expect(download.headers.get("content-disposition")).toBe(attachmentDocumentDisposition(numbered))
      const bytes = new Uint8Array(await download.arrayBuffer())
      expect(documentText(bytes)).toContain("within ten days")
      expect(bytes).toEqual(new Uint8Array(readFileSync(path.join(matter.dir, numbered))))
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(accepted)
    })
  }, 120_000)

  test("an accept never writes over a combined file changed since it was saved", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const saved = path.join(matter.dir, SAVED)
      rewriteDocumentXml(matter.dir, SAVED, xml => xml.replace("Memo marker", "Lawyer marker"))
      const changed = readFileSync(saved)
      const accepted = await request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
      expect(await accepted.json()).toMatchObject({ ok: true, combined_updated: false, combined_skip_reason: "combined_changed" })
      expect(readFileSync(saved)).toEqual(changed)
    })
  }, 120_000)

  test("an accept whose combined file cannot be rebuilt says so, and the stale file is never served as current", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      // Registered notes the combined file cannot carry.
      const notes = readFileSync(path.join(matter.dir, NOTES_NAME))
      const parts = unzipSync(notes)
      parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace("</w:t></w:r>", `</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`))
      await writeDeliverableRevision(matter.dir, NOTES_NAME, Buffer.from(zipSync(parts)))
      const saved = readFileSync(path.join(matter.dir, SAVED))
      const id = propose(matter.dir, BODY_NAME)
      const accepted = await request(`${matter.route}/redlines/${id}/accept`, { method: "POST" })
      expect(accepted.status).toBe(200)
      expect(await accepted.json()).toMatchObject({ ok: true, combined_updated: false, combined_skip_reason: "build_refused" })
      // The accepted body stands; the file saved before it is left as it was, and the download refuses rather than serve it.
      expect(documentText(readFileSync(path.join(matter.dir, BODY_NAME)))).toContain("within ten days")
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
      const download = await request(`${matter.route}/documents/combined`)
      expect({ status: download.status, body: await download.json() }).toEqual({ status: 409, body: {
        error: "unsupported_memo_dependency",
        message: "The memo holds content the combined file cannot carry. Download the draft and the memo separately.",
      } })
    })
  }, 120_000)

  test("a new revision of a part uploaded in place rebuilds the saved combined file", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const form = new FormData()
      form.append("file", new File([(await renderMarkdownToDocx(BODY.replace("delivers the goods", "delivers the goods by rail"))).bytes], BODY_NAME))
      const uploaded = await request(`${matter.route}/documents/revisions`, { method: "POST", body: form })
      expect({ status: uploaded.status, body: await uploaded.json() }).toMatchObject({ status: 200, body: { combined_updated: true } })
      const download = await request(`${matter.route}/documents/combined`)
      expect(download.status).toBe(200)
      const bytes = new Uint8Array(await download.arrayBuffer())
      expect(documentText(bytes)).toContain("by rail")
      expect(bytes).toEqual(new Uint8Array(readFileSync(path.join(matter.dir, SAVED))))
    })
  }, 120_000)

  test("a run that is not completed has its saved combined file left as it is, like its redline and notes", async () => {
    await withIngest(async request => {
      for (const status of ["running", "blocked", "failed"] as const) {
        const matter = await completedMatter(request)
        writeLeanState(matter.dir, { ...readLeanState(matter.dir)!, status })
        const saved = readFileSync(path.join(matter.dir, SAVED))
        const accepted = await request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
        expect({ status, code: accepted.status, body: await accepted.json() }).toEqual({ status, code: 200, body: { ok: true } })
        expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
      }
    })
  }, 120_000)

  test("an accept on a delivery that is not combined is answered as before", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request, [SIDE])
      const id = propose(matter.dir, matter.body)
      const accepted = await request(`${matter.route}/redlines/${id}/accept`, { method: "POST" })
      expect({ status: accepted.status, body: await accepted.json() })
        .toEqual({ status: 200, body: { ok: true, notes_updated: true, redline_updated: false, redline_skip_reason: "rebuild_failed" } })
      expect((await request(`${matter.route}/documents/combined`)).status).toBe(404)
    })
  }, 120_000)
})
