import { afterEach, describe, expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { withParagraphPointers } from "../../../dochaus/lib/drafting-amendment"
import { renderMarkdownToDocx } from "../../../dochaus/lib/render"
import * as deliverables from "../../../dochaus/workflows/lean/deliverables"
import { memoNotesFile, readLeanDocuments, recordLeanDocuments } from "../../../dochaus/workflows/lean/deliverables"
import { writeLeanPlan, type PlannedDocument } from "../../../dochaus/workflows/lean/plan"
import { initLeanState, readLeanState, writeLeanBytes, writeLeanState, writeLeanText } from "../../../dochaus/workflows/lean/state"
import { leanFixtureMatter, scriptedLeanModel } from "../../../dochaus/workflows/lean/test-support"
import { ownedDeliverable, readSystemDocx, systemDocxKey } from "../../../dochaus/lib/system-docx"
import { pendingRedlinesForDoc, recordRedline } from "../../../dochaus/lib/redlines"
import { getRedline, openDb, setRedlineStatus } from "./db"
import { writeDeliverableRevision } from "./deliverable"
import { leanDraftingStatus } from "./lean-drafting"
import { registerLeanDeliverable } from "./lean-deliverables"

// The combined file: a document planned as "combined" with a requested memo is
// delivered as the one file its delivery saved and registered; an accept into
// its body or notes rebuilds it while the system's last write of it is
// unchanged. Real registration through the matter registry; nothing is
// scripted but the model port, which registration never calls.

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
  return { directory }
}

/** The combined file the matter's status reports, as saved in the matter. */
function savedCombined(directory: string) {
  const file = leanDraftingStatus(directory)?.combined_output?.file
  if (!file) throw new Error("no combined file is reported")
  return new Uint8Array(readFileSync(path.join(directory, file)))
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

test("without a combined document there is no combined file", async () => {
  for (const documents of [null, [SUPPLY], [{ ...COMBINED, deliverable: "body" as const }], [COMBINED, { ...SIDE, deliverable: "combined" as const }]]) {
    const f = await deliveredMatter(documents)
    expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
  }
  // A combined document without a requested memo is delivered as its body.
  const unrequested = await deliveredMatter([COMBINED], false)
  expect(leanDraftingStatus(unrequested.directory)?.combined_output).toBeNull()
}, 60_000)

test("a combined document is saved in the matter as one file, memo first, and registered", async () => {
  const f = await deliveredMatter([COMBINED])
  expect(leanDraftingStatus(f.directory)?.combined_output).toEqual(COMBINED_OUTPUT)
  expect(readLeanDocuments(f.directory)?.documents.map(row => row.deliverable))
    .toEqual(["supply-pack - clean.docx", "supply-pack - redline.docx", "supply-pack - cover memo.docx", null, "supply-pack.docx"])
  const xml = documentText(savedCombined(f.directory))
  expect(xml.indexOf("Memo marker")).toBeGreaterThan(-1)
  expect(xml.indexOf("Memo marker")).toBeLessThan(xml.indexOf("Body marker"))
}, 60_000)

test("the checklist's combined row names, types and builds exactly the file saved in the matter", async () => {
  // Delivered without a gate, so the registered parts are the run's own copies.
  const f = await deliveredMatter([COMBINED])
  const row = deliverables.deliveryArtifactEvidence(f.directory, { build: true }).at(-1)
  if (row?.kind !== "combined") throw new Error("no combined row")
  expect(row.download_filename).toBe("supply-pack.docx")
  expect(row.content_type).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
  expect(row.files_delivered).toBe(1)
  expect(row.built_at_check).toMatchObject({ sha256: createHash("sha256").update(savedCombined(f.directory)).digest("hex"), paragraphs_equal_memo_then_body: true })
}, 60_000)

test("a combined document without its registered notes offers no combined file", async () => {
  const f = await deliveredMatter([COMBINED])
  const summary = readLeanDocuments(f.directory)!
  writeLeanText(f.directory, "documents.json", `${JSON.stringify({ ...summary, documents: summary.documents.filter(row => row.kind !== "drafting_notes") }, null, 2)}\n`)
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
}, 60_000)

test("with several planned documents the memo is paired with the combined document's own body", async () => {
  const f = await deliveredMatter([SIDE, COMBINED])
  expect(leanDraftingStatus(f.directory)?.combined_output).toEqual({ ...COMBINED_OUTPUT, document: "d2" })
  const xml = documentText(savedCombined(f.directory))
  expect(xml).toContain("Memo marker")
  expect(xml).toContain("Body marker")
  expect(xml).not.toContain("Side marker")
}, 60_000)

test("the delivery records the bytes it saved as the combined file, and any other write removes that record", async () => {
  const f = await deliveredMatter([COMBINED])
  const file = path.join(f.directory, "supply-pack.docx")
  const recorded = readSystemDocx(f.directory, "supply-pack.docx")
  expect(recorded).toEqual({ version: 1, sha256: createHash("sha256").update(readFileSync(file)).digest("hex") })
  expect(ownedDeliverable(f.directory, "supply-pack.docx")).toBe(true)
  // A change written into the saved file itself is kept, and the file is then no longer the system's.
  rewriteDocumentXml(f.directory, "supply-pack.docx", xml => xml.replace("Memo marker", "Lawyer marker"))
  await writeDeliverableRevision(f.directory, "supply-pack.docx", readFileSync(file))
  expect(readSystemDocx(f.directory, "supply-pack.docx")).toBeUndefined()
  expect(ownedDeliverable(f.directory, "supply-pack.docx")).toBe(false)
  const saved = savedCombined(f.directory)
  expect(documentText(saved)).toContain("Lawyer marker")
  // A part written outside an accept leaves the saved file as it was.
  rewriteDocumentXml(f.directory, "supply-pack - clean.docx", xml => xml.replace("Body marker", "Revised marker"))
  expect(savedCombined(f.directory)).toEqual(saved)
}, 60_000)

test("without a registered combined row nothing is offered, and no file is guessed from the name", async () => {
  const f = await deliveredMatter([COMBINED])
  const summary = readLeanDocuments(f.directory)!
  writeLeanText(f.directory, "documents.json", `${JSON.stringify({ ...summary, documents: summary.documents.filter(row => row.kind !== "combined") }, null, 2)}\n`)
  // The file the run saved is still in the matter under the name the plan gives it.
  expect(readFileSync(path.join(f.directory, "supply-pack.docx")).length).toBeGreaterThan(0)
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
}, 60_000)

/** The run's next outputs, recorded as a run stopped at its gate records them. */
async function blockedRerun(directory: string, notes: string) {
  writeLeanBytes(directory, "drafting-notes.docx", (await renderMarkdownToDocx(notes)).bytes)
  writeLeanText(directory, "drafting-notes.md", `${notes}\n`)
  await recordLeanDocuments(directory, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, false)
}

function refused(f: { directory: string }) {
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()
}

test("combined output is reported once both parts are delivered, not from the plan alone", async () => {
  const f = await deliveredMatter(null)
  writeLeanPlan(f.directory, [COMBINED], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  expect(leanDraftingStatus(f.directory)?.combined_output).toBeNull()

  const delivered = await deliveredMatter([COMBINED])
  expect(leanDraftingStatus(delivered.directory)?.combined_output).toEqual(COMBINED_OUTPUT)
}, 60_000)

test("an earlier delivery under another plan is not reported as the combined file", async () => {
  // Run 1 delivered the document as its body, without a memo; a re-outline
  // then asks for it combined with a memo under the same title and file name.
  const f = await deliveredMatter([{ ...COMBINED, deliverable: "body" }], false)
  expect(readLeanDocuments(f.directory)?.documents.map(row => row.deliverable))
    .toEqual(["supply-pack.docx", "supply-pack - redline.docx", "supply-pack - drafting notes.docx", null])
  writeLeanPlan(f.directory, [COMBINED], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  refused(f)
  // Still refused after run 2 stops at its gate with the same body.
  await blockedRerun(f.directory, MEMO)
  refused(f)
}, 60_000)

test("a combined file renamed after delivery is not reported from the file delivered under the old name", async () => {
  const f = await deliveredMatter([COMBINED])
  writeLeanPlan(f.directory, [{ ...COMBINED, file: "supply-final.docx" }], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  refused(f)
  const renamed = [{ ...COMBINED, file: "supply-final.docx" }]
  writeLeanPlan(f.directory, renamed, undefined, undefined, { name: "cover memo", file: memoNotesFile(renamed, "cover memo") })
  refused(f)
}, 60_000)

test("parts delivered for a document the plan has since retitled or moved are not paired", async () => {
  // The same file and memo under another title: the body was drafted for the old one.
  const retitled = await deliveredMatter([COMBINED])
  writeLeanPlan(retitled.directory, [{ ...COMBINED, title: "Supply schedule" }], undefined, undefined, { name: "cover memo", file: memoNotesFile([COMBINED], "cover memo") })
  refused(retitled)
  // The same document planned second: its parts were delivered as the first.
  const moved = await deliveredMatter([COMBINED])
  const plan = [SIDE, COMBINED]
  writeLeanPlan(moved.directory, plan, undefined, undefined, { name: "cover memo", file: memoNotesFile(plan, "cover memo") })
  refused(moved)
}, 60_000)

test("a rerun stopped at its gate with changed notes has no combined file until it is delivered", async () => {
  const f = await deliveredMatter([COMBINED])
  await blockedRerun(f.directory, `${MEMO}\n\nChanged.`)
  expect(readLeanDocuments(f.directory)?.documents.map(row => [row.kind, row.deliverable])).toContainEqual(["drafting_notes", null])
  refused(f)
}, 60_000)

test("a body planned under the combined file's name takes a free name and leaves the combined file and its proposals", async () => {
  const f = await deliveredMatter([COMBINED])
  const saved = () => createHash("sha256").update(readFileSync(path.join(f.directory, "supply-pack.docx"))).digest("hex")
  const before = saved()
  recordRedline(f.directory, { docPath: "supply-pack.docx", docName: "supply-pack.docx", scope: "phrase", findText: "goods", oldText: "goods", newText: "products", author: "reviewer", reviewComment: "", anchorId: "p000000000001" })
  const body = [{ ...COMBINED, deliverable: "body" as const }]
  writeLeanPlan(f.directory, body, undefined, undefined, { name: "cover memo", file: memoNotesFile(body, "cover memo") })
  deliverables.forgetUnplannedDocuments(f.directory)
  await recordLeanDocuments(f.directory, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
  expect(saved()).toBe(before)
  expect(pendingRedlinesForDoc(f.directory, "supply-pack.docx")).toHaveLength(1)
  expect(readLeanDocuments(f.directory)?.documents.find(row => row.kind === "body")?.deliverable).toBe("supply-pack (2).docx")
}, 60_000)

describe("combined file rebuild: lock, contention and ownership", () => {
  const BODY_NAME = "supply-pack - clean.docx"
  const NOTES_NAME = "supply-pack - cover memo.docx"
  const SAVED = "supply-pack.docx"

  /** Holds `names`' mutation locks from an async context of its own, so a
   * request started by the test still waits for them, until `release`. */
  async function holdLocks(directory: string, names: string[]) {
    let release!: () => void
    const released = new Promise<void>(resolve => { release = resolve })
    let held!: () => void
    const holding = new Promise<void>(resolve => { held = resolve })
    const done = withArtifactMutationLocksAsync(directory, names, async () => { held(); await released })
    await holding
    return { release: () => { release(); return done } }
  }

  test("names other than plain matter file names give no combined file", async () => {
    // combinedParts only returns names registration gives, which are plain;
    // registeredCombinedParts holds every name to a plain one on its own. A
    // relative or absolute alias of a registered file is not one.
    const f = await deliveredMatter([COMBINED])
    const parts = deliverables.combinedParts(f.directory)!
    expect(parts).toEqual({ document: "d1", name: "supply-pack", file: SAVED, body: BODY_NAME, notes: NOTES_NAME })
    for (const part of ["file", "body", "notes"] as const) {
      for (const name of [`./${parts[part]}`, `sub/../${parts[part]}`, path.join(f.directory, parts[part]), "..", `../${path.basename(f.directory)}/${parts[part]}`]) {
        const stand = spyOn(deliverables, "combinedParts").mockReturnValue({ ...parts, [part]: name })
        try {
          expect({ name, output: leanDraftingStatus(f.directory)?.combined_output }).toEqual({ name, output: null })
        } finally {
          stand.mockRestore()
        }
      }
    }
    // A combined row registered under a name outside the matter is no row for it.
    for (const name of ["..", `../${SAVED}`, `sub/${SAVED}`]) {
      const outside = await deliveredMatter([COMBINED])
      const summary = readLeanDocuments(outside.directory)!
      writeLeanText(outside.directory, "documents.json", `${JSON.stringify({ ...summary, documents: summary.documents.map(row => row.kind === "combined" ? { ...row, deliverable: name } : row) }, null, 2)}\n`)
      expect({ name, output: leanDraftingStatus(outside.directory)?.combined_output }).toEqual({ name, output: null })
    }
    expect(leanDraftingStatus(f.directory)?.combined_output).toEqual(COMBINED_OUTPUT)
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

      // The rebuilt file pairs the accepted body with the notes that record it, memo first.
      const xml = documentText(readFileSync(path.join(matter.dir, SAVED)))
      expect(xml).toContain(ADDENDUM)
      expect(xml.indexOf(ADDENDUM)).toBeLessThan(xml.indexOf("within ten days"))
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
      expect(documentText(readFileSync(path.join(matter.dir, SAVED)))).toContain("within ten days")
    })
  }, 120_000)

  test("a combined file saved under a numbered name is the one an accept rebuilds and the status names", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      await request(`${matter.route}/redlines/${propose(matter.dir, SAVED, "Memo marker", "Lawyer marker")}/accept`, { method: "POST" })
      const accepted = readFileSync(path.join(matter.dir, SAVED))
      // A re-delivery keeps the combined file the lawyer changed and saves the new one under the next free name.
      await recordLeanDocuments(matter.dir, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
      const numbered = "supply-pack (2).docx"
      const status = leanDraftingStatus(matter.dir)?.combined_output
      expect(status?.file).toBe(numbered)
      const body = status!.body
      const response = await request(`${matter.route}/redlines/${propose(matter.dir, body)}/accept`, { method: "POST" })
      expect(await response.json()).toMatchObject({ ok: true, combined_updated: true })
      expect(leanDraftingStatus(matter.dir)?.combined_output?.file).toBe(numbered)
      expect(documentText(readFileSync(path.join(matter.dir, numbered)))).toContain("within ten days")
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

  test("an accept whose combined file cannot be rebuilt says so, and the file saved before it is kept", async () => {
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
      // The accepted body stands; the file saved before it is left as it was.
      expect(documentText(readFileSync(path.join(matter.dir, BODY_NAME)))).toContain("within ten days")
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
    })
  }, 120_000)

  test("a new revision of a part uploaded in place rebuilds the saved combined file", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const form = new FormData()
      form.append("file", new File([(await renderMarkdownToDocx(BODY.replace("delivers the goods", "delivers the goods by rail"))).bytes], BODY_NAME))
      const uploaded = await request(`${matter.route}/documents/revisions`, { method: "POST", body: form })
      expect({ status: uploaded.status, body: await uploaded.json() }).toMatchObject({ status: 200, body: { combined_updated: true } })
      expect(documentText(readFileSync(path.join(matter.dir, SAVED)))).toContain("by rail")
      // The rebuilt file's version says what made it: a rebuild, not an accept (#268).
      const history = await (await request(`${matter.route}/documents/versions?${new URLSearchParams({ name: SAVED })}`)).json() as { versions: { origin: string; label: string | null }[] }
      expect(history.versions[0]).toMatchObject({ origin: "system", label: "Combined file rebuilt" })
    })
  }, 120_000)

  test("a redaction of a part rebuilds the saved combined file and deletes the combined file's earlier versions", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const removed = "the supplier delivers the goods"
      type Listed = { document: { id: string }; versions: { id: string; origin: string; removed: string | null; label: string | null; current: boolean }[] }
      const versions = async (name: string) => await (await request(`${matter.route}/documents/versions?${new URLSearchParams({ name })}`)).json() as Listed
      expect(documentText(readFileSync(path.join(matter.dir, SAVED)))).toContain(removed)
      // As the redaction tool saves: the part in place, then its revision.
      const redacted = (await renderMarkdownToDocx(BODY.replace(removed, "[REDACTED]"))).bytes
      writeFileSync(path.join(matter.dir, BODY_NAME), redacted)
      const form = new FormData()
      form.append("file", new File([redacted], BODY_NAME))
      const saved = await request(`${matter.route}/documents/revisions`, { method: "POST", body: form, headers: { "x-dochaus-origin": "redaction", "x-dochaus-actor": "agent:redactor" } })
      const body = await saved.json() as { combined_updated?: boolean; purge_failed?: string }
      expect({ status: saved.status, combined_updated: body.combined_updated, purge_failed: body.purge_failed }).toEqual({ status: 200, combined_updated: true, purge_failed: undefined })
      expect(documentText(readFileSync(path.join(matter.dir, SAVED)))).not.toContain(removed)
      // The rebuild is the redaction's write: every earlier combined version is removed, unreadable and not restorable.
      const combined = await versions(SAVED)
      expect(combined.versions.map(row => [row.origin, row.removed, row.current])).toEqual([
        ["redaction", null, true], ...combined.versions.slice(1).map(() => [expect.any(String), "redaction", false]),
      ])
      expect(combined.versions.length).toBeGreaterThan(1)
      for (const row of combined.versions.slice(1)) {
        expect((await request(`${matter.route}/documents/versions/${row.id}/content`)).status).toBe(410)
        expect((await request(`${matter.route}/documents/versions/${row.id}/restore`, { method: "POST" })).status).toBe(410)
      }
      // Neither the part's nor the combined file's kept versions hold the text.
      for (const id of [combined.document.id, (await versions(BODY_NAME)).document.id]) {
        const store = path.join(matter.dir, ".dochaus", "revisions", id)
        expect(readdirSync(store).filter(file => documentText(readFileSync(path.join(store, file))).includes(removed))).toEqual([])
      }
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
      expect(leanDraftingStatus(matter.dir)?.combined_output).toBeNull()
    })
  }, 120_000)

  const sha256 = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")

  test("a rebuild skipped for a proposal on the combined file is made when that proposal is rejected, one by one or all at once", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const saved = path.join(matter.dir, SAVED)
      const resolve = async (route: string) => {
        const response = await request(`${matter.route}/redlines/${route}`, { method: "POST" })
        return { status: response.status, body: await response.json() }
      }
      expect(await resolve(`${propose(matter.dir, BODY_NAME)}/accept`)).toMatchObject({ status: 200, body: { ok: true, combined_updated: true } })
      // A reject while nothing was skipped writes nothing and reports nothing about the combined file.
      const current = readFileSync(saved)
      expect(await resolve(`${propose(matter.dir, BODY_NAME, "Body marker", "Other marker")}/reject`)).toEqual({ status: 200, body: { ok: true } })
      expect(readFileSync(saved)).toEqual(current)

      // A change proposed on the combined file itself holds the next rebuild back.
      const onCombined = propose(matter.dir, SAVED, "Memo marker", "Memo note")
      expect(await resolve(`${propose(matter.dir, BODY_NAME, "the supplier", "the seller")}/accept`))
        .toMatchObject({ status: 200, body: { ok: true, combined_updated: false, combined_skip_reason: "pending_combined_proposals" } })
      expect(readFileSync(saved)).toEqual(current)
      // Rejecting it makes the skipped rebuild, from the parts as they are now.
      expect(await resolve(`${onCombined}/reject`)).toEqual({ status: 200, body: { ok: true, combined_updated: true } })
      const rebuilt = readFileSync(saved)
      expect(documentText(rebuilt)).toContain("the seller")
      expect(documentText(rebuilt)).not.toContain("Memo note")
      expect(readSystemDocx(matter.dir, SAVED)).toEqual({ version: 1, sha256: sha256(rebuilt) })

      // Likewise when every change on the combined file is rejected at once.
      propose(matter.dir, SAVED, "Memo marker", "Memo note")
      expect(await resolve(`${propose(matter.dir, BODY_NAME, "Body marker", "Main marker")}/accept`))
        .toMatchObject({ status: 200, body: { ok: true, combined_updated: false, combined_skip_reason: "pending_combined_proposals" } })
      expect(readFileSync(saved)).toEqual(rebuilt)
      expect(await resolve(`reject-all?name=${encodeURIComponent(SAVED)}`)).toEqual({ status: 200, body: { ok: true, rejected: 1, combined_updated: true } })
      const all = readFileSync(saved)
      expect(documentText(all)).toContain("Main marker")
      expect(readSystemDocx(matter.dir, SAVED)).toEqual({ version: 1, sha256: sha256(all) })

      // A change written to the file outside doc.haus is then kept, by an accept and by a reject.
      rewriteDocumentXml(matter.dir, SAVED, xml => xml.replace("Memo marker", "Outside marker"))
      const outside = readFileSync(saved)
      expect(await resolve(`${propose(matter.dir, BODY_NAME, "delivers the goods", "delivers the goods by rail")}/accept`))
        .toMatchObject({ status: 200, body: { ok: true, combined_updated: false, combined_skip_reason: "combined_changed" } })
      expect(await resolve(`${propose(matter.dir, SAVED, "Outside marker", "Inside marker")}/reject`))
        .toEqual({ status: 200, body: { ok: true, combined_updated: false, combined_skip_reason: "combined_changed" } })
      expect(readFileSync(saved)).toEqual(outside)
    })
  }, 120_000)

  test("a reject that owes no rebuild reports nothing about the combined file, even while a change proposed on that file is pending", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const saved = readFileSync(path.join(matter.dir, SAVED))
      const onCombined = propose(matter.dir, SAVED, "Memo marker", "Other marker")
      const resolve = async (route: string) => {
        const response = await request(`${matter.route}/redlines/${route}`, { method: "POST" })
        return { status: response.status, body: await response.json() }
      }
      expect(await resolve(`${propose(matter.dir, BODY_NAME, "Body marker", "Other marker")}/reject`)).toEqual({ status: 200, body: { ok: true } })
      propose(matter.dir, NOTES_NAME, "Memo marker", "Other marker")
      expect(await resolve(`reject-all?name=${encodeURIComponent(NOTES_NAME)}`)).toEqual({ status: 200, body: { ok: true, rejected: 1 } })
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
      const db = openDb(matter.dir)
      try { expect(getRedline(db, onCombined)?.status).toBe("pending") } finally { db.close() }
    })
  }, 120_000)

  test("a reject of a change on a part of the combined file waits for the combined file's lock before it reads the change", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const saved = readFileSync(path.join(matter.dir, SAVED))
      const id = propose(matter.dir, BODY_NAME, "Body marker", "Other marker")
      const status = () => {
        const db = openDb(matter.dir)
        try { return getRedline(db, id)?.status } finally { db.close() }
      }
      const hold = await holdLocks(matter.dir, [SAVED])
      const rejecting = request(`${matter.route}/redlines/${id}/reject`, { method: "POST" })
      for (const until = Date.now() + 1_500; Date.now() < until;) {
        expect(status()).toBe("pending")
        await Bun.sleep(50)
      }
      // A write holding that lock resolves the change meanwhile, as an accept into the body does.
      const db = openDb(matter.dir)
      try { setRedlineStatus(db, id, "accepted") } finally { db.close() }
      await hold.release()
      const rejected = await rejecting
      expect({ status: rejected.status, body: await rejected.json() }).toEqual({ status: 404, body: { error: "Redline not found or already resolved" } })
      expect(status()).toBe("accepted")
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
    })
  }, 120_000)

  test("a combined file rebuilt after an accept is still the system's: a rerun's delivery replaces it in place", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const saved = path.join(matter.dir, SAVED)
      const accepted = await request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
      expect(await accepted.json()).toMatchObject({ ok: true, combined_updated: true })
      const rebuilt = readFileSync(saved)
      expect(documentText(rebuilt)).toContain("within ten days")
      expect(readSystemDocx(matter.dir, SAVED)).toEqual({ version: 1, sha256: sha256(rebuilt) })

      // A rerun delivers new parts; its combined file goes to the same name.
      const rerun = (await renderMarkdownToDocx(BODY.replace("Body marker", "Rerun marker"))).bytes
      writeLeanBytes(matter.dir, "d1/body-applied.docx", rerun)
      writeLeanBytes(matter.dir, "d1/redline.docx", rerun)
      await recordLeanDocuments(matter.dir, { sessions: scriptedLeanModel().port, registerDeliverable: registerLeanDeliverable }, true)
      expect(leanDraftingStatus(matter.dir)?.combined_output?.file).toBe(SAVED)
      expect(readLeanDocuments(matter.dir)!.documents.filter(row => row.kind === "combined").map(row => row.deliverable)).toEqual([SAVED])
      const delivered = readFileSync(saved)
      expect(documentText(delivered)).toContain("Rerun marker")
      expect(delivered).not.toEqual(rebuilt)
      expect(readSystemDocx(matter.dir, SAVED)).toEqual({ version: 1, sha256: sha256(delivered) })
      expect(existsSync(path.join(matter.dir, "supply-pack (2).docx"))).toBe(false)
    })
  }, 120_000)

  test("a change the lawyer accepts into the combined file is kept by later accepts", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const saved = path.join(matter.dir, SAVED)
      const lawyer = await request(`${matter.route}/redlines/${propose(matter.dir, SAVED, "Memo marker", "Lawyer marker")}/accept`, { method: "POST" })
      expect({ status: lawyer.status, body: await lawyer.json() }).toEqual({ status: 200, body: { ok: true } })
      const accepted = readFileSync(saved)
      expect(documentText(accepted)).toContain("Lawyer marker")
      const part = await request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
      expect(await part.json()).toMatchObject({ ok: true, combined_updated: false, combined_skip_reason: "combined_changed" })
      expect(readFileSync(saved)).toEqual(accepted)
      // The accept removed the record of the system's write.
      expect(readSystemDocx(matter.dir, SAVED)).toBeUndefined()
    })
  }, 120_000)

  test("a combined file holding the bytes a build gives, without the record of the system writing them, is kept", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const db = openDb(matter.dir)
      try { db.run("DELETE FROM meta WHERE key = ?", [systemDocxKey(matter.dir, SAVED)]) } finally { db.close() }
      expect(readSystemDocx(matter.dir, SAVED)).toBeUndefined()
      const saved = readFileSync(path.join(matter.dir, SAVED))
      const accepted = await request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
      expect(await accepted.json()).toMatchObject({ ok: true, combined_updated: false, combined_skip_reason: "combined_changed" })
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
    })
  }, 120_000)

  test("a rebuild is saved with paragraph pointers as the run's save gives them, and recorded by the bytes saved", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      // A body revision saved without paragraph pointers; an accept into the notes then rebuilds the combined file from it.
      const body = unzipSync(readFileSync(path.join(matter.dir, BODY_NAME)))
      body["word/document.xml"] = strToU8(strFromU8(body["word/document.xml"]!).replace(/<w:bookmark(?:Start|End)\b[^>]*\/>/g, ""))
      await writeDeliverableRevision(matter.dir, BODY_NAME, Buffer.from(zipSync(body)))
      const accepted = await request(`${matter.route}/redlines/${propose(matter.dir, NOTES_NAME, "Memo marker", "Memo note")}/accept`, { method: "POST" })
      expect(await accepted.json()).toMatchObject({ ok: true, combined_updated: true })
      const saved = new Uint8Array(readFileSync(path.join(matter.dir, SAVED)))
      expect(documentText(saved)).toContain("Memo note")
      expect(withParagraphPointers(saved)).toEqual(saved)
      expect(readSystemDocx(matter.dir, SAVED)).toEqual({ version: 1, sha256: sha256(saved) })
    })
  }, 120_000)

  test("the rebuild checks the run again under its lock: a run no longer completed leaves the combined file as it is", async () => {
    await withIngest(async request => {
      const matter = await completedMatter(request)
      const REDLINE_NAME = "supply-pack - redline.docx"
      const saved = readFileSync(path.join(matter.dir, SAVED))
      // The accept starts while the run is not completed, so it takes no lock on the run's files up front.
      writeLeanState(matter.dir, { ...readLeanState(matter.dir)!, status: "running" })
      const body = await holdLocks(matter.dir, [BODY_NAME])
      const redline = await holdLocks(matter.dir, [REDLINE_NAME])
      const accepting = request(`${matter.route}/redlines/${propose(matter.dir, BODY_NAME)}/accept`, { method: "POST" })
      await Bun.sleep(500)
      // Completed when it reads the combined file to rebuild, then running again before the rebuild.
      writeLeanState(matter.dir, { ...readLeanState(matter.dir)!, status: "completed" })
      await body.release()
      await Bun.sleep(1_500)
      writeLeanState(matter.dir, { ...readLeanState(matter.dir)!, status: "running" })
      await redline.release()
      const accepted = await accepting
      expect({ status: accepted.status, body: await accepted.json() }).toEqual({ status: 200, body: {
        ok: true, redline_updated: false, redline_skip_reason: "rebuild_failed", combined_updated: false, combined_skip_reason: "combined_not_deliverable",
      } })
      expect(documentText(readFileSync(path.join(matter.dir, BODY_NAME)))).toContain("within ten days")
      expect(readFileSync(path.join(matter.dir, SAVED))).toEqual(saved)
    })
  }, 120_000)
})
