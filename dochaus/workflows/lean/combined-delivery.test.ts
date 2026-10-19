import { afterEach, describe, expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { createHash } from "node:crypto"
import { existsSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { buildCombinedFile } from "../../lib/combined-docx"
import { withParagraphPointers } from "../../lib/drafting-amendment"
import { pendingRedlinesForDoc, recordRedline } from "../../lib/redlines"
import { renderMarkdownToDocx } from "../../lib/render"
import { forgetSystemDocx, ownedDeliverable, recordSystemDocx, systemDocxKey } from "../../lib/system-docx"
import { workingDocxKey } from "../../lib/working-docx"
import { combinedDelivery, combinedParts, DELIVERED_NOTES, deliveryArtifactEvidence, forgetUnplannedDocuments, memoNotesFile, readLeanDocuments, recordLeanDocuments, type LeanDocumentKind } from "./deliverables"
import { checkDeliverySelections, selectedDeliveryPlan, type DeliverySelection } from "./delivery-selection"
import { packageReviewPrompt } from "./package-review"
import { writeLeanPlan, type PlannedDocument } from "./plan"
import { openItem } from "./quotes"
import { LEAN_REQUESTED_FILES_SCHEMA, requestedFilesItems, requestedFilesPlanned, writeRequestedFiles } from "./requested-files"
import { disclosableItems, type LeanDeps } from "./stages"
import { leanFile, writeLeanBytes, writeLeanText } from "./state"
import { cleanupMatters, matter, registering, scriptedModel } from "./test-support"

// A document delivered with the requested memo in one file ("Deliverable:
// combined"). The body is registered as "<stem> - clean" and the notes under
// the memo's name; once both are registered, the combined file built from them
// (buildCombinedFile) is registered under the requested name. A rerun replaces
// that file only while it holds the bytes the system last saved there (the
// registry's record of its own write, system-docx.ts) and no proposal on it is
// pending review; any other file there is kept and the new one takes the next
// free name. Real registration path; the
// registry writes are local and record the system's writes as ingest does
// (registering).

afterEach(cleanupMatters)

type Document = Omit<PlannedDocument, "id">
const SUPPLY: Document = { title: "Supply note", purpose: "Operative terms", parts: [{ heading: "Delivery", scope: "Delivery terms" }] }
const SCHEDULE: Document = { title: "Delivery schedule", purpose: "Delivery dates", parts: [{ heading: "Dates", scope: "Delivery dates" }] }
const COMBINED: Document = { ...SUPPLY, file: "supply-pack.docx", deliverable: "combined" }
const memo = (documents: Document[]) => ({ name: "cover memo", file: memoNotesFile(documents, "cover memo") })
const PACKAGE_FINDING = "Distinct named primary artifacts need distinct filenames. Report a conflicting naming instruction as a package finding."

/** Writes each planned document's body and redline and the notes, then
 * registers them as a passed gate does; returns the registered names. */
async function registered(directory: string, documents: Document[]) {
  const files = [...documents.flatMap((_, index) => [`d${index + 1}/body-applied.docx`, `d${index + 1}/redline.docx`]), "drafting-notes.docx", "drafting-notes.md"]
  for (const [index, file] of files.entries())
    writeLeanBytes(directory, file, file.endsWith(".md") ? new TextEncoder().encode(`# Notes ${index}\n`) : (await renderMarkdownToDocx(`# Record ${index}\n\nGenerated record ${index}.`)).bytes)
  const registration = registering()
  const summary = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: registration.registerDeliverable }, true)
  return { names: registration.names, rows: summary.documents.map(row => [row.kind, row.document, row.deliverable]) }
}

const matterFile = (directory: string, name: string) => new Uint8Array(readFileSync(path.join(directory, name)))
const sha256 = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
const docx = async (markdown: string) => (await renderMarkdownToDocx(markdown)).bytes

/** Writes the run's copy of each given file, rendered from Markdown. */
async function writeCopies(directory: string, files: Record<string, string>) {
  for (const [file, markdown] of Object.entries(files)) writeLeanBytes(directory, file, await docx(markdown))
}

/** Registers the run's current copies as a passed gate does, recording each
 * registration's name and the sha256 it replaces. */
async function deliver(directory: string) {
  const registration = registering()
  const summary = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: registration.registerDeliverable }, true)
  return { calls: registration.calls, rows: summary.documents.map(row => [row.kind, row.document, row.deliverable]) }
}

/** Runs `write` in a registry write transaction, as ingest's writes do. */
function registry(directory: string, write: (db: Database) => void) {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  try {
    db.run("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
    db.transaction(() => write(db)).immediate()
  } finally {
    db.close()
  }
}

/** The accepted Word revision's record of the matter file `file`, bound to the
 * bytes the accept wrote there, as ingest records it. */
function bindWorkingRevision(db: Database, directory: string, file: string) {
  db.query("INSERT INTO meta (key, value) VALUES (?, ?)").run(workingDocxKey(directory, file), JSON.stringify({ version: 1, source_sha256: sha256(new Uint8Array(readFileSync(file))), projection_sha256: "2".repeat(64) }))
}

/** The gate's delivered notes, as a gate writes them in place of the drafting notes. */
async function writeDelivered(directory: string) {
  const text = "# Cover memo\n\nThe delivery terms follow the supply terms.\n\n## Requirements left open\n\nNone."
  writeLeanBytes(directory, DELIVERED_NOTES.docx, await docx(text))
  writeLeanText(directory, DELIVERED_NOTES.markdown, `${text}\n`)
}

/** The gate's record of `items`, each disclosed by the notes unless recorded with another `status`. */
function discloseAtGate(directory: string, items: ReturnType<typeof openItem>[], status: ReturnType<typeof openItem>["status"] = "disclosed") {
  writeLeanText(directory, "gate.json", `${JSON.stringify({ items: items.map(item => ({ ...item, status })) }, null, 2)}\n`)
}

const COPIES = {
  "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each month.",
  "d1/redline.docx": "# Supply note\n\nThe supplier delivers the goods each month.",
  "drafting-notes.docx": "# Cover memo\n\nThe delivery terms follow the supply terms.",
}
const unsaved = (reason: string) => openItem("finalize", "omission", `The combined file "supply-pack.docx" (the memo, then "Supply note") is not saved in the matter: ${reason}. ` +
  'The memo and "Supply note" are delivered as separate files.', { repair_target: "memo" })

describe("registration of a combined document", () => {
  test("registers the body as '<stem> - clean', the notes under the memo's name, then the file built from them under the requested name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    const { names, rows } = await registered(directory, [COMBINED])
    expect(rows).toEqual([
      ["body", "d1", "supply-pack - clean.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["drafting_notes", null, "supply-pack - cover memo.docx"],
      ["drafting_notes_markdown", null, null],
      ["combined", "d1", "supply-pack.docx"],
    ])
    expect(names).toEqual(["supply-pack - clean.docx", "supply-pack - redline.docx", "supply-pack - cover memo.docx", "supply-pack.docx"])
    // Exactly the bytes the build gives from the two registered parts.
    const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean.docx") })
    expect("bytes" in built && sha256(built.bytes)).toBe(sha256(matterFile(directory, "supply-pack.docx")))
    expect(readLeanDocuments(directory)?.documents.at(-1)).toEqual({ kind: "combined", document: "d1", title: "Supply note", path: "supply-pack.docx",
      sha256: sha256(matterFile(directory, "supply-pack.docx")), deliverable: "supply-pack.docx" })
  }, 60_000)

  test("without combined, registered names are unchanged", async () => {
    const plain = matter()
    writeLeanPlan(plain, [SUPPLY])
    expect((await registered(plain, [SUPPLY])).rows).toEqual([
      ["body", "d1", "Supply note.docx"],
      ["redline", "d1", "Supply note - redline.docx"],
      ["drafting_notes", null, "Supply note - drafting notes.docx"],
      ["drafting_notes_markdown", null, null],
    ])

    const named = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx", deliverable: "body" }, { ...SCHEDULE, file: "schedule-review.docx", deliverable: "redline" }]
    writeLeanPlan(named, documents, undefined, undefined, memo(documents))
    expect((await registered(named, documents)).rows).toEqual([
      ["body", "d1", "supply-pack.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["body", "d2", "schedule-review - clean.docx"],
      ["redline", "d2", "schedule-review.docx"],
      ["drafting_notes", null, "supply-pack - cover memo.docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)
})

describe("the combined file saved in the matter", () => {
  test("a memo the build refuses: one finalize item, no combined file, and registration fails visibly unless the gate disclosed it", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    // A numbered list whose definition links to a numbering style.
    const numbered = unzipSync(await docx("# Cover memo\n\n1. First point.\n2. Second point."))
    numbered["word/numbering.xml"] = strToU8(strFromU8(numbered["word/numbering.xml"]!).replace(/(<w:abstractNum\b[^>]*>)/g, '$1<w:numStyleLink w:val="MemoList"/>'))
    writeLeanBytes(directory, "drafting-notes.docx", zipSync(numbered))
    expect(buildCombinedFile({ memo: zipSync(numbered), body: await docx(COPIES["d1/body-applied.docx"]) })).toEqual({ refused: "unsupported_memo_numbering" })

    const item = unsaved("the memo uses list numbering the combined file cannot carry")
    expect(disclosableItems(directory)).toEqual([item])
    await expect(deliver(directory)).rejects.toThrow(item.detail)
    expect(existsSync(path.join(directory, "supply-pack.docx"))).toBe(false)
    discloseAtGate(directory, [item])
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).toEqual(["supply-pack - clean.docx", "supply-pack - redline.docx", "supply-pack - cover memo.docx"])
    expect(rows.map(([kind]) => kind)).not.toContain("combined")
    expect(existsSync(path.join(directory, "supply-pack.docx"))).toBe(false)
  }, 60_000)

  test("a source under the requested name is untouched: the combined file takes the next free name, with no item", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    const source = await docx("# Supply terms\n\nThe goods are delivered monthly.")
    writeFileSync(path.join(directory, "supply-pack.docx"), source)
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    db.run("CREATE TABLE IF NOT EXISTS documents (doc_path TEXT, role TEXT)")
    db.query("INSERT INTO documents (doc_path, role) VALUES (?, 'source')").run(path.join(directory, "supply-pack.docx"))
    db.close()

    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack (2).docx" })
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows.filter(([kind]) => kind === "combined")).toEqual([["combined", "d1", "supply-pack (2).docx"]])
    expect(sha256(matterFile(directory, "supply-pack.docx"))).toBe(sha256(source))
  }, 60_000)

  test("a rerun rebuilds the file from the current parts, replacing the file the system saved there; unchanged parts give the same bytes", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    expect(disclosableItems(directory)).toEqual([])
    await deliver(directory)
    const first = sha256(matterFile(directory, "supply-pack.docx"))

    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })
    const second = await deliver(directory)
    expect(second.calls).toContainEqual(["supply-pack.docx", first])
    const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean.docx") })
    expect("bytes" in built && sha256(built.bytes)).toBe(sha256(matterFile(directory, "supply-pack.docx")))
    const rebuilt = sha256(matterFile(directory, "supply-pack.docx"))
    expect(rebuilt).not.toBe(first)

    const third = await deliver(directory)
    expect(third.calls).toContainEqual(["supply-pack.docx", rebuilt])
    expect(sha256(matterFile(directory, "supply-pack.docx"))).toBe(rebuilt)
  }, 60_000)

  /** The lawyer's accept of a change into the registered body, as ingest saves
   * it: the body rewritten in the matter, which removes the record of the
   * system's write of it (and adds the accepted Word revision's record when
   * `bound`); then, while the saved combined file still holds the system's
   * last write, that file rebuilt from the parts in the matter, with paragraph
   * pointers, and recorded as the system's own write. */
  async function acceptIntoBody(directory: string, bound: boolean) {
    const body = path.join(directory, "supply-pack - clean.docx")
    writeFileSync(body, withParagraphPointers(await docx("# Supply note\n\nThe supplier delivers the goods within ten days.")))
    registry(directory, db => {
      forgetSystemDocx(db, directory, body)
      if (bound) bindWorkingRevision(db, directory, body)
    })
    const combined = path.join(directory, "supply-pack.docx")
    if (ownedDeliverable(directory, "supply-pack.docx")) {
      const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean.docx") })
      if (!("bytes" in built)) throw new Error(built.refused)
      const bytes = withParagraphPointers(built.bytes, readFileSync(combined))
      writeFileSync(combined, bytes)
      registry(directory, db => recordSystemDocx(db, directory, combined, sha256(bytes)))
    }
    return { body: matterFile(directory, "supply-pack - clean.docx"), combined: matterFile(directory, "supply-pack.docx") }
  }

  /** The lawyer's accept of a change into the saved combined file itself, as
   * ingest saves it: the file rewritten, which removes the record of the
   * system's write of it (and adds the accepted Word revision's record when
   * `bound`). */
  async function acceptIntoCombined(directory: string, bound = false) {
    const combined = path.join(directory, "supply-pack.docx")
    writeFileSync(combined, withParagraphPointers(await docx("# Cover memo\n\nThe delivery terms follow the supply terms.\n\n# Supply note\n\nThe supplier delivers the goods on request.")))
    registry(directory, db => {
      forgetSystemDocx(db, directory, combined)
      if (bound) bindWorkingRevision(db, directory, combined)
    })
    return matterFile(directory, "supply-pack.docx")
  }

  const combinedRows = (directory: string) => readLeanDocuments(directory)!.documents.filter(row => row.kind === "combined")

  test("an accept into the body rebuilds the combined file as the system's own write: a rerun replaces it in place, and keeps the accepted body as before", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    expect(combinedParts(directory)?.file).toBe("supply-pack.docx")
    const accepted = await acceptIntoBody(directory, false)
    // The rerun's memo replaces the registered one before the combined file is saved.
    await writeCopies(directory, REVISED_MEMO)

    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack.docx" })
    const { calls, rows } = await deliver(directory)
    expect(rows).toContainEqual(["body", "d1", "supply-pack - clean (2).docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack.docx"])
    expect(calls).toContainEqual(["supply-pack.docx", sha256(accepted.combined)])
    expect(matterFile(directory, "supply-pack - clean.docx")).toEqual(accepted.body)
    const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean (2).docx") })
    expect("bytes" in built && sha256(built.bytes)).toBe(sha256(matterFile(directory, "supply-pack.docx")))
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack.docx", deliverable: "supply-pack.docx" })])
    expect(combinedParts(directory)?.file).toBe("supply-pack.docx")
  }, 60_000)

  test("an accept into the combined file itself: a rerun keeps that file, listed as an earlier combined file, and saves the new one under the next free name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const [first] = combinedRows(directory)
    const accepted = await acceptIntoCombined(directory)
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack (2).docx" })
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows).toContainEqual(["body", "d1", "supply-pack - clean.docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted)
    const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean.docx") })
    expect("bytes" in built && sha256(built.bytes)).toBe(sha256(matterFile(directory, "supply-pack (2).docx")))
    // The kept file stays listed as a combined file, with its deliverable; the new one is the run's.
    expect(combinedRows(directory)).toEqual([first, expect.objectContaining({ path: "supply-pack (2).docx", deliverable: "supply-pack (2).docx" })])
    expect(combinedParts(directory)?.file).toBe("supply-pack (2).docx")

    // A further rerun replaces its own file there and still lists the kept one.
    const saved = sha256(matterFile(directory, "supply-pack (2).docx"))
    expect(disclosableItems(directory)).toEqual([])
    const again = await deliver(directory)
    expect(again.calls).toContainEqual(["supply-pack (2).docx", saved])
    expect(combinedRows(directory)).toEqual([first, expect.objectContaining({ path: "supply-pack (2).docx", deliverable: "supply-pack (2).docx" })])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted)
  }, 60_000)

  test("the same bytes without the record of the system's write are not the system's to replace: the combined file takes the next free name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const delivered = matterFile(directory, "supply-pack.docx")
    // Saved again by another write, with the same bytes.
    registry(directory, db => forgetSystemDocx(db, directory, path.join(directory, "supply-pack.docx")))

    expect(disclosableItems(directory)).toEqual([])
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(delivered)
    expect(combinedRows(directory).map(row => row.deliverable)).toEqual(["supply-pack.docx", "supply-pack (2).docx"])
  }, 60_000)

  test("an accept into the combined file, then a delivery that fails to save the new one: the retry keeps the accepted file and saves the new one under the next free name", async () => {
    for (const memoChanged of [false, true]) {
      const directory = matter()
      writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
      await writeCopies(directory, COPIES)
      await deliver(directory)
      const accepted = await acceptIntoCombined(directory)
      if (memoChanged) await writeCopies(directory, REVISED_MEMO)

      // The parts are registered, then the matter fails to write the combined file; nothing discloses it.
      const failing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes) => {
        if (/^supply-pack(?: \(\d+\))?\.docx$/.test(name)) throw new Error("The matter could not write the file.")
        writeFileSync(path.join(directory, name), bytes)
        return name
      }
      const fail = () => expect(recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: failing }, true)).rejects
        .toThrow('The combined file "supply-pack (2).docx" is not saved in the matter: The matter could not write the file.')
      const kept = expect.objectContaining({ kind: "combined", path: "supply-pack.docx", deliverable: "supply-pack.docx" })
      const unsavedRow = expect.objectContaining({ kind: "combined", path: "supply-pack (2).docx", deliverable: null })
      await fail()
      expect(readLeanDocuments(directory)?.documents).toContainEqual(expect.objectContaining({ kind: "body", deliverable: "supply-pack - clean.docx" }))
      expect(combinedRows(directory)).toEqual([kept, unsavedRow])
      // The combined file is the one not yet saved, never the kept one.
      expect(combinedParts(directory)?.file).toBe("supply-pack (2).docx")
      // Neither a second failure nor a stop before delivery loses that name.
      await fail()
      await recordLeanDocuments(directory, { sessions: scriptedModel({}).port }, false)
      expect(combinedRows(directory)).toEqual([kept, unsavedRow])

      expect(disclosableItems(directory)).toEqual([])
      expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack (2).docx" })
      const { calls, rows } = await deliver(directory)
      expect(calls).toContainEqual(["supply-pack (2).docx", null])
      expect(rows).toContainEqual(["body", "d1", "supply-pack - clean.docx"])
      expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
      expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted)
      const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean.docx") })
      expect("bytes" in built && sha256(built.bytes)).toBe(sha256(matterFile(directory, "supply-pack (2).docx")))
    }
  }, 60_000)

  test("a matter that fails to write the combined file under its own name: the summary is still written, and the run stops with that reason in plain words", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    const registration = registering()
    const failing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current) => {
      if (name === "supply-pack.docx") throw new Error("disk full")
      return registration.registerDeliverable(directory, name, bytes, current)
    }
    await expect(recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: failing }, true)).rejects
      .toThrow('The combined file "supply-pack.docx" is not saved in the matter: disk full')
    const documents = readLeanDocuments(directory)!.documents
    expect(documents.find(row => row.kind === "body")?.sha256).toBe(sha256(matterFile(directory, "supply-pack - clean.docx")))
    expect(documents.at(-1)).toMatchObject({ kind: "combined", path: "supply-pack.docx", deliverable: null })
  }, 60_000)

  test("an accept with a working record into the body: a rerun keeps the body and replaces the combined file the system rebuilt in place, with no item", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const accepted = await acceptIntoBody(directory, true)

    expect(disclosableItems(directory)).toEqual([])
    const { calls, rows } = await deliver(directory)
    expect(rows).toContainEqual(["body", "d1", "supply-pack - clean.docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack.docx"])
    expect(calls).toContainEqual(["supply-pack.docx", sha256(accepted.combined)])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted.combined)
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
  }, 60_000)

  test("with the gate's delivered notes, an accept with a working record into the body: the gate raises no item, and the rerun replaces the combined file in place", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await writeDelivered(directory)
    await deliver(directory)
    expect(readLeanDocuments(directory)?.documents).toContainEqual(expect.objectContaining({ kind: "drafting_notes", path: path.relative(directory, leanFile(directory, DELIVERED_NOTES.docx)), deliverable: "supply-pack - cover memo.docx" }))
    const accepted = await acceptIntoBody(directory, true)

    expect(disclosableItems(directory)).toEqual([])
    const { calls, rows } = await deliver(directory)
    expect(rows).toContainEqual(["combined", "d1", "supply-pack.docx"])
    expect(calls).toContainEqual(["supply-pack.docx", sha256(accepted.combined)])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted.combined)
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
  }, 60_000)

  test("a stop at the gate after a delivery keeps the combined file the system saved: the next delivery replaces it in place, with no item", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await writeDelivered(directory)
    await deliver(directory)
    const accepted = await acceptIntoBody(directory, false)
    // The rerun stops at its gate, which leaves the delivered notes as they were.
    await recordLeanDocuments(directory, { sessions: scriptedModel({}).port }, false)
    expect(readLeanDocuments(directory)?.documents).toContainEqual(expect.objectContaining({ kind: "drafting_notes", deliverable: "supply-pack - cover memo.docx" }))
    expect(readLeanDocuments(directory)?.documents.at(-1)).toMatchObject({ kind: "combined", path: "supply-pack.docx", deliverable: "supply-pack.docx" })
    expect(combinedParts(directory)?.file).toBe("supply-pack.docx")

    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })
    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack.docx" })
    const { calls, rows } = await deliver(directory)
    expect(calls).toContainEqual(["supply-pack.docx", sha256(accepted.combined)])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack.docx"])
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
  }, 60_000)

  test("a new run after the summary is archived keeps the combined file an earlier run saved: the new one takes the next free name, as the body does", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const earlier = matterFile(directory, "supply-pack.docx")
    rmSync(leanFile(directory, "documents.json"))
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack (2).docx" })
    const { calls, rows } = await deliver(directory)
    expect(rows).toContainEqual(["body", "d1", "supply-pack - clean (2).docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(matterFile(directory, "supply-pack.docx")).toEqual(earlier)
  }, 60_000)

  test("a file an earlier plan delivered another output under is not the combined file's to replace: the combined file takes the next free name", async () => {
    const directory = matter()
    const body = { ...COMBINED, deliverable: "body" as const }
    writeLeanPlan(directory, [body], undefined, undefined, memo([body]))
    await writeCopies(directory, COPIES)
    expect((await deliver(directory)).rows).toContainEqual(["body", "d1", "supply-pack.docx"])
    const earlier = matterFile(directory, "supply-pack.docx")
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))

    expect(disclosableItems(directory)).toEqual([])
    const { calls, rows } = await deliver(directory)
    expect(rows).toContainEqual(["body", "d1", "supply-pack - clean.docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(matterFile(directory, "supply-pack.docx")).toEqual(earlier)
  }, 60_000)

  test("a file saved under the combined name by another write is kept but no longer listed as a combined file: the combined file takes the next free name, with no item", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const other = await docx("# Supply pack\n\nA synthetic file saved under this name by hand.")
    writeFileSync(path.join(directory, "supply-pack.docx"), other)
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    expect(disclosableItems(directory)).toEqual([])
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(other)
    // The file is no longer the system's last write: no row lists it as created by doc.haus.
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack (2).docx", deliverable: "supply-pack (2).docx" })])
  }, 60_000)

  test("an accept with a working record into the combined file itself: the file is held with its item, never moved to the next free name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const accepted = await acceptIntoCombined(directory, true)
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    const item = unsaved("the file of that name holds the lawyer's accepted Word changes, which the run does not write over")
    expect(disclosableItems(directory)).toEqual([item])
    await expect(deliver(directory)).rejects.toThrow(item.detail)
    discloseAtGate(directory, [item])
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name).filter(name => /^supply-pack(?: \(\d+\))?\.docx$/.test(name))).toEqual([])
    expect(rows.filter(([kind, , deliverable]) => kind === "combined" && deliverable)).toEqual([])
    // The run keeps its name there without claiming a delivered file.
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack.docx", deliverable: null })])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted)
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
  }, 60_000)

  const REVISED_MEMO = { "drafting-notes.docx": "# Cover memo\n\nThe delivery terms follow the revised supply terms." }

  test("an accept with a working record into the body and a changed memo: the combined file the system rebuilt is replaced in place, with no item", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const accepted = await acceptIntoBody(directory, true)
    await writeCopies(directory, REVISED_MEMO)

    expect(disclosableItems(directory)).toEqual([])
    const { calls, rows } = await deliver(directory)
    expect(rows).toContainEqual(["body", "d1", "supply-pack - clean.docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack.docx"])
    expect(calls).toContainEqual(["supply-pack.docx", sha256(accepted.combined)])
    expect(matterFile(directory, "supply-pack - clean.docx")).toEqual(accepted.body)
    const built = buildCombinedFile({ memo: matterFile(directory, "supply-pack - cover memo.docx"), body: matterFile(directory, "supply-pack - clean.docx") })
    expect("bytes" in built && sha256(built.bytes)).toBe(sha256(matterFile(directory, "supply-pack.docx")))
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
  }, 60_000)

  /** A redline proposal on the matter file `name`, pending the lawyer's review. */
  function propose(directory: string, name: string) {
    recordRedline(directory, { docPath: name, docName: name, scope: "phrase", findText: "goods", oldText: "goods", newText: "products", author: "reviewer", reviewComment: "", anchorId: "p000000000001" })
  }
  const pendingOn = (directory: string, name: string) => pendingRedlinesForDoc(directory, name).length

  test("a proposal pending on the combined file the system saved: a rerun keeps that file and the proposal, and saves the new one under the next free name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const [first] = combinedRows(directory)
    const delivered = matterFile(directory, "supply-pack.docx")
    propose(directory, "supply-pack.docx")
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack (2).docx" })
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(delivered)
    expect(pendingOn(directory, "supply-pack.docx")).toBe(1)
    // The file is kept and listed as an earlier combined file, as a lawyer's change to it is.
    expect(combinedRows(directory)).toEqual([first, expect.objectContaining({ path: "supply-pack (2).docx", deliverable: "supply-pack (2).docx" })])
  }, 60_000)

  test("a save that fails leaves no delivered combined file: a file the lawyer then puts under that name is kept and never listed as a combined file", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    const registration = registering()
    const failing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current) => {
      if (name === "supply-pack.docx") throw new Error("disk full")
      return registration.registerDeliverable(directory, name, bytes, current)
    }
    await expect(recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: failing }, true)).rejects.toThrow("disk full")
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack.docx", deliverable: null })])
    const own = await docx("# Supply pack\n\nA synthetic file the lawyer saved under this name.")
    writeFileSync(path.join(directory, "supply-pack.docx"), own)

    const { rows } = await deliver(directory)
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(own)
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack (2).docx", deliverable: "supply-pack (2).docx" })])
  }, 60_000)

  test("a stop at the gate after an accept into the combined file itself: the summary no longer lists that file as the run's delivered combined file", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const accepted = await acceptIntoCombined(directory)
    await recordLeanDocuments(directory, { sessions: scriptedModel({}).port }, false)
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack.docx", deliverable: null })])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted)
  }, 60_000)

  test("a combined file another write saves between the run's check and its save is kept: the save is refused under the write's lock and the file takes the next free name", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const delivered = matterFile(directory, "supply-pack.docx")
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    const registration = registering()
    // Another write saves the same bytes after the run found the file its own.
    const racing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current, options) => {
      if (name === "supply-pack.docx" && current) registry(directory, db => forgetSystemDocx(db, directory, path.join(directory, name)))
      return registration.registerDeliverable(directory, name, bytes, current, options)
    }
    const summary = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: racing }, true)
    expect(registration.calls).toContainEqual(["supply-pack.docx", sha256(delivered)])
    expect(summary.documents.filter(row => row.kind === "combined").map(row => row.deliverable)).toEqual(["supply-pack.docx", "supply-pack (2).docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(delivered)
  }, 60_000)

  test("the next free name skips a name another planned output is requested under, while that output is not yet written", async () => {
    const directory = matter()
    const documents: Document[] = [COMBINED, { ...SCHEDULE, file: "supply-pack (2).docx", deliverable: "body" }]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    await acceptIntoCombined(directory)
    await writeCopies(directory, REVISED_MEMO)

    expect(disclosableItems(directory)).toEqual([])
    expect(deliveryArtifactEvidence(directory).at(-1)).toMatchObject({ kind: "combined", download_filename: "supply-pack (3).docx" })
    expect((await deliver(directory)).rows).toContainEqual(["combined", "d1", "supply-pack (3).docx"])
    expect(existsSync(path.join(directory, "supply-pack (2).docx"))).toBe(false)
    // Once written, that output takes its requested name.
    await writeCopies(directory, { "d2/body-applied.docx": "# Delivery schedule\n\nGoods arrive on the first day.", "d2/redline.docx": "# Delivery schedule\n\nGoods arrive on the first day." })
    const { rows } = await deliver(directory)
    expect(rows).toContainEqual(["body", "d2", "supply-pack (2).docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (3).docx"])
  }, 60_000)

  test("an accept into the combined file while every numbered name is taken: a finalize item before delivery, and no file is written over", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const accepted = await acceptIntoCombined(directory)
    const others = new Map<string, Uint8Array<ArrayBuffer>>()
    for (let index = 2; index <= 9; index++) {
      const bytes = await docx(`# Supply pack ${index}\n\nA synthetic file saved under this name by hand.`)
      others.set(`supply-pack (${index}).docx`, bytes)
      writeFileSync(path.join(directory, `supply-pack (${index}).docx`), bytes)
    }
    await writeCopies(directory, REVISED_MEMO)

    const item = unsaved("another file already has that name, which the run keeps (it replaces only a combined file it saved there that still holds what it saved, with no change pending review), " +
      'and other files already have each numbered name the run can give the new file, up to "supply-pack (9).docx"')
    expect(disclosableItems(directory)).toEqual([item])
    // Registration finds the same item, which the gate's notes disclose.
    discloseAtGate(directory, [item])
    const { rows } = await deliver(directory)
    expect(rows.filter(([kind, , deliverable]) => kind === "combined" && deliverable)).toEqual([])
    // The run keeps its name there without claiming a delivered file.
    expect(combinedRows(directory)).toEqual([expect.objectContaining({ path: "supply-pack.docx", deliverable: null })])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(accepted)
    for (const [name, bytes] of others) expect(matterFile(directory, name)).toEqual(bytes)
  }, 60_000)

  test("a request naming the combined file is checked against the name registration gives it: once it moves to the next free name, the mismatch is disclosed", async () => {
    const directory = matter()
    const words = "Send the supply note with the cover memo in one file named supply-pack.docx"
    const plan = writeLeanPlan(directory, [COMBINED], "basis-1", undefined, memo([COMBINED]))
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
      files: [{ holds: ["memo", "d1"], name: "supply-pack.docx", quote: words }] })
    const evidence = () => deliveryArtifactEvidence(directory).at(-1)
    await writeCopies(directory, COPIES)
    expect(requestedFilesItems(directory)).toEqual([])
    await deliver(directory)
    expect(requestedFilesItems(directory)).toEqual([])
    expect(evidence()).toMatchObject({ kind: "combined", download_filename: "supply-pack.docx" })
    await acceptIntoCombined(directory)
    await writeCopies(directory, REVISED_MEMO)

    const item = openItem("finalize", "omission", `Requested files: the lawyer's words "${words}" [source: request] ask for the memo and "Supply note" in one file named "supply-pack.docx", ` +
      'but the run delivers the memo and "Supply note" in one file named "supply-pack (2).docx", since a file already in the matter holds the name "supply-pack.docx". ' +
      "The lawyer should confirm how these files are to be delivered.", { repair_target: "memo", requested_file: 0, delivered_names: ["supply-pack (2).docx"] })
    // Predicted before the delivery moves the file, and read from the summary after it.
    expect(requestedFilesItems(directory)).toEqual([item])
    expect(evidence()).toMatchObject({ download_filename: "supply-pack (2).docx" })
    expect((await deliver(directory)).rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(requestedFilesItems(directory)).toEqual([item])
    expect(evidence()).toMatchObject({ download_filename: "supply-pack (2).docx" })
  }, 60_000)

  test("a request for one file without a name: once a kept file holds the combined name, the file the run delivers instead is disclosed with the reason", async () => {
    const directory = matter()
    const words = "Send the supply note with the cover memo in one file"
    const plan = writeLeanPlan(directory, [COMBINED], "basis-1", undefined, memo([COMBINED]))
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
      files: [{ holds: ["memo", "d1"], name: null, quote: words }] })
    await writeCopies(directory, COPIES)
    await deliver(directory)
    expect(requestedFilesItems(directory)).toEqual([])
    propose(directory, "supply-pack.docx")
    await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

    const item = openItem("finalize", "omission", `Requested files: the lawyer's words "${words}" [source: request] ask for the memo and "Supply note" in one file, ` +
      'but the run delivers the memo and "Supply note" in one file named "supply-pack (2).docx", since a file already in the matter holds the name "supply-pack.docx". ' +
      "The lawyer should confirm how these files are to be delivered.", { repair_target: "memo", requested_file: 0, delivered_names: ["supply-pack (2).docx"] })
    expect(requestedFilesItems(directory)).toEqual([item])
    discloseAtGate(directory, [item])
    expect((await deliver(directory)).rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
    expect(requestedFilesItems(directory)).toEqual([item])
  }, 60_000)

  test("a request for one file while a kept file holds the combined name and the combined file will not be saved: the items agree that the files are delivered separately", async () => {
    for (const named of [true, false]) {
      const directory = matter()
      const words = named ? "Send the supply note with the cover memo in one file named supply-pack.docx" : "Send the supply note with the cover memo in one file"
      const plan = writeLeanPlan(directory, [COMBINED], "basis-1", undefined, memo([COMBINED]))
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
        files: [{ holds: ["memo", "d1"], name: named ? "supply-pack.docx" : null, quote: words }] })
      await writeCopies(directory, COPIES)
      writeFileSync(path.join(directory, "supply-pack.docx"), await docx("# Supply pack\n\nA synthetic file saved under this name by hand."))
      discloseAtGate(directory, disclosableItems(directory))
      expect((await deliver(directory)).rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])
      // A later run's memo cannot be read as Word, so the combined file is not saved.
      writeLeanBytes(directory, "drafting-notes.docx", new Uint8Array([1, 2, 3]))
      await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })

      const unsavedItem = openItem("finalize", "omission", 'The combined file "supply-pack (2).docx" (the memo, then "Supply note") is not saved in the matter: ' +
        'the memo could not be read as a Word document. The memo and "Supply note" are delivered as separate files.', { repair_target: "memo" })
      const separate = openItem("finalize", "omission", `Requested files: the lawyer's words "${words}" [source: request] ask for the memo and "Supply note" in one file named "supply-pack.docx", ` +
        'but the run delivers the memo in "supply-pack - cover memo.docx" and "Supply note" in "supply-pack - clean.docx". ' +
        "The lawyer should confirm how these files are to be delivered.", { repair_target: "memo", requested_file: 0, delivered_names: ["supply-pack - cover memo.docx", "supply-pack - clean.docx"] })
      expect({ named, items: disclosableItems(directory).filter(item => /combined file|Requested files/.test(item.detail)) })
        .toEqual({ named, items: named ? expect.arrayContaining([unsavedItem, separate]) : [unsavedItem] })
      expect(disclosableItems(directory).some(item => item.detail.includes("in one file named \"supply-pack (2).docx\""))).toBe(false)
    }
  }, 60_000)

  test("a save refused under the write's lock after the gate: the request is checked against the name the file is saved under, and an undisclosed mismatch with a requested name fails the run", async () => {
    const words = "Send the supply note with the cover memo in one file named supply-pack.docx"
    const unnamed = "Send the supply note with the cover memo in one file"
    const item = openItem("finalize", "omission", `Requested files: the lawyer's words "${words}" [source: request] ask for the memo and "Supply note" in one file named "supply-pack.docx", ` +
      'but the run delivers the memo and "Supply note" in one file named "supply-pack (2).docx", since a file already in the matter holds the name "supply-pack.docx". ' +
      "The lawyer should confirm how these files are to be delivered.", { repair_target: "memo", requested_file: 0, delivered_names: ["supply-pack (2).docx"] })
    /** A delivered matter whose rerun's gate finds nothing to disclose, then a
     * rerun whose save of the combined file is refused under the write's lock. */
    const raced = async (disclosed: ReturnType<typeof openItem>[], named = true, status: ReturnType<typeof openItem>["status"] = "disclosed") => {
      const directory = matter()
      const plan = writeLeanPlan(directory, [COMBINED], "basis-1", undefined, memo([COMBINED]))
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
        files: [{ holds: ["memo", "d1"], name: named ? "supply-pack.docx" : null, quote: named ? words : unnamed }] })
      await writeCopies(directory, COPIES)
      await deliver(directory)
      await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })
      // The gate finds the file the system's to replace in place.
      expect(requestedFilesItems(directory)).toEqual([])
      discloseAtGate(directory, disclosed, status)
      const registration = registering()
      const racing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current, options) => {
        if (name === "supply-pack.docx" && current) registry(directory, db => forgetSystemDocx(db, directory, path.join(directory, name)))
        return registration.registerDeliverable(directory, name, bytes, current, options)
      }
      return { directory, delivered: recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: racing }, true) }
    }

    const { directory, delivered } = await raced([])
    await expect(delivered).rejects.toThrow(item.detail)
    // The parts and the moved file are delivered and listed.
    expect(combinedRows(directory).map(row => row.deliverable)).toEqual(["supply-pack.docx", "supply-pack (2).docx"])
    expect(requestedFilesItems(directory)).toEqual([item])
    // Once the gate's notes disclose it, the next delivery replaces the moved file in place.
    discloseAtGate(directory, [item])
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).toContain("supply-pack (2).docx")
    expect(rows).toContainEqual(["combined", "d1", "supply-pack (2).docx"])

    // An item the gate recorded but its notes did not disclose does not count.
    for (const status of ["open", "resolved"] as const) await expect((await raced([item], true, status)).delivered).rejects.toThrow(item.detail)

    // Notes that already disclose that mismatch let the moved delivery complete.
    const disclosed = await raced([item])
    expect((await disclosed.delivered).documents.filter(row => row.kind === "combined").map(row => row.deliverable)).toEqual(["supply-pack.docx", "supply-pack (2).docx"])

    // One file asked for without a name is the one file delivered: the move
    // completes the run, and the next finalize discloses where the file went.
    const one = await raced([], false)
    expect((await one.delivered).documents.filter(row => row.kind === "combined").map(row => row.deliverable)).toEqual(["supply-pack.docx", "supply-pack (2).docx"])
    expect(requestedFilesItems(one.directory)).toEqual([openItem("finalize", "omission", `Requested files: the lawyer's words "${unnamed}" [source: request] ask for the memo and "Supply note" in one file, ` +
      'but the run delivers the memo and "Supply note" in one file named "supply-pack (2).docx", since a file already in the matter holds the name "supply-pack.docx". ' +
      "The lawyer should confirm how these files are to be delivered.", { repair_target: "memo", requested_file: 0, delivered_names: ["supply-pack (2).docx"] })])
  }, 60_000)

  test("a save refused under the write's lock: only a disclosed item about the one file itself counts as the notes saying which name it takes", async () => {
    const words = "Send the supply note with the cover memo in one file and the delivery schedule as dates.docx"
    for (const shared of [true, false]) {
      const directory = matter()
      const documents = [COMBINED, SCHEDULE]
      const plan = writeLeanPlan(directory, documents, "basis-1", undefined, memo(documents))
      // Two files asked for in one sentence, or in two: the one file without a
      // name, and the schedule under a name the plan does not give it.
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
        files: [{ holds: ["memo", "d1"], name: null, quote: words }, { holds: ["d2"], name: "dates.docx", quote: shared ? words : "the delivery schedule as dates.docx" }] })
      const schedule = "# Delivery schedule\n\nGoods arrive on the first day."
      await writeCopies(directory, { ...COPIES, "d2/body-applied.docx": schedule, "d2/redline.docx": schedule })
      await deliver(directory)
      await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })
      // The notes disclose the schedule's file only: the one file is the system's to replace in place.
      const gate = disclosableItems(directory).filter(item => item.detail.startsWith("Requested files:"))
      expect({ shared, gate: gate.map(item => item.detail.includes('ask for "Delivery schedule" as a file named "dates.docx"')) }).toEqual({ shared, gate: [true] })
      discloseAtGate(directory, gate)
      const registration = registering()
      const racing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current, options) => {
        if (name === "supply-pack.docx" && current) registry(directory, db => forgetSystemDocx(db, directory, path.join(directory, name)))
        return registration.registerDeliverable(directory, name, bytes, current, options)
      }
      const delivered = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: racing }, true).then(summary => summary, (error: Error) => error.message)
      expect({ shared, combined: typeof delivered === "string" ? delivered : delivered.documents.filter(row => row.kind === "combined").map(row => row.deliverable) })
        .toEqual({ shared, combined: ["supply-pack.docx", "supply-pack (2).docx"] })
      // The next finalize discloses where the one file went.
      expect(requestedFilesItems(directory).filter(item => item.detail.includes('in one file named "supply-pack (2).docx"'))).toHaveLength(1)
    }
  }, 60_000)

  test("a save refused under the write's lock after the gate's notes said which name the file takes: the run fails for a request without a usable name too", async () => {
    for (const request of ["refused name", "held name"] as const) {
      const directory = matter()
      const words = request === "refused name" ? "Send the supply note with the cover memo in one file named supply-pack.pdf" : "Send the supply note with the cover memo in one file"
      const plan = writeLeanPlan(directory, [COMBINED], "basis-1", undefined, memo([COMBINED]))
      writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
        files: [{ holds: ["memo", "d1"], name: null, quote: words, ...(request === "refused name" ? { refused_name: "supply-pack.pdf" } : {}) }] })
      await writeCopies(directory, COPIES)
      const lawyer = await docx("# Supply pack\n\nA synthetic file saved under this name by hand.")
      // A name the run could not use: the notes say the file the system saved
      // is replaced in place. One file asked for without a name: the notes say
      // the file takes "(2)", since a kept file holds the combined name.
      if (request === "refused name") {
        await deliver(directory)
        await writeCopies(directory, { "d1/body-applied.docx": "# Supply note\n\nThe supplier delivers the goods each week." })
      } else writeFileSync(path.join(directory, "supply-pack.docx"), lawyer)
      const told = request === "refused name" ? "supply-pack.docx" : "supply-pack (2).docx"
      const gate = disclosableItems(directory).filter(item => item.detail.startsWith("Requested files:"))
      expect({ request, told: gate.map(item => item.detail.includes(`in one file named "${told}"`)) }).toEqual({ request, told: [true] })
      discloseAtGate(directory, gate)
      // Under the write's lock: the system's record of its write is gone, or a
      // lawyer's file now holds the name the notes gave.
      const registration = registering()
      const racing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current, options) => {
        if (name === told && current) registry(directory, db => forgetSystemDocx(db, directory, path.join(directory, name)))
        if (name === told && !current && !existsSync(path.join(directory, name))) writeFileSync(path.join(directory, name), lawyer)
        return registration.registerDeliverable(directory, name, bytes, current, options)
      }
      const error = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: racing }, true).then(() => undefined, (error: Error) => error.message)
      const moved = request === "refused name" ? "supply-pack (2).docx" : "supply-pack (3).docx"
      expect({ request, error }).toEqual({ request, error: requestedFilesItems(directory)[0]?.detail })
      expect(error).toContain(`in one file named "${moved}"`)
      expect(combinedRows(directory).at(-1)?.deliverable).toBe(moved)
      if (request === "held name") expect([matterFile(directory, "supply-pack.docx"), matterFile(directory, told)]).toEqual([lawyer, lawyer])
    }
  }, 60_000)

  test("a disclosed item counts for the requested file it is about, never for another file asked for in the same words", () => {
    const directory = matter()
    const words = "Send each supply note on its own as supply-note.pdf"
    // An outline gives each document its own title; this plan is written
    // directly, so two files asked for in the same words have items that open
    // alike and only the file each item is about tells them apart.
    const plan = writeLeanPlan(directory, [SUPPLY, { ...SUPPLY, deliverable: "redline" }], "basis-1")
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
      files: ["d1", "d2"].map(id => ({ holds: [id], name: null, quote: words, refused_name: "supply-note.pdf" })) })
    const items = requestedFilesItems(directory)
    const opening = (detail: string) => detail.slice(0, detail.indexOf(", but the run "))
    expect(items.map(item => opening(item.detail))).toEqual([opening(items[0]!.detail), opening(items[0]!.detail)])
    for (const item of items) expect(requestedFilesItems(directory, [{ ...item, status: "disclosed" }])).toEqual([item])
    // An item from a gate's record before items carried the reference is about no file.
    expect(requestedFilesItems(directory, items.map(item => ({ ...item, ref: { repair_target: "memo" }, status: "disclosed" })))).toEqual([])
  })

  test("a save refused under the write's lock after a matter source named for the request is added: the item the notes disclosed still finds its file, and the run fails", async () => {
    const directory = matter()
    const words = "Send the supply note with the cover memo in one file"
    const plan = writeLeanPlan(directory, [COMBINED], "basis-1", undefined, memo([COMBINED]))
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
      files: [{ holds: ["memo", "d1"], name: null, quote: words }] })
    await writeCopies(directory, COPIES)
    const lawyer = await docx("# Supply pack\n\nA synthetic file saved under this name by hand.")
    writeFileSync(path.join(directory, "supply-pack.docx"), lawyer)
    // The notes say the one file takes "(2)", since a kept file holds the combined name.
    const gate = disclosableItems(directory).filter(item => item.detail.startsWith("Requested files:"))
    expect(gate.map(item => item.detail.includes(`"${words}" [source: request]`) && item.detail.includes('in one file named "supply-pack (2).docx"'))).toEqual([true])
    discloseAtGate(directory, gate)
    // A matter source named "request" then quotes the lawyer's words without "[source: request]".
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    try { db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("request", "1", words, 0, words.length) } finally { db.close() }
    // Under the write's lock a lawyer's file takes the name the notes gave.
    const registration = registering()
    const racing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current, options) => {
      if (name === "supply-pack (2).docx" && !current && !existsSync(path.join(directory, name))) writeFileSync(path.join(directory, name), lawyer)
      return registration.registerDeliverable(directory, name, bytes, current, options)
    }
    const error = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: racing }, true).then(() => undefined, (error: Error) => error.message)
    expect(error).toBe(requestedFilesItems(directory)[0]?.detail)
    expect(error).toContain('in one file named "supply-pack (3).docx"')
    expect(error).not.toContain("[source:")
    // It fails as the item the notes disclosed is about this file and gave another name.
    const about = (items: ReturnType<typeof openItem>[]) => items.map(item => [item.ref.requested_file, item.ref.delivered_names])
    expect([about(gate), about(requestedFilesItems(directory))]).toEqual([[[0, ["supply-pack (2).docx"]]], [[0, ["supply-pack (3).docx"]]]])
  }, 60_000)

  test("a save refused under the write's lock after a matter source named for the request is added: an item the notes disclosed about another file, still delivered where they said, does not fail the run", async () => {
    const directory = matter()
    const documents = [COMBINED, SCHEDULE]
    const plan = writeLeanPlan(directory, documents, "basis-1", undefined, memo(documents))
    // One file asked for without a name, which the notes do not place, and the
    // schedule asked for under a name the plan does not give it.
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-1", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true, files: [
      { holds: ["memo", "d1"], name: null, quote: "Send the supply note with the cover memo in one file" },
      { holds: ["d2"], name: "schedule-review.docx", quote: "Send the delivery schedule as schedule-review.docx" },
    ] })
    await writeCopies(directory, { ...COPIES, "d2/body-applied.docx": "# Delivery schedule\n\nThe goods arrive on Mondays.", "d2/redline.docx": "# Delivery schedule\n\nThe goods arrive on Mondays." })
    const gate = disclosableItems(directory).filter(item => item.detail.startsWith("Requested files:"))
    expect(gate.map(item => [item.ref.requested_file, item.detail.includes("[source: request]"), item.detail.includes('in "Delivery schedule.docx"')])).toEqual([[1, true, true]])
    discloseAtGate(directory, gate)
    // A matter source named "request" then quotes the lawyer's words without "[source: request]".
    const db = new Database(path.join(directory, ".dochaus", "legal.db"))
    try { db.query("INSERT INTO chunks VALUES (?, ?, ?, ?, ?)").run("request", "1", "Send the delivery schedule", 0, 26) } finally { db.close() }
    // Under the write's lock a lawyer's file takes the combined name.
    const lawyer = await docx("# Supply pack\n\nA synthetic file saved under this name by hand.")
    const registration = registering()
    const racing: NonNullable<LeanDeps["registerDeliverable"]> = async (directory, name, bytes, current, options) => {
      if (name === "supply-pack.docx" && !current && !existsSync(path.join(directory, name))) writeFileSync(path.join(directory, name), lawyer)
      return registration.registerDeliverable(directory, name, bytes, current, options)
    }
    const error = await recordLeanDocuments(directory, { sessions: scriptedModel({}).port, registerDeliverable: racing }, true).then(() => undefined, (error: Error) => error.message)
    expect(combinedRows(directory).at(-1)?.deliverable).toBe("supply-pack (2).docx")
    // The schedule's item now reads without the suffix, so its id differs from
    // the one the notes disclosed; it is still about the same file, delivered
    // under the same name, so the notes still say where it went.
    const schedule = requestedFilesItems(directory).find(item => item.ref.requested_file === 1)!
    expect([schedule.id === gate[0]!.id, schedule.detail.includes("[source:"), schedule.detail.includes('in "Delivery schedule.docx"')]).toEqual([false, false, true])
    expect(error).toBeUndefined()
  }, 60_000)

  test("a later plan without the combined delivery handles the saved file as a dropped deliverable's", async () => {
    const dropped = matter()
    writeLeanPlan(dropped, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(dropped, COPIES)
    await deliver(dropped)
    const saved = sha256(matterFile(dropped, "supply-pack.docx"))
    // Without the memo there is no combined file of this run: the file stays,
    // listed as the earlier combined file it is.
    writeLeanPlan(dropped, [COMBINED])
    forgetUnplannedDocuments(dropped)
    const { calls, rows } = await deliver(dropped)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows.filter(([kind]) => kind === "combined")).toEqual([["combined", "d1", "supply-pack.docx"]])
    expect(sha256(matterFile(dropped, "supply-pack.docx"))).toBe(saved)
    // Once the file is gone, so is its row.
    rmSync(path.join(dropped, "supply-pack.docx"))
    expect((await deliver(dropped)).rows.map(([kind]) => kind)).not.toContain("combined")

    // A retitled combined document: its new file takes the next free name, and
    // the earlier one stays listed as an earlier combined file, at a stop at the gate too.
    const retitled = matter()
    writeLeanPlan(retitled, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(retitled, COPIES)
    await deliver(retitled)
    const terms = { ...COMBINED, title: "Supply terms" }
    writeLeanPlan(retitled, [terms], undefined, undefined, memo([terms]))
    forgetUnplannedDocuments(retitled)
    await recordLeanDocuments(retitled, { sessions: scriptedModel({}).port }, false)
    expect(combinedRows(retitled)).toEqual([expect.objectContaining({ title: "Supply note", path: "supply-pack.docx", deliverable: "supply-pack.docx" })])
    await writeCopies(retitled, { "d1/body-applied.docx": "# Supply terms\n\nThe supplier delivers the goods each week." })
    await deliver(retitled)
    expect(combinedRows(retitled).map(row => [row.title, row.deliverable])).toEqual([["Supply note", "supply-pack.docx"], ["Supply terms", "supply-pack (2).docx"]])
    expect(combinedParts(retitled)?.file).toBe("supply-pack (2).docx")

    // Unlike an earlier file of the same document, a combined file is never
    // replaced by a body or redline now planned under its name: only the
    // combined save, which keeps pending proposals, writes over it. The part
    // takes the next free name; the file and its pending proposal stay, listed
    // as an earlier combined file.
    for (const deliverable of ["body", "redline"] as const) {
      for (const forget of [false, true]) {
        const renamed = matter()
        writeLeanPlan(renamed, [COMBINED], undefined, undefined, memo([COMBINED]))
        await writeCopies(renamed, COPIES)
        await deliver(renamed)
        const combined = sha256(matterFile(renamed, "supply-pack.docx"))
        propose(renamed, "supply-pack.docx")
        const part = { ...COMBINED, deliverable }
        writeLeanPlan(renamed, [part], undefined, undefined, memo([part]))
        if (forget) forgetUnplannedDocuments(renamed)
        const again = await deliver(renamed)
        expect(again.calls.map(([name]) => name)).not.toContain("supply-pack.docx")
        expect(again.rows).toContainEqual([deliverable, "d1", "supply-pack (2).docx"])
        expect(again.rows.filter(([kind]) => kind === "combined")).toEqual([["combined", "d1", "supply-pack.docx"]])
        expect(sha256(matterFile(renamed, "supply-pack.docx"))).toBe(combined)
        expect(pendingOn(renamed, "supply-pack.docx")).toBe(1)
      }
    }
  }, 60_000)

  test("a document requested by name under an earlier combined file's name: the gate discloses that a kept file holds the name before the run saves the document under another", async () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    await writeCopies(directory, COPIES)
    await deliver(directory)
    const combined = matterFile(directory, "supply-pack.docx")
    const body = { ...COMBINED, deliverable: "body" as const }
    const words = "Send the supply note as supply-pack.docx"
    const plan = writeLeanPlan(directory, [body], "basis-2", undefined, memo([body]))
    writeRequestedFiles(directory, { schema: LEAN_REQUESTED_FILES_SCHEMA, basis: "basis-2", planned: requestedFilesPlanned(plan.documents, plan.memo), stated: true,
      files: [{ holds: ["d1"], name: "supply-pack.docx", quote: words }] })
    forgetUnplannedDocuments(directory)

    const item = openItem("finalize", "omission", `Requested files: the lawyer's words "${words}" [source: request] ask for "Supply note" as a file named "supply-pack.docx", ` +
      'but the run delivers "Supply note" in a file of another name, since a file already in the matter holds the name "supply-pack.docx". ' +
      "The lawyer should confirm how these files are to be delivered.", { repair_target: "memo", requested_file: 0, delivered_names: [null] })
    // Predicted at the gate, and the same item once the document is saved.
    expect(disclosableItems(directory)).toContainEqual(item)
    discloseAtGate(directory, [item])
    const { calls, rows } = await deliver(directory)
    expect(calls.map(([name]) => name)).not.toContain("supply-pack.docx")
    expect(rows).toContainEqual(["body", "d1", "supply-pack (2).docx"])
    expect(matterFile(directory, "supply-pack.docx")).toEqual(combined)
    expect(requestedFilesItems(directory)).toEqual([item])
  }, 60_000)

  test("an earlier combined file the plan no longer delivers stays listed only while it is the system's, or its accepted revision", async () => {
    for (const change of ["retitle", "drop"] as const) {
      for (const write of ["outside", "accept", "bound accept", "bound accept, then outside", "unreadable record", "none"] as const) {
        const directory = matter()
        writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
        await writeCopies(directory, COPIES)
        await deliver(directory)
        const terms = { ...COMBINED, title: "Supply terms" }
        if (change === "retitle") writeLeanPlan(directory, [terms], undefined, undefined, memo([terms]))
        else writeLeanPlan(directory, [COMBINED])
        forgetUnplannedDocuments(directory)
        await deliver(directory)
        const file = path.join(directory, "supply-pack.docx")
        if (write.startsWith("bound accept") || write === "accept") await acceptIntoCombined(directory, write !== "accept")
        // A save outside doc.haus leaves the record of the system's write, or of
        // the lawyer's accepted Word revision, which no longer matches.
        if (write.endsWith("outside")) writeFileSync(file, await docx("# Supply pack\n\nA synthetic file saved under this name by hand."))
        if (write === "unreadable record") registry(directory, db => db.query("UPDATE meta SET value = ? WHERE key = ?").run(JSON.stringify({ version: 2 }), systemDocxKey(directory, file)))
        const saved = sha256(matterFile(directory, "supply-pack.docx"))
        const listed = (await deliver(directory)).rows.filter(([kind, , name]) => kind === "combined" && name === "supply-pack.docx")
        const kept = ["accept", "bound accept", "none"].includes(write)
        expect({ change, write, listed }).toEqual({ change, write, listed: kept ? [["combined", "d1", "supply-pack.docx"]] : [] })
        expect(sha256(matterFile(directory, "supply-pack.docx"))).toBe(saved)
      }
    }
  }, 60_000)

  test("without a combined delivery, registration calls, the summary and the items are as before", async () => {
    const plain = matter()
    writeLeanPlan(plain, [SUPPLY])
    await writeCopies(plain, COPIES)
    expect(disclosableItems(plain)).toEqual([])
    expect((await deliver(plain)).calls).toEqual([["Supply note.docx", null], ["Supply note - redline.docx", null], ["Supply note - drafting notes.docx", null]])
    const row = (kind: LeanDocumentKind, file: string, deliverable: string) => ({ kind, document: "d1" as const, title: "Supply note", path: `.dochaus/lean/${file}`, sha256: sha256(matterFile(plain, deliverable)), deliverable })
    expect(readLeanDocuments(plain)).toEqual({ schema: "dochaus.lean.documents/2", status: "completed", documents: [
      row("body", "d1/body-applied.docx", "Supply note.docx"),
      row("redline", "d1/redline.docx", "Supply note - redline.docx"),
      { ...row("drafting_notes", "drafting-notes.docx", "Supply note - drafting notes.docx"), document: null, title: null },
    ] })
    const registered = readLeanDocuments(plain)!.documents.map(row => [row.deliverable, row.sha256])
    expect((await deliver(plain)).calls).toEqual(registered.map(([name, sha]) => [name!, sha!]))

    // A memo with each document in its own file.
    const named = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx", deliverable: "body" }, { ...SCHEDULE, file: "schedule-review.docx", deliverable: "redline" }]
    writeLeanPlan(named, documents, undefined, undefined, memo(documents))
    await writeCopies(named, { ...COPIES, "d2/body-applied.docx": "# Delivery schedule\n\nGoods arrive on the first day.", "d2/redline.docx": "# Delivery schedule\n\nGoods arrive on the first day." })
    expect(disclosableItems(named)).toEqual([])
    expect((await deliver(named)).calls).toEqual([["supply-pack.docx", null], ["supply-pack - redline.docx", null], ["schedule-review - clean.docx", null],
      ["schedule-review.docx", null], ["supply-pack - cover memo.docx", null]])
    expect(readLeanDocuments(named)!.documents.map(row => [row.kind, row.document, row.title, row.path, row.deliverable, row.sha256 === sha256(matterFile(named, row.deliverable!))])).toEqual([
      ["body", "d1", "Supply note", ".dochaus/lean/d1/body-applied.docx", "supply-pack.docx", true],
      ["redline", "d1", "Supply note", ".dochaus/lean/d1/redline.docx", "supply-pack - redline.docx", true],
      ["body", "d2", "Delivery schedule", ".dochaus/lean/d2/body-applied.docx", "schedule-review - clean.docx", true],
      ["redline", "d2", "Delivery schedule", ".dochaus/lean/d2/redline.docx", "schedule-review.docx", true],
      ["drafting_notes", null, null, ".dochaus/lean/drafting-notes.docx", "supply-pack - cover memo.docx", true],
    ])
  }, 60_000)
})

describe("delivery names that collide", () => {
  const named = (documents: Document[], name: string) => ({ name, file: memoNotesFile(documents, name) })

  test("another document named like the combined document does not take the combined file's name", async () => {
    const directory = matter()
    const documents: Document[] = [COMBINED, { ...SCHEDULE, title: "supply-pack" }]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    const { names, rows } = await registered(directory, documents)
    expect(rows).toContainEqual(["body", "d2", "supply-pack (2).docx"])
    expect(rows).toContainEqual(["combined", "d1", "supply-pack.docx"])
    expect(names.filter(name => name === "supply-pack.docx")).toHaveLength(1)
  }, 60_000)

  test("the memo does not take a later combined document's name", async () => {
    // A memo file not named after the combined document (an earlier plan's, or a
    // lawyer's edit) can repeat its name.
    const directory = matter()
    const documents: Document[] = [SCHEDULE, { ...SUPPLY, file: "Delivery schedule - cover memo.docx", deliverable: "combined" }]
    writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo", file: "Delivery schedule - cover memo.docx" })
    const { names, rows } = await registered(directory, documents)
    expect(rows).toContainEqual(["drafting_notes", null, "Delivery schedule - cover memo (2).docx"])
    expect(rows).toContainEqual(["combined", "d2", "Delivery schedule - cover memo.docx"])
    expect(names.filter(name => name === "Delivery schedule - cover memo.docx")).toHaveLength(1)
  }, 60_000)

  test("a memo named like the redline yields: the redline keeps its name", async () => {
    const directory = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx" }]
    writeLeanPlan(directory, documents, undefined, undefined, named(documents, "redline"))
    expect((await registered(directory, documents)).rows).toEqual([
      ["body", "d1", "supply-pack.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["drafting_notes", null, "supply-pack - redline (2).docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)

  test("a memo named like the clean body yields: the combined or redline document's body keeps its name", async () => {
    const combined = matter()
    writeLeanPlan(combined, [COMBINED], undefined, undefined, named([COMBINED], "clean"))
    expect((await registered(combined, [COMBINED])).rows).toEqual([
      ["body", "d1", "supply-pack - clean.docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["drafting_notes", null, "supply-pack - clean (2).docx"],
      ["drafting_notes_markdown", null, null],
      ["combined", "d1", "supply-pack.docx"],
    ])

    const redline = matter()
    const documents: Document[] = [{ ...SUPPLY, file: "supply-pack.docx", deliverable: "redline" }]
    writeLeanPlan(redline, documents, undefined, undefined, named(documents, "clean"))
    expect((await registered(redline, documents)).rows).toEqual([
      ["body", "d1", "supply-pack - clean.docx"],
      ["redline", "d1", "supply-pack.docx"],
      ["drafting_notes", null, "supply-pack - clean (2).docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)

  test("without a memo or a combined document, colliding names resolve as before", async () => {
    // A requested name stays its document's; an untitled document with the same title yields.
    const requested = matter()
    const documents: Document[] = [{ ...SCHEDULE, title: "supply-pack" }, { ...SUPPLY, file: "supply-pack.docx" }]
    writeLeanPlan(requested, documents)
    expect((await registered(requested, documents)).rows).toEqual([
      ["body", "d1", "supply-pack (2).docx"],
      ["redline", "d1", "supply-pack - redline.docx"],
      ["body", "d2", "supply-pack.docx"],
      ["redline", "d2", "supply-pack - redline (2).docx"],
      ["drafting_notes", null, "supply-pack - drafting notes.docx"],
      ["drafting_notes_markdown", null, null],
    ])

    // A notes name the delivery selection requested keeps its name.
    const notes = matter()
    const plan = writeLeanPlan(notes, [SUPPLY, SCHEDULE])
    writeLeanText(notes, "plan.json", `${JSON.stringify({ ...plan, notes_file: "Delivery schedule.docx" }, null, 2)}\n`)
    expect((await registered(notes, [SUPPLY, SCHEDULE])).rows).toEqual([
      ["body", "d1", "Supply note.docx"],
      ["redline", "d1", "Supply note - redline.docx"],
      ["body", "d2", "Delivery schedule (2).docx"],
      ["redline", "d2", "Delivery schedule - redline.docx"],
      ["drafting_notes", null, "Delivery schedule.docx"],
      ["drafting_notes_markdown", null, null],
    ])
  }, 60_000)
})

describe("combinedDelivery", () => {
  test("is the combined document's id and its requested delivery stem, given a memo", () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(combinedDelivery(directory)).toEqual({ document: "d1", name: "supply-pack" })
    // Without a requested file name the stem is the title, as the body's would be.
    const untitled = { ...SUPPLY, deliverable: "combined" as const }
    writeLeanPlan(directory, [SCHEDULE, untitled], undefined, undefined, memo([SCHEDULE, untitled]))
    expect(combinedDelivery(directory)).toEqual({ document: "d2", name: "Supply note" })
  })

  test("is undefined without a memo, without a combined document, or with more than one", () => {
    const directory = matter()
    writeLeanPlan(directory, [COMBINED])
    expect(combinedDelivery(directory)).toBeUndefined()
    const body = { ...COMBINED, deliverable: "body" as const }
    writeLeanPlan(directory, [body], undefined, undefined, memo([body]))
    expect(combinedDelivery(directory)).toBeUndefined()
    const both = [COMBINED, { ...SCHEDULE, deliverable: "combined" as const }]
    writeLeanPlan(directory, both, undefined, undefined, memo(both))
    expect(combinedDelivery(directory)).toBeUndefined()
  })

  test("a name cut at the length limit keeps whole characters", () => {
    // 119 letters, then a character outside the Basic Multilingual Plane at the cut.
    const long = { ...COMBINED, file: `${"a".repeat(119)}\u{1F600}tail.docx` }
    const directory = matter()
    writeLeanPlan(directory, [long], undefined, undefined, memo([long]))
    const name = combinedDelivery(directory)!.name
    expect(name).toBe(`${"a".repeat(119)}\u{1F600}`)
    expect(encodeURIComponent(`${name}.docx`)).toEndWith("%F0%9F%98%80.docx")
    // The memo's file shortens the document's stem, at the same character here.
    const notes = memoNotesFile([{ title: "x", file: `${"a".repeat(106)}\u{1F600}tail.docx` }], "cover memo")
    expect(notes).toBe(`${"a".repeat(106)}\u{1F600} - cover memo.docx`)
  })
})

describe("the requested memo's notes file in a combined package", () => {
  test("is named after the combined document wherever it stands, before another document's requested file", () => {
    expect(memoNotesFile([SCHEDULE, COMBINED], "cover memo")).toBe("supply-pack - cover memo.docx")
    // An earlier document titled like the memo, without a file, does not name it.
    expect(memoNotesFile([{ ...SCHEDULE, title: "Cover memo" }, COMBINED], "cover memo")).toBe("supply-pack - cover memo.docx")
    expect(memoNotesFile([{ ...SCHEDULE, file: "schedule.docx" }, { ...SUPPLY, deliverable: "combined" }], "cover memo")).toBe("Supply note - cover memo.docx")
    // More than one combined document builds no combined file (combinedDelivery).
    expect(memoNotesFile([SCHEDULE, { ...SUPPLY, deliverable: "combined" }, { ...COMBINED, title: "Annex", file: "annex.docx" }], "cover memo")).toBe("annex - cover memo.docx")
  })

  test("registration and the combined file's parts bind the notes to that name", async () => {
    const directory = matter()
    const documents: Document[] = [SCHEDULE, COMBINED]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    const { names, rows } = await registered(directory, documents)
    expect(rows).toContainEqual(["drafting_notes", null, "supply-pack - cover memo.docx"])
    expect(names.at(-1)).toBe("supply-pack.docx")
    expect(combinedParts(directory)).toEqual({ document: "d2", name: "supply-pack", file: "supply-pack.docx", body: "supply-pack - clean.docx", notes: "supply-pack - cover memo.docx" })
  }, 60_000)
})

describe("package review's selection on a combined document", () => {
  const INSTRUCTION = "Send the note and the cover memo together as Supply Pack Final.docx"
  const selection = (role: "body" | "redline", file = "Supply Pack Final.docx") => ({ document: "d1" as const, role, file, instruction_quote: INSTRUCTION })

  test("a body selection names the combined file and keeps it combined; the same name is no change", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(checkDeliverySelections(directory, [selection("body")])).toEqual([])

    const selected = selectedDeliveryPlan(directory, [selection("body")])!
    expect(selected.documents[0]).toEqual({ ...COMBINED, id: "d1", file: "Supply Pack Final.docx", deliverable: "combined" })
    expect(selected).toMatchObject({ memo: "cover memo", notes_file: "Supply Pack Final - cover memo.docx" })
    writeLeanText(directory, "plan.json", `${JSON.stringify(selected, null, 2)}\n`)
    expect(combinedDelivery(directory)).toEqual({ document: "d1", name: "Supply Pack Final" })
    expect(selectedDeliveryPlan(directory, [selection("body")])).toBeUndefined()
  })

  test("a redline selection still makes it a redline deliverable", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(selectedDeliveryPlan(directory, [selection("redline")])?.documents[0]).toMatchObject({ file: "Supply Pack Final.docx", deliverable: "redline" })
    expect(selectedDeliveryPlan(directory, [selection("redline", COMBINED.file)])?.documents[0]).toMatchObject({ file: "supply-pack.docx", deliverable: "redline" })
  })

  test("a body selection on a combined document that is not first moves the memo's file with it; renaming another document leaves it", () => {
    const directory = matter()
    const documents: Document[] = [{ ...SCHEDULE, file: "schedule.docx" }, COMBINED]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    expect(selectedDeliveryPlan(directory, [{ ...selection("body"), document: "d2" }])).toMatchObject({ memo: "cover memo", notes_file: "Supply Pack Final - cover memo.docx" })
    expect(selectedDeliveryPlan(directory, [selection("body", "Schedule Final.docx")])).toMatchObject({ memo: "cover memo", notes_file: "supply-pack - cover memo.docx" })
    // A memo file not named this way is left alone.
    writeLeanPlan(directory, documents, undefined, undefined, { name: "cover memo", file: "schedule - cover memo.docx" })
    expect(selectedDeliveryPlan(directory, [{ ...selection("body"), document: "d2" }])).toMatchObject({ memo: "cover memo", notes_file: "schedule - cover memo.docx" })
  })

  const ONE_ROW = '"Supply Pack Final.docx" is the one file that holds the memo and document d1. Select this file by one row: document "d1", role body, filename "Supply Pack Final.docx". Keep the other rows. The code builds it from the memo and the clean body at delivery; give no drafting_notes row for it.'
  const ONE_DOCUMENT_ROW = "The plan supports one named primary artifact per document and one name for the run's notes. Report conflicting or additional requested outputs as a package finding rather than choosing silently."
  const notesRow = { document: null, role: "drafting_notes" as const, file: "Supply Pack Final.docx", instruction_quote: INSTRUCTION }

  test("a notes row beside the body row under the combined file's name is refused with the one-row rule, not as a package finding", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(checkDeliverySelections(directory, [notesRow, selection("body")])).toEqual([ONE_ROW])
  })

  test("a notes row under the plan's combined file name is refused with the one-row rule when no body row names the file", () => {
    // Registration would otherwise name the notes after the combined file and suffix them.
    const SCHEDULE_INSTRUCTION = "Deliver the schedule as Schedule Final.docx"
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}. ${SCHEDULE_INSTRUCTION}.`)
    const documents: Document[] = [{ ...COMBINED, file: "supply pack final.docx" }, SCHEDULE]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    expect(checkDeliverySelections(directory, [notesRow])).toEqual([ONE_ROW])
    // A row for another document does not name the combined file.
    expect(checkDeliverySelections(directory, [{ document: "d2", role: "body", file: "Schedule Final.docx", instruction_quote: SCHEDULE_INSTRUCTION }, notesRow])).toEqual([ONE_ROW])
  })

  test("a notes row under a name the reply leaves free, or without a combined file, keeps today's result", () => {
    // The reply renames the combined file, or delivers that document as a redline (selectedDeliveryPlan).
    const NOTE_INSTRUCTION = "Deliver the note itself as Supply Note Final.docx"
    const combined = matter()
    writeLeanText(combined, "request.md", `${INSTRUCTION}. ${NOTE_INSTRUCTION}.`)
    const documents: Document[] = [{ ...COMBINED, file: "Supply Pack Final.docx" }]
    writeLeanPlan(combined, documents, undefined, undefined, memo(documents))
    for (const role of ["body", "redline"] as const)
      expect(checkDeliverySelections(combined, [{ ...selection(role, "Supply Note Final.docx"), instruction_quote: NOTE_INSTRUCTION }, notesRow])).toEqual([])
    // Another notes name in a combined package.
    expect(checkDeliverySelections(combined, [{ ...notesRow, file: "Supply Note Final.docx", instruction_quote: NOTE_INSTRUCTION }])).toEqual([])
    // A combined document without a memo builds no combined file.
    const plain = matter()
    writeLeanText(plain, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(plain, documents)
    expect(checkDeliverySelections(plain, [notesRow])).toEqual([])
  })

  test("a row under a name the plan already gives another artifact is refused as a package finding", () => {
    // The applied plan (selectedDeliveryPlan) would name two artifacts alike; registration would suffix one.
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const documents: Document[] = [{ ...COMBINED, file: "Supply Pack Final.docx" }, SCHEDULE]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    const other = (role: "body" | "redline") => ({ ...selection(role), document: "d2" as const })
    for (const role of ["body", "redline"] as const) expect(checkDeliverySelections(directory, [other(role)])).toEqual([PACKAGE_FINDING])
    expect(checkDeliverySelections(directory, [other("body"), notesRow])).toEqual([PACKAGE_FINDING])
    // A notes row under a plain body document's file, beside the memo.
    const plain = matter()
    writeLeanText(plain, "request.md", `${INSTRUCTION}.`)
    const body: Document[] = [{ ...SUPPLY, file: "Supply Pack Final.docx", deliverable: "body" }]
    writeLeanPlan(plain, body, undefined, undefined, memo(body))
    expect(checkDeliverySelections(plain, [notesRow])).toEqual([PACKAGE_FINDING])
    // A repeat the plan has without any row of the reply stays registration's to suffix.
    writeLeanPlan(plain, [body[0]!, { ...SCHEDULE, file: "Supply Pack Final.docx" }], undefined, undefined, memo(body))
    expect(checkDeliverySelections(plain, [])).toEqual([])
  })

  test("the checker reads the row selectedDeliveryPlan applies when a document has two", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(checkDeliverySelections(directory, [selection("redline"), selection("body"), notesRow])).toEqual([ONE_DOCUMENT_ROW, ONE_ROW])
  })

  test("a notes name the saved plan already gives the combined file is left to registration: a reply without rows changes nothing", async () => {
    // No reply sets it, so applying the reply is no delivery change (reconcilePackageReview charges no repair).
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const documents: Document[] = [{ ...COMBINED, file: "Supply Pack Final.docx" }]
    const plan = writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...plan, notes_file: "Supply Pack Final.docx" }, null, 2)}\n`)
    expect(checkDeliverySelections(directory, [])).toEqual([])
    expect(selectedDeliveryPlan(directory, [])).toBeUndefined()
    expect((await registered(directory, documents)).rows).toContainEqual(["drafting_notes", null, "Supply Pack Final (2).docx"])
  }, 60_000)

  test("two passes: a notes row under the requested name, then the body row naming the combined file so, leave the notes the memo's default name", async () => {
    const fallback = "Supply Pack Final - cover memo.docx"
    const apply = (directory: string, reply: DeliverySelection[]) => {
      expect(checkDeliverySelections(directory, reply)).toEqual([])
      const plan = selectedDeliveryPlan(directory, reply)
      if (plan) writeLeanText(directory, "plan.json", `${JSON.stringify(plan, null, 2)}\n`)
      return plan
    }
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(apply(directory, [notesRow])?.notes_file).toBe("Supply Pack Final.docx")
    expect(apply(directory, [selection("body")])).toMatchObject({ documents: [{ file: "Supply Pack Final.docx", deliverable: "combined" }], notes_file: fallback })
    // Applying the second reply again changes nothing.
    expect(selectedDeliveryPlan(directory, [selection("body")])).toBeUndefined()
    const { names, rows } = await registered(directory, [COMBINED])
    expect(rows).toContainEqual(["drafting_notes", null, fallback])
    expect(names.filter(name => / \(\d+\)\.docx$/.test(name))).toEqual([])
  }, 60_000)

  test("a reply with no row for the combined document leaves a saved notes name untouched", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const documents: Document[] = [COMBINED, SCHEDULE]
    const plan = writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...plan, notes_file: "Supply Pack Final.docx" }, null, 2)}\n`)
    // Another document's body row, even under the notes name.
    expect(selectedDeliveryPlan(directory, [{ ...selection("body"), document: "d2" }])).toMatchObject({ documents: [{ file: "supply-pack.docx" }, { file: "Supply Pack Final.docx" }], notes_file: "Supply Pack Final.docx" })
  })

  test("a body row repeating the combined file's name renames nothing: a saved notes name equal to it stays", () => {
    // A notes name an earlier contract saved under the combined file's name.
    const directory = matter()
    const documents: Document[] = [{ ...COMBINED, file: "Supply Pack Final.docx" }]
    const plan = writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...plan, notes_file: "Supply Pack Final.docx" }, null, 2)}\n`)
    expect(selectedDeliveryPlan(directory, [selection("body")])).toBeUndefined()
  })

  test("a redline row onto the saved notes name leaves it: the notes and the redline keep the package-finding refusal", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const plan = writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...plan, notes_file: "Supply Pack Final.docx" }, null, 2)}\n`)
    expect(selectedDeliveryPlan(directory, [selection("redline")])).toMatchObject({ documents: [{ file: "Supply Pack Final.docx", deliverable: "redline" }], notes_file: "Supply Pack Final.docx" })
    expect(checkDeliverySelections(directory, [selection("redline")])).toEqual([PACKAGE_FINDING])
  })

  test("a body row renaming the combined file onto the saved notes name in another case gives the notes the memo's default name", () => {
    const directory = matter()
    writeLeanText(directory, "request.md", `${INSTRUCTION}.`)
    const plan = writeLeanPlan(directory, [COMBINED], undefined, undefined, memo([COMBINED]))
    writeLeanText(directory, "plan.json", `${JSON.stringify({ ...plan, notes_file: "supply pack final.docx" }, null, 2)}\n`)
    expect(checkDeliverySelections(directory, [selection("body")])).toEqual([])
    expect(selectedDeliveryPlan(directory, [selection("body")])).toMatchObject({ documents: [{ file: "Supply Pack Final.docx", deliverable: "combined" }], notes_file: "Supply Pack Final - cover memo.docx" })
  })

  test("any other repeated file name keeps the package-finding refusal", () => {
    // Two documents under one name in a combined package.
    const combined = matter()
    writeLeanText(combined, "request.md", `${INSTRUCTION}.`)
    const documents: Document[] = [COMBINED, SCHEDULE]
    writeLeanPlan(combined, documents, undefined, undefined, memo(documents))
    expect(checkDeliverySelections(combined, [selection("body"), { ...selection("body"), document: "d2" }])).toEqual([PACKAGE_FINDING])
    // A notes row under a redline row's name.
    expect(checkDeliverySelections(combined, [notesRow, selection("redline")])).toEqual([PACKAGE_FINDING])
    // The same two rows without a combined document.
    const plain = matter()
    writeLeanText(plain, "request.md", `${INSTRUCTION}.`)
    writeLeanPlan(plain, [SUPPLY], undefined, undefined, memo([SUPPLY]))
    expect(checkDeliverySelections(plain, [notesRow, selection("body")])).toEqual([PACKAGE_FINDING])
  })

  test("the selection contract states the one-row rule only for a combined package", () => {
    const rule = 'If the lawyer names the file that holds the memo and document d1, select it by one row: document "d1", role body, that filename. The code builds that file from the memo and the clean body at delivery, so give no drafting_notes row for it.'
    const combined = matter()
    writeLeanPlan(combined, [COMBINED], undefined, undefined, memo([COMBINED]))
    expect(packageReviewPrompt(combined)).toContain(rule)
    const plain = matter()
    writeLeanPlan(plain, [SUPPLY], undefined, undefined, memo([SUPPLY]))
    expect(packageReviewPrompt(plain)).not.toContain("holds the memo and document")
    writeLeanPlan(plain, [COMBINED])
    expect(packageReviewPrompt(plain)).not.toContain("holds the memo and document")
  })
})

describe("package review's selection names, as registration gives them", () => {
  const check = (directory: string, request: string, reply: DeliverySelection[]) => {
    writeLeanText(directory, "request.md", request)
    return checkDeliverySelections(directory, reply)
  }
  const notesRow = (file: string, instruction_quote: string) => ({ document: null, role: "drafting_notes" as const, file, instruction_quote })

  test("a body row that moves the memo's name onto another document's requested name is refused", () => {
    // The memo follows its anchor's rename (selectedDeliveryPlan) to "New - cover memo.docx", which d2 was requested under.
    const directory = matter()
    const instruction = "Deliver the supply note as New.docx"
    const documents: Document[] = [{ ...SUPPLY, file: "Supply.docx" }, { ...SCHEDULE, file: "New - cover memo.docx" }]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    expect(check(directory, `${instruction}. Deliver the schedule as New - cover memo.docx.`, [{ document: "d1", role: "body", file: "New.docx", instruction_quote: instruction }])).toEqual([PACKAGE_FINDING])
  })

  test("a memo's notes row under a document's unrequested name is refused: the memo would yield it", () => {
    const directory = matter()
    const instruction = "Deliver the memo as Supply note.docx"
    writeLeanPlan(directory, [SUPPLY], undefined, undefined, memo([SUPPLY]))
    expect(check(directory, `${instruction}.`, [notesRow("Supply note.docx", instruction)])).toEqual([PACKAGE_FINDING])
  })

  test("a memo's notes row under the combined document's unrequested redline name is refused", () => {
    const directory = matter()
    const instruction = "Deliver the memo as Pack - redline.docx"
    const documents: Document[] = [{ ...SUPPLY, file: "Pack.docx", deliverable: "combined" }]
    writeLeanPlan(directory, documents, undefined, undefined, memo(documents))
    expect(check(directory, `${instruction}.`, [notesRow("Pack - redline.docx", instruction)])).toEqual([PACKAGE_FINDING])
  })

  test("a row for one document passes while two other planned files already share a name", () => {
    const directory = matter()
    const instruction = "Deliver the annex as Annex Final.docx"
    writeLeanPlan(directory, [{ ...SUPPLY, file: "Same.docx" }, { ...SCHEDULE, file: "Same.docx" }, { ...SCHEDULE, title: "Annex" }])
    expect(check(directory, `${instruction}.`, [{ document: "d3", role: "body", file: "Annex Final.docx", instruction_quote: instruction }])).toEqual([])
  })

  test("a notes name only an existing matter document holds passes: registration keeps that document", async () => {
    const directory = matter()
    const instruction = "Deliver the memo as Supply advice.docx"
    writeFileSync(path.join(directory, "Supply advice.docx"), (await renderMarkdownToDocx("# Source\n\nInput document.")).bytes)
    writeLeanPlan(directory, [SUPPLY], undefined, undefined, memo([SUPPLY]))
    expect(check(directory, `${instruction}.`, [notesRow("Supply advice.docx", instruction)])).toEqual([])
  })
})
