// A source ingest extracted from an email or zip is shown with its container
// chain beside its plain name; its name, text and offsets are unchanged.
// Run from dochaus/: bun test test/source-provenance.test.ts
//
// Every fixture below is invented test content: no actual person,
// organization, transaction or obligation.
import { Database } from "bun:sqlite"
import { afterEach, expect, mock, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { formatCitations, type DocumentCitation } from "../lib/citations"
import { liveText } from "../lib/extract"
import { provenanceChains, sourceProvenance } from "../lib/source-provenance"
import { LegalPlugin } from "../plugin/legal/policy"

mock.module("@huggingface/transformers", () => ({
  env: {},
  pipeline: async () => async () => ({ data: new Float32Array(384) }),
}))

const directories: string[] = []
const fetchOriginal = globalThis.fetch
afterEach(() => {
  globalThis.fetch = fetchOriginal
  directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true }))
})

function matterDirectory() {
  const directory = mkdtempSync(path.join(tmpdir(), "source-provenance-"))
  directories.push(directory)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "provenance-test" }))
  return directory
}

// Thread.eml carried Exhibits.zip, which holds Exhibit A.eml; Notes.docx is a plain upload.
const sources = [
  { document: "Thread.eml" },
  { document: "Schedule.docx", attachment_of: "Thread.eml" },
  { document: "Exhibits.zip", attachment_of: "Thread.eml" },
  { document: "Exhibit A.eml", attachment_of: "Exhibits.zip" },
  { document: "Notes.docx" },
]

function writeIndex(directory: string) {
  writeFileSync(path.join(directory, "source-index.json"), JSON.stringify({
    sources: sources.map((source) => ({ ...source, extracted_text: `sources/${source.document}.txt` })),
  }))
}

test("an attachment, a zip entry inside an attachment and a plain upload", () => {
  const chains = provenanceChains(sources)
  expect(chains.get("Schedule.docx")).toBe("attachment of Thread.eml")
  expect(chains.get("Exhibit A.eml")).toBe("in Exhibits.zip, attachment of Thread.eml")
  expect(chains.has("Notes.docx")).toBe(false)
  expect(chains.has("Thread.eml")).toBe(false)
})

test("a container cycle ends the chain instead of looping", () => {
  const chains = provenanceChains([
    { document: "A.eml", attachment_of: "B.zip" },
    { document: "B.zip", attachment_of: "A.eml" },
    { document: "Self.eml", attachment_of: "Self.eml" },
  ])
  expect(chains.get("A.eml")).toBe("in B.zip")
  expect(chains.get("B.zip")).toBe("attachment of A.eml")
  expect(chains.has("Self.eml")).toBe(false)
})

test("a container no longer in the matter is named as removed and ends the chain", () => {
  expect(provenanceChains([{ document: "Kept.pdf", attachment_of: "Gone.eml" }]).get("Kept.pdf"))
    .toBe("attachment of Gone.eml (removed from matter)")
})

test("a missing or malformed index gives no provenance", () => {
  const directory = matterDirectory()
  expect(sourceProvenance(directory).size).toBe(0)
  writeFileSync(path.join(directory, "source-index.json"), "{not json")
  expect(sourceProvenance(directory).size).toBe(0)
})

test("citations show provenance once per document, outside the heading", () => {
  const passage = (section: string): DocumentCitation => ({
    documentName: "Schedule.docx", docPath: "/m/Schedule.docx", section, excerpt: `Schedule text ${section}.`,
    charStart: 0, charEnd: 10, provenance: "attachment of Thread.eml",
  })
  const rendered = formatCitations([passage("1"), passage("2")])
  expect(rendered).toContain("1. [Schedule.docx § 1]\n   provenance: attachment of Thread.eml\nSchedule text 1.")
  expect(rendered).toContain("2. [Schedule.docx § 2]\nSchedule text 2.")
  expect(rendered.match(/provenance:/g)).toHaveLength(1)
})

test("search-document carries provenance without changing the cited name or offsets", async () => {
  const searchDocument = (await import("../tool/search-document")).default
  const directory = matterDirectory()
  writeIndex(directory)
  const docPath = path.join(directory, "Exhibit A.eml")
  writeFileSync(docPath, "Subject: Exhibit A\r\nFrom: seller@example.invalid\r\nTo: buyer@example.invalid\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nThe purchase price is payable in two instalments.\r\n")
  const excerpt = "The purchase price is payable in two instalments."
  const start = (await liveText(docPath)).indexOf(excerpt)
  expect(start).toBeGreaterThan(0)

  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  for (const sql of [
    "CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT, created_at INTEGER)",
    "CREATE TABLE chunks (id INTEGER PRIMARY KEY, document_id INTEGER, doc_path TEXT, doc_name TEXT, section TEXT, structural_path TEXT, party TEXT, chunk_index INTEGER, text TEXT, char_start INTEGER, char_end INTEGER, embedding BLOB, flagged INTEGER)",
    "CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, term TEXT, definition TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE doc_relations (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, relation TEXT, target_name TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE section_ops (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section_label TEXT, operation TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE redlines (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, scope TEXT, find_text TEXT, old_text TEXT, new_text TEXT, author TEXT, review_comment TEXT NOT NULL DEFAULT '', anchor_id TEXT, status TEXT, created_at INTEGER)",
    "CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)",
    "CREATE VIRTUAL TABLE chunks_fts USING fts5(text, section, doc_name, content='chunks', content_rowid='id')",
  ]) db.run(sql)
  db.run("INSERT INTO meta (key, value) VALUES ('embedding_model', 'onnx-community/granite-embedding-small-english-r2-ONNX')")
  db.run("INSERT INTO documents (id, doc_path, name, created_at) VALUES (1, ?, 'Exhibit A.eml', 1)", [docPath])
  db.run(
    "INSERT INTO chunks (id, document_id, doc_path, doc_name, section, structural_path, party, text, char_start, char_end, embedding, flagged) VALUES (1, 1, ?, 'Exhibit A.eml', '1', 'Exhibit A.eml › 1', '', ?, ?, ?, ?, 0)",
    [docPath, excerpt, start, start + excerpt.length, Buffer.alloc(384 * 4)],
  )
  db.run("INSERT INTO chunks_fts(rowid, text, section, doc_name) VALUES (1, ?, '1', 'Exhibit A.eml')", [excerpt])
  db.close()

  const result = await searchDocument.execute({ query: "purchase price" }, { directory, sessionID: `provenance-${crypto.randomUUID()}` } as never)
  if (typeof result === "string") throw new Error(result)
  expect(result.output).toContain(`1. [Exhibit A.eml › 1] (score `)
  expect(result.output).toContain(`\n   provenance: in Exhibits.zip, attachment of Thread.eml\n${excerpt}`)
  const citation = result.metadata.citations[0]!
  expect(citation).toMatchObject({ documentName: "Exhibit A.eml", docPath, excerpt, charStart: start, charEnd: start + excerpt.length })

  // The legal plugin verifies the span against the live document unchanged.
  const output = { ...result, metadata: { ...result.metadata } }
  const hooks = await LegalPlugin({ directory } as Parameters<typeof LegalPlugin>[0])
  await hooks["tool.execute.after"]!({ tool: "search-document", sessionID: "provenance", callID: "search", args: {} }, output as never)
  expect(output.output).toBe(result.output)
  expect(output.metadata.citations[0]).toMatchObject({ verified: true, charStart: start, documentName: "Exhibit A.eml" })
})

test("grep-matter and read-document name the chain beside the plain name", async () => {
  const directory = matterDirectory()
  writeIndex(directory)
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  db.run("CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)")
  for (const [name, text] of [["Schedule.docx", "Completion occurs on the agreed date."], ["Notes.docx", "Completion checklist."]])
    db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, '1', ?, 0, ?)", [`/m/${name}`, name, text, text.length])
  db.close()

  const grepMatter = (await import("../tool/grep-matter")).default
  const grep = await grepMatter.execute({ pattern: "Completion" }, { directory } as never) as string
  expect(grep).toContain("Schedule.docx [1] @ 0: Completion occurs on the agreed date.")
  expect(grep).toContain("Notes.docx [1] @ 0: Completion checklist.")
  expect(grep).toEndWith("\n\nProvenance:\nSchedule.docx: attachment of Thread.eml")

  globalThis.fetch = (async () => Response.json({ text: "Completion occurs on the agreed date.", injection: null })) as unknown as typeof fetch
  const readDocument = (await import("../tool/read-document")).default
  const page = await readDocument.execute({ document: "Schedule.docx" }, { directory } as never)
  if (typeof page === "string") throw new Error(page)
  expect(page.output).toContain('<untrusted-document name="Schedule.docx">\nCompletion occurs on the agreed date.\n</untrusted-document>')
  expect(page.output).toContain('[document-page] {"document":"Schedule.docx","provenance":"attachment of Thread.eml","offset":0,')
  expect(page.metadata.document).toBe("Schedule.docx")
})
