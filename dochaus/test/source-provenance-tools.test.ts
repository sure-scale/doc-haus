// The structure tools (parties, defined terms, key dates, amendment chain,
// current section) show a child source's container chain after their output,
// once per document, leaving every line's document name plain. Plain matters'
// output has no footer. Run from dochaus/: bun test test/source-provenance-tools.test.ts
//
// Every fixture below is invented test content.
import { Database } from "bun:sqlite"
import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { provenanceFooter, provenanceMapLine } from "../lib/source-provenance"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true })))

const CHAIN = "attachment of Thread.eml"

function matter(children: string[]) {
  const directory = mkdtempSync(path.join(tmpdir(), "provenance-tools-"))
  directories.push(directory)
  writeFileSync(path.join(directory, "matter.json"), JSON.stringify({ id: "provenance-tools" }))
  writeFileSync(path.join(directory, "source-index.json"), JSON.stringify({
    sources: [{ document: "Thread.eml" }, { document: "Plain.docx" }, ...children.map((document) => ({ document, attachment_of: "Thread.eml" }))],
  }))
  mkdirSync(path.join(directory, ".dochaus"))
  const db = new Database(path.join(directory, ".dochaus", "legal.db"))
  for (const sql of [
    "CREATE TABLE documents (id INTEGER PRIMARY KEY, doc_path TEXT UNIQUE, name TEXT, created_at INTEGER)",
    "CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, term TEXT, definition TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE parties (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, name TEXT, role TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE doc_relations (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, relation TEXT, target_name TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE section_ops (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section_label TEXT, operation TEXT, char_start INTEGER, char_end INTEGER)",
    "CREATE TABLE key_dates (id INTEGER PRIMARY KEY, doc_name TEXT, kind TEXT, date_iso TEXT, date_formula TEXT, notice_days INTEGER, cancel_by_iso TEXT, party TEXT, quote TEXT, reason TEXT, author TEXT, created_at INTEGER)",
  ]) db.run(sql)
  return { directory, db }
}

const ctx = (directory: string) => ({ directory, sessionID: "provenance-tools" }) as never
const text = (result: unknown) => (typeof result === "string" ? result : (result as { output: string }).output)

test("footer and map line are empty unless a named source is a child", () => {
  const provenance = new Map([["Schedule.docx", CHAIN]])
  expect(provenanceFooter(["Plain.docx"], provenance)).toBe("")
  expect(provenanceMapLine(["Plain.docx"], provenance)).toBe("")
  expect(provenanceFooter(["Schedule.docx", "Plain.docx", "Schedule.docx"], provenance)).toBe(`\n\nProvenance:\nSchedule.docx: ${CHAIN}`)
  expect(JSON.parse(provenanceMapLine(["Schedule.docx", "Plain.docx"], provenance).split("): ")[1]!)).toEqual({ "Schedule.docx": CHAIN })
})

test("matter-parties, define and key-dates add one provenance block for child documents only", async () => {
  const { directory, db } = matter(["Schedule.docx"])
  for (const name of ["Schedule.docx", "Plain.docx"]) {
    db.run("INSERT INTO parties (doc_path, doc_name, name, role, char_start, char_end) VALUES (?, ?, 'Acme Ltd', 'Seller', 0, 8)", [`/m/${name}`, name])
    db.run("INSERT INTO defined_terms (doc_path, doc_name, term, definition, char_start, char_end) VALUES (?, ?, 'Price', 'The Price means 10.', 0, 20)", [`/m/${name}`, name])
    db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, '2', 'The Price is due.', 0, 17)", [`/m/${name}`, name])
    db.run("INSERT INTO key_dates (doc_name, kind, date_iso, quote, reason, author, created_at) VALUES (?, 'expiry', '2027-01-01', 'q', 'r', 'a', 1)", [name])
  }
  db.close()
  const footer = `\n\nProvenance:\nSchedule.docx: ${CHAIN}`

  const parties = text(await (await import("../tool/matter-parties")).default.execute({}, ctx(directory)))
  expect(parties).toContain("Schedule.docx\n- Acme Ltd (Seller)")
  expect(parties.match(/Provenance:/g)).toHaveLength(1)
  expect(parties).toContain(footer)

  const define = text(await (await import("../tool/define")).default.execute({ term: "Price" }, ctx(directory)))
  expect(define).toContain("## Schedule.docx\n")
  expect(define.match(/Provenance:/g)).toHaveLength(1)
  expect(define).toContain(footer)

  const dates = text(await (await import("../tool/key-dates")).default.execute({}, ctx(directory)))
  expect(dates).toContain("| expiry | Schedule.docx")
  expect(dates).toEndWith(footer)

  // A plain document alone: output carries no provenance at all.
  expect(text(await (await import("../tool/matter-parties")).default.execute({ document: "Plain.docx" }, ctx(directory)))).not.toContain("Provenance")
  expect(text(await (await import("../tool/define")).default.execute({ term: "Price", document: "Plain.docx" }, ctx(directory)))).not.toContain("Provenance")
  expect(text(await (await import("../tool/key-dates")).default.execute({ document: "Plain.docx" }, ctx(directory)))).not.toContain("Provenance")
})

test("amendment-chain and current-section name the amender's chain beside its plain name", async () => {
  const { directory, db } = matter(["Amendment No. 1"])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES ('/m/base.docx', 'Original Agreement', 1)")
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES ('/m/am1.docx', 'Amendment No. 1', 2)")
  db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES ('/m/am1.docx', 'Amendment No. 1', '', 'Amendment No. 1 to the Original Agreement dated 2026-01-01. Section 4 is restated as follows.', 0, 88)")
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES ('/m/am1.docx', 'Amendment No. 1', 'amends', 'Original Agreement', 0, 59)")
  db.run("INSERT INTO section_ops (doc_path, doc_name, section_label, operation, char_start, char_end) VALUES ('/m/am1.docx', 'Amendment No. 1', '4', 'restate', 60, 88)")
  db.close()
  const footer = `\n\nProvenance:\nAmendment No. 1: ${CHAIN}`

  const chain = text(await (await import("../tool/amendment-chain")).default.execute({}, ctx(directory)))
  expect(chain).toContain('- Amendment No. 1 amends "Original Agreement"')
  expect(chain).toEndWith(footer)
  expect(chain.match(/Provenance:/g)).toHaveLength(1)

  const current = text(await (await import("../tool/current-section")).default.execute({ section: "4" }, ctx(directory)))
  expect(current).toContain("Latest operative source: Amendment No. 1")
  expect(current).toEndWith(footer)
  expect(current.match(/Provenance:/g)).toHaveLength(1)
})

test("amendment-chain on a matter whose documents are all plain has no provenance block", async () => {
  const { directory, db } = matter([])
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES ('/m/base.docx', 'Original Agreement', 1)")
  db.run("INSERT INTO documents (doc_path, name, created_at) VALUES ('/m/am1.docx', 'Plain.docx', 2)")
  db.run("INSERT INTO doc_relations (doc_path, doc_name, relation, target_name, char_start, char_end) VALUES ('/m/am1.docx', 'Plain.docx', 'amends', 'Original Agreement', 0, 5)")
  db.close()
  const chain = text(await (await import("../tool/amendment-chain")).default.execute({}, ctx(directory)))
  expect(chain).toContain("Plain.docx amends")
  expect(chain).not.toContain("Provenance")
})
