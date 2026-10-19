import { Database } from "bun:sqlite"
import { afterEach, describe, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import define from "../tool/define"

const dirs: string[] = []

afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

function matter(input: {
  terms: Array<{ term: string; definition: string }>
  chunks: Array<{ section: string; text: string; start: number }>
}) {
  const dir = mkdtempSync(path.join(tmpdir(), "define-usage-"))
  dirs.push(dir)
  const dbPath = path.join(dir, ".dochaus", "legal.db")
  mkdirSync(path.dirname(dbPath), { recursive: true })
  const db = new Database(dbPath)
  db.run(
    "CREATE TABLE chunks (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, section TEXT, text TEXT, char_start INTEGER, char_end INTEGER)",
  )
  db.run(
    "CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, term TEXT, definition TEXT, char_start INTEGER, char_end INTEGER)",
  )
  for (const row of input.terms) {
    db.run("INSERT INTO defined_terms (doc_path, doc_name, term, definition, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [
      "/m/supply.docx",
      "Supply Agreement",
      row.term,
      row.definition,
      0,
      row.definition.length,
    ])
  }
  for (const row of input.chunks) {
    db.run("INSERT INTO chunks (doc_path, doc_name, section, text, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [
      "/m/supply.docx",
      "Supply Agreement",
      row.section,
      row.text,
      row.start,
      row.start + row.text.length,
    ])
  }
  db.close()
  return dir
}

async function lookup(dir: string, term: string) {
  return (await define.execute({ term }, { directory: dir } as never)) as string
}

function usageSites(output: string) {
  const at = output.indexOf("Usage sites")
  return at === -1 ? "" : output.slice(at)
}

describe("define usage sites", () => {
  test("lists only case-sensitive whole-word uses of the recorded term", async () => {
    const dir = matter({
      terms: [{ term: "Term", definition: '"Term" means the period stated in Section 2.' }],
      chunks: [
        { section: "3", text: "Either party may terminate this agreement on notice.", start: 100 },
        { section: "4", text: "The fee is determined by the board.", start: 200 },
        { section: "5", text: "The long-term plan is attached.", start: 300 },
        { section: "6", text: "During the Term, the supplier delivers monthly.", start: 400 },
      ],
    })

    const sites = usageSites(await lookup(dir, "Term"))

    expect(sites).toContain("During the Term, the supplier delivers monthly.")
    expect(sites).not.toContain("terminate")
    expect(sites).not.toContain("determined")
    expect(sites).not.toContain("long-term")
  })

  test("does not count lowercase prose as a use of a capitalized term", async () => {
    const dir = matter({
      terms: [{ term: "Party", definition: '"Party" means each signatory.' }],
      chunks: [
        { section: "7", text: "No third party may rely on this clause.", start: 100 },
        { section: "8", text: "Each Party's notice address is in Schedule 1.", start: 200 },
      ],
    })

    const sites = usageSites(await lookup(dir, "party"))

    expect(sites).not.toContain("third party")
    // An apostrophe is not a word character, so the possessive still counts.
    expect(sites).toContain("Each Party's notice address")
  })

  test("finds a real use after more prefilter matches than the usage limit", async () => {
    const dir = matter({
      terms: [{ term: "Term", definition: '"Term" means the period stated in Section 2.' }],
      chunks: [
        ...Array.from({ length: 6 }, (_, i) => ({ section: `1.${i}`, text: `Clause ${i} may terminate early.`, start: i * 10 })),
        { section: "9", text: "Renewal extends the Term by one year.", start: 500 },
      ],
    })

    const sites = usageSites(await lookup(dir, "Term"))

    expect(sites).toContain("Renewal extends the Term by one year.")
    expect(sites).not.toContain("terminate")
  })

  test("centers the snippet on the real use, not an earlier non-use", async () => {
    const dir = matter({
      terms: [{ term: "Term", definition: '"Term" means the period stated in Section 2.' }],
      chunks: [{ section: "2", text: "Either side may terminate early.\nDuring the Term, rent is fixed.", start: 0 }],
    })

    const sites = usageSites(await lookup(dir, "Term"))

    expect(sites).toContain("During the Term, rent is fixed.")
    expect(sites).not.toContain("terminate")
  })
})
