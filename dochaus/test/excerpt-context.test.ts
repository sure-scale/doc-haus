import { Database } from "bun:sqlite"
import { describe, expect, test } from "bun:test"
import type { ChunkRow } from "../lib/search/channels"
import { attachDefinitions } from "../lib/search/excerpt-context"

function excerpt(text: string): { row: ChunkRow } {
  return {
    row: {
      id: 1,
      doc_name: "Supply Agreement",
      doc_path: "/m/supply.docx",
      section: "4",
      structural_path: "",
      party: "",
      text,
      char_start: 0,
      char_end: text.length,
      flagged: 0,
    },
  }
}

function termsDb(terms: string[]) {
  const db = new Database(":memory:")
  db.run(
    "CREATE TABLE defined_terms (id INTEGER PRIMARY KEY, doc_path TEXT, doc_name TEXT, term TEXT, definition TEXT, char_start INTEGER, char_end INTEGER)",
  )
  for (const term of terms) {
    db.run("INSERT INTO defined_terms (doc_path, doc_name, term, definition, char_start, char_end) VALUES (?, ?, ?, ?, ?, ?)", [
      "/m/supply.docx",
      "Supply Agreement",
      term,
      `"${term}" has the meaning given in Section 1.`,
      0,
      10,
    ])
  }
  return db
}

describe("attachDefinitions term matching", () => {
  test("finds a whole-word use that follows an earlier longer word", () => {
    const db = termsDb(["Agreement"])

    const note = attachDefinitions(db, [excerpt("The Agreements in Schedule 2 remain. This Agreement governs.")], {})

    expect(note).toContain('"Agreement" (Supply Agreement)')
  })

  test("still ignores a passage whose only occurrences are inside longer words", () => {
    const db = termsDb(["Agreement"])

    const note = attachDefinitions(db, [excerpt("The Agreements in Schedule 2 remain. Prior Agreements lapse.")], {})

    expect(note).toBe("")
  })
})
