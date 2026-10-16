import { Database } from "bun:sqlite"
import { existsSync, mkdirSync } from "node:fs"
import path from "node:path"

// Key-dates register for the record-key-date and key-dates tools: typed
// renewal/termination/expiry/payment rows that reviews emit as a side-effect,
// so deadlines survive the chat session and anything downstream can read them
// deterministically. A row carries either a resolved calendar date (date_iso)
// or the document's verbatim relative formula (date_formula) when no calendar
// date is resolvable — never a guessed one.
//
// The register lives in the matter's index DB (.dochaus/legal.db) next to the
// redactions and redlines tables; only the dochaus tools read and write it, so
// the DDL lives here alone.

export const KEY_DATE_KINDS = ["renewal", "termination_notice", "expiry", "payment", "deliverable", "other"] as const
export type KeyDateKind = (typeof KEY_DATE_KINDS)[number]

export function recordKeyDate(
  matterDir: string,
  row: {
    docName: string
    kind: KeyDateKind
    dateIso?: string
    dateFormula?: string
    noticeDays?: number
    cancelByIso?: string
    party?: string
    quote: string
    reason: string
    author: string
  },
) {
  const dir = path.join(matterDir, ".dochaus")
  mkdirSync(dir, { recursive: true })
  const db = new Database(path.join(dir, "legal.db"))
  db.run("PRAGMA journal_mode = WAL")
  db.run(`
    CREATE TABLE IF NOT EXISTS key_dates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_name TEXT NOT NULL,
      kind TEXT NOT NULL,
      date_iso TEXT,
      date_formula TEXT,
      notice_days INTEGER,
      cancel_by_iso TEXT,
      party TEXT,
      quote TEXT NOT NULL,
      reason TEXT NOT NULL,
      author TEXT NOT NULL,
      created_at INTEGER NOT NULL
    )
  `)
  const result = db.run(
    "INSERT INTO key_dates (doc_name, kind, date_iso, date_formula, notice_days, cancel_by_iso, party, quote, reason, author, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [
      row.docName,
      row.kind,
      row.dateIso ?? null,
      row.dateFormula ?? null,
      row.noticeDays ?? null,
      row.cancelByIso ?? null,
      row.party ?? null,
      row.quote,
      row.reason,
      row.author,
      Date.now(),
    ],
  )
  db.close()
  return Number(result.lastInsertRowid)
}

export type KeyDateRow = {
  id: number
  doc_name: string
  kind: KeyDateKind
  date_iso: string | null
  date_formula: string | null
  notice_days: number | null
  cancel_by_iso: string | null
  party: string | null
  quote: string
  reason: string
  author: string
  created_at: number
}

export function listKeyDates(matterDir: string, docName?: string): KeyDateRow[] {
  const dbFile = path.join(matterDir, ".dochaus", "legal.db")
  if (!existsSync(dbFile)) return []
  const db = new Database(dbFile, { readonly: true })
  const hasTable = db.query("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'key_dates'").get()
  if (!hasTable) {
    db.close()
    return []
  }
  // Soonest calendar date first; formula-only rows (date_iso NULL) sort last so
  // the next real deadline always leads the register.
  const order = "ORDER BY date_iso IS NULL, date_iso ASC, id ASC"
  const rows = docName
    ? db.query(`SELECT * FROM key_dates WHERE doc_name = ? ${order}`).all(docName)
    : db.query(`SELECT * FROM key_dates ${order}`).all()
  db.close()
  return rows as KeyDateRow[]
}
