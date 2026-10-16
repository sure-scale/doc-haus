import { Database } from "bun:sqlite"
import { existsSync, mkdirSync, realpathSync } from "node:fs"
import path from "node:path"
import {
  canonicalArtifactIdentity,
  withArtifactMutationLockAsync,
} from "../capabilities/artifacts/mutation-lock"

// Record and read redline proposals in the matter's index DB. The redline tools
// propose changes here instead of writing tracked changes into the .docx; the
// ingest service reads these rows to render the redlined view and to accept
// (bake) or reject them. The canonical .docx stays clean until a reviewer accepts.
//
// Ingest owns this database and creates the same `redlines` table in its openDb;
// keep this DDL in sync with services/ingest/src/db.ts. We create-if-not-exists
// here too so a tool can propose before the document is ever (re-)ingested.

export function recordRedline(
  matterDir: string,
  row: {
    docPath: string
    docName: string
    scope: "phrase" | "clause"
    findText: string
    oldText: string
    newText: string
    author: string
    reviewComment: string
    anchorId: string
  },
) {
  const normalized = { ...row, docPath: canonicalRedlinePath(matterDir, row.docPath) }
  const db = openRedlineDb(matterDir)
  try {
    return db.transaction(() => insertRedline(db, normalized)).immediate()
  } finally {
    db.close()
  }
}

// A pending proposal, as the propose-time conflict check sees it. anchor_id is the
// Docxodus block token the proposal targets; the canonical .docx stays clean while
// proposals are pending, so two proposals carrying the same anchor_id address the
// same paragraph and the replay would collide.
export type PendingRedline = {
  id: number
  scope: "phrase" | "clause"
  find_text: string
  new_text: string
  author: string
  review_comment: string
  anchor_id: string | null
}

export function pendingRedlinesForDoc(matterDir: string, docPath: string): PendingRedline[] {
  const dbFile = path.join(matterDir, ".dochaus", "legal.db")
  if (!existsSync(dbFile)) return []
  const db = new Database(dbFile, { readonly: true })
  const rows = pendingRedlines(db, matterDir, docPath)
  db.close()
  return rows
}

export function withRedlineDocumentLock<T>(matterDir: string, docPath: string, fn: () => Promise<T>) {
  return withArtifactMutationLockAsync(matterDir, docPath, fn)
}

// Conflict discovery, insertion, and supersession are one immediate SQLite
// transaction. The caller holds the document artifact lock while deriving the
// anchor and invoking this function, so the proposal is bound to the same bytes
// against which its conflict set was computed.
export function recordRedlineProposal(
  matterDir: string,
  row: {
    docPath: string
    docName: string
    scope: "phrase" | "clause"
    findText: string
    oldText: string
    newText: string
    author: string
    reviewComment: string
    anchorId: string
  },
) {
  const normalized = { ...row, docPath: canonicalRedlinePath(matterDir, row.docPath) }
  const db = openRedlineDb(matterDir)
  try {
    return db.transaction(() => {
      const superseded = conflictingRedlines(pendingRedlines(db, matterDir, normalized.docPath), normalized)
      const id = insertRedline(db, normalized)
      supersedeRedlineRows(db, superseded.map((candidate) => candidate.id))
      return { id, superseded }
    }).immediate()
  } finally {
    db.close()
  }
}

export function supersedeMatchingRedlines(
  matterDir: string,
  docPath: string,
  matches: (row: PendingRedline) => boolean,
) {
  const db = openRedlineDb(matterDir)
  try {
    return db.transaction(() => {
      const superseded = pendingRedlines(db, matterDir, docPath).filter(matches)
      supersedeRedlineRows(db, superseded.map((row) => row.id))
      return superseded
    }).immediate()
  } finally {
    db.close()
  }
}

// Pending proposals that collide with the one about to be recorded. Two proposals
// collide when they target the same paragraph (same anchor_id) AND either rewrites
// the whole clause — a clause rewrite replaces the entire paragraph, so it voids
// every other edit on it — or their find texts overlap (one contains the other),
// so the replay of one erases the text the other anchors to. Independent phrase
// edits in the same paragraph do not collide and coexist.
export function conflictingRedlines(
  pending: PendingRedline[],
  next: { anchorId: string; scope: "phrase" | "clause"; findText: string },
) {
  return pending.filter((p) => {
    if (p.anchor_id !== next.anchorId) return false
    if (p.scope === "clause" || next.scope === "clause") return true
    const existing = p.find_text.trim()
    const incoming = next.findText.trim()
    return existing.includes(incoming) || incoming.includes(existing)
  })
}

// Retire superseded proposals so only the newest edit on a paragraph stays pending.
// A separate 'superseded' status (not 'rejected') records that the reviewer never
// declined the edit — a later proposal replaced it — without surfacing it in the
// pending queue or the redlined view.
export function supersedeRedlines(matterDir: string, ids: number[]) {
  if (!ids.length) return
  const db = openRedlineDb(matterDir)
  try {
    db.transaction(() => supersedeRedlineRows(db, ids)).immediate()
  } finally {
    db.close()
  }
}

function openRedlineDb(matterDir: string) {
  const dir = path.join(matterDir, ".dochaus")
  mkdirSync(dir, { recursive: true })
  const db = new Database(path.join(dir, "legal.db"))
  db.run("PRAGMA journal_mode = WAL")
  db.run(`
    CREATE TABLE IF NOT EXISTS redlines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      doc_path TEXT NOT NULL,
      doc_name TEXT NOT NULL,
      scope TEXT NOT NULL,
      find_text TEXT NOT NULL,
      old_text TEXT NOT NULL,
      new_text TEXT NOT NULL,
      author TEXT NOT NULL,
      review_comment TEXT NOT NULL DEFAULT '',
      anchor_id TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      created_at INTEGER NOT NULL
    )
  `)
  if (!hasColumn(db, "redlines", "review_comment"))
    db.run("ALTER TABLE redlines ADD COLUMN review_comment TEXT NOT NULL DEFAULT ''")
  return db
}

function pendingRedlines(db: Database, matterDir: string, docPath: string) {
  const identity = canonicalArtifactIdentity(matterDir, docPath)
  return (db
    .query(
      "SELECT id, doc_path, scope, find_text, new_text, author, review_comment, anchor_id FROM redlines WHERE status = 'pending' ORDER BY created_at, id",
    )
    .all() as Array<PendingRedline & { doc_path: string }>).filter(
      (row) => canonicalArtifactIdentity(matterDir, row.doc_path) === identity,
    )
}

function insertRedline(
  db: Database,
  row: {
    docPath: string
    docName: string
    scope: "phrase" | "clause"
    findText: string
    oldText: string
    newText: string
    author: string
    reviewComment: string
    anchorId: string
  },
) {
  return Number(db.run(
    "INSERT INTO redlines (doc_path, doc_name, scope, find_text, old_text, new_text, author, review_comment, anchor_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
    [row.docPath, row.docName, row.scope, row.findText, row.oldText, row.newText, row.author, row.reviewComment, row.anchorId, Date.now()],
  ).lastInsertRowid)
}

function hasColumn(db: Database, table: string, column: string) {
  return (db.query(`PRAGMA table_info(${table})`).all() as { name: string }[]).some((item) => item.name === column)
}

function supersedeRedlineRows(db: Database, ids: number[]) {
  ids.forEach((id) => db.run("UPDATE redlines SET status = 'superseded' WHERE id = ? AND status = 'pending'", [id]))
}

function canonicalRedlinePath(matterDir: string, docPath: string) {
  const root = realpathSync.native(matterDir)
  const document = realpathSync.native(path.resolve(matterDir, docPath))
  const relative = path.relative(root, document)
  if (!relative || relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative))
    throw new Error("A redline proposal must target a document inside its matter.")
  return document
}
