import { expect, test } from "bun:test"
import { Database } from "bun:sqlite"
import { mkdtempSync, mkdirSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { canonicalArtifactIdentity } from "../capabilities/artifacts/mutation-lock"
import { assertMatterWordEditsAllowed, assertWordEditAllowed, beginWordNativeOperation, claimWordEditLease, clearWordEditLease, createWordEditLeaseTable, EditingLeaseError, readWordEditLease, readWordNativeOperation, releaseWordEditLease, renewWordEditLease, requireWordEditLease, settleWordNativeOperation, wordLeaseState, wordProposalApplied } from "./word-edit-lease"

function fixture(legacy = false) {
  const db = new Database(":memory:")
  db.run("CREATE TABLE versioned_documents (id TEXT PRIMARY KEY, removed_at INTEGER)")
  db.run("CREATE TABLE word_grants (token_hash TEXT PRIMARY KEY, document_id TEXT, revoked_at INTEGER)")
  db.run("INSERT INTO versioned_documents VALUES ('document', NULL)")
  db.run("INSERT INTO word_grants VALUES ('grant', 'document', NULL)")
  if (legacy) db.run("CREATE TABLE word_edit_leases (document_id TEXT PRIMARY KEY, token_hash TEXT NOT NULL, session TEXT NOT NULL, generation TEXT NOT NULL, expires_at INTEGER NOT NULL)")
  createWordEditLeaseTable(db)
  const identity = { document: "document", tokenHash: "grant", session: "pane" }
  return { db, identity, owner: claimWordEditLease(db, identity, 1000) }
}

test("migration preserves an existing lease's identity and expiry", () => {
  const db = new Database(":memory:")
  try {
    db.run("CREATE TABLE word_edit_leases (document_id TEXT PRIMARY KEY, token_hash TEXT NOT NULL, session TEXT NOT NULL, generation TEXT NOT NULL, expires_at INTEGER NOT NULL)")
    db.run("INSERT INTO word_edit_leases VALUES ('document', 'grant', 'pane', 'existing-generation', 1234)")
    createWordEditLeaseTable(db); createWordEditLeaseTable(db)
    expect(readWordEditLease(db, "document")).toMatchObject({ document_id: "document", token_hash: "grant", session: "pane", generation: "existing-generation", expires_at: 1234 })
    expect(readWordNativeOperation(db, "document")).toEqual({ sequence: 0, operationId: null, state: "idle" })
  } finally { db.close() }
})

test("legacy leases migrate idempotently and retain sequence across logical clear and reclaim", () => {
  const { db, owner, identity } = fixture(true)
  try {
    createWordEditLeaseTable(db)
    expect(readWordNativeOperation(db, owner.document)).toEqual({ sequence: 0, operationId: null, state: "idle" })
    expect(beginWordNativeOperation(db, owner, 0, "op-1", 1100)).toEqual({ sequence: 1, operationId: "op-1", state: "pending" })
    expect(beginWordNativeOperation(db, owner, 0, "op-1", 1100)).toEqual({ sequence: 1, operationId: "op-1", state: "pending" })
    expect(settleWordNativeOperation(db, owner, 1, "op-1", "confirmed").state).toBe("confirmed")
    expect(settleWordNativeOperation(db, owner, 1, "op-1", "confirmed").state).toBe("confirmed")
    expect(() => settleWordNativeOperation(db, owner, 1, "op-1", "not-dispatched")).toThrow(EditingLeaseError)
    clearWordEditLease(db, owner.document)
    expect(readWordEditLease(db, owner.document)).toBeNull()
    expect(readWordNativeOperation(db, owner.document).sequence).toBe(1)
    const next = claimWordEditLease(db, { ...identity, session: "next-pane" }, 1200)
    expect(next.generation).not.toBe(owner.generation)
    expect(() => beginWordNativeOperation(db, next, 0, "op-1", 1200)).toThrow(EditingLeaseError)
    expect(beginWordNativeOperation(db, next, 1, "op-2", 1200)).toEqual({ sequence: 2, operationId: "op-2", state: "pending" })
    expect(() => settleWordNativeOperation(db, owner, 1, "op-1", "confirmed")).toThrow(EditingLeaseError)
    expect(readWordNativeOperation(db, owner.document).state).toBe("pending")
  } finally { db.close() }
})

test("unresolved native work survives expiry and revocation and never yields takeover", () => {
  const { db, owner, identity } = fixture()
  try {
    beginWordNativeOperation(db, owner, 0, "op-1", 1100)
    expect(wordLeaseState(db, owner.document, identity, 1200).state).toBe("owned")
    expect(() => requireWordEditLease(db, owner, 1200)).not.toThrow()
    expect(() => renewWordEditLease(db, owner, 1200)).not.toThrow()
    expect(() => renewWordEditLease(db, owner, 500000)).not.toThrow()
    expect(() => renewWordEditLease(db, { ...owner, session: "other" }, 500000)).toThrow(EditingLeaseError)
    db.run("UPDATE word_grants SET revoked_at = 1 WHERE token_hash='grant'")
    expect(() => renewWordEditLease(db, owner, 500001)).toThrow(EditingLeaseError)
    expect(wordLeaseState(db, owner.document, identity, 500000)).toMatchObject({ state: "other", expiresAt: null, generation: owner.generation })
    expect(() => claimWordEditLease(db, { ...identity, session: "second" }, 500000)).toThrow(EditingLeaseError)
    expect(() => releaseWordEditLease(db, owner)).toThrow(EditingLeaseError)
    expect(() => clearWordEditLease(db, owner.document)).toThrow(EditingLeaseError)
    expect(() => beginWordNativeOperation(db, owner, 0, "op-1", 500000)).toThrow(EditingLeaseError)
    expect(readWordNativeOperation(db, owner.document).state).toBe("pending")
    // A trusted original callback may settle after expiry; revocation and
    // authentication of the callback are enforced by the calling route.
    expect(settleWordNativeOperation(db, owner, 1, "op-1", "not-dispatched").state).toBe("not-dispatched")
  } finally { db.close() }
})

test("operation identities and stale owners cannot settle or replay a newer operation", () => {
  const { db, owner } = fixture()
  try {
    expect(() => beginWordNativeOperation(db, owner, -1, "bad", 1100)).toThrow(EditingLeaseError)
    expect(() => beginWordNativeOperation(db, owner, 0, "bad id", 1100)).toThrow(EditingLeaseError)
    beginWordNativeOperation(db, owner, 0, "one", 1100)
    expect(() => beginWordNativeOperation(db, owner, 0, "different", 1100)).toThrow(EditingLeaseError)
    expect(() => settleWordNativeOperation(db, { ...owner, session: "impostor" }, 1, "one", "confirmed")).toThrow(EditingLeaseError)
    settleWordNativeOperation(db, owner, 1, "one", "not-dispatched")
    beginWordNativeOperation(db, owner, 1, "two", 1100)
    expect(() => settleWordNativeOperation(db, owner, 1, "one", "not-dispatched")).toThrow(EditingLeaseError)
    expect(() => beginWordNativeOperation(db, owner, 0, "one", 1100)).toThrow(EditingLeaseError)
    expect(readWordNativeOperation(db, owner.document)).toEqual({ sequence: 2, operationId: "two", state: "pending" })
  } finally { db.close() }
})


test("unresolved operation fences matter and file writes after expiry and grant revocation", () => {
  const dir = mkdtempSync(path.join(tmpdir(), "word-native-lease-")), file = path.join(dir, "Draft.docx")
  mkdirSync(path.join(dir, ".dochaus")); const db = new Database(path.join(dir, ".dochaus", "legal.db"))
  try {
    db.run("CREATE TABLE versioned_documents (id TEXT PRIMARY KEY, name_key TEXT, removed_at INTEGER)")
    db.run("CREATE TABLE word_grants (token_hash TEXT PRIMARY KEY, document_id TEXT, revoked_at INTEGER)")
    db.run("INSERT INTO versioned_documents VALUES (?, ?, NULL)", ["document", path.basename(canonicalArtifactIdentity(dir, file))])
    db.run("INSERT INTO word_grants VALUES ('grant', 'document', NULL)"); createWordEditLeaseTable(db)
    const owner = claimWordEditLease(db, { document: "document", tokenHash: "grant", session: "pane" }, 1000)
    beginWordNativeOperation(db, owner, 0, "pending", 1100)
    db.run("UPDATE word_grants SET revoked_at = 1 WHERE token_hash = 'grant'")
    expect(() => assertWordEditAllowed(dir, file, undefined, 500000)).toThrow(EditingLeaseError)
    expect(() => assertWordEditAllowed(dir, file, owner, 500000)).toThrow(EditingLeaseError)
    expect(() => assertMatterWordEditsAllowed(dir, 500000)).toThrow(EditingLeaseError)
    settleWordNativeOperation(db, owner, 1, "pending", "not-dispatched")
    expect(() => assertWordEditAllowed(dir, file, undefined, 500000)).not.toThrow()
    expect(() => assertMatterWordEditsAllowed(dir, 500000)).not.toThrow()
  } finally { db.close(); rmSync(dir, { recursive: true, force: true }) }
})


test("read-only guards accept an old lease schema without mutating it", () => {
  const dir = mkdtempSync(path.join(tmpdir(), "word-legacy-lease-")), file = path.join(dir, "Draft.docx")
  mkdirSync(path.join(dir, ".dochaus")); const db = new Database(path.join(dir, ".dochaus", "legal.db"))
  try {
    db.run("CREATE TABLE versioned_documents (id TEXT PRIMARY KEY, name_key TEXT, removed_at INTEGER)")
    db.run("CREATE TABLE word_grants (token_hash TEXT PRIMARY KEY, document_id TEXT, revoked_at INTEGER)")
    db.run("CREATE TABLE word_edit_leases (document_id TEXT PRIMARY KEY, token_hash TEXT NOT NULL, session TEXT NOT NULL, generation TEXT NOT NULL, expires_at INTEGER NOT NULL)")
    db.run("INSERT INTO versioned_documents VALUES (?, ?, NULL)", ["document", path.basename(canonicalArtifactIdentity(dir, file))])
    db.run("INSERT INTO word_grants VALUES ('grant', 'document', NULL)")
    db.run("INSERT INTO word_edit_leases VALUES ('document', 'grant', 'pane', 'gen', 2000)")
    expect(() => assertWordEditAllowed(dir, file, undefined, 1000)).toThrow(EditingLeaseError)
    expect(() => assertMatterWordEditsAllowed(dir, 1000)).toThrow(EditingLeaseError)
    expect(() => assertWordEditAllowed(dir, file, undefined, 3000)).not.toThrow()
    expect((db.query("PRAGMA table_info(word_edit_leases)").all() as unknown[])).toHaveLength(5)
  } finally { db.close(); rmSync(dir, { recursive: true, force: true }) }
})


test("confirmed proposal survives later operations and no-dispatch remains retryable", () => {
  const { db, owner } = fixture()
  const proposal = { conversationId: "conversation", turnId: "turn" }
  try {
    db.transaction(() => beginWordNativeOperation(db, owner, 0, "first", 1001, proposal)).immediate()
    db.transaction(() => settleWordNativeOperation(db, owner, 1, "first", "not-dispatched")).immediate()
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(false)
    db.transaction(() => beginWordNativeOperation(db, owner, 1, "second", 1002, proposal)).immediate()
    db.transaction(() => settleWordNativeOperation(db, owner, 2, "second", "confirmed")).immediate()
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(true)
    expect(wordProposalApplied(db, owner.document, { ...proposal, conversationId: "other" })).toBe(false)
    expect(() => db.transaction(() => beginWordNativeOperation(db, owner, 2, "third", 1003, proposal)).immediate()).toThrow("already applied")
    db.transaction(() => beginWordNativeOperation(db, owner, 2, "refresh", 1003)).immediate()
    db.transaction(() => settleWordNativeOperation(db, owner, 3, "refresh", "confirmed")).immediate()
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(true)
  } finally { db.close() }
})

test("proposal settlement and fence roll back together and admission binds exact turn", () => {
  const { db, owner } = fixture()
  const proposal = { conversationId: "conversation", turnId: "turn" }
  try {
    db.transaction(() => beginWordNativeOperation(db, owner, 0, "op", 1001, proposal)).immediate()
    expect(() => db.transaction(() => beginWordNativeOperation(db, owner, 0, "op", 1002, { ...proposal, turnId: "other" })).immediate()).toThrow()
    expect(() => db.transaction(() => { settleWordNativeOperation(db, owner, 1, "op", "confirmed"); throw new Error("rollback") }).immediate()).toThrow("rollback")
    expect(readWordNativeOperation(db, owner.document).state).toBe("pending")
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(false)
    db.transaction(() => settleWordNativeOperation(db, owner, 1, "op", "confirmed")).immediate()
    db.transaction(() => settleWordNativeOperation(db, owner, 1, "op", "confirmed")).immediate()
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(true)
  } finally { db.close() }
})

test("a failed outcome insert rolls back direct settlement and exact retry succeeds", () => {
  const { db, owner } = fixture()
  const proposal = { conversationId: "conversation", turnId: "turn" }
  try {
    beginWordNativeOperation(db, owner, 0, "operation", 1001, proposal)
    db.run("CREATE TRIGGER reject_outcome BEFORE INSERT ON word_chat_applied BEGIN SELECT RAISE(ABORT, 'outcome unavailable'); END")
    expect(() => settleWordNativeOperation(db, owner, 1, "operation", "confirmed")).toThrow("outcome unavailable")
    expect(readWordNativeOperation(db, owner.document).state).toBe("pending")
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(false)
    db.run("DROP TRIGGER reject_outcome")
    expect(settleWordNativeOperation(db, owner, 1, "operation", "confirmed").state).toBe("confirmed")
    expect(wordProposalApplied(db, owner.document, proposal)).toBe(true)
  } finally { db.close() }
})
