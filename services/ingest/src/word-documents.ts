import { Hono, type Context } from "hono"
import { createHash, randomBytes } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import { linkedWordDocx, MAX_WORD_UPLOAD_BYTES, stripWordLink, wordLastModifiedBy, WordPackageError } from "../../../dochaus/lib/word-linked-docx"
import type { WordOpenDocument, WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { getDocumentRole, openDb } from "./db"
import { readDocumentTarget } from "./registered-document"
import { readCurrentVersion, readVersion, sha256Hex, type CapturedVersion, type VersionWrite } from "./document-versions"
import { activeWordSessions, type WordGrant } from "./word-grants"
import { attachmentDocumentDisposition } from "./content-disposition"
import { DeliverableConflictError } from "./deliverable"
import { SourceDocumentConflictError } from "./ingest"
import { wordHttpsConfiguration } from "./word-configuration"
import { wordHttpsFailure } from "./word-https"

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
export const wordTokenPath = (pathname: string) => /^\/word\/(saves|heartbeat|done)$/.test(pathname)

export function wordConfiguration() {
  const failure = wordHttpsFailure()
  if (failure) return { available: false, origin: null, manifestUrl: null, reason: failure }
  let config: ReturnType<typeof wordHttpsConfiguration>
  try { config = wordHttpsConfiguration() } catch {
    return { available: false, origin: null, manifestUrl: null, reason: "The Word add-in settings are invalid. Run Word setup again." }
  }
  const origin = config?.origin ?? null
  return {
    available: Boolean(config), origin, manifestUrl: origin ? `${origin}/word-addin/manifest.xml` : null,
    ...(!config ? { reason: "Configure this instance's trusted HTTPS Word origin and certificate before opening a document in Word." } : {}),
  }
}

type Matter = { id: string; title: string; dir: string }
type Dependencies = {
  matter: (id: string) => Matter
  matters: () => Matter[]
  locks: (directory: string, file: string) => string[]
  write: (directory: string, name: string, bytes: Buffer, prior: string, version: VersionWrite) => Promise<{ version: CapturedVersion }>
}
const digest = (token: string) => createHash("sha256").update(token).digest("hex")
const failure = (c: Context, status: 400 | 401 | 403 | 404 | 409 | 410 | 413 | 503, code: string, error: string) => c.json({ error, code }, status)
const sessionId = (c: Context) => {
  const id = c.req.header("x-dochaus-word-session")
  return id && /^[a-zA-Z0-9-]{8,128}$/.test(id) ? id : null
}

/** Exact pane-origin CORS applies before global localhost CORS, including errors. */
export async function wordCors(c: Context, next: () => Promise<void>) {
  const origin = wordConfiguration().origin
  if (!origin || c.req.header("Origin") !== origin)
    return failure(c, 403, "word_origin_refused", "This Word request is not from this instance's configured pane.")
  c.header("Access-Control-Allow-Origin", origin)
  c.header("Vary", "Origin")
  if (c.req.method === "OPTIONS") {
    c.header("Access-Control-Allow-Methods", "POST, OPTIONS")
    c.header("Access-Control-Allow-Headers", "Authorization, Content-Type, X-DocHaus-Word-Base-Version, X-DocHaus-Word-Session")
    return c.body(null, 204)
  }
  if (c.req.method !== "POST") return failure(c, 403, "word_method_refused", "The Word connection can only save and report its own editing status.")
  await next()
}

export function wordDocumentRoutes(deps: Dependencies) {
  const app = new Hono()
  app.onError((error, c) => {
    if (error instanceof WordPackageError) return failure(c, error.status, error.code, error.message)
    if (error instanceof DeliverableConflictError || error instanceof SourceDocumentConflictError)
      return failure(c, 409, "word_document_changed", error.message)
    console.error("[word] request failed", { name: error.name })
    return failure(c, 503, "word_save_failed", "The Word connection could not complete this request. Try again after reconnecting.")
  })
  app.get("/word/config", c => c.json(wordConfiguration()))

  app.post("/matters/:id/documents/word", async c => {
    const config = wordConfiguration()
    if (!config.available || !config.origin) return failure(c, 503, "word_unavailable", config.reason ?? "Word is not configured.")
    const matter = deps.matter(c.req.param("id"))
    const name = c.req.query("name") ?? ""
    if (!name.toLowerCase().endsWith(".docx") || /[\\/\0]/.test(name)) return failure(c, 400, "word_requires_docx", "Choose a Word document in this matter, without a path.")
    let input: unknown
    try { input = await c.req.json() } catch { return failure(c, 400, "invalid_word_origin", "Provide the doc.haus browser origin.") }
    const browserOrigin = typeof input === "object" && input && "origin" in input ? input.origin : undefined
    try {
      const url = new URL(typeof browserOrigin === "string" ? browserOrigin : "")
      if (!/^https?:$/.test(url.protocol) || url.origin !== browserOrigin || url.username || url.password) throw new Error()
    } catch { return failure(c, 400, "invalid_word_origin", "Provide a valid doc.haus browser origin.") }
    const target = readDocumentTarget(matter.dir, name)
    if (!target || !existsSync(target.docPath)) return failure(c, 404, "word_document_missing", "That document is not in this matter.")
    return withArtifactMutationLocksAsync(matter.dir, deps.locks(matter.dir, target.docPath), async () => {
      const registry = openDb(matter.dir)
      try {
        if (!getDocumentRole(registry, target.docPath)) return failure(c, 404, "word_document_missing", "That document is not registered in this matter.")
      } finally { registry.close() }
      const bytes = readFileSync(target.docPath)
      const plain = Buffer.from(stripWordLink(bytes))
      let current = readCurrentVersion(matter.dir, target.name)
      // Old installations or an outside edit may not have a captured current
      // version. Reuse the one writer to capture those exact current bytes.
      if (!current || current.current.sha256 !== sha256Hex(bytes)) {
        await deps.write(matter.dir, target.name, plain, sha256Hex(bytes), { origin: "unknown" })
        current = readCurrentVersion(matter.dir, target.name)
      }
      if (!current) return failure(c, 409, "word_version_missing", "The document's current version could not be captured.")
      const token = randomBytes(32).toString("base64url")
      const linked = linkedWordDocx(plain, { version: 1, origin: browserOrigin as string, apiOrigin: config.origin!,
        matter: matter.id, matterName: matter.title, document: current.document.id, documentName: target.name, baseVersion: current.current.id, token })
      const db = openDb(matter.dir)
      try { db.run("INSERT INTO word_grants (token_hash, document_id, origin, issued_version, created_at) VALUES (?, ?, ?, ?, ?)",
        [digest(token), current.document.id, config.origin!, current.current.id, Date.now()]) } finally { db.close() }
      return new Response(linked, { headers: { "Content-Type": DOCX, "Content-Disposition": attachmentDocumentDisposition(target.name), "Cache-Control": "no-store" } })
    })
  })

  app.get("/matters/:id/documents/word", c => {
    const matter = deps.matter(c.req.param("id"))
    const db = openDb(matter.dir)
    try {
      const rows = db.query(`SELECT d.id, d.name, d.current_version, MAX(g.overlap) AS overlap FROM versioned_documents d
        JOIN word_grants g ON g.document_id = d.id WHERE d.removed_at IS NULL AND g.revoked_at IS NULL GROUP BY d.id`).all() as
        { id: string; name: string; current_version: string | null; overlap: number }[]
      const documents: WordOpenDocument[] = rows.map(row => {
        const sessions = activeWordSessions(db, row.id)
        return { name: row.name, open: sessions > 0, currentVersion: row.current_version, overlap: Boolean(row.overlap) || sessions > 1 }
      })
      return c.json({ documents })
    } finally { db.close() }
  })

  app.delete("/matters/:id/documents/word", async c => {
    const matter = deps.matter(c.req.param("id"))
    const name = c.req.query("name") ?? ""
    if (!name || /[\\/\0]/.test(name)) return failure(c, 400, "word_requires_docx", "Choose a Word document in this matter, without a path.")
    const target = readDocumentTarget(matter.dir, name)
    if (!target) return failure(c, 404, "word_document_missing", "That document is not in this matter.")
    return withArtifactMutationLocksAsync(matter.dir, [target.name], async () => {
      const current = readCurrentVersion(matter.dir, target.name)
      const db = openDb(matter.dir)
      try { if (current) db.run("UPDATE word_grants SET revoked_at = ? WHERE document_id = ? AND revoked_at IS NULL", [Date.now(), current.document.id]) }
      finally { db.close() }
      return c.json({ ok: true })
    })
  })

  function authenticate(c: Context) {
    const match = /^Bearer ([A-Za-z0-9_-]{43})$/.exec(c.req.header("Authorization") ?? "")
    if (!match) return
    const hash = digest(match[1])
    for (const matter of deps.matters()) {
      const db = openDb(matter.dir)
      try {
        const grant = db.query("SELECT * FROM word_grants WHERE token_hash = ?").get(hash) as WordGrant | null
        if (grant && !grant.revoked_at && grant.origin === c.req.header("Origin")) {
          const issued = readVersion(matter.dir, grant.issued_version)
          if (!issued || issued.version.removed || issued.document.removed_at != null || issued.document.id !== grant.document_id) return
          const target = readDocumentTarget(matter.dir, issued.document.name)
          if (!target || !existsSync(target.docPath) || !getDocumentRole(db, target.docPath)) return
          return { matter, grant, target }
        }
      } finally { db.close() }
    }
  }
  function touch(matter: Matter, grant: WordGrant, session: string, done = false) {
    const db = openDb(matter.dir)
    try {
      db.run(`INSERT INTO word_sessions (token_hash, session, seen_at, done) VALUES (?, ?, ?, ?)
        ON CONFLICT(token_hash, session) DO UPDATE SET seen_at = excluded.seen_at, done = excluded.done`,
      [grant.token_hash, session, Date.now(), Number(done)])
      return activeWordSessions(db, grant.document_id) > 1
    } finally { db.close() }
  }

  app.post("/word/saves", async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    const base = c.req.header("x-dochaus-word-base-version") ?? ""
    if (!base) return failure(c, 400, "word_base_required", "The Word save must identify the version it was edited from.")
    const size = Number(c.req.header("content-length"))
    if (size > MAX_WORD_UPLOAD_BYTES) return failure(c, 413, "word_too_large", "The Word file exceeds the 128 MB upload limit.")
    // Bound the stream as well: chunked requests cannot bypass content-length.
    const reader = c.req.raw.body?.getReader()
    if (!reader) return failure(c, 400, "invalid_docx", "A Word document is required.")
    const chunks: Uint8Array[] = []
    let received = 0
    try {
      for (;;) {
        const { done, value } = await reader.read()
        if (done) break
        received += value.byteLength
        if (received > MAX_WORD_UPLOAD_BYTES) { await reader.cancel(); return failure(c, 413, "word_too_large", "The Word file exceeds the 128 MB upload limit.") }
        chunks.push(value)
      }
    } finally { reader.releaseLock() }
    const raw = Buffer.concat(chunks)
    const actor = wordLastModifiedBy(raw)
    const bytes = Buffer.from(stripWordLink(raw))
    return withArtifactMutationLocksAsync(auth.matter.dir, deps.locks(auth.matter.dir, auth.target.docPath), async () => {
      // Re-check after awaiting upload/locks: deletion, revoke and redaction may
      // have happened since the connection was first authenticated.
      const live = authenticate(c)
      if (!live || live.grant.document_id !== auth.grant.document_id)
        return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
      const parent = readVersion(live.matter.dir, base)
      if (!parent || parent.document.id !== live.grant.document_id || parent.document.removed_at != null)
        return failure(c, 409, "word_base_invalid", "This save's base version does not belong to this document.")
      if (parent.version.removed) return failure(c, 410, "removed_by_redaction", "This Word copy predates a redaction. Reopen the redacted document from doc.haus.")
      const current = readCurrentVersion(live.matter.dir, live.target.name)
      const session = sessionId(c)
      let overlap = current?.current.id !== base
      if (session) overlap = touch(live.matter, live.grant, session) || overlap
      const working = readFileSync(live.target.docPath)
      const workingHash = sha256Hex(working)
      if (current && current.current.sha256 !== workingHash) overlap = true
      // A retried exact export acknowledges the existing version and does not
      // reindex, rebuild, or revert a later writer to an older branch.
      const hash = sha256Hex(bytes)
      const db = openDb(live.matter.dir)
      let sessionOverlap = false
      const priorSave = (() => { try {
        sessionOverlap = activeWordSessions(db, live.grant.document_id) > 1
        overlap = sessionOverlap || overlap
        return db.query("SELECT id, number, parent FROM document_versions WHERE document_id = ? AND sha256 = ? AND origin = 'word' AND removed IS NULL AND (parent = ? OR id = ?) ORDER BY number DESC LIMIT 1")
          .get(live.grant.document_id, hash, base, base) as { id: string; number: number; parent: string | null } | null
      } finally { db.close() } })()
      const recordOverlap = () => {
        if (!overlap) return
        const registry = openDb(live.matter.dir)
        try { registry.run("UPDATE word_grants SET overlap = 1 WHERE document_id = ?", [live.grant.document_id]) } finally { registry.close() }
      }
      if (priorSave) {
        // A lost-ack retry against the version it just wrote is not itself an
        // overlapping edit. Preserve a real earlier overlap or active panes.
        if (current?.current.id === priorSave.id && current.current.sha256 === hash && workingHash === hash)
          overlap = sessionOverlap || Boolean(live.grant.overlap)
        recordOverlap()
        return c.json({ version: { ...priorSave, created: false }, overlap } satisfies WordSaveResponse)
      }
      if (current?.current.sha256 === hash && workingHash === hash) {
        recordOverlap()
        return c.json({ version: { id: current.current.id, number: current.current.number, parent: current.current.parent, created: false }, overlap } satisfies WordSaveResponse)
      }
      const written = await deps.write(live.matter.dir, live.target.name, bytes, workingHash, { origin: "word", actor, parent: base })
      recordOverlap()
      return c.json({ version: { id: written.version.id, number: written.version.number, parent: written.version.parent, created: written.version.created }, overlap } satisfies WordSaveResponse)
    })
  })

  for (const route of ["heartbeat", "done"] as const) app.post(`/word/${route}`, async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    const session = sessionId(c)
    if (!session) return failure(c, 400, "word_session_required", "A Word pane session identifier is required.")
    const overlap = touch(auth.matter, auth.grant, session, route === "done")
    const current = readCurrentVersion(auth.matter.dir, auth.target.name)
    return c.json({ ok: true, currentVersion: current?.current.id ?? null, overlap: overlap || Boolean(auth.grant.overlap) })
  })
  return app
}
