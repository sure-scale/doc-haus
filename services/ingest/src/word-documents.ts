import { Hono, type Context } from "hono"
import { createHash, randomBytes } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import { linkedWordDocx, MAX_WORD_UPLOAD_BYTES, stripWordLink, wordLastModifiedBy, WordPackageError } from "../../../dochaus/lib/word-linked-docx"
import { WORD_ADDIN_VERSION, type WordOpenDocument, type WordSaveResponse, type WordConnectionState } from "../../../dochaus/lib/word-addin-contract"
import { wordChatPath } from "../../../dochaus/lib/word-chat-contract"
import { claimWordEditLease, clearWordEditLease, EditingLeaseError, readWordEditLease, releaseWordEditLease, renewWordEditLease, requireWordEditLease, wordLeaseState, WORD_EDIT_LEASE_TTL, type WordLeaseIdentity, type WordLeaseOwner } from "../../../dochaus/lib/word-edit-lease"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { getDocumentRole, openDb } from "./db"
import { readDocumentTarget } from "./registered-document"
import { readCurrentVersion, readVersion, sha256Hex, type CapturedVersion, type VersionWrite } from "./document-versions"
import { activeWordSessions, type WordGrant } from "./word-grants"
import { attachmentDocumentDisposition } from "./content-disposition"
import { DeliverableConflictError } from "./deliverable"
import { SourceDocumentConflictError } from "./ingest"
import { wordHttpsConfiguration } from "./word-configuration"
import { wordHttpsAvailable, wordHttpsFailure } from "./word-https"
import { installWordSourceDesk, readDeskSource } from "./word-source-desk"
import { installWordChat } from "./word-chat"
import { WordOpenTickets, wordOpenPath, wordOpenResultPath } from "./word-open-tickets"
import { readWordHandout, writeWordHandout, type WordHandoutIdentity } from "./word-handouts"

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
export const wordTokenPath = (pathname: string) => /^\/word\/(saves|heartbeat|done|lease|latest|desk\/(sources|source|versions|compare|open))$/.test(pathname) || wordChatPath(pathname)

export function wordConfiguration() {
  const failure = wordHttpsFailure()
  if (failure) return { available: false, origin: null, manifestUrl: null, reason: failure }
  let config: ReturnType<typeof wordHttpsConfiguration>
  try { config = wordHttpsConfiguration() } catch {
    return { available: false, origin: null, manifestUrl: null, reason: "Word setup needs repair. Open Settings → Add-ins." }
  }
  const available = Boolean(config && wordHttpsAvailable(config))
  const origin = available ? config!.origin : null
  return {
    available, origin, manifestUrl: origin ? `${origin}/word-addin/manifest.xml` : null,
    ...(!available ? { reason: "Set up or repair Microsoft Word in Settings → Add-ins." } : {}),
  }
}

type Matter = { id: string; title: string; dir: string }
type Dependencies = {
  matter: (id: string) => Matter
  matters: () => Matter[]
  locks: (directory: string, file: string) => string[]
  write: (directory: string, name: string, bytes: Buffer, prior: string, version: VersionWrite, owner?: WordLeaseOwner) => Promise<{ version: CapturedVersion }>
}
const digest = (token: string) => createHash("sha256").update(token).digest("hex")
const failure = (c: Context, status: 400 | 401 | 403 | 404 | 409 | 410 | 413 | 423 | 426 | 503, code: string, error: string) => c.json({ error, code }, status)
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
    c.header("Access-Control-Allow-Headers", "Authorization, Content-Type, X-DocHaus-Word-Base-Version, X-DocHaus-Word-Session, X-DocHaus-Word-Version, X-DocHaus-Word-Lease")
    return c.body(null, 204)
  }
  if (c.req.method !== "POST") return failure(c, 403, "word_method_refused", "The Word connection only accepts scoped document commands.")
  if (c.req.header("x-dochaus-word-version") !== WORD_ADDIN_VERSION)
    return failure(c, 426, "word_addin_update_required", "This Word add-in is out of date. Update Word setup in doc.haus, restart Word, then open a fresh copy.")
  await next()
}

export function wordDocumentRoutes(deps: Dependencies) {
  const app = new Hono()
  const tickets = new WordOpenTickets<{ identity: WordHandoutIdentity; tokenHash: string; name: string; previousOwner?: WordLeaseOwner; requestedBy?: { tokenHash: string; sourceId: string } }>()
  app.use("/word/open/*", async (c, next) => {
    c.header("Cache-Control", "no-store")
    c.header("Referrer-Policy", "no-referrer")
    await next()
  })
  app.onError((error, c) => {
    if (wordOpenPath(c.req.path)) tickets.result(c.req.param("nonce") ?? "", "failed")
    if (error instanceof EditingLeaseError) {
      const auth = authenticate(c)
      return c.json({ error: error.message, code: error.code, expiresAt: error.expiresAt,
        ...(auth ? connection(auth, c.req.header("x-dochaus-word-base-version") ?? auth.grant.issued_version, sessionId(c)) : {}) }, 423)
    }
    if (error instanceof WordPackageError) return failure(c, error.status, error.code, error.message)
    if (error instanceof DeliverableConflictError || error instanceof SourceDocumentConflictError) {
      const auth = authenticate(c)
      return c.json({ error: error.message, code: "word_document_changed",
        ...(auth ? connection(auth, c.req.header("x-dochaus-word-base-version") ?? auth.grant.issued_version, sessionId(c)) : {}) }, 409)
    }
    console.error("[word] request failed", { name: error.name })
    return failure(c, 503, "word_save_failed", "The Word connection could not complete this request. Try again after reconnecting.")
  })
  app.get("/word/config", c => c.json(wordConfiguration()))

  function sendHandout(c: Context, native: boolean, matter: Matter, target: { name: string },
    current: NonNullable<ReturnType<typeof readCurrentVersion>>, plain: Buffer, browserOrigin: string, apiOrigin: string, previousOwner?: WordLeaseOwner, requestedBy?: { tokenHash: string; sourceId: string }) {
    const identity: WordHandoutIdentity = { origin: browserOrigin, apiOrigin,
      matter: matter.id, document: current.document.id, baseVersion: current.current.id }
    const db = openDb(matter.dir)
    let linked: Uint8Array, tokenHash: string
    try {
      const cached = readWordHandout(matter.dir, identity, hash => Boolean(db.query(`SELECT token_hash FROM word_grants
        WHERE token_hash = ? AND document_id = ? AND origin = ? AND issued_version = ? AND revoked_at IS NULL`)
        .get(hash, identity.document, identity.apiOrigin, identity.baseVersion)))
      if (cached && cached.link.documentName === target.name) {
        linked = cached.bytes
        tokenHash = digest(cached.link.token)
      } else {
        const token = randomBytes(32).toString("base64url")
        linked = linkedWordDocx(plain, { version: 1, ...identity, matterName: matter.title, documentName: target.name, token })
        tokenHash = digest(token)
        db.run("INSERT INTO word_grants (token_hash, document_id, origin, issued_version, created_at) VALUES (?, ?, ?, ?, ?)",
          [tokenHash, current.document.id, apiOrigin, current.current.id, Date.now()])
        writeWordHandout(matter.dir, identity, linked)
      }
    } finally { db.close() }
    if (native) {
      const { nonce, id } = tickets.issue({ identity, tokenHash, name: target.name, ...(previousOwner ? { previousOwner } : {}), ...(requestedBy ? { requestedBy } : {}) })
      c.header("Cache-Control", "no-store")
      return c.json({ uri: `doc.haus://word/open?source=${encodeURIComponent(`${apiOrigin}/word/open/${nonce}`)}`, id })
    }
    return new Response(new Uint8Array(linked), { headers: { "Content-Type": DOCX, "Content-Disposition": attachmentDocumentDisposition(target.name), "Cache-Control": "no-store" } })
  }

  const handout = (native: boolean) => async (c: Context) => {
    const config = wordConfiguration()
    if (!config.available || !config.origin) return failure(c, 503, "word_unavailable", config.reason ?? "Word is not configured.")
    const matter = deps.matter(c.req.param("id") ?? "")
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
      return sendHandout(c, native, matter, target, current, plain, browserOrigin as string, config.origin!)
    })
  }
  // A pane capability may open only its own document's recorded current copy.
  // Prefetching never rebases, saves, closes or releases the older Word session.
  app.post("/word/latest", async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    const input = await c.req.json().catch(() => null) as { origin?: unknown; mode?: unknown } | null
    const browserOrigin = input?.origin
    try {
      const url = new URL(typeof browserOrigin === "string" ? browserOrigin : "")
      if (!/^https?:$/.test(url.protocol) || url.origin !== browserOrigin || url.username || url.password) throw new Error()
    } catch { return failure(c, 400, "invalid_word_origin", "Provide a valid doc.haus browser origin.") }
    return withArtifactMutationLocksAsync(auth.matter.dir, deps.locks(auth.matter.dir, auth.target.docPath), async () => {
      const live = authenticate(c)
      if (!live || live.grant.document_id !== auth.grant.document_id) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid.")
      const session = sessionId(c)
      const state = connection(live, c.req.header("x-dochaus-word-base-version") ?? live.grant.issued_version, session)
      if (!state.version.base) return c.json({ error: "This copy's version does not belong to this document.", code: "word_base_invalid", ...state }, 409)
      if (state.version.state === "removed" || state.version.state === "unrecorded-change") return requireCurrent(c, state)!
      const config = wordConfiguration()
      if (!config.available || !config.origin) return failure(c, 503, "word_unavailable", config.reason ?? "Word is not configured.")
      const current = readCurrentVersion(live.matter.dir, live.target.name)
      if (!current || current.document.id !== live.grant.document_id) return failure(c, 404, "word_document_missing", "That document is no longer in this matter.")
      if (input?.mode === "inline") {
        const bytes = Buffer.from(stripWordLink(readFileSync(live.target.docPath)))
        // Word's whole-document insertion API supports packages up to 4 MiB.
        if (bytes.length > 4 * 1024 * 1024) return failure(c, 413, "word_refresh_too_large", "This document is too large to refresh here. Open the latest copy separately.")
        c.header("Cache-Control", "no-store")
        return c.json({ version: { id: current.current.id, number: current.current.number }, bytes: bytes.toString("base64"), sha256: sha256Hex(bytes) })
      }
      // The explicit native read may hand over only this stale pane's own
      // lease. Another editor's lease is never changed, including after a race.
      const previousOwner = session && state.version.state === "stale" && state.lease.state === "owned" && state.lease.generation
        ? { ...identity(live, session), generation: state.lease.generation } : undefined
      return sendHandout(c, true, live.matter, live.target, current, Buffer.from(stripWordLink(readFileSync(live.target.docPath))), browserOrigin as string, config.origin, previousOwner)
    })
  })
  app.post("/word/desk/open", async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    c.header("Cache-Control", "no-store")
    const text = await c.req.text()
    let input: Record<string, unknown> | undefined
    try {
      const value: unknown = text.length <= 2048 ? JSON.parse(text) : undefined
      if (value && typeof value === "object" && !Array.isArray(value)) input = value as Record<string, unknown>
    } catch { /* Reject malformed identifiers without selecting a document. */ }
    if (!input || Object.keys(input).length !== 3 || Object.keys(input).some(key => !["sourceId", "expectedHash", "origin"].includes(key)) ||
      typeof input.sourceId !== "string" || !/^[1-9]\d{0,14}$/.test(input.sourceId) ||
      typeof input.expectedHash !== "string" || !/^[a-f0-9]{64}$/.test(input.expectedHash))
      return failure(c, 400, "invalid_desk_open", "Choose an identified matter document from the list.")
    const { sourceId, expectedHash, origin } = input
    try {
      const url = new URL(typeof origin === "string" ? origin : "")
      if (!/^https?:$/.test(url.protocol) || url.origin !== origin || url.username || url.password) throw new Error()
    } catch { return failure(c, 400, "invalid_word_origin", "Provide a valid doc.haus browser origin.") }
    const selected = readDeskSource(auth, sourceId)
    if (!selected || selected.hash !== expectedHash) return failure(c, 409, "desk_source_changed", "This document changed or is unavailable. Refresh the document list and choose it again.")
    if (!selected.name.toLowerCase().endsWith(".docx")) return failure(c, 400, "word_requires_docx", "Choose a Word document from this matter.")
    const target = readDocumentTarget(auth.matter.dir, selected.name)
    if (!target) return failure(c, 409, "desk_source_changed", "This document is no longer available.")
    return withArtifactMutationLocksAsync(auth.matter.dir, deps.locks(auth.matter.dir, target.docPath), async () => {
      const live = authenticate(c)
      if (!live || live.matter.id !== auth.matter.id || live.grant.token_hash !== auth.grant.token_hash) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid.")
      const after = readDeskSource(live, sourceId)
      const current = readCurrentVersion(live.matter.dir, target.name)
      if (!after || after.name !== target.name || after.hash !== expectedHash || !current || after.version?.id !== current.current.id)
        return failure(c, 409, "desk_source_changed", "This document changed while opening. Refresh the document list and choose it again.")
      const config = wordConfiguration()
      if (!config.available || !config.origin) return failure(c, 503, "word_unavailable", config.reason ?? "Word is not configured.")
      // Opening a matter document creates only a native handout. Existing Word
      // content, editing leases and save sessions remain untouched.
      return sendHandout(c, true, live.matter, target, current, Buffer.from(stripWordLink(readFileSync(target.docPath))), origin as string, config.origin, undefined, { tokenHash: live.grant.token_hash, sourceId })
    })
  })
  app.post("/matters/:id/documents/word", handout(false))
  app.post("/matters/:id/documents/word/open", handout(true))
  app.get("/matters/:id/documents/word/open/status", c => {
    const matter = deps.matter(c.req.param("id")), id = c.req.query("id") ?? ""
    if (!/^[a-f0-9-]{36}$/.test(id)) return failure(c, 400, "word_open_status_invalid", "Provide a Word handoff identifier.")
    const state = tickets.status(id, ticket => ticket.identity.matter === matter.id)
    c.header("Cache-Control", "no-store")
    return state ? c.json({ state }) : failure(c, 404, "word_open_status_missing", "This Word handoff is no longer available.")
  })

  app.post("/word/open/:nonce/result", async c => {
    if (!wordOpenResultPath(c.req.path) || c.req.raw.headers.has("Origin") || c.req.raw.headers.has("Authorization"))
      return failure(c, 403, "word_open_refused", "This Word handoff is only available to the desktop helper.")
    if (!/^application\/json(?:\s*;|$)/i.test(c.req.header("Content-Type") ?? ""))
      return failure(c, 400, "word_open_result_invalid", "Provide the Word launch result.")
    const reader = c.req.raw.body?.getReader()
    if (!reader) return failure(c, 400, "word_open_result_invalid", "Provide the Word launch result.")
    const chunks: Uint8Array[] = []
    let size = 0
    while (true) {
      const next = await reader.read()
      if (next.done) break
      size += next.value.byteLength
      if (size > 64) { void reader.cancel().catch(() => {}); return failure(c, 400, "word_open_result_invalid", "Provide the Word launch result.") }
      chunks.push(next.value)
    }
    const text = new TextDecoder().decode(Buffer.concat(chunks))
    let input: unknown
    try { if (text.length > 64) throw new Error(); input = JSON.parse(text) } catch { /* Refuse every additional or unbounded field. */ }
    if (!input || typeof input !== "object" || Object.keys(input).length !== 1 || !("state" in input) || !["opened", "failed"].includes(String(input.state)))
      return failure(c, 400, "word_open_result_invalid", "Provide the Word launch result.")
    if (!tickets.result(c.req.param("nonce"), input.state as "opened" | "failed"))
      return failure(c, 410, "word_open_expired", "This Word handoff is no longer available.")
    c.header("Cache-Control", "no-store")
    return c.json({ ok: true })
  })

  app.get("/word/open/:nonce", async c => {
    if (c.req.method !== "GET") return failure(c, 403, "word_open_refused", "This Word handoff only accepts a desktop read.")
    if (!wordOpenPath(c.req.path) || c.req.raw.headers.has("Origin") || c.req.raw.headers.has("Authorization"))
      return failure(c, 403, "word_open_refused", "This Word handoff is only available to the desktop helper.")
    const ticket = tickets.redeem(c.req.param("nonce"))
    if (!ticket) return failure(c, 410, "word_open_expired", "This Word handoff expired or was already used. Choose Open in Word again.")
    const config = wordConfiguration()
    if (!config.available || config.origin !== ticket.identity.apiOrigin) {
      tickets.result(c.req.param("nonce"), "failed")
      return failure(c, 503, "word_unavailable", "The Word connection is not available. Open doc.haus and try again.")
    }
    const matter = deps.matter(ticket.identity.matter)
    const target = readDocumentTarget(matter.dir, ticket.name)
    if (!target || !existsSync(target.docPath)) { tickets.result(c.req.param("nonce"), "failed"); return failure(c, 410, "word_open_expired", "This document changed. Choose Open in Word again.") }
    return withArtifactMutationLocksAsync(matter.dir, deps.locks(matter.dir, target.docPath), async () => {
      const current = readCurrentVersion(matter.dir, target.name)
      const db = openDb(matter.dir)
      try {
        const requester = ticket.requestedBy ? authenticateToken(ticket.requestedBy.tokenHash, ticket.identity.apiOrigin) : undefined
        const selected = requester && ticket.requestedBy ? readDeskSource(requester, ticket.requestedBy.sourceId) : undefined
        const allowed = !ticket.requestedBy || (requester?.matter.id === ticket.identity.matter && selected?.name === target.name && selected.version?.id === ticket.identity.baseVersion)
        const valid = allowed && current?.document.id === ticket.identity.document && current.current.id === ticket.identity.baseVersion &&
          current.current.sha256 === sha256Hex(readFileSync(target.docPath)) && Boolean(getDocumentRole(db, target.docPath))
        const cached = valid ? readWordHandout(matter.dir, ticket.identity, hash => hash === ticket.tokenHash && Boolean(db.query(`SELECT token_hash FROM word_grants
          WHERE token_hash = ? AND document_id = ? AND origin = ? AND issued_version = ? AND revoked_at IS NULL`)
          .get(hash, ticket.identity.document, ticket.identity.apiOrigin, ticket.identity.baseVersion))) : undefined
        if (!cached || cached.link.documentName !== target.name) {
          tickets.result(c.req.param("nonce"), "failed")
          return failure(c, 410, "word_open_expired", "This document or its Word link changed. Choose Open in Word again.")
        }
        const previous = ticket.previousOwner
        if (previous) {
          const lease = readWordEditLease(db, previous.document)
          if (lease?.token_hash === previous.tokenHash && lease.session === previous.session && lease.generation === previous.generation) {
            releaseWordEditLease(db, previous)
            db.run("UPDATE word_sessions SET done = 1 WHERE token_hash = ? AND session = ?", [previous.tokenHash, previous.session])
          }
        }
        return new Response(new Uint8Array(cached.bytes), { headers: { "Content-Type": DOCX, "Content-Disposition": attachmentDocumentDisposition(target.name),
          "Cache-Control": "no-store", "Referrer-Policy": "no-referrer" } })
      } finally { db.close() }
    })
  })

  app.get("/matters/:id/documents/word", c => {
    const matter = deps.matter(c.req.param("id"))
    const db = openDb(matter.dir)
    try {
      const rows = db.query(`SELECT d.id, d.name, d.current_version, v.number, MAX(g.overlap) AS overlap FROM versioned_documents d
        JOIN word_grants g ON g.document_id = d.id LEFT JOIN document_versions v ON v.id = d.current_version
        WHERE d.removed_at IS NULL AND g.revoked_at IS NULL GROUP BY d.id`).all() as
        { id: string; name: string; current_version: string | null; number: number | null; overlap: number }[]
      const documents: WordOpenDocument[] = rows.map(row => {
        const sessions = activeWordSessions(db, row.id)
        return { name: row.name, open: sessions > 0, currentVersion: row.current_version, ...(row.number ? { currentNumber: row.number } : {}),
          lease: wordLeaseState(db, row.id), overlap: Boolean(row.overlap) || sessions > 1 }
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
      try { if (current) db.transaction(() => {
        db.run("UPDATE word_grants SET revoked_at = ? WHERE document_id = ? AND revoked_at IS NULL", [Date.now(), current.document.id])
        clearWordEditLease(db, current.document.id)
      }).immediate() }
      finally { db.close() }
      return c.json({ ok: true })
    })
  })

  function authenticate(c: Context) {
    const match = /^Bearer ([A-Za-z0-9_-]{43})$/.exec(c.req.header("Authorization") ?? "")
    if (!match) return
    return authenticateToken(digest(match[1]), c.req.header("Origin"))
  }
  function authenticateToken(hash: string, origin: string | undefined) {
    for (const matter of deps.matters()) {
      const db = openDb(matter.dir)
      try {
        const grant = db.query("SELECT * FROM word_grants WHERE token_hash = ?").get(hash) as WordGrant | null
        if (grant && !grant.revoked_at && grant.origin === origin) {
          const issued = readVersion(matter.dir, grant.issued_version)
          if (!issued || issued.version.removed || issued.document.removed_at != null || issued.document.id !== grant.document_id) return
          const target = readDocumentTarget(matter.dir, issued.document.name)
          if (!target || !existsSync(target.docPath) || !getDocumentRole(db, target.docPath)) return
          return { matter, grant, target }
        }
      } finally { db.close() }
    }
  }
  type Authenticated = NonNullable<ReturnType<typeof authenticate>>
  installWordSourceDesk(app, authenticate)
  installWordChat(app, authenticate)
  function identity(auth: Authenticated, session: string): WordLeaseIdentity {
    return { document: auth.grant.document_id, tokenHash: auth.grant.token_hash, session }
  }
  function owner(auth: Authenticated, session: string, c: Context): WordLeaseOwner {
    const generation = c.req.header("x-dochaus-word-lease") ?? ""
    if (!generation) throw new EditingLeaseError("Check editing status and acquire this document's Word editing session before saving.")
    return { ...identity(auth, session), generation }
  }
  function connection(auth: Authenticated, base: string, session: string | null): WordConnectionState {
    const parent = readVersion(auth.matter.dir, base)
    const belongs = parent?.document.id === auth.grant.document_id
    const current = readCurrentVersion(auth.matter.dir, auth.target.name)
    const currentMatches = current?.document.id === auth.grant.document_id
    const removed = Boolean(belongs && (parent.version.removed || parent.document.removed_at != null))
    const unrecorded = !currentMatches || !existsSync(auth.target.docPath) || current.current.sha256 !== sha256Hex(readFileSync(auth.target.docPath))
    const db = openDb(auth.matter.dir)
    try { return {
      version: {
        base: belongs ? { id: parent.version.id, number: parent.version.number } : null,
        current: currentMatches ? { id: current.current.id, number: current.current.number } : null,
        state: removed ? "removed" : unrecorded ? "unrecorded-change" : base === current.current.id ? "current" : "stale",
      }, lease: wordLeaseState(db, auth.grant.document_id, session ? identity(auth, session) : undefined),
    } } finally { db.close() }
  }
  function requireCurrent(c: Context, state: WordConnectionState) {
    if (state.version.state === "removed")
      return c.json({ error: "This Word copy predates a redaction. Keep your local file and reopen the redacted document from doc.haus.", code: "removed_by_redaction", ...state }, 410)
    if (state.version.state === "unrecorded-change")
      return c.json({ error: "The matter's file changed outside its recorded version. Keep your local Word file and refresh the document in doc.haus before continuing.", code: "word_document_changed", ...state }, 409)
    if (state.version.state !== "current")
      return c.json({ error: "This Word copy is behind the current doc.haus version. Keep your local file and compare or reopen the current copy before saving.", code: "word_base_stale", ...state }, 409)
  }
  app.post("/word/lease", async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    const session = sessionId(c)
    if (!session) return failure(c, 400, "word_session_required", "A Word pane session identifier is required.")
    const input = await c.req.json().catch(() => null) as { action?: unknown; baseVersion?: unknown; expectedCurrent?: unknown; expectedGeneration?: unknown } | null
    if (!input || typeof input.baseVersion !== "string" || !["inspect", "claim", "takeover", "release"].includes(String(input.action)))
      return failure(c, 400, "invalid_word_lease", "Provide an editing action and this copy's base version.")
    return withArtifactMutationLocksAsync(auth.matter.dir, deps.locks(auth.matter.dir, auth.target.docPath), async () => {
      const live = authenticate(c)
      if (!live || live.grant.document_id !== auth.grant.document_id) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid.")
      const base = input.baseVersion as string
      const state = connection(live, base, session)
      if (input.action === "inspect") return c.json(state)
      if (!state.version.base) return c.json({ error: "This save's base version does not belong to this document.", code: "word_base_invalid", ...state }, 409)
      const db = openDb(live.matter.dir)
      try {
        if (input.action === "release") {
          releaseWordEditLease(db, owner(live, session, c))
          touch(live.matter, live.grant, session, true)
        } else {
          const refused = requireCurrent(c, state)
          if (refused) return refused
          if (input.action === "takeover") {
            if (input.expectedCurrent !== state.version.current?.id || typeof input.expectedGeneration !== "string" || input.expectedGeneration !== readWordEditLease(db, live.grant.document_id)?.generation)
              return c.json({ error: "Editing status changed after it was displayed. Check it again before taking over.", code: "word_recovery_changed", ...state }, 409)
            clearWordEditLease(db, live.grant.document_id)
            retireSessions(db, live.grant.document_id)
          }
          claimWordEditLease(db, identity(live, session))
          touch(live.matter, live.grant, session)
        }
      } catch (error) {
        if (error instanceof EditingLeaseError) return c.json({ error: error.message, code: error.code, expiresAt: error.expiresAt, ...connection(live, base, session) }, 423)
        throw error
      } finally { db.close() }
      return c.json(connection(live, base, session))
    })
  })

  app.post("/matters/:id/documents/word/recover", async c => {
    const matter = deps.matter(c.req.param("id"))
    const name = c.req.query("name") ?? ""
    if (!name || /[\\/\0]/.test(name)) return failure(c, 400, "word_requires_docx", "Choose a Word document in this matter, without a path.")
    const target = readDocumentTarget(matter.dir, name)
    if (!target) return failure(c, 404, "word_document_missing", "That document is not in this matter.")
    const input = await c.req.json().catch(() => null) as { action?: unknown; expectedCurrent?: unknown; expectedGeneration?: unknown } | null
    if (!input || !["release", "takeover"].includes(String(input.action)) || typeof input.expectedCurrent !== "string" ||
      !(typeof input.expectedGeneration === "string" || input.expectedGeneration === null))
      return failure(c, 400, "invalid_word_recovery", "Provide the displayed current version, lease generation and recovery action.")
    return withArtifactMutationLocksAsync(matter.dir, deps.locks(matter.dir, target.docPath), async () => {
      const current = readCurrentVersion(matter.dir, target.name)
      if (!current) return failure(c, 404, "word_document_missing", "That document is no longer in this matter.")
      const db = openDb(matter.dir)
      try {
        const observed = readWordEditLease(db, current.document.id)
        if (input.expectedCurrent !== current.current.id || input.expectedGeneration !== (observed?.generation ?? null))
          return c.json({ error: "Editing status changed after it was displayed. Refresh it before recovering this document.", code: "word_recovery_changed", currentVersion: current.current.id, lease: wordLeaseState(db, current.document.id) }, 409)
        db.transaction(() => {
          if (input.action === "takeover") db.run("UPDATE word_grants SET revoked_at = ? WHERE document_id = ? AND revoked_at IS NULL", [Date.now(), current.document.id])
          clearWordEditLease(db, current.document.id)
          retireSessions(db, current.document.id)
        }).immediate()
        return c.json({ ok: true, currentVersion: current.current.id, lease: wordLeaseState(db, current.document.id) })
      } finally { db.close() }
    })
  })

  function touch(matter: Matter, grant: WordGrant, session: string, done = false) {
    const db = openDb(matter.dir)
    try {
      db.run(`INSERT INTO word_sessions (token_hash, session, seen_at, done) VALUES (?, ?, ?, ?)
        ON CONFLICT(token_hash, session) DO UPDATE SET seen_at = excluded.seen_at, done = excluded.done`,
      [grant.token_hash, session, Date.now(), Number(done)])
      return activeWordSessions(db, grant.document_id) > 1
    } finally { db.close() }
  }
  function retireSessions(db: ReturnType<typeof openDb>, documentId: string) {
    db.run("UPDATE word_sessions SET done = 1 WHERE token_hash IN (SELECT token_hash FROM word_grants WHERE document_id = ?)", [documentId])
  }

  app.post("/word/saves", async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    const base = c.req.header("x-dochaus-word-base-version") ?? ""
    if (!base) return failure(c, 400, "word_base_required", "The Word save must identify the version it was edited from.")
    const session = sessionId(c)
    if (!session) return failure(c, 400, "word_session_required", "A Word pane session identifier is required.")
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
        return c.json({ error: "This save's base version does not belong to this document.", code: "word_base_invalid", ...connection(live, base, session) }, 409)
      if (parent.version.removed) return c.json({ error: "This Word copy predates a redaction. Reopen the redacted document from doc.haus.", code: "removed_by_redaction", ...connection(live, base, session) }, 410)
      const current = readCurrentVersion(live.matter.dir, live.target.name)
      const leaseOwner = owner(live, session, c)
      const working = readFileSync(live.target.docPath)
      const workingHash = sha256Hex(working)
      // A retried exact export acknowledges the existing version and does not
      // reindex, rebuild, or revert a later writer to an older branch.
      const hash = sha256Hex(bytes)
      const db = openDb(live.matter.dir)
      const priorSave = (() => { try {
        requireWordEditLease(db, leaseOwner)
        return db.query("SELECT id, number, parent FROM document_versions WHERE document_id = ? AND sha256 = ? AND origin = 'word' AND removed IS NULL AND (parent = ? OR id = ?) ORDER BY number DESC LIMIT 1")
          .get(live.grant.document_id, hash, base, base) as { id: string; number: number; parent: string | null } | null
      } finally { db.close() } })()
      const renew = () => {
        const registry = openDb(live.matter.dir)
        try { renewWordEditLease(registry, leaseOwner) } finally { registry.close() }
        touch(live.matter, live.grant, session)
      }
      const overlap = Boolean(live.grant.overlap)
      if (priorSave && current?.current.id === priorSave.id && current.current.sha256 === hash && workingHash === hash) {
        renew()
        return c.json({ version: { ...priorSave, created: false }, overlap, connection: connection(live, priorSave.id, session) } satisfies WordSaveResponse)
      }
      const refused = requireCurrent(c, connection(live, base, session))
      if (refused) return refused
      if (current?.current.sha256 === hash && workingHash === hash) {
        renew()
        return c.json({ version: { id: current.current.id, number: current.current.number, parent: current.current.parent, created: false }, overlap,
          connection: connection(live, current.current.id, session) } satisfies WordSaveResponse)
      }
      renew()
      const written = await deps.write(live.matter.dir, live.target.name, bytes, workingHash, { origin: "word", actor, parent: base }, leaseOwner)
      // The document lock prevented another owner taking over while a source
      // save was being indexed. Give its successful acknowledgement a full TTL.
      const registry = openDb(live.matter.dir)
      try { registry.run("UPDATE word_edit_leases SET expires_at = ? WHERE document_id = ? AND generation = ?", [Date.now() + WORD_EDIT_LEASE_TTL, leaseOwner.document, leaseOwner.generation]) } finally { registry.close() }
      return c.json({ version: { id: written.version.id, number: written.version.number, parent: written.version.parent, created: written.version.created }, overlap,
        connection: connection(live, written.version.id, session) } satisfies WordSaveResponse)
    })
  })

  for (const route of ["heartbeat", "done"] as const) app.post(`/word/${route}`, async c => {
    const auth = authenticate(c)
    if (!auth) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid. Reopen the document from doc.haus.")
    const session = sessionId(c)
    if (!session) return failure(c, 400, "word_session_required", "A Word pane session identifier is required.")
    const base = c.req.header("x-dochaus-word-base-version") ?? ""
    if (!base) return failure(c, 400, "word_base_required", "Report the Word copy's acknowledged base version.")
    return withArtifactMutationLocksAsync(auth.matter.dir, deps.locks(auth.matter.dir, auth.target.docPath), async () => {
      const live = authenticate(c)
      if (!live || live.grant.document_id !== auth.grant.document_id) return failure(c, 401, "word_link_revoked", "This Word connection is no longer valid.")
      const leaseOwner = owner(live, session, c)
      const state = connection(live, base, session)
      const db = openDb(live.matter.dir)
      try {
        if (route === "done") {
          const refused = requireCurrent(c, state)
          if (refused) return refused
          releaseWordEditLease(db, leaseOwner)
        } else renewWordEditLease(db, leaseOwner)
      } finally { db.close() }
      const overlap = touch(live.matter, live.grant, session, route === "done")
      return c.json({ ok: true, currentVersion: state.version.current?.id ?? null, overlap: overlap || Boolean(live.grant.overlap), ...connection(live, base, session) })
    })
  })
  return app
}
