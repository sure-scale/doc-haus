import { Hono, type Context } from "hono"
import { Database } from "bun:sqlite"
import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { canonicalArtifactIdentity } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { createDocxLocalization, openRevision } from "../../../dochaus/lib/docx-amendments"
import { paragraphTargets, resolveParagraphTarget } from "../../../dochaus/lib/drafting-amendment"
import { correspondHistoryParagraphs, historyParagraph, resolveHistoryAnchor } from "../../../dochaus/lib/clause-history"
import type { ClauseHistoryChange, ClauseHistoryReason, ClauseHistoryRequest, ClauseHistoryResult, ClauseHistoryRound, ClauseHistorySourceRequest, ClauseHistorySourceResult, ClauseHistoryVersion } from "../../../dochaus/lib/clause-history-contract"
import { readDocxMarkup, type DocxMarkup } from "../../../dochaus/lib/docx-markup"
import { compareParagraphs, readComparisonParagraphs } from "../../../dochaus/lib/version-compare"
import type { ComparisonParagraph } from "../../../dochaus/lib/version-compare-contract"
import { readCurrentVersion, readDocumentVersions, readVersion, readVersionBytes, sha256Hex, type LinkedDocumentBranch, type VersionEventView, type VersionRow, type VersionView } from "./document-versions"
import { readDocumentReceivedBranch } from "./document-exchanges"
import { readDocumentTarget } from "./registered-document"
import type { WordGrant } from "./word-grants"
import type { RedlineRow } from "./db"
import { binding } from "./redline"

type HistoryAuth = { matter: { id: string; title: string; dir: string }; grant: WordGrant; target: { name: string; docPath: string } }
const SCOPES = ["body", "header", "footer", "footnote", "endnote"]
const MAX_FILE = 32 * 1024 * 1024
const MAX_COMPARISON_BYTES = 128 * 1024 * 1024

export class ClauseHistoryError extends Error {
  constructor(readonly status: 400 | 404 | 409 | 410 | 413 | 422, readonly code: string, message: string) { super(message) }
}
const fail = (status: ClauseHistoryError["status"], code: string, message: string): never => { throw new ClauseHistoryError(status, code, message) }
const actor = (name: string | undefined | null) => !name || name === "unknown" ? null : name
const validId = (id: unknown): id is string => typeof id === "string" && id.length > 0 && id.length <= 256 && !/[\u0000-\u001f]/.test(id)
const record = (value: unknown): value is Record<string, unknown> => Boolean(value) && typeof value === "object" && !Array.isArray(value)

export function clauseHistoryRequest(value: unknown): ClauseHistoryRequest | undefined {
  if (!record(value) || Object.keys(value).some(key => !["anchorVersion", "paragraph", "expectedCurrentHash"].includes(key)) || !validId(value.anchorVersion)
    || !record(value.paragraph) || Object.keys(value.paragraph).some(key => !["scope", "text", "anchorId"].includes(key))
    || typeof value.paragraph.scope !== "string" || !SCOPES.includes(value.paragraph.scope)
    || typeof value.paragraph.text !== "string" || !value.paragraph.text.length || value.paragraph.text.length > 128_000
    || value.paragraph.anchorId !== undefined && !validId(value.paragraph.anchorId)
    || value.expectedCurrentHash !== undefined && (typeof value.expectedCurrentHash !== "string" || !/^[a-f0-9]{64}$/.test(value.expectedCurrentHash))) return
  return value as ClauseHistoryRequest
}
function sourceRequest(value: unknown): ClauseHistorySourceRequest | undefined {
  if (!record(value) || Object.keys(value).some(key => !["versionId", "hash", "kind", "id", "exchangeId", "recordHash"].includes(key)) || !validId(value.versionId) || !validId(value.id)
    || value.exchangeId !== undefined && (!Number.isSafeInteger(value.exchangeId) || (value.exchangeId as number) < 1)
    || typeof value.hash !== "string" || !/^[a-f0-9]{64}$/.test(value.hash) || !["comment", "revision", "decision", "proposal"].includes(String(value.kind))
    || value.recordHash !== undefined && (typeof value.recordHash !== "string" || !/^[a-f0-9]{64}$/.test(value.recordHash))) return
  return value as ClauseHistorySourceRequest
}

function versionView(version: VersionRow | VersionView, events: VersionEventView[]): ClauseHistoryVersion {
  const parties = [...new Set(events.filter(event => event.version_id === version.id && (event.kind === "sent" || event.kind === "received"))
    .flatMap(event => typeof event.detail?.party === "string" ? [event.detail.party] : []))]
  return { id: version.id, number: version.number, hash: version.sha256, parent: version.parent, createdAt: version.created_at,
    actor: actor(version.actor), origin: version.origin, label: version.label, party: parties.length === 1 ? parties[0]! : null }
}
function targetHistory(directory: string, name: string) {
  if (!name || /[\/\\\0]/.test(name) || !/\.docx$/i.test(name)) return fail(400, "history_requires_word", "Choose a Word document in this matter, without a path.")
  const target = readDocumentTarget(directory, name)
  const history = target && readDocumentVersions(directory, target.name)
  const current = target && readCurrentVersion(directory, target.name)
  if (!target || !history?.document || !current || !existsSync(target.docPath)) return fail(404, "history_document_missing", "This document's saved history is unavailable.")
  if (current.current.removed) return fail(410, "history_removed", "This document's current version is unavailable.")
  if (sha256Hex(readFileSync(target.docPath)) !== current.current.sha256)
    return fail(409, "history_working_changed", "The working file differs from its saved current version. History has not captured or saved it.")
  return { target, history, current }
}

function records(version: VersionRow | VersionView, markup: DocxMarkup, paragraphs: ComparisonParagraph[], events: VersionEventView[], versions: VersionView[]) {
  const commentIds = new Set(paragraphs.flatMap(paragraph => paragraph.comments))
  // Replies inherit the root's exact native comment association, not a text match.
  for (let changed = true; changed;) {
    changed = false
    for (const comment of markup.comments) if (comment.parent !== null && commentIds.has(comment.parent) && !commentIds.has(comment.id)) { commentIds.add(comment.id); changed = true }
  }
  const comments: ClauseHistoryReason[] = markup.comments.filter(comment => commentIds.has(comment.id) && (!comment.detached || comment.parent !== null && commentIds.has(comment.parent))).map(comment => ({
    kind: "comment", text: comment.text, author: actor(comment.author), date: comment.date, parent: comment.parent === null ? null : String(comment.parent), resolved: comment.resolved,
    source: { versionId: version.id, hash: version.sha256, kind: "comment", id: String(comment.id) },
  }))
  const revisions = new Set(paragraphs.flatMap(paragraph => paragraph.revisions ?? []))
  const changes: ClauseHistoryChange[] = markup.revisions.filter(revision => revisions.has(revision.id)).map(revision => ({
    kind: revision.kind, text: revision.text, author: actor(revision.author), date: revision.date,
    source: { versionId: version.id, hash: version.sha256, kind: "revision", id: revision.id },
  }))
  const decisions: ClauseHistoryReason[] = events.flatMap(event => {
    const written = versions.find(candidate => candidate.id === event.version_id)
    if (!written || written.removed || written.parent !== version.id || !Array.isArray(event.detail?.markup)) return []
    return event.detail.markup.flatMap(value => !record(value) || typeof value.id !== "string" || !revisions.has(value.id)
      || !["accept", "reject", "hold"].includes(String(value.decision)) ? [] : [{
        kind: "decision" as const, text: "", author: actor(event.actor), date: new Date(event.created_at).toISOString(),
        decision: value.decision as "accept" | "reject" | "hold", ...(typeof value.outcome === "string" ? { outcome: value.outcome } : {}),
        source: { versionId: version.id, hash: version.sha256, kind: "decision" as const, id: `${event.id}:${value.id}` },
      }])
  })
  return { reasons: [...comments, ...decisions], changes, reasonStatus: comments.length ? "recorded" as const : "not-recorded" as const }
}

function proposalRow(directory: string, id: number) {
  const db = new Database(path.join(directory, ".dochaus", "legal.db"), { readonly: true })
  try { return db.query("SELECT * FROM redlines WHERE id=?").get(id) as RedlineRow | null } finally { db.close() }
}
function proposalBase(event: VersionEventView, versions: VersionView[]) {
  const written = versions.find(version => version.id === event.version_id)
  return !written || written.removed ? undefined : event.kind === "accept" ? written.parent : event.kind === "reject" ? written.id : undefined
}
const proposalPin = (event: VersionEventView, row: RedlineRow) => sha256Hex(new TextEncoder().encode(JSON.stringify([
  event.id, event.kind, event.version_id, event.actor, event.created_at, row.id, row.doc_path, row.doc_name, row.scope,
  row.find_text, row.old_text, row.new_text, row.author, row.review_comment, row.anchor_id, row.target_json ?? null, row.formatting_json ?? null, row.status, row.created_at,
])))

/** These are existing proposal/decision records, not reconstructed negotiation
 * rationale. Resolve only on the exact event-bound saved bytes through an
 * actual source anchor or the same exact Prepare binding used by the writer. */
async function proposalReasons(directory: string, version: VersionRow | VersionView, bytes: Uint8Array, selected: ComparisonParagraph[], events: VersionEventView[], versions: VersionView[]): Promise<ClauseHistoryReason[]> {
  const decisions = events.filter(event => proposalBase(event, versions) === version.id && Array.isArray(event.detail?.redline_ids))
  if (!decisions.length) return []
  const owner = readVersion(directory, version.id)
  if (!owner) return []
  const opened = await openRevision(bytes, { pendingRevisions: "layer" })
  let targets: Awaited<ReturnType<typeof paragraphTargets>> | undefined
  try {
    const reasons: ClauseHistoryReason[] = []
    for (const event of decisions) for (const id of event.detail!.redline_ids as unknown[]) {
      if (!Number.isSafeInteger(id) || (id as number) < 1) continue
      const row = proposalRow(directory, id as number)
      if (!row || !["phrase", "clause"].includes(row.scope) || row.status !== (event.kind === "accept" ? "accepted" : "rejected")
        || !row.find_text || canonicalArtifactIdentity(directory, row.doc_path) !== canonicalArtifactIdentity(directory, path.join(directory, owner.document.name))) continue
      let anchors: string[] = []
      try {
        if (row.target_json) {
          const target = binding(row)
          if (target.version === 1 && target.source_sha256 !== version.sha256) continue
          const packet = await createDocxLocalization(bytes, [{ id: target.amendment_id, before_quote: row.find_text, replacement: row.new_text }], {}, { pendingRevisions: "layer" })
          let chosen: typeof packet.candidates[number] | undefined
          if (target.version === 1) chosen = packet.candidates.find(candidate => candidate.target_handle === target.target_handle)
          else {
            targets ??= await paragraphTargets(bytes, { pendingRevisions: "layer" })
            const pointed = "pid" in target ? resolveParagraphTarget(targets, target.pid) : targets.find(candidate => candidate.para_id === target.para_id)
            if (!pointed || pointed.sha256 !== target.paragraph_sha256) continue
            chosen = packet.candidates.filter(candidate => candidate.paragraph.paragraph_handle === pointed.paragraph_handle)[target.occurrence]
          }
          const at = chosen ? opened.revision.paragraphs.findIndex(paragraph => paragraph.paragraph_handle === chosen!.paragraph.paragraph_handle) : -1
          if (at >= 0) anchors = [opened.paragraphs[at]!.match.enclosingAnchor.id]
        } else if (row.anchor_id) {
          const anchored = opened.paragraphs.find(paragraph => paragraph.match.enclosingAnchor.id === row.anchor_id)
          if (!anchored || !anchored.match.text.includes(row.find_text)) continue
          if (row.scope === "clause") {
            const exact = opened.session.projectAnchor(row.anchor_id).markdown.replace(/\{#[^}]*\}/g, "").trim()
            if (exact !== row.old_text) continue
            anchors = [row.anchor_id]
          } else if (row.old_text === row.find_text) {
            // Legacy phrase acceptance replaces every exact occurrence. The
            // recorded first anchor must exist; no similar passage is admitted.
            anchors = opened.paragraphs.filter(paragraph => paragraph.match.text.includes(row.find_text)).map(paragraph => paragraph.match.enclosingAnchor.id)
          }
        }
      } catch { continue }
      if (!selected.some(paragraph => anchors.includes(paragraph.anchorId) && paragraph.text.includes(row.find_text))) continue
      const source = { versionId: version.id, hash: version.sha256, id: `redline:${event.id}:${row.id}`, recordHash: proposalPin(event, row) }
      reasons.push({ kind: "proposal", text: row.review_comment, author: actor(row.author), date: new Date(row.created_at).toISOString(),
        source: { ...source, kind: "proposal" }, proposal: { id: row.id, scope: row.scope, oldText: row.old_text, newText: row.new_text } })
      reasons.push({ kind: "decision", text: "", author: actor(event.actor), date: new Date(event.created_at).toISOString(), decision: event.kind as "accept" | "reject",
        outcome: event.kind === "accept" ? "applied" : "not-applied", source: { ...source, kind: "decision" } })
    }
    return reasons
  } finally { opened.session.close() }
}

function proposalSourceCurrent(directory: string, source: ClauseHistorySourceRequest) {
  const match = /^redline:(\d+):(\d+)$/.exec(source.id)
  if (!match) return source.kind !== "proposal"
  if (!source.recordHash) return false
  const found = readVersion(directory, source.versionId)
  if (!found || found.version.removed || found.document.removed_at !== null || found.version.sha256 !== source.hash) return false
  const history = readDocumentVersions(directory, found.document.name)
  const event = history.events.find(event => event.id === Number(match[1]))
  const row = proposalRow(directory, Number(match[2]))
  return Boolean(event && row && Array.isArray(event.detail?.redline_ids) && event.detail.redline_ids.includes(row.id)
    && proposalBase(event, history.versions) === source.versionId && proposalPin(event, row) === source.recordHash)
}
async function historyRecords(directory: string, version: VersionRow | VersionView, bytes: Uint8Array, markup: DocxMarkup, paragraphs: ComparisonParagraph[], events: VersionEventView[], versions: VersionView[]) {
  const base = records(version, markup, paragraphs, events, versions)
  const proposals = await proposalReasons(directory, version, bytes, paragraphs, events, versions)
  return { ...base, reasons: [...base.reasons, ...proposals], reasonStatus: base.reasonStatus === "recorded" || proposals.some(reason => reason.kind === "proposal" && reason.text.trim()) ? "recorded" as const : "not-recorded" as const }
}

function unavailableRound(version: ClauseHistoryVersion, reason: string): ClauseHistoryRound {
  return { version, status: "unavailable", reason, paragraphs: [], reasons: [], changes: [], reasonStatus: "unavailable" }
}
function linkedMetadata(version: VersionRow, events: VersionEventView[], branch: LinkedDocumentBranch): ClauseHistoryVersion {
  return { ...versionView(version, events), linkedReturn: { exchangeId: branch.exchangeId, sourceDocumentId: branch.sourceDocumentId, sourceDocumentName: branch.sourceName, targetDocumentId: branch.targetDocumentId, baseVersionId: branch.baseVersionId } }
}
function linkedRecords(value: ReturnType<typeof records>, exchangeId: number) {
  return { ...value, reasons: value.reasons.map(reason => ({ ...reason, source: { ...reason.source, exchangeId } })), changes: value.changes.map(change => ({ ...change, source: { ...change.source, exchangeId } })) }
}

/** Read-only computed history. No backfill, capture, lease, session or model. */
export async function readClauseHistory(directory: string, name: string, request: ClauseHistoryRequest): Promise<ClauseHistoryResult> {
  const { target, history, current } = targetHistory(directory, name)
  if (request.expectedCurrentHash && request.expectedCurrentHash !== current.current.sha256) return fail(409, "history_current_changed", "The current saved version changed. Refresh history.")
  const anchor = readVersion(directory, request.anchorVersion)
  if (!anchor || anchor.document.id !== history.document!.id || anchor.document.removed_at !== null) return fail(404, "history_anchor_missing", "The acknowledged version is not available for this document.")
  if (anchor.version.removed) return fail(410, "history_anchor_removed", "The acknowledged version was removed by redaction. Its wording is no longer kept.")
  const anchorBytes = readVersionBytes(directory, anchor.version)
  if (!anchorBytes) return fail(409, "history_anchor_damaged", "The acknowledged version's retained bytes are unavailable.")
  if (anchorBytes.length > MAX_FILE) return fail(413, "history_large", "This saved version is too large to inspect as clause history.")
  let inventory: ComparisonParagraph[]
  try { inventory = await readComparisonParagraphs(anchorBytes) }
  catch { return fail(422, "history_unreadable", "The acknowledged Word version could not be read for clause history.") }
  const selected = resolveHistoryAnchor(inventory, request.paragraph)
  let comparedBytes = 0
  const rounds: ClauseHistoryRound[] = []
  for (const version of [...history.versions].sort((left, right) => left.created_at - right.created_at || left.number - right.number)) {
    const metadata = versionView(version, history.events)
    if (version.removed) { rounds.push(unavailableRound(metadata, "Removed by redaction. The wording and its recorded reasons are no longer available.")); continue }
    if (selected.status !== "verified") { rounds.push({ ...unavailableRound(metadata, selected.reason), status: "unresolved" }); continue }
    const bytes = readVersionBytes(directory, { ...version, document_id: history.document!.id })
    if (!bytes) { rounds.push(unavailableRound(metadata, "The retained version bytes are missing or damaged.")); continue }
    if (bytes.length > MAX_FILE || version.id !== anchor.version.id && comparedBytes + bytes.length + anchorBytes.length > MAX_COMPARISON_BYTES) {
      rounds.push(unavailableRound(metadata, "This version exceeds the bounded history comparison workload; its wording has not been compared.")); continue
    }
    comparedBytes += bytes.length + anchorBytes.length
    try {
      const own = version.id === anchor.version.id ? inventory : await readComparisonParagraphs(bytes)
      const relation = version.id === anchor.version.id ? selected : correspondHistoryParagraphs(await compareParagraphs(anchorBytes, bytes), selected.paragraphs)
      if (relation.status !== "verified") { rounds.push({ ...unavailableRound(metadata, relation.reason), status: "unresolved" }); continue }
      const paragraphs = relation.paragraphs.map(paragraph => own.find(candidate => candidate.anchorId === paragraph.anchorId) ?? paragraph)
      const markup = await readDocxMarkup(bytes)
      rounds.push({ version: metadata, status: "verified", paragraphs: paragraphs.map(historyParagraph), evidence: relation.evidence,
        ...await historyRecords(directory, version, bytes, markup, paragraphs, history.events, history.versions) })
    } catch { rounds.push(unavailableRound(metadata, "This Word version could not be read or compared. No clause correspondence is asserted.")) }
  }
  for (const branch of history.linked_returns) {
    const found = readVersion(directory, branch.versionId)
    if (!found) continue // No author/time is invented for missing registry data.
    const sourceHistory = readDocumentVersions(directory, branch.sourceName)
    const metadata = linkedMetadata(found.version, sourceHistory.events, branch)
    if (!branch.available) { rounds.push(unavailableRound(metadata, "This linked return's document, base or kept bytes are no longer available. Its wording and recorded reasons are not shown.")); continue }
    if (selected.status !== "verified") { rounds.push({ ...unavailableRound(metadata, selected.reason), status: "unresolved" }); continue }
    try {
      const { bytes } = readDocumentReceivedBranch(directory, branch.exchangeId)
      if (bytes.length > MAX_FILE || comparedBytes + bytes.length + anchorBytes.length > MAX_COMPARISON_BYTES) { rounds.push(unavailableRound(metadata, "This linked return exceeds the bounded history comparison workload.")); continue }
      comparedBytes += bytes.length + anchorBytes.length
      // A linked branch retains its physical document identity. A coincident
      // Word paragraph ID across those documents does not prove correspondence.
      const comparison = await compareParagraphs(anchorBytes, bytes)
      const withoutNative = (paragraph: ComparisonParagraph): ComparisonParagraph => ({ ...paragraph, nativeId: undefined })
      const relation = correspondHistoryParagraphs({ ...comparison, left: comparison.left.map(withoutNative), right: comparison.right.map(withoutNative) }, selected.paragraphs.map(withoutNative))
      if (relation.status !== "verified") { rounds.push({ ...unavailableRound(metadata, relation.reason), status: "unresolved" }); continue }
      const own = await readComparisonParagraphs(bytes)
      const paragraphs = relation.paragraphs.map(paragraph => own.find(candidate => candidate.anchorId === paragraph.anchorId) ?? paragraph)
      rounds.push({ version: metadata, status: "verified", paragraphs: paragraphs.map(historyParagraph), evidence: relation.evidence,
        ...linkedRecords(await historyRecords(directory, found.version, bytes, await readDocxMarkup(bytes), paragraphs, sourceHistory.events, sourceHistory.versions), branch.exchangeId) })
    } catch { rounds.push(unavailableRound(metadata, "This linked return could not be read or its association is no longer available.")) }
  }
  rounds.sort((left, right) => left.version.createdAt - right.version.createdAt || left.version.number - right.version.number)
  const after = targetHistory(directory, target.name)
  if (after.history.document!.id !== history.document!.id || after.current.current.id !== current.current.id || after.current.current.sha256 !== current.current.sha256)
    return fail(409, "history_current_changed", "The document changed while history loaded. Refresh history.")
  const anchorAfter = readVersion(directory, anchor.version.id)
  if (!anchorAfter || anchorAfter.version.removed || anchorAfter.document.removed_at !== null || !readVersionBytes(directory, anchorAfter.version))
    return fail(410, "history_anchor_removed", "The acknowledged version is no longer available. Refresh history.")
  // Redaction may have removed older versions while comparison was running.
  for (let index = 0; index < rounds.length; index++) {
    const round = rounds[index]!, fresh = readVersion(directory, round.version.id)
    const keptReasons = round.reasons.filter(reason => proposalSourceCurrent(directory, reason.source))
    if (keptReasons.length !== round.reasons.length) {
      round.reasons = keptReasons
      round.reasonStatus = keptReasons.some(reason => (reason.kind === "comment" || reason.kind === "proposal") && reason.text.trim()) ? "recorded" : "not-recorded"
    }
    if (round.version.linkedReturn) {
      try {
        const { branch } = readDocumentReceivedBranch(directory, round.version.linkedReturn.exchangeId)
        if (branch.targetDocumentId !== history.document!.id || branch.versionId !== round.version.id || branch.baseVersionId !== round.version.linkedReturn.baseVersionId) throw new Error("association_changed")
      } catch { rounds[index] = unavailableRound(round.version, "This linked return is no longer available through this document. Its wording and recorded reasons are not shown.") }
      continue
    }
    if (!fresh || fresh.document.id !== history.document!.id || fresh.version.removed || fresh.document.removed_at !== null || !readVersionBytes(directory, fresh.version))
      rounds[index] = unavailableRound(round.version, "This version is no longer available. Its wording and recorded reasons are not shown.")
  }
  return { document: { id: history.document!.id, name: target.name, currentVersion: current.current.id }, anchorVersion: anchor.version.id, currentVersion: current.current.id,
    anchor: { status: selected.status, ...(selected.status === "unresolved" ? { reason: selected.reason } : {}), paragraphs: selected.paragraphs.map(historyParagraph) },
    rounds, wordingView: "current-text", limitations: ["Wording shows the current-text projection of each saved version. Pending tracked edits remain proposals in the original version.",
      "Comments are recorded notes; a local decision is not evidence of external agreement. Missing reasons are not reconstructed from changed wording."] }
}

export async function readClauseHistorySource(directory: string, name: string, request: ClauseHistorySourceRequest): Promise<ClauseHistorySourceResult> {
  const { history } = targetHistory(directory, name)
  const selected = readVersion(directory, request.versionId)
  let linked: LinkedDocumentBranch | undefined
  if (request.exchangeId) {
    try { linked = readDocumentReceivedBranch(directory, request.exchangeId).branch } catch { return fail(410, "history_source_removed", "This linked history source is no longer available.") }
    if (linked.targetDocumentId !== history.document!.id || linked.versionId !== request.versionId) return fail(404, "history_source_missing", "That linked history source is not associated with this document.")
  }
  if (!selected || selected.document.removed_at !== null || selected.document.id !== history.document!.id && !linked) return fail(404, "history_source_missing", "That history source is not in this document.")
  if (selected.version.removed) return fail(410, "history_source_removed", "This source version was removed by redaction.")
  if (selected.version.sha256 !== request.hash) return fail(409, "history_source_changed", "The source version does not match this history entry.")
  const bytes = readVersionBytes(directory, selected.version)
  if (!bytes) return fail(410, "history_source_unavailable", "This history source is no longer available.")
  if (bytes.length > MAX_FILE) return fail(413, "history_source_large", "This source is too large to inspect here.")
  const markup = await readDocxMarkup(bytes)
  const paragraphs = await readComparisonParagraphs(bytes)
  const sourceHistory = linked ? readDocumentVersions(directory, selected.document.name) : history
  const recorded = await historyRecords(directory, selected.version, bytes, markup, paragraphs, sourceHistory.events, sourceHistory.versions)
  const available = linked ? linkedRecords(recorded, linked.exchangeId) : recorded
  const reason = available.reasons.find(reason => reason.source.kind === request.kind && reason.source.id === request.id)
  const change = available.changes.find(change => change.source.kind === request.kind && change.source.id === request.id)
  if (!reason && !change) return fail(404, "history_record_missing", "This recorded source could not be located in the exact saved version.")
  if (request.recordHash !== (reason?.source ?? change?.source)?.recordHash || !proposalSourceCurrent(directory, request))
    return fail(409, "history_source_changed", "This proposal or decision record changed. Refresh history before reading its source.")
  const fresh = readVersion(directory, request.versionId)
  if (!fresh || fresh.version.removed || fresh.document.removed_at !== null || !readVersionBytes(directory, fresh.version)) return fail(410, "history_source_removed", "This history source is no longer available.")
  if (linked) {
    try {
      const latest = readDocumentReceivedBranch(directory, linked.exchangeId).branch
      if (latest.targetDocumentId !== history.document!.id || latest.versionId !== request.versionId || latest.baseVersionId !== linked.baseVersionId) throw new Error("association_changed")
    } catch { return fail(410, "history_source_removed", "This linked history source is no longer available through this document.") }
  }
  return { version: linked ? linkedMetadata(selected.version, sourceHistory.events, linked) : versionView(selected.version, history.events), source: request, ...(reason ? { reason } : {}), ...(change ? { change } : {}) }
}

async function body(c: Context): Promise<unknown> {
  const text = await c.req.text()
  if (new TextEncoder().encode(text).length > 160_000) return
  try { return JSON.parse(text) } catch { return }
}
async function respond(c: Context, run: () => Promise<unknown>) {
  c.header("Cache-Control", "no-store")
  try { const result = await run(); return result instanceof Response ? result : c.json(result) }
  catch (error) {
    if (error instanceof ClauseHistoryError) return c.json({ code: error.code, error: error.message }, error.status)
    return c.json({ code: "history_unavailable", error: "Clause history could not be read. The document was not changed." }, 422)
  }
}
export function clauseHistoryRoutes(options: { directory: (id: string) => string }) {
  const app = new Hono()
  app.post("/:id/documents/clause-history", c => respond(c, async () => {
    const request = clauseHistoryRequest(await body(c))
    if (!request) return fail(400, "invalid_clause_history", "Choose one complete paragraph and its acknowledged saved version.")
    return readClauseHistory(options.directory(c.req.param("id")), c.req.query("name") ?? "", request)
  }))
  app.post("/:id/documents/clause-history/source", c => respond(c, async () => {
    const request = sourceRequest(await body(c))
    if (!request) return fail(400, "invalid_history_source", "Choose an exact saved history source.")
    return readClauseHistorySource(options.directory(c.req.param("id")), c.req.query("name") ?? "", request)
  }))
  return app
}

export function installWordClauseHistory(app: Hono, authenticate: (c: Context) => HistoryAuth | undefined) {
  for (const source of [false, true]) app.post(source ? "/word/chat/history-source" : "/word/chat/history", c => respond(c, async () => {
    const auth = authenticate(c)
    if (!auth) return c.json({ code: "word_link_revoked", error: "This Word connection is no longer valid. Reopen the document from doc.haus." }, 401)
    const request = source ? sourceRequest(await body(c)) : clauseHistoryRequest(await body(c))
    if (!request) return fail(400, "invalid_clause_history", "Choose a saved clause or exact history source.")
    const granted = readVersion(auth.matter.dir, auth.grant.issued_version)
    if (!granted || granted.document.id !== auth.grant.document_id || granted.document.name !== auth.target.name)
      return fail(404, "history_document_missing", "The linked document is no longer available.")
    const result = source ? await readClauseHistorySource(auth.matter.dir, auth.target.name, request as ClauseHistorySourceRequest)
      : await readClauseHistory(auth.matter.dir, auth.target.name, request as ClauseHistoryRequest)
    const live = authenticate(c)
    if (!live || live.matter.id !== auth.matter.id || live.grant.document_id !== auth.grant.document_id || live.grant.token_hash !== auth.grant.token_hash)
      return c.json({ code: "word_link_revoked", error: "This Word connection is no longer valid. Reopen the document from doc.haus." }, 401)
    return result
  }))
}
