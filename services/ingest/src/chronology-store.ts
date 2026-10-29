import { createHash, randomUUID } from "node:crypto"
import { existsSync, lstatSync, mkdirSync, readdirSync, realpathSync } from "node:fs"
import path from "node:path"
import { isDeepStrictEqual } from "node:util"
import { withArtifactMutationLock } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import { writeDurableJson } from "../../../dochaus/lib/durable-json"
import { readRegularFile } from "../../../dochaus/lib/regular-file"
import { verifyCitation } from "../../../dochaus/plugin/legal/citations"
import {
  chronologyCalendarDayOffset,
  supportedNameAssociationQuote,
  isSuppliedChronologyName,
  SUPPLIED_CHRONOLOGY_EVIDENCE_ISSUE,
} from "../../../dochaus/lib/chronology"
import type {
  Chronology,
  ChronologyAccount,
  ChronologyAction,
  ChronologyDate,
  ChronologyDecision,
  ChronologyEvent,
  ChronologyEvidence,
  ChronologyProposal,
} from "../../../dochaus/lib/chronology"
import { openDb } from "./db"
import { documentRegistryIdentity } from "./document-identity"
import { readDocumentTarget, registeredDocumentName } from "./registered-document"

export class ChronologyError extends Error {
  constructor(
    message: string,
    readonly code = "chronology_invalid",
    readonly status: 400 | 404 | 409 | 423 = 400,
  ) {
    super(message)
  }
}
const record = (x: unknown): x is Record<string, unknown> => !!x && typeof x === "object" && !Array.isArray(x)
const string = (x: unknown): x is string => typeof x === "string"
const text = (x: unknown, max = 100_000): x is string => string(x) && x.length <= max
const strings = (x: unknown): x is string[] => Array.isArray(x) && x.every(string)
const timestamp = (x: unknown): x is number => typeof x === "number" && Number.isFinite(x) && x >= 0
const integer = (x: unknown): x is number => Number.isSafeInteger(x) && (x as number) >= 0
const choices = (x: unknown, values: string[]) => string(x) && values.includes(x)
const bool = (x: unknown): x is boolean => typeof x === "boolean"

export function validChronologyDate(x: unknown, depth = 0): x is ChronologyDate {
  if (
    depth > 2 ||
    !record(x) ||
    !text(x.original, 2000) ||
    !choices(x.precision, ["exact", "approximate", "month", "year", "ambiguous", "relative", "undated"])
  )
    return false
  for (const k of ["start", "end", "time", "timezone", "qualifier", "documentDate"])
    if (x[k] !== undefined && !text(x[k], 2000)) return false
  if (x.rangeKind !== undefined && !choices(x.rangeKind, ["duration", "uncertainty"])) return false
  if (x.anchorFreshness !== undefined && !choices(x.anchorFreshness, ["current", "stale", "unavailable"])) return false
  if (x.anchorIssue !== undefined && !text(x.anchorIssue)) return false
  for (const field of ["start", "end"])
    if (x[field] !== undefined && !validNormalizedDate(x[field] as string)) return false
  if (["undated", "ambiguous"].includes(x.precision as string) && (x.start !== undefined || x.end !== undefined))
    return false
  if (x.end !== undefined && (x.rangeKind === undefined || x.start === undefined)) return false
  if (x.start !== undefined && x.end !== undefined && String(x.end) < String(x.start)) return false
  const dateLength = x.precision === "exact" ? 10 : x.precision === "month" ? 7 : x.precision === "year" ? 4 : undefined
  if (
    dateLength !== undefined &&
    [x.start, x.end].some((value) => value !== undefined && String(value).length !== dateLength)
  )
    return false
  if (x.relative !== undefined) {
    const r = x.relative
    if (
      !record(r) ||
      !text(r.expression, 2000) ||
      !text(r.anchorDocument, 1000) ||
      (r.anchorDate !== undefined && !text(r.anchorDate, 2000)) ||
      (r.anchorSourceId !== undefined && !text(r.anchorSourceId, 1000)) ||
      (r.anchorVersion !== undefined && !text(r.anchorVersion, 1000)) ||
      (r.offsetDays !== undefined &&
        (!Number.isSafeInteger(r.offsetDays) || Math.abs(r.offsetDays as number) > 366000)) ||
      (r.calculation !== undefined && !text(r.calculation, 2000)) ||
      (r.interpretation !== undefined && !choices(r.interpretation, ["document-stated", "legal-interpretation"]))
    )
      return false
  }
  if (
    x.correction !== undefined &&
    (!record(x.correction) ||
      !validChronologyDate(x.correction.original, depth + 1) ||
      !timestamp(x.correction.at) ||
      !string(x.correction.actor))
  )
    return false
  return true
}
export function validChronologyEvidence(x: unknown): x is ChronologyEvidence {
  return (
    record(x) &&
    text(x.id, 1000) &&
    text(x.documentName, 1000) &&
    text(x.sourceId, 1000) &&
    text(x.version, 1000) &&
    text(x.excerpt) &&
    string(x.section) &&
    integer(x.charStart) &&
    integer(x.charEnd) &&
    x.charEnd >= x.charStart &&
    bool(x.verified) &&
    choices(x.freshness, ["current", "stale", "unavailable", "unverified"]) &&
    choices(x.provenance, ["source", "supplied-chronology"]) &&
    (x.sourceKind === undefined || string(x.sourceKind)) &&
    (x.issue === undefined || string(x.issue))
  )
}
export function validChronologyAccount(x: unknown): x is ChronologyAccount {
  return (
    record(x) &&
    text(x.id, 1000) &&
    text(x.title, 5000) &&
    text(x.detail) &&
    validChronologyDate(x.date) &&
    Array.isArray(x.evidence) &&
    x.evidence.every(validChronologyEvidence) &&
    choices(x.privilege, ["cleared", "uncertain", "privileged", "not-privileged"]) &&
    choices(x.significance, ["key", "relevant", "background"]) &&
    choices(x.status, ["reported", "scheduled", "required"]) &&
    strings(x.people) &&
    (x.nameAssociations === undefined ||
      (Array.isArray(x.nameAssociations) &&
        x.nameAssociations.every(
          (association) =>
            record(association) &&
            text(association.name, 500) &&
            !!association.name.trim() &&
            text(association.canonicalName, 500) &&
            !!association.canonicalName.trim() &&
            choices(association.status, ["supported", "uncertain"]) &&
            strings(association.evidenceIds) &&
            new Set(association.evidenceIds).size === association.evidenceIds.length &&
            association.evidenceIds.length > 0 &&
            association.evidenceIds.every((id) =>
              (x.evidence as ChronologyEvidence[]).some((evidence) => evidence.id === id),
            ) &&
            (association.status !== "supported" ||
              (x.evidence as ChronologyEvidence[]).some(
                (evidence) =>
                  (association.evidenceIds as string[]).includes(evidence.id) &&
                  evidence.verified &&
                  supportedNameAssociationQuote(
                    association.name as string,
                    association.canonicalName as string,
                    evidence.excerpt,
                  ),
              )),
        ))) &&
    strings(x.gaps) &&
    (x.attribution === undefined || string(x.attribution)) &&
    (x.obligationId === undefined || string(x.obligationId))
  )
}
export function validChronologyEvent(x: unknown): x is ChronologyEvent {
  return (
    record(x) &&
    text(x.id, 1000) &&
    !!x.id &&
    text(x.title, 5000) &&
    !!x.title.trim() &&
    text(x.detail) &&
    validChronologyDate(x.date) &&
    Array.isArray(x.accounts) &&
    x.accounts.every(validChronologyAccount) &&
    choices(x.origin, ["generated", "manual"]) &&
    ["reviewed", "opened", "hidden", "key", "disputed", "lawyerEdited"].every((k) => bool(x[k])) &&
    text(x.note) &&
    timestamp(x.createdAt) &&
    timestamp(x.updatedAt) &&
    strings(x.previousIds) &&
    (x.original === undefined ||
      (record(x.original) &&
        string(x.original.title) &&
        string(x.original.detail) &&
        validChronologyDate(x.original.date))) &&
    (x.dispute === undefined || string(x.dispute)) &&
    (x.reviewedAt === undefined || timestamp(x.reviewedAt)) &&
    (x.reviewedBy === undefined || string(x.reviewedBy)) &&
    (x.supersededBy === undefined || strings(x.supersededBy))
  )
}
export function validChronologyProposal(x: unknown): x is ChronologyProposal {
  return (
    record(x) &&
    text(x.id, 1000) &&
    !!x.id &&
    choices(x.kind, ["addition", "amendment", "conflict", "removal", "merge", "split"]) &&
    strings(x.targetIds) &&
    record(x.baseEventHashes) &&
    Object.values(x.baseEventHashes).every(string) &&
    Array.isArray(x.events) &&
    x.events.every(validChronologyEvent) &&
    text(x.reason) &&
    text(x.recommendation) &&
    choices(x.status, ["pending", "deferred", "accepted", "kept"]) &&
    choices(x.origin, ["refresh", "chat", "source-date"]) &&
    timestamp(x.createdAt) &&
    (x.decidedAt === undefined || timestamp(x.decidedAt)) &&
    (x.actor === undefined || string(x.actor)) &&
    (x.sourceSetKey === undefined || string(x.sourceSetKey))
  )
}
export function validChronology(x: unknown): x is Chronology {
  if (
    !record(x) ||
    x.schemaVersion !== 1 ||
    !integer(x.revision) ||
    !bool(x.created) ||
    !Array.isArray(x.events) ||
    !x.events.every(validChronologyEvent) ||
    !Array.isArray(x.proposals) ||
    !x.proposals.every(validChronologyProposal) ||
    !record(x.sourceDates) ||
    !Array.isArray(x.coverage) ||
    !Array.isArray(x.history) ||
    !string(x.framing) ||
    !strings(x.gaps)
  )
    return false
  if (
    new Set(x.events.map((e) => e.id)).size !== x.events.length ||
    new Set(x.proposals.map((p) => p.id)).size !== x.proposals.length
  )
    return false
  if (
    !Object.values(x.sourceDates).every(
      (d) =>
        record(d) &&
        string(d.date) &&
        timestamp(d.at) &&
        string(d.actor) &&
        ((d.sourceId === undefined && d.version === undefined) ||
          (text(d.sourceId, 1000) && !!d.sourceId && text(d.version, 1000) && !!d.version)) &&
        (d.freshness === undefined || choices(d.freshness, ["current", "stale", "unavailable"])) &&
        (d.issue === undefined || string(d.issue)),
    )
  )
    return false
  if (
    !x.coverage.every(
      (c) =>
        record(c) &&
        string(c.sourceId) &&
        string(c.documentName) &&
        string(c.version) &&
        choices(c.status, [
          "pending",
          "processing",
          "complete",
          "partial",
          "failed",
          "unreadable",
          "unsupported",
          "removed",
        ]) &&
        bool(c.uploaded) &&
        bool(c.textExtracted) &&
        integer(c.sections) &&
        Array.isArray(c.processedSections) &&
        c.processedSections.every(integer) &&
        strings(c.completedKeys) &&
        integer(c.events) &&
        strings(c.issues) &&
        bool(c.suppliedChronology),
    )
  )
    return false
  if (
    !record(x.suggestion) ||
    !choices(x.suggestion.status, ["unknown", "relevant", "irrelevant", "insufficient", "hidden", "later"]) ||
    !string(x.suggestion.reason) ||
    ["assessedKey", "dismissedKey"].some(
      (k) => x.suggestion && record(x.suggestion) && x.suggestion[k] !== undefined && !string(x.suggestion[k]),
    )
  )
    return false
  if (
    !x.history.every(
      (h) =>
        record(h) &&
        string(h.id) &&
        timestamp(h.at) &&
        string(h.actor) &&
        string(h.action) &&
        strings(h.eventIds) &&
        (h.detail === undefined || string(h.detail)),
    )
  )
    return false
  if (
    x.run !== undefined &&
    (!record(x.run) ||
      !string(x.run.id) ||
      !choices(x.run.mode, ["initial", "refresh"]) ||
      !choices(x.run.status, ["running", "stopped", "failed", "complete"]) ||
      !choices(x.run.privilege, ["all-cleared", "mixed"]) ||
      !timestamp(x.run.startedAt) ||
      !strings(x.run.sourceKeys) ||
      !strings(x.run.completedKeys) ||
      !strings(x.run.errors))
  )
    return false
  return true
}

export function freshChronology(): Chronology {
  return {
    schemaVersion: 1,
    revision: 0,
    created: false,
    events: [],
    coverage: [],
    proposals: [],
    sourceDates: {},
    suggestion: { status: "unknown", reason: "" },
    history: [],
    framing: "Client framing is unknown; the account is neutral.",
    gaps: [],
  }
}
function chronologyFile(directory: string) {
  const root = realpathSync.native(directory)
  const data = path.join(root, ".dochaus")
  if (existsSync(data) && (lstatSync(data).isSymbolicLink() || !lstatSync(data).isDirectory()))
    throw new ChronologyError("The chronology folder is unsafe.", "chronology_unavailable", 423)
  const file = path.join(data, "chronology.json")
  if (existsSync(file) && (lstatSync(file).isSymbolicLink() || !lstatSync(file).isFile()))
    throw new ChronologyError("The saved chronology is unsafe.", "chronology_unavailable", 423)
  return file
}
export function readChronology(directory: string): Chronology {
  const file = chronologyFile(directory)
  const bytes = readRegularFile(file)
  if (!bytes) return freshChronology()
  let parsed: unknown
  try {
    parsed = JSON.parse(bytes.toString("utf8"))
  } catch {
    throw new ChronologyError(
      "The saved chronology is invalid. It has been preserved; restore or repair it before making changes.",
      "chronology_unavailable",
      423,
    )
  }
  if (!validChronology(parsed))
    throw new ChronologyError(
      "The saved chronology has an unsupported or invalid structure. It has been preserved.",
      "chronology_unavailable",
      423,
    )
  return parsed
}
export function mutateChronology(
  directory: string,
  expectedRevision: number | undefined,
  action: string,
  callback: (state: Chronology) => void,
  actor = "local operator",
): Chronology {
  const file = chronologyFile(directory)
  mkdirSync(path.dirname(file), { recursive: true })
  return withArtifactMutationLock(directory, ".dochaus/chronology.json", () => {
    const current = readChronology(directory)
    if (expectedRevision !== undefined && current.revision !== expectedRevision)
      throw new ChronologyError(
        "The chronology changed. Reload it before applying this decision.",
        "chronology_conflict",
        409,
      )
    if (current.revision >= Number.MAX_SAFE_INTEGER)
      throw new ChronologyError("The chronology revision limit has been reached.", "chronology_unavailable", 423)
    const next = structuredClone(current)
    callback(next)
    stripProjectedDateFreshness(next)
    if (JSON.stringify(next) === JSON.stringify(current)) return current
    next.revision = current.revision + 1
    if (!validChronology(next))
      throw new ChronologyError("The proposed chronology change is invalid; the saved chronology is unchanged.")
    if (action !== "chronology source progress")
      next.history.push({
        id: randomUUID(),
        at: Date.now(),
        actor,
        action,
        eventIds: next.events
          .filter((e) => !current.events.some((old) => old.id === e.id && JSON.stringify(old) === JSON.stringify(e)))
          .map((e) => e.id),
      })
    writeDurableJson(file, next)
    return next
  })
}
export function chronologyEventHash(event: ChronologyEvent) {
  // Changes in evidence, wording, dates or visibility invalidate a pending
  // proposal. Merely opening an event and adding a note do not change its basis.
  const { opened: _opened, note: _note, key: _key, updatedAt: _updatedAt, ...basis } = event
  return createHash("sha256").update(JSON.stringify(basis)).digest("hex")
}
export function newChronologyEvent(input: {
  title: string
  detail: string
  date: ChronologyDate
  accounts?: ChronologyAccount[]
  id?: string
  origin?: "generated" | "manual"
}): ChronologyEvent {
  const now = Date.now()
  return {
    id: input.id ?? randomUUID(),
    title: input.title,
    detail: input.detail,
    date: structuredClone(input.date),
    accounts: structuredClone(input.accounts ?? []),
    origin: input.origin ?? "generated",
    original:
      input.origin === "manual"
        ? undefined
        : { title: input.title, detail: input.detail, date: structuredClone(input.date) },
    reviewed: false,
    opened: false,
    hidden: false,
    key: false,
    disputed: false,
    note: "",
    lawyerEdited: input.origin === "manual",
    createdAt: now,
    updatedAt: now,
    previousIds: [],
  }
}

type CurrentSource = {
  sourceId: string
  documentName: string
  version: string
  docPath: string
  textExtracted: boolean
}
export function currentChronologySources(directory: string): CurrentSource[] {
  chronologyFile(directory)
  const db = openDb(directory)
  try {
    const rows = db
      .query(
        "SELECT id,name,doc_path,EXISTS(SELECT 1 FROM chunks WHERE document_id=documents.id AND length(trim(text))>0) AS text_extracted FROM documents WHERE role='source'",
      )
      .all() as {
      id: number
      name: string
      doc_path: string
      text_extracted: number
    }[]
    const registered = rows.flatMap((row) => {
      if (!registeredDocumentName(row.name)) return []
      const target = readDocumentTarget(directory, row.name)
      if (
        !target ||
        !existsSync(target.docPath) ||
        documentRegistryIdentity(row.doc_path) !== documentRegistryIdentity(target.docPath)
      )
        return []
      try {
        const bytes = readRegularFile(target.docPath)
        return bytes
          ? [
              {
                sourceId: String(row.id),
                documentName: row.name,
                docPath: target.docPath,
                version: createHash("sha256").update(bytes).digest("hex"),
                textExtracted: !!row.text_extracted,
              },
            ]
          : []
      } catch {
        return []
      }
    })
    const allNames = db.query("SELECT name FROM documents").all() as { name: string }[]
    const supplied = readdirSync(directory)
      .filter(
        (name) =>
          /^chronology\.md$/i.test(name) && !allNames.some((row) => row.name.toLowerCase() === name.toLowerCase()),
      )
      .flatMap((name) => {
        const target = readDocumentTarget(directory, name)
        if (!target) return []
        try {
          const bytes = readRegularFile(target.docPath)
          return bytes
            ? [
                {
                  sourceId: `supplied:${name}`,
                  documentName: name,
                  docPath: target.docPath,
                  version: createHash("sha256").update(bytes).digest("hex"),
                  textExtracted: false,
                },
              ]
            : []
        } catch {
          return []
        }
      })
    return [...registered, ...supplied]
  } finally {
    db.close()
  }
}
export function projectChronologySources(directory: string, state = readChronology(directory)): Chronology {
  const next = structuredClone(state)
  const sources = currentChronologySources(directory)
  for (const [documentName, correction] of Object.entries(next.sourceDates)) {
    const source = sources.find((source) => source.documentName === documentName)
    if (!source) {
      correction.freshness = "unavailable"
      correction.issue =
        "The dated source is no longer available. Reassert its date only after checking the current source."
    } else if (
      !correction.sourceId ||
      correction.sourceId !== source.sourceId ||
      correction.version !== source.version
    ) {
      correction.freshness = "stale"
      correction.issue =
        "This date belongs to an earlier or unbound source version. Check and reassert it before using it as a date anchor."
    } else {
      correction.freshness = "current"
      delete correction.issue
    }
  }
  for (const event of [...next.events, ...next.proposals.flatMap((proposal) => proposal.events)]) {
    projectAnchorDate(event.date, next, sources)
    for (const account of event.accounts) projectAnchorDate(account.date, next, sources)
  }
  for (const event of [...next.events, ...next.proposals.flatMap((p) => p.events)])
    for (const account of event.accounts)
      for (const evidence of account.evidence) {
        const source = sources.find((s) => s.sourceId === evidence.sourceId && s.documentName === evidence.documentName)
        if (!source) {
          evidence.freshness = "unavailable"
          evidence.excerpt = ""
          evidence.issue = "Source is no longer available in this matter."
        } else if (source.version !== evidence.version) {
          evidence.freshness = "stale"
          evidence.issue = "Source bytes have changed since this evidence was verified. Historical review is preserved."
        } else {
          // Byte-identical restoration restores the original evidential basis;
          // changed bytes remain stale and never renew a historical review.
          evidence.freshness = evidence.verified ? "current" : "unverified"
        }
      }
  for (const coverage of next.coverage) {
    const source = sources.find((s) => s.sourceId === coverage.sourceId && s.documentName === coverage.documentName)
    if (!source) {
      coverage.uploaded = false
      coverage.textExtracted = false
      coverage.status = "removed"
      if (!coverage.issues.includes("Source is no longer available."))
        coverage.issues.push("Source is no longer available.")
    } else if (source.version !== coverage.version) {
      coverage.uploaded = true
      coverage.textExtracted = source.textExtracted
      coverage.status = "pending"
      if (!coverage.issues.includes("Source changed; refresh required."))
        coverage.issues.push("Source changed; refresh required.")
    } else {
      coverage.uploaded = true
      coverage.textExtracted ||= source.textExtracted
    }
  }
  for (const source of sources)
    if (!next.coverage.some((c) => c.sourceId === source.sourceId && c.documentName === source.documentName))
      next.coverage.push({
        ...source,
        status: "pending",
        uploaded: true,
        textExtracted: source.textExtracted,
        sections: 0,
        processedSections: [],
        completedKeys: [],
        events: 0,
        issues: ["Added since the chronology source set was enumerated."],
        suppliedChronology: isSuppliedChronologyName(source.documentName),
      })
  return next
}

export async function verifyChronologyProposals(directory: string, proposals: ChronologyProposal[]) {
  const sources = currentChronologySources(directory),
    state = readChronology(directory)
  for (const p of proposals)
    for (const e of p.events) {
      if (e.origin !== "generated" || !e.accounts.length)
        throw new ChronologyError("AI proposals require sourced accounts. Manual additions belong in Add event.")
      for (const a of e.accounts) {
        if (provenanceCorrectionAccount(state, p, a, sources) || historicalAccount(state, p, a)) continue
        if (!a.evidence.length)
          throw new ChronologyError("Every proposed account needs verified supporting source text.")
        for (const evidence of a.evidence) {
          const source = sources.find(
            (s) => s.sourceId === evidence.sourceId && s.documentName === evidence.documentName,
          )
          if (!source || source.version !== evidence.version)
            throw new ChronologyError(
              "A proposed source changed or is unavailable. Refresh the evidence before proposing it.",
              "chronology_source_changed",
              409,
            )
          assertEvidenceProvenance(evidence)
          const citation = await verifyCitation({ ...evidence, docPath: source.docPath }).catch(() => undefined)
          if (
            !citation ||
            citation.charStart !== evidence.charStart ||
            citation.charEnd !== evidence.charEnd ||
            citation.excerpt !== evidence.excerpt
          )
            throw new ChronologyError(
              "A proposed quotation does not match its exact source passage. Repeated or displaced wording must be resolved first.",
            )
          evidence.verified = true
          evidence.freshness = "current"
        }
      }
    }
}
function assertCurrentProposalSources(directory: string, p: ChronologyProposal, state = readChronology(directory)) {
  const sources = currentChronologySources(directory)
  if (p.origin === "source-date")
    for (const event of p.events) {
      const relative = event.date.relative
      if (
        !relative ||
        state.sourceDates[relative.anchorDocument]?.date !== relative.anchorDate ||
        !sources.some(
          (source) =>
            source.documentName === relative.anchorDocument &&
            !!state.sourceDates[relative.anchorDocument]?.sourceId &&
            source.sourceId === state.sourceDates[relative.anchorDocument]?.sourceId &&
            source.version === state.sourceDates[relative.anchorDocument]?.version &&
            p.sourceSetKey === `${source.sourceId}:${source.version}`,
        )
      )
        throw new ChronologyError(
          "The source date or its source changed after this calculation was proposed. Review the latest source-date proposal.",
          "chronology_source_changed",
          409,
        )
    }
  if (p.kind === "removal" && p.sourceSetKey !== chronologySourceSetKey(directory))
    throw new ChronologyError(
      "The source set changed after this removal was proposed. Refresh before deciding whether to hide the event.",
      "chronology_source_changed",
      409,
    )
  for (const event of p.events)
    for (const account of event.accounts) {
      if (provenanceCorrectionAccount(state, p, account, sources) || historicalAccount(state, p, account)) continue
      for (const evidence of account.evidence) {
        assertEvidenceProvenance(evidence)
        if (
          !evidence.verified ||
          !sources.some(
            (s) =>
              s.sourceId === evidence.sourceId &&
              s.documentName === evidence.documentName &&
              s.version === evidence.version,
          )
        )
          throw new ChronologyError(
            "The proposal evidence changed. Refresh and review the new evidence before deciding.",
            "chronology_source_changed",
            409,
          )
      }
    }
}
export function stageChronologyProposals(
  directory: string,
  expectedRevision: number,
  proposals: ChronologyProposal[],
  origin: ChronologyProposal["origin"] = "chat",
) {
  if (!proposals.length || proposals.length > 100 || !proposals.every(validChronologyProposal))
    throw new ChronologyError("Supply valid chronology proposals.")
  return mutateChronology(
    directory,
    expectedRevision,
    "propose changes",
    (state) => {
      for (const input of proposals) {
        const old = state.proposals.find((p) => p.id === input.id)
        if (old) {
          if (proposalContent(old) !== proposalContent(input))
            throw new ChronologyError("This proposal id already identifies a different change.")
          continue
        }
        const p = structuredClone(input)
        p.origin = origin
        p.status = "pending"
        p.createdAt = Date.now()
        delete p.decidedAt
        delete p.actor
        if (
          new Set(p.targetIds).size !== p.targetIds.length ||
          p.targetIds.some((id) => !state.events.some((e) => e.id === id))
        )
          throw new ChronologyError("A proposal targets an event outside this chronology.")
        p.baseEventHashes = Object.fromEntries(
          p.targetIds.map((id) => [id, chronologyEventHash(state.events.find((e) => e.id === id)!)]),
        )
        for (const e of p.events) {
          e.reviewed = false
          delete e.reviewedAt
          delete e.reviewedBy
          e.opened = false
          e.hidden = false
          e.key = false
          e.note = ""
          e.lawyerEdited = false
          e.origin = "generated"
        }
        if (p.kind === "removal") p.sourceSetKey = chronologySourceSetKey(directory)
        validateProposalShape(p)
        assertCurrentProposalSources(directory, p, state)
        state.proposals.push(p)
      }
    },
    "AI proposal",
  )
}
function validateProposalShape(p: ChronologyProposal) {
  if (p.kind === "addition" && (p.targetIds.length !== 0 || p.events.length !== 1))
    throw new ChronologyError("An addition creates exactly one new event.")
  if (["amendment", "conflict"].includes(p.kind) && (p.targetIds.length !== 1 || p.events.length !== 1))
    throw new ChronologyError("An amendment or conflicting account changes one event.")
  if (p.kind === "removal" && (p.targetIds.length !== 1 || p.events.length !== 0))
    throw new ChronologyError("A removal must identify one saved event.")
  if (p.kind === "merge" && (p.targetIds.length < 2 || p.events.length !== 1))
    throw new ChronologyError("A merge combines at least two events.")
  if (p.kind === "split" && (p.targetIds.length !== 1 || p.events.length < 2))
    throw new ChronologyError("A split creates at least two events.")
}
export function decideChronologyProposal(
  directory: string,
  expectedRevision: number,
  id: string,
  decision: ChronologyDecision,
  batch = false,
) {
  const current = readChronology(directory)
  const old = current.proposals.find((p) => p.id === id)
  // A lost response can be recovered safely by repeating the same decision.
  if (
    old &&
    ((decision === "accept" && old.status === "accepted") ||
      (decision === "keep" && old.status === "kept") ||
      (decision === "defer" && old.status === "deferred"))
  )
    return current
  return mutateChronology(directory, expectedRevision, `proposal ${decision}`, (state) =>
    applyProposalDecision(directory, state, id, decision, batch),
  )
}
function applyProposalDecision(
  directory: string,
  state: Chronology,
  id: string,
  decision: ChronologyDecision,
  batch = false,
) {
  const p = state.proposals.find((p) => p.id === id)
  if (!p) throw new ChronologyError("Proposal not found.", "chronology_not_found", 404)
  if (["accepted", "kept"].includes(p.status))
    throw new ChronologyError(
      "This proposal has already been decided. Reload the chronology.",
      "chronology_conflict",
      409,
    )
  if (decision === "defer") {
    p.status = "deferred"
    return
  }
  if (decision === "keep") {
    p.status = "kept"
    p.decidedAt = Date.now()
    p.actor = "local operator"
    return
  }
  validateProposalShape(p)
  for (const target of p.targetIds) {
    const event = state.events.find((e) => e.id === target)
    if (!event || chronologyEventHash(event) !== p.baseEventHashes[target])
      throw new ChronologyError(
        "The current event changed after this proposal. Keep it or refresh before accepting.",
        "chronology_conflict",
        409,
      )
  }
  if (batch && (p.kind !== "addition" || p.events.some((e) => e.disputed)))
    throw new ChronologyError("Only straightforward additions can be accepted together.")
  assertCurrentProposalSources(directory, p, state)
  const targets = p.targetIds.map((id) => state.events.find((e) => e.id === id)!)
  if (p.kind === "merge")
    p.events[0]!.accounts = preserveAccounts(
      targets.flatMap((e) => e.accounts),
      p.events[0]!.accounts,
    )
  if (
    p.kind === "split" &&
    targets[0]!.accounts.some(
      (a) =>
        !p.events.some((e) =>
          e.accounts.some((candidate) => candidate.id === a.id && accountMaterial(candidate) === accountMaterial(a)),
        ),
    )
  )
    throw new ChronologyError("A split must preserve every original source account in its resulting events.")
  if (p.kind === "removal") targets[0]!.hidden = true
  else if (p.kind === "amendment" || p.kind === "conflict") {
    const old = targets[0]!,
      proposed = p.events[0]!
    const sources = currentChronologySources(directory)
    const correctedAccounts = old.accounts.map((original) => {
      const candidate = proposed.accounts.find(
        (account) => account.id === original.id && isProvenanceOnlyCorrection(original, account, sources),
      )
      return candidate ?? original
    })
    const accounts = [
      ...correctedAccounts,
      ...proposed.accounts
        .filter((a) => !correctedAccounts.some((b) => a.id === b.id && accountMaterial(a) === accountMaterial(b)))
        .map((a) =>
          correctedAccounts.some((b) => a.id === b.id) ? { ...a, id: stableId(a.id, accountMaterial(a)) } : a,
        ),
    ]
    Object.assign(old, {
      title: proposed.title,
      detail: proposed.detail,
      date: structuredClone(proposed.date),
      accounts: structuredClone(accounts),
      disputed: p.kind === "conflict" || old.disputed,
      updatedAt: Date.now(),
      reviewed: true,
      reviewedAt: Date.now(),
      reviewedBy: "local operator",
    })
  } else {
    const resulting = p.events.map((e) => ({
      ...structuredClone(e),
      id: p.kind === "addition" ? e.id : stableId(p.id, e.id),
      reviewed: !batch,
      reviewedAt: batch ? undefined : Date.now(),
      reviewedBy: batch ? undefined : "local operator",
      previousIds: [...new Set([...e.previousIds, ...p.targetIds])],
      updatedAt: Date.now(),
    }))
    for (const e of resulting) {
      if (state.events.some((old) => old.id === e.id))
        throw new ChronologyError("The proposed event identity is already in use.")
      if (targets.length) {
        e.note = targets
          .map((t) => t.note)
          .filter(Boolean)
          .join("\n\n")
        e.key = targets.some((t) => t.key)
        e.hidden = targets.every((t) => t.hidden)
      }
    }
    if (p.kind === "merge" || p.kind === "split")
      for (const target of targets) {
        target.hidden = true
        target.supersededBy = resulting.map((e) => e.id)
      }
    state.events.push(...resulting)
  }
  p.status = "accepted"
  p.decidedAt = Date.now()
  p.actor = "local operator"
  state.created = true
}
export function batchChronologyProposals(directory: string, expectedRevision: number, ids: string[]) {
  const before = readChronology(directory)
  if (
    ids.length &&
    ids.every((id) => before.proposals.some((p) => p.id === id && p.kind === "addition" && p.status === "accepted"))
  )
    return before
  if (!ids.length || ids.length > 100 || new Set(ids).size !== ids.length)
    throw new ChronologyError("Choose distinct additions to accept together.")
  return mutateChronology(directory, expectedRevision, "batch additions", (state) => {
    for (const id of ids) applyProposalDecision(directory, state, id, "accept", true)
  })
}
const stableId = (...parts: string[]) => createHash("sha256").update(parts.join("\0")).digest("hex").slice(0, 24)

function actionEvent(state: Chronology, id: string) {
  const event = state.events.find((e) => e.id === id)
  if (!event) throw new ChronologyError("Event not found in this matter.", "chronology_not_found", 404)
  return event
}
function eventInput(
  input: unknown,
): input is Record<string, unknown> & { title: string; detail: string; date: ChronologyDate } {
  return (
    record(input) &&
    text(input.title, 5000) &&
    !!input.title.trim() &&
    text(input.detail) &&
    validChronologyDate(input.date)
  )
}
export function validChronologyAction(a: unknown): a is ChronologyAction {
  if (!record(a) || !string(a.type)) return false
  if (a.type === "add") return eventInput(a)
  if (a.type === "source-date")
    return (
      string(a.documentName) &&
      /^\d{4}-\d{2}-\d{2}$/.test(String(a.date)) &&
      !Number.isNaN(Date.parse(String(a.date))) &&
      new Date(String(a.date)).toISOString().slice(0, 10) === a.date
    )
  if (a.type === "invitation") return choices(a.choice, ["hide", "later", "show"])
  if (a.type === "discard-refresh") return true
  if (a.type === "merge")
    return (
      strings(a.eventIds) &&
      a.eventIds.length >= 2 &&
      new Set(a.eventIds).size === a.eventIds.length &&
      eventInput(a) &&
      (a.preview === undefined || bool(a.preview))
    )
  if (!string(a.eventId)) return false
  if (a.type === "split")
    return (
      Array.isArray(a.parts) &&
      a.parts.length >= 2 &&
      a.parts.length <= 100 &&
      a.parts.every((p) => eventInput(p) && record(p) && strings(p.accountIds)) &&
      (a.preview === undefined || bool(a.preview))
    )
  if (a.type === "open") return true
  if (a.type === "review") return bool(a.reviewed)
  if (a.type === "edit") return eventInput(a)
  if (a.type === "note") return text(a.note)
  if (a.type === "key") return bool(a.key)
  if (a.type === "hide") return bool(a.hidden)
  if (a.type === "dispute") return bool(a.disputed) && text(a.dispute)
  return false
}
function transformationPreview(
  state: Chronology,
  action: Extract<ChronologyAction, { type: "merge" | "split" }>,
): ChronologyEvent[] {
  if (action.type === "merge") {
    const originals = action.eventIds.map((id) => actionEvent(state, id))
    if (originals.some((e) => e.supersededBy?.length))
      throw new ChronologyError("This event has already been merged or split. Choose its current entry.")
    const accounts = [...new Map(originals.flatMap((e) => e.accounts).map((a) => [a.id, a])).values()]
    const next = newChronologyEvent({
      title: action.title,
      detail: action.detail,
      date: action.date,
      accounts,
      origin: accounts.length ? "generated" : "manual",
      id: stableId("merge", ...action.eventIds, JSON.stringify([action.title, action.detail, action.date])),
    })
    next.previousIds = [...new Set(originals.flatMap((e) => [e.id, ...e.previousIds]))]
    next.note = originals
      .map((e) => e.note)
      .filter(Boolean)
      .join("\n\n")
    next.key = originals.some((e) => e.key)
    next.hidden = originals.every((e) => e.hidden)
    next.lawyerEdited = true
    next.disputed = originals.some((e) => e.disputed)
    return [next]
  }
  const original = actionEvent(state, action.eventId)
  if (original.supersededBy?.length)
    throw new ChronologyError("This event has already been merged or split. Choose its current entry.")
  const all = action.parts.flatMap((p) => p.accountIds)
  if (
    original.accounts.some((a) => !all.includes(a.id)) ||
    all.some((id) => !original.accounts.some((a) => a.id === id))
  )
    throw new ChronologyError("Every source account must remain attached to at least one resulting event.")
  return action.parts.map((p, index) => {
    const accounts = original.accounts.filter((a) => p.accountIds.includes(a.id))
    const next = newChronologyEvent({
      ...p,
      accounts,
      origin: accounts.length ? "generated" : "manual",
      id: stableId("split", original.id, String(index), JSON.stringify(p)),
    })
    return {
      ...next,
      previousIds: [original.id, ...original.previousIds],
      note: original.note,
      key: original.key,
      hidden: original.hidden,
      lawyerEdited: true,
      disputed: original.disputed,
    }
  })
}
export function actOnChronology(
  directory: string,
  expectedRevision: number,
  action: ChronologyAction,
): Chronology | { preview: ChronologyEvent[] } {
  if (!validChronologyAction(action)) throw new ChronologyError("Invalid chronology action.")
  if ((action.type === "merge" || action.type === "split") && action.preview) {
    const current = readChronology(directory)
    if (current.revision !== expectedRevision)
      throw new ChronologyError("Reload before previewing this change.", "chronology_conflict", 409)
    return { preview: transformationPreview(current, action) }
  }
  return mutateChronology(directory, expectedRevision, action.type, (state) => {
    if (action.type === "invitation") {
      state.suggestion = {
        ...state.suggestion,
        status: action.choice === "hide" ? "hidden" : action.choice === "later" ? "later" : "relevant",
        dismissedKey: action.choice === "show" ? undefined : state.suggestion.assessedKey,
      }
      return
    }
    if (action.type === "discard-refresh") {
      if (state.run?.status === "running")
        throw new ChronologyError("Stop the refresh before discarding its proposals.", "chronology_conflict", 409)
      state.proposals = state.proposals.filter(
        (p) => p.origin !== "refresh" || p.status === "accepted" || p.status === "kept",
      )
      return
    }
    if (action.type === "source-date") {
      const source = currentChronologySources(directory).find((s) => s.documentName === action.documentName)
      if (!source) throw new ChronologyError("Choose a source currently available in this matter.")
      state.sourceDates[action.documentName] = {
        date: action.date,
        at: Date.now(),
        actor: "local operator",
        sourceId: source.sourceId,
        version: source.version,
      }
      for (const event of state.events) {
        const relative = event.date.relative
        if (
          !relative ||
          relative.anchorDocument !== action.documentName ||
          relative.offsetDays === undefined ||
          relative.interpretation === "legal-interpretation"
        )
          continue
        const offset = chronologyCalendarDayOffset(relative)
        if (offset === undefined) continue
        const anchor = Date.parse(`${action.date}T00:00:00Z`)
        const result = new Date(anchor + offset * 86400000).toISOString().slice(0, 10)
        if (
          event.date.start === result &&
          relative.anchorDate === action.date &&
          relative.anchorSourceId === source.sourceId &&
          relative.anchorVersion === source.version
        )
          continue
        const next = structuredClone(event)
        next.reviewed = false
        delete next.reviewedAt
        delete next.reviewedBy
        next.date = {
          ...event.date,
          start: result,
          relative: {
            ...relative,
            anchorDate: action.date,
            anchorSourceId: source.sourceId,
            anchorVersion: source.version,
            calculation: `${action.date} + ${relative.offsetDays} calendar days (lawyer-supplied source date)`,
          },
        }
        const p: ChronologyProposal = {
          id: stableId("source-date", action.documentName, action.date, source.sourceId, source.version, event.id),
          sourceSetKey: `${source.sourceId}:${source.version}`,
          kind: "amendment",
          targetIds: [event.id],
          baseEventHashes: { [event.id]: chronologyEventHash(event) },
          events: [next],
          reason:
            "A lawyer supplied or corrected the source date supporting this relative event date. Confirm the affected date calculation.",
          recommendation: "Check the anchor, original expression and calculation before accepting.",
          status: "pending",
          origin: "source-date",
          createdAt: Date.now(),
        }
        if (!state.proposals.some((old) => old.id === p.id)) state.proposals.push(p)
      }
      return
    }
    if (action.type === "add") {
      const next = newChronologyEvent({ ...action, origin: "manual" })
      next.opened = true
      state.events.push(next)
      state.created = true
      return
    }
    if (action.type === "merge" || action.type === "split") {
      const next = transformationPreview(state, action)
      if (next.some((n) => state.events.some((e) => e.id === n.id)))
        throw new ChronologyError(
          "This transformation has already been applied. Reload to see its resulting events.",
          "chronology_conflict",
          409,
        )
      for (const id of action.type === "merge" ? action.eventIds : [action.eventId]) {
        const original = actionEvent(state, id)
        original.hidden = true
        original.supersededBy = next.map((e) => e.id)
      }
      state.events.push(...next)
      return
    }
    const event = actionEvent(state, action.eventId)
    if (action.type === "open") {
      event.opened = true
      return
    }
    if (action.type === "review") {
      event.reviewed = action.reviewed
      event.reviewedAt = action.reviewed ? Date.now() : undefined
      event.reviewedBy = action.reviewed ? "local operator" : undefined
      return
    }
    if (action.type === "edit") {
      event.original ??= { title: event.title, detail: event.detail, date: structuredClone(event.date) }
      const correction =
        JSON.stringify(event.date) !== JSON.stringify(action.date)
          ? {
              original: structuredClone(event.date.correction?.original ?? event.date),
              at: Date.now(),
              actor: "local operator",
            }
          : event.date.correction
      event.title = action.title
      event.detail = action.detail
      event.date = { ...structuredClone(action.date), ...(correction ? { correction } : {}) }
      event.lawyerEdited = true
      event.reviewed = false
      delete event.reviewedAt
      delete event.reviewedBy
    }
    if (action.type === "note") event.note = action.note
    if (action.type === "key") event.key = action.key
    if (action.type === "hide") event.hidden = action.hidden
    if (action.type === "dispute") {
      event.disputed = action.disputed
      event.dispute = action.dispute
    }
    event.updatedAt = Date.now()
  })
}

/** Accept only content from chat; review and decision metadata is application owned. */
export function chronologyProposalDrafts(value: unknown, state?: Chronology): ChronologyProposal[] {
  if (!Array.isArray(value) || !value.length || value.length > 100)
    throw new ChronologyError("Supply one or more bounded change proposals.")
  return value.map((input) => {
    if (
      !record(input) ||
      !choices(input.kind, ["addition", "amendment", "conflict", "removal", "merge", "split"]) ||
      !strings(input.targetIds) ||
      !Array.isArray(input.events) ||
      !text(input.reason) ||
      !text(input.recommendation)
    )
      throw new ChronologyError("Invalid proposed change.")
    const targetIds = input.targetIds
    const events = input.events.map((e: unknown) => {
      if (
        !eventInput(e) ||
        !record(e) ||
        (e.accounts !== undefined && (!Array.isArray(e.accounts) || !e.accounts.every(validChronologyAccount))) ||
        (e.accountIds !== undefined && (!strings(e.accountIds) || new Set(e.accountIds).size !== e.accountIds.length))
      )
        throw new ChronologyError("Proposed events require wording, qualified date and sourced accounts.")
      const references = (e.accountIds ?? []) as string[]
      const historical = references.map((id) => {
        const candidates =
          state?.events
            .filter((event) => targetIds.includes(event.id))
            .flatMap((event) => event.accounts)
            .filter((account) => account.id === id) ?? []
        if (
          !candidates.length ||
          candidates.some((candidate) => accountMaterial(candidate) !== accountMaterial(candidates[0]!))
        )
          throw new ChronologyError(
            "A referenced source account is not uniquely available on the targeted events in this matter.",
          )
        return structuredClone(candidates[0]!)
      })
      const accounts = preserveAccounts(historical, (e.accounts ?? []) as ChronologyAccount[])
      if (!accounts.length)
        throw new ChronologyError(
          "Every proposed event needs a verified source account or a preserved account reference.",
        )
      return newChronologyEvent({
        ...e,
        id: string(e.id) ? e.id : stableId("proposed-event", JSON.stringify([e.title, e.detail, e.date, accounts])),
        accounts,
        origin: "generated",
      })
    })
    return {
      id:
        string(input.id) && input.id
          ? input.id
          : stableId(
              "proposal",
              JSON.stringify([
                input.kind,
                input.targetIds,
                events.map((e) => e.id),
                input.reason,
                input.recommendation,
              ]),
            ),
      kind: input.kind as ChronologyProposal["kind"],
      targetIds: input.targetIds,
      events,
      baseEventHashes: {},
      reason: input.reason,
      recommendation: input.recommendation,
      status: "pending",
      origin: "chat",
      createdAt: Date.now(),
    }
  })
}

function proposalContent(p: ChronologyProposal) {
  return JSON.stringify([
    p.kind,
    p.targetIds,
    p.events.map((e) => [e.title, e.detail, e.date, e.accounts]),
    p.reason,
    p.recommendation,
  ])
}

function historicalAccount(state: Chronology, p: ChronologyProposal, account: ChronologyAccount) {
  return p.targetIds.some((id) =>
    state.events
      .find((e) => e.id === id)
      ?.accounts.some((a) => a.id === account.id && accountMaterial(a) === accountMaterial(account)),
  )
}

function accountMaterial(a: ChronologyAccount) {
  return JSON.stringify({ ...a, evidence: a.evidence.map(({ freshness: _f, issue: _i, ...e }) => e) })
}

export function chronologySourceSetKey(directory: string) {
  return createHash("sha256")
    .update(
      JSON.stringify(
        currentChronologySources(directory)
          .map((s) => [s.sourceId, s.documentName, s.version])
          .sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
      ),
    )
    .digest("hex")
}

function preserveAccounts(old: ChronologyAccount[], proposed: ChronologyAccount[]) {
  return [
    ...old,
    ...proposed
      .filter((a) => !old.some((b) => a.id === b.id && accountMaterial(a) === accountMaterial(b)))
      .map((a) => (old.some((b) => a.id === b.id) ? { ...a, id: stableId(a.id, accountMaterial(a)) } : a)),
  ]
}

function validNormalizedDate(value: string) {
  if (!/^\d{4}(?:-\d{2}(?:-\d{2})?)?$/.test(value)) return false
  if (value.length === 4) return true
  const month = Number(value.slice(5, 7))
  if (month < 1 || month > 12) return false
  if (value.length === 7) return true
  const date = new Date(`${value}T00:00:00Z`)
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value
}

function projectAnchorDate(date: ChronologyDate, state: Chronology, sources: CurrentSource[]) {
  delete date.anchorFreshness
  delete date.anchorIssue
  const relative = date.relative
  if (!relative) return
  const correction = state.sourceDates[relative.anchorDocument]
  if (!correction && !relative.anchorSourceId && !relative.anchorVersion) return
  const current = sources.find((source) => source.documentName === relative.anchorDocument)
  if (!current) {
    date.anchorFreshness = "unavailable"
    date.anchorIssue =
      "The source supporting this relative-date anchor is unavailable. The saved date and historical review are preserved."
    return
  }
  if (
    !relative.anchorSourceId ||
    relative.anchorSourceId !== current.sourceId ||
    relative.anchorVersion !== current.version
  ) {
    date.anchorFreshness = "stale"
    date.anchorIssue =
      "The relative-date anchor belongs to an earlier or unbound source version. Review a fresh date proposal before relying on this calculation."
    return
  }
  if (correction && (correction.freshness !== "current" || correction.date !== relative.anchorDate)) {
    date.anchorFreshness = "stale"
    date.anchorIssue =
      "The lawyer-supplied source date changed or needs reassertion. Confirm the affected relative-date proposal."
    return
  }
  date.anchorFreshness = "current"
}
function stripProjectedDateFreshness(state: Chronology) {
  const strip = (date: ChronologyDate) => {
    delete date.anchorFreshness
    delete date.anchorIssue
    if (date.correction) strip(date.correction.original)
  }
  for (const event of [...state.events, ...state.proposals.flatMap((proposal) => proposal.events)]) {
    strip(event.date)
    if (event.original) strip(event.original.date)
    for (const account of event.accounts) strip(account.date)
  }
  for (const correction of Object.values(state.sourceDates)) {
    delete correction.freshness
    delete correction.issue
  }
}

function assertEvidenceProvenance(evidence: ChronologyEvidence) {
  const expected = isSuppliedChronologyName(evidence.documentName) ? "supplied-chronology" : "source"
  if (evidence.provenance !== expected)
    throw new ChronologyError("The proposed evidence provenance does not match the authorised source classification.")
}
function provenanceCorrectionAccount(
  state: Chronology,
  proposal: ChronologyProposal,
  proposed: ChronologyAccount,
  sources: CurrentSource[],
) {
  if (proposal.kind !== "amendment") return false
  return proposal.targetIds.some((id) =>
    state.events
      .find((event) => event.id === id)
      ?.accounts.some((original) => isProvenanceOnlyCorrection(original, proposed, sources)),
  )
}
/** Reclassifying an existing passage is metadata, not a competing report.
 * Only current authorised source identity and the shared classifier can justify
 * this exact delta; substantive account edits retain append/history semantics. */
function isProvenanceOnlyCorrection(
  original: ChronologyAccount,
  proposed: ChronologyAccount,
  sources: CurrentSource[],
): boolean {
  if (original.id !== proposed.id || original.evidence.length !== proposed.evidence.length) return false
  const basis = (account: ChronologyAccount) => ({
    ...account,
    evidence: account.evidence.map(({ provenance: _p, issue: _i, freshness: _f, ...evidence }) => evidence),
  })
  if (!isDeepStrictEqual(basis(original), basis(proposed))) return false
  let changed = false
  for (let index = 0; index < original.evidence.length; index++) {
    const before = original.evidence[index]!,
      after = proposed.evidence[index]!
    if (before.provenance === after.provenance && before.issue === after.issue) continue
    changed = true
    const current = sources.find(
      (source) =>
        source.sourceId === after.sourceId &&
        source.documentName === after.documentName &&
        source.version === after.version,
    )
    const expected = isSuppliedChronologyName(after.documentName) ? "supplied-chronology" : "source"
    if (!current)
      throw new ChronologyError(
        "The provenance correction evidence changed. Refresh its current source before deciding.",
        "chronology_source_changed",
        409,
      )
    if (after.provenance !== expected)
      throw new ChronologyError("The proposed evidence provenance does not match the authorised source classification.")
    const issueUnchanged = before.issue === after.issue
    const issueAdded =
      expected === "supplied-chronology" &&
      before.issue === undefined &&
      after.issue === SUPPLIED_CHRONOLOGY_EVIDENCE_ISSUE
    const issueRemoved =
      expected === "source" && before.issue === SUPPLIED_CHRONOLOGY_EVIDENCE_ISSUE && after.issue === undefined
    if (!issueUnchanged && !issueAdded && !issueRemoved)
      throw new ChronologyError("A provenance correction must preserve unrelated evidence issues.")
  }
  return changed
}
