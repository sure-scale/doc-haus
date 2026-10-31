import path from "node:path"
import { randomUUID } from "node:crypto"
import { withArtifactMutationLocksAsync } from "../../../dochaus/capabilities/artifacts/mutation-lock"
import type { ChronologyCoverage } from "../../../dochaus/lib/chronology"
import {
  CHRONOLOGY_EXPORT_LABEL,
  type ChronologyExportRequest,
  type ChronologyExportSnapshot,
} from "../../../dochaus/lib/chronology-export"
import { readRegularFile } from "../../../dochaus/lib/regular-file"
import { ChronologyError, projectChronologySources, readChronology } from "./chronology-store"
import { writeDeliverable } from "./deliverable"

/** Re-read evidence availability for every preview/download. Never serialize the
 * browser's cached state or cached excerpts from a changed/redacted source. */
export function chronologyExportSnapshot(
  directory: string,
  matterId: string,
  input: ChronologyExportRequest,
  now = Date.now(),
): ChronologyExportSnapshot {
  const saved = readChronology(directory)
  if (saved.revision !== input.expectedRevision)
    throw new ChronologyError(
      "The chronology changed. Refresh the export preview before downloading.",
      "chronology_conflict",
      409,
    )
  const state = projectChronologySources(directory, saved)
  const current = state.events.filter((event) => !event.supersededBy?.length)
  const byId = new Map(current.map((event) => [event.id, event]))
  const selected =
    input.scope === "filtered"
      ? input.eventIds!.map((id) => {
          const event = byId.get(id)
          if (!event)
            throw new ChronologyError(
              "An export entry is no longer in the saved chronology. Refresh the preview.",
              "chronology_conflict",
              409,
            )
          if (event.hidden && !input.includeHidden)
            throw new ChronologyError("Choose to include hidden entries explicitly before exporting them.")
          return event
        })
      : current.filter((event) => input.includeHidden || !event.hidden)
  selected.sort(
    (a, b) =>
      (a.date.start ? (b.date.start ? a.date.start.localeCompare(b.date.start) : -1) : b.date.start ? 1 : 0) ||
      a.id.localeCompare(b.id),
  )
  if (!selected.length) throw new ChronologyError("There are no saved entries in this export scope.")
  let omittedExcerpts = 0
  for (const event of selected)
    for (const account of event.accounts)
      for (const evidence of account.evidence) {
        if (evidence.freshness === "current" && evidence.verified) continue
        if (evidence.excerpt) omittedExcerpts++
        evidence.excerpt = ""
        evidence.issue =
          evidence.freshness === "stale"
            ? "Historical excerpt omitted because the source changed. Check the current source."
            : evidence.issue ||
              "Excerpt omitted because current source access and quotation verification are unavailable."
      }
  const pendingProposals = state.proposals.filter((proposal) =>
    ["pending", "deferred"].includes(proposal.status),
  ).length
  const warnings: string[] = []
  if (state.run?.status === "running")
    warnings.push("Chronology building is still in progress; this snapshot contains only entries already saved.")
  if (state.run?.status === "failed" || state.run?.status === "stopped")
    warnings.push(
      `Chronology building ${state.run.status === "failed" ? "failed" : "was stopped"}; saved entries may be incomplete.`,
    )
  for (const error of new Set(state.run?.errors ?? [])) warnings.push(`Source processing issue: ${error}`)
  for (const gap of new Set(state.gaps)) warnings.push(`Chronology qualification: ${gap}`)
  if (state.coverage.some((source) => source.status !== "complete" || source.issues.length))
    warnings.push("Source coverage is incomplete or needs checking. This snapshot is not a complete record of events.")
  if (pendingProposals)
    warnings.push(
      `${pendingProposals} pending or deferred proposal${pendingProposals === 1 ? " is" : "s are"} excluded from this snapshot.`,
    )
  if (
    selected.some((event) => event.accounts.some((account) => ["uncertain", "privileged"].includes(account.privilege)))
  )
    warnings.push(
      "Some source accounts have privilege qualifications. Review them before using or sharing this snapshot.",
    )
  if (omittedExcerpts)
    warnings.push(
      `${omittedExcerpts} cached source excerpt${omittedExcerpts === 1 ? " was" : "s were"} omitted because current support could not be verified.`,
    )
  warnings.push(
    "A verified quotation is a text match, not proof that an event occurred. This is a dated snapshot, not a live synchronized record.",
  )
  let title = matterId
  try {
    const metadata = JSON.parse(readRegularFile(path.join(directory, "matter.json"))!.toString("utf8"))
    if (typeof metadata.title === "string" && metadata.title.trim()) title = metadata.title.trim()
  } catch {
    throw new ChronologyError("The matter details could not be read. Reload the matter before exporting.")
  }
  const selectedIds = new Set(selected.map((event) => event.id))
  return {
    schemaVersion: 1,
    matterId,
    title,
    snapshotAt: now,
    revision: saved.revision,
    scope: input.scope,
    includeHidden: input.includeHidden,
    events: selected,
    // Source projection can include internal file paths in newly enumerated
    // coverage rows. Export only the public coverage contract.
    coverage: state.coverage.map(publicCoverage),
    totalEvents: current.length,
    excludedHidden: current.filter((event) => event.hidden && !selectedIds.has(event.id)).length,
    pendingProposals,
    warnings,
  }
}

function publicCoverage(source: ChronologyCoverage): ChronologyCoverage {
  return {
    sourceId: source.sourceId,
    documentName: source.documentName,
    version: source.version,
    status: source.status,
    uploaded: source.uploaded,
    textExtracted: source.textExtracted,
    sections: source.sections,
    processedSections: source.processedSections,
    completedKeys: source.completedKeys,
    events: source.events,
    issues: source.issues,
    suppliedChronology: source.suppliedChronology,
  }
}

/** The existing source and chronology locks keep this short serialization and
 * write on the same source-access view. The writer owns atomic publication and
 * version capture; no second state store or export receipt is needed. */
export async function writeChronologyExport(
  directory: string,
  matterId: string,
  input: ChronologyExportRequest,
  serialize: (snapshot: ChronologyExportSnapshot) => Promise<Buffer>,
  actor?: string,
) {
  const now = Date.now()
  const stamp = new Date(now).toISOString().replace(/[-:.]/g, "").slice(0, 15)
  const name = `Chronology-${stamp}-${randomUUID().slice(0, 8)}.docx`
  return withArtifactMutationLocksAsync(
    directory,
    [name, ".dochaus/chronology.json", ".dochaus/source-corpus"],
    async () => {
      const snapshot = chronologyExportSnapshot(directory, matterId, input, now)
      const buffer = await serialize(snapshot)
      const written = await writeDeliverable(directory, name, buffer, {
        precondition: { kind: "create" },
        version: { origin: "generated", actor, label: CHRONOLOGY_EXPORT_LABEL },
      })
      // The deliverable writer adds stable paragraph pointers. Return the exact
      // canonical bytes whose hash was captured in document_versions.
      const bytes = readRegularFile(written.docPath)
      if (!bytes) throw new Error("The saved chronology export could not be read.")
      return { ...written, bytes, revision: snapshot.revision }
    },
  )
}
