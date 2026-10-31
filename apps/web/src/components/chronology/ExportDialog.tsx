import { useEffect, useMemo, useState } from "react"
import {
  ChronologyConflict,
  exportChronology,
  getChronology,
  previewChronologyExport,
  type Chronology,
  type ChronologyExportRequest,
  type ChronologyExportSnapshot,
} from "../../api/ingest/chronology"
import ChronologyDialog from "./Dialog"
import { eventEvidence } from "./model"
import {
  chronologyExportDate,
  chronologyExportQualifications,
  chronologyExportEventNotes,
  chronologyExportCoverageNotes,
} from "../../../../../dochaus/lib/chronology-export-format"
import "./export.css"

export default function ChronologyExportDialog({
  matterId,
  revision,
  filteredIds,
  filteredIncludingHiddenIds,
  onState,
  onDocumentsChanged,
  onClose,
}: {
  matterId: string
  revision: number
  filteredIds: string[]
  filteredIncludingHiddenIds: string[]
  onState: (state: Chronology) => void
  onDocumentsChanged?: () => void
  onClose: () => void
}) {
  const [scope, setScope] = useState<ChronologyExportRequest["scope"]>("all-visible")
  const [includeHidden, setIncludeHidden] = useState(false)
  const [snapshot, setSnapshot] = useState<ChronologyExportSnapshot>()
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState("")
  const [saved, setSaved] = useState("")
  const [savedRevision, setSavedRevision] = useState(0)
  const [refresh, setRefresh] = useState(0)
  const idsKey = JSON.stringify(includeHidden ? filteredIncludingHiddenIds : filteredIds)
  const selection = useMemo<ChronologyExportRequest>(
    () => ({
      expectedRevision: revision,
      scope,
      includeHidden,
      ...(scope === "filtered" ? { eventIds: JSON.parse(idsKey) as string[] } : {}),
    }),
    [revision, scope, includeHidden, idsKey],
  )
  useEffect(() => {
    let active = true
    setLoading(true)
    setSnapshot(undefined)
    setSaved("")
    setError("")
    previewChronologyExport(matterId, selection)
      .then(
        (next) => {
          if (active) setSnapshot(next)
        },
        (failure) => {
          if (active) setError(failure instanceof Error ? failure.message : "Could not preview the export.")
        },
      )
      .finally(() => {
        if (active) setLoading(false)
      })
    return () => {
      active = false
    }
  }, [matterId, selection, refresh])
  async function reload() {
    setError("")
    setLoading(true)
    try {
      onState(await getChronology(matterId))
      setRefresh((value) => value + 1)
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : "Could not refresh the chronology.")
      setLoading(false)
    }
  }
  async function save() {
    if (saving || loading || !snapshot) return
    setSaving(true)
    setError("")
    try {
      const { blob, name } = await exportChronology(matterId, selection)
      const url = URL.createObjectURL(blob)
      const link = document.createElement("a")
      link.href = url
      link.download = name
      link.click()
      setTimeout(() => URL.revokeObjectURL(url), 60_000)
      setSaved(name)
      setSavedRevision(selection.expectedRevision)
      onDocumentsChanged?.()
    } catch (failure) {
      if (failure instanceof ChronologyConflict) {
        setSnapshot(undefined)
        if (failure.current) onState(failure.current)
      }
      setError(failure instanceof Error ? failure.message : "Could not save the export.")
    } finally {
      setSaving(false)
    }
  }
  return (
    <ChronologyDialog
      title="Export chronology to Word"
      onClose={() => {
        if (!saving) onClose()
      }}
      footer={
        <div className="chronology-form-actions chronology-export-actions">
          <button type="button" disabled={saving} onClick={onClose}>
            Close
          </button>
          <button
            type="button"
            className="primary"
            disabled={loading || saving || !snapshot || (!!saved && savedRevision === snapshot.revision)}
            onClick={() => void save()}
          >
            {saving
              ? "Saving…"
              : saved && savedRevision === snapshot?.revision
                ? "Saved to matter"
                : "Save and download Word"}
          </button>
        </div>
      }
    >
      <p>
        A dated Word snapshot, saved to this matter as derived work. The compact table is followed by full evidence
        notes and source coverage.
      </p>
      <fieldset disabled={saving}>
        <legend>Entries to export</legend>
        <label className="chronology-check">
          <input
            type="radio"
            name="chronology-export-scope"
            checked={scope === "all-visible"}
            onChange={() => setScope("all-visible")}
          />
          All visible entries
        </label>
        <label className="chronology-check">
          <input
            type="radio"
            name="chronology-export-scope"
            checked={scope === "filtered"}
            onChange={() => setScope("filtered")}
          />
          Current filtered selection
        </label>
        <label className="chronology-check">
          <input type="checkbox" checked={includeHidden} onChange={(event) => setIncludeHidden(event.target.checked)} />
          Include hidden entries explicitly
        </label>
        <p className="chronology-muted">
          Filters select entries, not individual accounts. Each selected entry retains all its accounts and warnings.
          Export order is earliest first, with undated entries last.
        </p>
      </fieldset>
      {loading && <p role="status">Preparing preview…</p>}
      {error && (
        <div className="chronology-error" role="alert">
          <span>{error}</span>
          <button type="button" disabled={saving || loading} onClick={() => void reload()}>
            Refresh preview
          </button>
        </div>
      )}
      {saved && (
        <p role="status">
          Saved {saved} from chronology revision {savedRevision} to the matter and started the download.
        </p>
      )}
      {snapshot && (
        <section aria-label="Word export preview" className="chronology-export-preview">
          <h3>Preview</h3>
          <p>
            {snapshot.events.length} of {snapshot.totalEvents} saved entries · {snapshot.excludedHidden} hidden{" "}
            {snapshot.excludedHidden === 1 ? "entry" : "entries"} excluded · revision {snapshot.revision}
          </p>
          <p className="chronology-muted">
            Snapshot {new Date(snapshot.snapshotAt).toLocaleString()}. Source access is checked again when saving;
            unavailable or stale excerpts are omitted.
          </p>
          <ul className="chronology-export-warnings">
            {snapshot.warnings.map((warning) => (
              <li key={warning}>{warning}</li>
            ))}
          </ul>
          <div
            className="chronology-export-table-scroll"
            tabIndex={0}
            role="region"
            aria-label="Exported chronology entries"
          >
            <table className="chronology-export-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Event</th>
                  <th>Source</th>
                </tr>
              </thead>
              <tbody>
                {snapshot.events.map((event, i) => (
                  <tr key={event.id}>
                    <td>{chronologyExportDate(event.date)}</td>
                    <td>
                      {i + 1}. {event.title}
                      <small>{chronologyExportQualifications(event).join(" · ")}</small>
                    </td>
                    <td>
                      {[
                        ...new Set(
                          eventEvidence(event).map(
                            (evidence) => `${evidence.documentName}${evidence.section ? ` — ${evidence.section}` : ""}`,
                          ),
                        ),
                      ].join("; ") || "Source not supplied"}
                      <small>Evidence notes {i + 1}</small>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <h3>Evidence notes</h3>
          {snapshot.events.map((event, i) => (
            <details key={event.id}>
              <summary>
                {i + 1}. {event.title}
              </summary>
              <div className="chronology-export-account">
                {chronologyExportEventNotes(event).map((note, index) => (
                  <p key={index}>{note.bold ? <strong>{note.text}</strong> : note.text}</p>
                ))}
              </div>
            </details>
          ))}
          <details>
            <summary>
              Source coverage ({snapshot.coverage.length} {snapshot.coverage.length === 1 ? "source" : "sources"})
            </summary>
            <p>
              Coverage includes matter sources with no selected entry. Complete processing does not establish
              completeness or accuracy of events.
            </p>
            {snapshot.coverage.map((source) => (
              <div key={source.sourceId} className="chronology-export-account">
                {chronologyExportCoverageNotes(source).map((note, i) => (
                  <p key={i}>{note.bold ? <strong>{note.text}</strong> : note.text}</p>
                ))}
              </div>
            ))}
          </details>
        </section>
      )}
    </ChronologyDialog>
  )
}
