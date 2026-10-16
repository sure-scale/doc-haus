import { useEffect, useState } from "react"
import { clearTelemetry, getTelemetryStorage, type TelemetryStorage } from "../../api/ingest/telemetry"
import { getPreferences, savePreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { listMatters } from "../../api/ingest/matters"
import { formatDateTime } from "../../lib/format"
import { FieldRow } from "./controls"

export default function StatisticsTab({ onSaved, onError }: { onSaved: (text: string) => void; onError: (text: string) => void }) {
  const [prefs, setPrefs] = useState<DraftingPreferences | null>(null)
  const [retention, setRetention] = useState("")
  const [storage, setStorage] = useState<TelemetryStorage | null>(null)
  const [saving, setSaving] = useState(false)
  const [clearing, setClearing] = useState(false)
  const [loadAttempt, setLoadAttempt] = useState(0)
  const [loadError, setLoadError] = useState(false)
  useEffect(() => {
    let cancelled = false
    setLoadError(false)
    Promise.all([getPreferences(), getTelemetryStorage()])
      .then(([loadedPrefs, loadedStorage]) => {
        if (cancelled) return
        setPrefs(loadedPrefs)
        setRetention(String(loadedPrefs.telemetry.retention_days))
        setStorage(loadedStorage)
      })
      .catch(() => {
        if (!cancelled) setLoadError(true)
      })
    return () => { cancelled = true }
  }, [loadAttempt])
  async function saveRetention() {
    if (!prefs) return
    const days = Number(retention)
    if (!Number.isFinite(days) || days <= 0) {
      setRetention(String(prefs.telemetry.retention_days))
      return
    }
    setSaving(true)
    try {
      const next = await savePreferences({
        ...prefs,
        telemetry: { ...prefs.telemetry, retention_days: Math.floor(days), retention_days_source: "user" },
      })
      setPrefs(next)
      setRetention(String(next.telemetry.retention_days))
      onSaved("Statistics preferences saved.")
    } catch {
      setRetention(String(prefs.telemetry.retention_days))
      onError("Could not save statistics preferences.")
    } finally {
      setSaving(false)
    }
  }
  async function clearAllTelemetry() {
    if (!confirm("Clear telemetry for every matter? This removes stored usage and cost rows, but keeps matters and documents.")) return
    setClearing(true)
    try {
      await Promise.all((await listMatters()).map((matter) => clearTelemetry(matter.id)))
      setStorage(await getTelemetryStorage())
      onSaved("Telemetry cleared.")
    } catch {
      onError("Could not clear telemetry.")
    } finally {
      setClearing(false)
    }
  }
  if (loadError) return (
    <section className="settings-section">
      <h3>Statistics</h3>
      <p className="settings-hint settings-warning" role="status">Could not load statistics.</p>
      <button type="button" className="ghost" onClick={() => setLoadAttempt((attempt) => attempt + 1)}>Try again</button>
    </section>
  )
  if (!prefs || !storage) return <p className="muted" role="status">Loading statistics...</p>
  return (
    <section className="settings-section">
      <h3>Statistics</h3>
      <p className="settings-hint muted">
        Local usage telemetry used for cost and token reporting. These controls affect stored telemetry only; they do
        not change matter documents or drafting instructions.
      </p>
      <dl className="settings-summary">
        <div className="settings-summary-row">
          <dt>Telemetry storage</dt>
          <dd>{formatBytes(storage.size_bytes)} across {storage.matter_count} matters</dd>
        </div>
        <div className="settings-summary-row">
          <dt>Oldest telemetry row</dt>
          <dd>{formatTelemetryDate(storage.oldest_ts)}</dd>
        </div>
        <div className="settings-summary-row">
          <dt>Default model rates</dt>
          <dd><a href={storage.rate_reference_url} target="_blank" rel="noreferrer">{storage.rate_source}</a></dd>
        </div>
      </dl>
      <p className="settings-hint muted">
        Input cost estimates use the bundled catalogue, with saved rate overrides taking precedence. Rates update
        with the installed doc.haus version; they are not fetched live. Estimates exclude output, cache discounts,
        subscription fees and hosting charges.
      </p>
      {storage.unknown_models.length > 0 && (
        <p className="settings-hint settings-warning">
          No input cost estimate: {storage.unknown_models.join(", ")}. These models remain unpriced until the catalogue
          supplies rates or a complete rate override is saved.
        </p>
      )}
      <FieldRow id="telemetry-retention" label="Keep telemetry (days)" hint="How long to retain usage and cost records. Saved when you leave this field.">
        <input
          id="telemetry-retention"
          aria-describedby="telemetry-retention-hint"
          type="number"
          min="1"
          step="1"
          value={retention}
          disabled={saving}
          onChange={(e) => setRetention(e.target.value)}
          onBlur={saveRetention}
        />
      </FieldRow>
      <div className="settings-actions">
        <button className={`danger${clearing ? " btn-loading" : ""}`} disabled={clearing} onClick={clearAllTelemetry}>
          Clear all telemetry
        </button>
      </div>
    </section>
  )
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`
  return `${(n / (1024 * 1024)).toFixed(1)} MB`
}

function formatTelemetryDate(ts: number | null): string {
  if (!ts) return "None"
  return formatDateTime(ts)
}
