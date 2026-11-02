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
    if (!confirm("Clear usage records for every matter? This removes tool and context records and the cost of jobs run outside a conversation. Conversation cost stays with conversation history.")) return
    setClearing(true)
    try {
      await Promise.all((await listMatters()).map((matter) => clearTelemetry(matter.id)))
      setStorage(await getTelemetryStorage())
      onSaved("Usage records cleared.")
    } catch {
      onError("Could not clear usage records.")
    } finally {
      setClearing(false)
    }
  }
  if (loadError) return (
    <section className="settings-section">
      <p className="settings-hint settings-warning" role="status">Could not load statistics.</p>
      <button type="button" className="ghost" onClick={() => setLoadAttempt((attempt) => attempt + 1)}>Try again</button>
    </section>
  )
  if (!prefs || !storage) return <p className="muted" role="status">Loading statistics...</p>
  return (
    <section className="settings-section">
      <p className="settings-hint muted">
        Usage records kept on this computer: tool use, context size and the cost of jobs run outside a conversation,
        such as Prepare changes. Conversation cost is reported by your model provider and stays with conversation
        history; clearing here does not remove it.
      </p>
      <dl className="settings-summary">
        <div className="settings-summary-row">
          <dt>Stored records</dt>
          <dd>{formatBytes(storage.size_bytes)} across {storage.matter_count} matters</dd>
        </div>
        <div className="settings-summary-row">
          <dt>Oldest record</dt>
          <dd>{formatTelemetryDate(storage.oldest_ts)}</dd>
        </div>
      </dl>
      <FieldRow id="telemetry-retention" label="Keep usage records (days)" hint="How long to keep these records. Saved when you leave this field.">
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
          Clear usage records
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
