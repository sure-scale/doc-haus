import { useEffect, useId, useState } from "react"
import { getConfig, patchConfig } from "../../api/opencode/config"
import { testConcurrency, type ConcurrencyTestResult } from "../../api/opencode/concurrency"
import { FieldRow } from "./controls"

const PRESETS = [1, 3, 5, 10, 20, 30, 40, 50, 60, 80, 100]
const message = (error: unknown, fallback: string) => error instanceof Error ? error.message : fallback
const failureLabel: Record<string, string> = { auth: "Sign-in error", model: "Model unavailable", rate_limit: "Provider rate limit", unavailable: "Connection unavailable", empty: "Empty response", timeout: "Timed out" }

export default function ConcurrencySettings({ onBusy }: { onBusy: (busy: boolean) => void }) {
  const inputId = useId()
  const presetsId = useId()
  const [saved, setSaved] = useState<number>()
  const [value, setValue] = useState("")
  const [loading, setLoading] = useState(true)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const [result, setResult] = useState<ConcurrencyTestResult>()
  const number = Number(value)
  const valid = value.trim() !== "" && Number.isInteger(number) && number >= 1 && number <= 100
  const changed = valid && number !== saved

  async function load() {
    setLoading(true); setError("")
    try {
      const config = await getConfig()
      const limit = config.model_concurrency ?? 4
      setSaved(limit); setValue(String(limit))
    } catch (failure) { setError(message(failure, "Could not load the request limit.")) }
    finally { setLoading(false) }
  }
  useEffect(() => { void load() }, [])

  async function persist(limit: number) {
    await patchConfig({ model_concurrency: limit })
    const readback = await getConfig()
    if (readback.model_concurrency !== limit) throw new Error("The limit could not be confirmed. Reload and try again.")
    setSaved(limit)
  }
  async function save() {
    if (!valid || !changed) return
    setBusy(true); onBusy(true); setError(""); setNotice(""); setResult(undefined)
    try { await persist(number); setNotice("Request limit saved.") }
    catch (failure) { setError(message(failure, "Could not save the request limit.")) }
    finally { setBusy(false); onBusy(false) }
  }
  async function runTest() {
    if (!valid || saved === undefined) return
    const limit = number
    setBusy(true); onBusy(true); setError(""); setNotice(changed ? "Saving request limit…" : `Testing ${limit} requests using your saved primary model…`); setResult(undefined)
    try {
      if (changed) await persist(limit)
      setNotice(`Testing ${limit} requests using your saved primary model…`)
      setResult(await testConcurrency(limit)); setNotice("")
    }
    catch (failure) { setNotice(""); setError(message(failure, "The test could not be completed.")) }
    finally { setBusy(false); onBusy(false) }
  }

  return <section className="settings-section" aria-labelledby={`${inputId}-heading`}>
    <h3 id={`${inputId}-heading`}>Concurrent AI requests</h3>
    <p className="settings-hint muted">Choose how many AI requests can run at once. Other requests wait until a place is available.</p>
    <FieldRow id={inputId} label="Request limit" hint="Choose a whole number from 1 to 100. Suggested values are available in the list.">
      <input id={inputId} type="number" min={1} max={100} step={1} list={presetsId} value={value} disabled={loading || busy || saved === undefined} aria-describedby={`${inputId}-hint`} aria-invalid={value !== "" && !valid} onChange={event => { setValue(event.target.value); setNotice(""); setResult(undefined) }} />
      <datalist id={presetsId}>{PRESETS.map(option => <option key={option} value={option} />)}</datalist>
    </FieldRow>
    {value !== "" && !valid && <p role="alert" className="settings-hint settings-warning">Enter a whole number between 1 and 100.</p>}
    <div className="settings-actions">
      <button type="button" className="btn-sm" disabled={busy || loading || !changed} onClick={() => void save()}>Save limit</button>
      <button type="button" className="btn-sm" disabled={busy || loading || !valid || saved === undefined} onClick={() => void runTest()}>{changed ? "Save & test" : "Test"} {valid ? number : ""} requests</button>
    </div>
    <p className="settings-hint muted">Testing saves the selected limit first, then sends that many short requests using your saved primary model. This may incur provider charges. Results reflect this test, not a guarantee of future capacity.</p>
    {notice && <p role="status" className="settings-hint">{notice}</p>}
    {error && <p role="alert" className="settings-hint settings-warning">{error} {saved === undefined && <button type="button" className="btn-sm" onClick={() => void load()}>Retry</button>}</p>}
    {result && <div role="status" className="settings-hint">
      <p>{result.succeeded} of {result.requested} requests succeeded; {result.failed} failed.</p>
      <p>Peak overlapping requests in doc.haus: {result.peak}. Elapsed: {(result.elapsed_ms / 1000).toFixed(1)} seconds.</p>
      <p>Model: {result.model}</p>
      {result.errors.length > 0 && <ul>{result.errors.map(({ kind, count }) => <li key={kind}>{failureLabel[kind] ?? "Other error"}: {count}</li>)}</ul>}
    </div>}
  </section>
}
