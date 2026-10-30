import { forwardRef, useEffect, useId, useImperativeHandle, useState } from "react"
import { getConfig, patchConfig } from "../../api/opencode/config"
import { testConcurrency, type ConcurrencyTestResult } from "../../api/opencode/concurrency"
import { FieldRow } from "./controls"
import { DEFAULT_MODEL_CONCURRENCY, MODEL_CONCURRENCY_PRESETS as PRESETS, nearestModelConcurrency } from "../../../../../dochaus/lib/model-concurrency-presets"

const message = (error: unknown, fallback: string) => error instanceof Error ? error.message : fallback
const failureLabel: Record<string, string> = { auth: "Sign-in error", model: "Model unavailable", rate_limit: "Provider rate limit", unavailable: "Connection unavailable", empty: "Empty response", timeout: "Timed out" }

export type ConcurrencySettingsHandle = { save: () => Promise<void> }

const ConcurrencySettings = forwardRef<ConcurrencySettingsHandle, { onBusy: (busy: boolean) => void; onDirty: (dirty: boolean) => void; disabled?: boolean }>(function ConcurrencySettings({ onBusy, onDirty, disabled = false }, ref) {
  const inputId = useId()
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
  const unavailable = loading || busy || disabled || saved === undefined
  const previous = valid ? [...PRESETS].reverse().find(option => option < number) : undefined
  const next = valid ? PRESETS.find(option => option > number) : undefined

  function setLimit(limit: string) {
    setValue(limit); setNotice(""); setResult(undefined)
  }

  async function load() {
    setLoading(true); setError("")
    try {
      const config = await getConfig()
      const limit = nearestModelConcurrency(config.model_concurrency ?? DEFAULT_MODEL_CONCURRENCY)
      setSaved(limit); setValue(String(limit))
    } catch (failure) { setError(message(failure, "Could not load the request limit.")) }
    finally { setLoading(false) }
  }
  useEffect(() => { void load() }, [])
  useEffect(() => { onDirty(changed) }, [changed, onDirty])

  async function persist(limit: number) {
    await patchConfig({ model_concurrency: limit })
    const readback = await getConfig()
    if (readback.model_concurrency !== limit) throw new Error("The limit could not be confirmed. Reload and try again.")
    setSaved(limit)
  }
  async function save() {
    if (!valid || !changed) return
    setBusy(true); onBusy(true); setError(""); setNotice(""); setResult(undefined)
    try { await persist(number) }
    catch (failure) { setError(message(failure, "Could not save the request limit.")); throw failure }
    finally { setBusy(false); onBusy(false) }
  }
  useImperativeHandle(ref, () => ({ save }))
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
    <FieldRow id={inputId} label="Request limit">
      <div className="settings-concurrency-slider">
        <output htmlFor={inputId}>{valid ? number : "—"}</output>
        <input id={inputId} type="range" min={0} max={PRESETS.length - 1} step={1} value={valid ? PRESETS.findIndex(option => option === number) : 0} disabled={unavailable} aria-valuetext={valid ? `${number} requests` : "Loading"} onChange={event => setLimit(String(PRESETS[Number(event.target.value)]))} onKeyDown={event => {
          if (["ArrowUp", "ArrowRight", "ArrowDown", "ArrowLeft", "Home", "End"].includes(event.key)) {
            event.preventDefault()
            const target = event.key === "Home" ? PRESETS[0] : event.key === "End" ? PRESETS[PRESETS.length - 1] : event.key === "ArrowUp" || event.key === "ArrowRight" ? next : previous
            if (target !== undefined) setLimit(String(target))
          }
        }} />
        <div className="settings-concurrency-stops" aria-hidden="true">{PRESETS.map((option, index) => <span key={option} style={{ left: `${index * 100 / (PRESETS.length - 1)}%` }}>{option}</span>)}</div>
      </div>
    </FieldRow>
    {value !== "" && !valid && <p role="alert" className="settings-hint settings-warning">Enter a whole number between 1 and 100.</p>}
    <div className="settings-actions">
      <button type="button" className="btn-sm" disabled={unavailable || !valid} onClick={() => void runTest()}>Test request capacity</button>
    </div>
    {notice && <p role="status" className="settings-hint">{notice}</p>}
    {error && <p role="alert" className="settings-hint settings-warning">{error} {saved === undefined && <button type="button" className="btn-sm" onClick={() => void load()}>Retry</button>}</p>}
    {result && <div role="status" className="settings-hint">
      <p>{result.succeeded} of {result.requested} requests succeeded; {result.failed} failed.</p>
      <p>Peak overlapping requests in doc.haus: {result.peak}. Elapsed: {(result.elapsed_ms / 1000).toFixed(1)} seconds.</p>
      <p>Model: {result.model}</p>
      {result.errors.length > 0 && <ul>{result.errors.map(({ kind, count }) => <li key={kind}>{failureLabel[kind] ?? "Other error"}: {count}</li>)}</ul>}
    </div>}
  </section>
})

export default ConcurrencySettings
