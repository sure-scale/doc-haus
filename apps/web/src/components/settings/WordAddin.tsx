import { useEffect, useRef, useState } from "react"
import { fetchWordChatPreferences, fetchWordSetup, saveWordChatPreferences, setupWordAddin, type WordChatPreferences, type WordSetupStatus } from "../../api/ingest/word"
import { humanizeError } from "../../api/ingest/errors"
import { FieldRow, SegRow } from "./controls"

export function useWordAddin() {
  const [status, setStatus] = useState<WordSetupStatus>()
  const [checking, setChecking] = useState(true)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const alive = useRef(true)
  const pending = useRef(false)
  async function refresh() {
    setChecking(true)
    try {
      const next = await fetchWordSetup()
      if (alive.current) {
        setStatus(next); setError("")
        if (next.available) window.dispatchEvent(new Event("dochaus:word-setup"))
      }
      return next
    } catch (cause) {
      if (alive.current) setError(humanizeError(cause, "Could not check Word setup."))
    } finally { if (alive.current) setChecking(false) }
  }
  useEffect(() => {
    alive.current = true
    void refresh()
    return () => { alive.current = false }
  }, [])
  async function prepare() {
    if (pending.current) return
    pending.current = true
    setBusy(true); setError("")
    try {
      const next = await setupWordAddin()
      if (alive.current) setStatus(next)
      if (!next.available) throw new Error(next.reason || "Word setup needs attention. Check its status and try again.")
      window.dispatchEvent(new Event("dochaus:word-setup"))
      return next
    } catch (cause) {
      // Native setup may have finished even if the response was lost.
      const confirmed = await fetchWordSetup().catch(() => undefined)
      if (alive.current && confirmed) setStatus(confirmed)
      if (confirmed?.available) { window.dispatchEvent(new Event("dochaus:word-setup")); return confirmed }
      if (alive.current) setError(humanizeError(cause, "Could not confirm Word setup. Check status before trying again."))
      return
    } finally {
      pending.current = false
      if (alive.current) setBusy(false)
    }
  }
  return { status, checking, busy, error, refresh, prepare }
}

export type WordAddinState = ReturnType<typeof useWordAddin>

export function WordAddinStatus({ state }: { state: WordAddinState }) {
  const { status, checking, busy, error } = state
  return <>
    {busy ? <p role="status">Setting up Word… Approve the local connection if your computer asks.</p>
      : checking ? <p role="status">Checking Word…</p>
      : error ? null
      : status?.available ? <p role="status">Ready{status.restartWord ? " — restart Word to load the add-in." : "."}</p>
      : status && <p className="settings-hint">{status.reason || "Set up Word to open and save your matter documents."}</p>}
    {error && <p role="alert" className="settings-warning">{error}</p>}
  </>
}

export function WordChatDefaults({ onBusy }: { onBusy?: (busy: boolean) => void }) {
  const [preferences, setPreferences] = useState<WordChatPreferences>()
  const [saved, setSaved] = useState<WordChatPreferences>()
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState("")
  const [message, setMessage] = useState("")
  const alive = useRef(true), pending = useRef(false)
  async function load() {
    setLoading(true); setError("")
    try {
      const value = await fetchWordChatPreferences()
      if (alive.current) { setPreferences(value); setSaved(value) }
    } catch (cause) {
      if (alive.current) setError(humanizeError(cause, "Could not load Word chat defaults."))
    } finally { if (alive.current) setLoading(false) }
  }
  useEffect(() => { alive.current = true; void load(); return () => { alive.current = false } }, [])
  useEffect(() => { onBusy?.(saving); return () => onBusy?.(false) }, [saving, onBusy])
  function change(update: Partial<WordChatPreferences>) {
    setPreferences(current => current ? { ...current, ...update } : current); setError(""); setMessage("")
  }
  async function save() {
    if (!preferences || pending.current) return
    pending.current = true; setSaving(true); setError(""); setMessage("")
    try {
      const value = await saveWordChatPreferences(preferences)
      if (alive.current) { setPreferences(value); setSaved(value); setMessage("Chat defaults saved.") }
    } catch (cause) {
      if (alive.current) setError(humanizeError(cause, "Could not save Word chat defaults. Your choices are still here; try again."))
    } finally { pending.current = false; if (alive.current) setSaving(false) }
  }
  const dirty = preferences && saved && (preferences.defaultScope !== saved.defaultScope || preferences.suggestions !== saved.suggestions)
  return <div className="settings-section">
    <h4>Chat defaults</h4>
    <p className="settings-hint muted">Starting choices for new Word chats.</p>
    {loading ? <p role="status">Loading chat defaults…</p> : preferences ? <>
      <SegRow label="Default source scope" value={preferences.defaultScope} disabled={saving}
        options={[{ value: "document", label: "This document" }, { value: "matter", label: "Whole matter" }]}
        hint="You can change source scope in each chat."
        onChange={value => change({ defaultScope: value as WordChatPreferences["defaultScope"] })} />
      <FieldRow id="word-chat-suggestions" label="Suggestions" hint="Show up to three ideas beside your document.">
        <button id="word-chat-suggestions" type="button" role="switch" aria-label="Suggestions" aria-checked={preferences.suggestions}
          aria-describedby="word-chat-suggestions-hint" className={`switch${preferences.suggestions ? " on" : ""}`} disabled={saving}
          onClick={() => change({ suggestions: !preferences.suggestions })} />
        <span className="switch-state muted">{preferences.suggestions ? "On" : "Off"}</span>
      </FieldRow>
      <div className="row settings-row"><button type="button" className="primary" disabled={saving || !dirty} onClick={() => void save()}>{saving ? "Saving…" : "Save chat defaults"}</button></div>
    </> : <button type="button" onClick={() => void load()}>Retry chat defaults</button>}
    {error && <p role="alert" className="settings-warning">{error}</p>}
    {message && <p role="status">{message}</p>}
  </div>
}

export default function WordAddin({ onBusy }: { onBusy?: (busy: boolean) => void }) {
  const state = useWordAddin()
  const { status, busy, checking, error } = state
  const [preferencesBusy, setPreferencesBusy] = useState(false)
  useEffect(() => { onBusy?.(busy || preferencesBusy); return () => onBusy?.(false) }, [busy, preferencesBusy, onBusy])
  return <section className="settings-section">
    <h3>Microsoft Word</h3>
    <p className="settings-hint muted">Work in desktop Word with your matter documents and saved versions beside you.</p>
    <WordAddinStatus state={state} />
    {status?.canSetup && !status.available && <p className="settings-hint muted">Setup installs the add-in for this account and asks your computer to trust its local connection.</p>}
    <div className="row settings-row">
      {status?.canSetup && <button type="button" className={status.available && !error ? "" : "primary"} disabled={busy || checking} onClick={() => void state.prepare()}>
        {busy ? "Setting up…" : status.configured || status.registered ? "Repair Word connection" : "Set up Word"}
      </button>}
      <button type="button" disabled={busy || checking} onClick={() => void state.refresh()}>Check status</button>
    </div>
    {status?.status === "external" && status.manifestUrl && <p><a href={status.manifestUrl} download="doc-haus-word.xml">Download Word add-in</a></p>}
    <details><summary>Opening a document in Word</summary>
      <p>In a matter, choose Open in Word. Allow your browser to open doc.haus if asked; it opens the linked document in desktop Word. For an ordinary Word file, add it to the matter first. Keep doc.haus running while using the add-in.</p>
      {status?.origin && <p className="settings-hint muted">Connection: {status.origin}</p>}
    </details>
    <WordChatDefaults onBusy={setPreferencesBusy} />
  </section>
}
