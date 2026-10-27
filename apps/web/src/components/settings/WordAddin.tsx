import { useEffect, useRef, useState } from "react"
import { fetchWordSetup, setupWordAddin, type WordSetupStatus } from "../../api/ingest/word"
import { humanizeError } from "../../api/ingest/errors"

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

export default function WordAddin({ onBusy }: { onBusy?: (busy: boolean) => void }) {
  const state = useWordAddin()
  const { status, busy, checking, error } = state
  useEffect(() => { onBusy?.(busy); return () => onBusy?.(false) }, [busy, onBusy])
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
      <p>In a matter, choose Open in Word and open the downloaded copy. For an ordinary Word file, add it to the matter first. Keep doc.haus running while using the add-in.</p>
      {status?.origin && <p className="settings-hint muted">Connection: {status.origin}</p>}
    </details>
  </section>
}
