import { useCallback, useEffect, useRef, useState } from "react"
import { listReusableSignIns, reuseSignIn, type ReusableSource } from "../../api/ingest/credential-reuse"

export type MatchingSignInCandidate = {
  sourceId: string
  connectionId: string
  integrationID: string
  label: string
  authType: "key" | "oauth"
  sourceLabel: string
}

export function useMatchingSignIns(integrationID: string) {
  const request = useRef(0)
  const [result, setResult] = useState<{ integrationID: string; candidates: MatchingSignInCandidate[]; loading: boolean; error: string }>({ integrationID, candidates: [], loading: true, error: "" })
  const retry = useCallback(async () => {
    const current = ++request.current
    setResult({ integrationID, candidates: [], loading: true, error: "" })
    try {
      const sources = await listReusableSignIns(integrationID)
      const candidates = sources.filter(source => source.status === "ready").flatMap(source => source.connections
        .filter(connection => connection.integrationID === integrationID)
        .map(connection => ({ sourceId: source.id, connectionId: connection.id, integrationID: connection.integrationID, label: connection.label, authType: connection.authType, sourceLabel: source.label })))
      if (current === request.current) setResult({ integrationID, candidates, loading: false, error: "" })
    } catch {
      if (current === request.current) setResult({ integrationID, candidates: [], loading: false, error: "Could not check for a saved sign-in for this provider. You can retry or connect directly." })
    }
  }, [integrationID])
  useEffect(() => {
    void retry()
    return () => { request.current++ }
  }, [retry])
  return { candidates: result.integrationID === integrationID ? result.candidates : [], loading: result.integrationID !== integrationID || result.loading, error: result.integrationID === integrationID ? result.error : "", retry }
}

export function MatchingSignIn({ candidates, providerName, onReused }: {
  candidates: MatchingSignInCandidate[]
  providerName: string
  onReused: (integrationID: string) => Promise<void>
}) {
  const [choice, setChoice] = useState("")
  const [pending, setPending] = useState(false)
  const [error, setError] = useState("")
  const [message, setMessage] = useState("")
  const selected = candidates.find(candidate => JSON.stringify([candidate.sourceId, candidate.connectionId]) === choice)
  if (!candidates.length) return null
  return <div className="onboarding-saved-sign-in">
    <p className="settings-hint muted">Choose a saved {providerName} sign-in to copy into this installation. Model access is checked separately.</p>
    <label>Saved sign-in<select style={{ width: "100%", minWidth: 0, maxWidth: "100%" }} value={choice} disabled={pending} onChange={event => { setChoice(event.target.value); setError(""); setMessage("") }}>
      <option value="">Choose a saved sign-in</option>
      {candidates.map(candidate => <option key={JSON.stringify([candidate.sourceId, candidate.connectionId])} value={JSON.stringify([candidate.sourceId, candidate.connectionId])}>{candidate.label} · {candidate.authType === "oauth" ? "OAuth" : "API key"} ({candidate.sourceLabel})</option>)}
    </select></label>
    <button className="primary" disabled={!selected || pending} onClick={async () => {
      if (!selected || pending) return
      setPending(true); setError(""); setMessage("")
      try {
        const copied = await reuseSignIn(selected.sourceId, selected.connectionId, selected.integrationID)
        setChoice(""); setMessage("The selected sign-in was copied. Model access has not been tested.")
        try { await onReused(copied.integrationID) }
        catch { setError("The sign-in was copied, but the provider catalog could not refresh. Refresh the provider catalog before continuing.") }
      } catch { setError("The selected sign-in could not be copied or checked. Refresh saved sign-ins before retrying, or connect your provider directly.") }
      finally { setPending(false) }
    }}>{pending ? "Copying selected sign-in…" : "Copy selected sign-in"}</button>
    {error && <p role="alert" className="settings-warning">{error}</p>}
    {message && <p role="status">{message}</p>}
  </div>
}

export function ExistingSignInDisclosure({ onReused }: { onReused: (integrationID: string) => Promise<void> }) {
  const [open, setOpen] = useState(false)
  return <details className="settings-section" onToggle={event => setOpen(event.currentTarget.open)}>
    <summary>Reuse a saved OpenCode sign-in (optional)</summary>
    {open && <ExistingSignIn onReused={onReused} />}
  </details>
}

export default function ExistingSignIn({ onReused }: { onReused: (integrationID: string) => Promise<void> }) {
  const [sources, setSources] = useState<ReusableSource[]>([])
  const [choice, setChoice] = useState("")
  const [loading, setLoading] = useState(true)
  const [pending, setPending] = useState(false)
  const [error, setError] = useState("")
  const [message, setMessage] = useState("")
  async function discover() {
    setLoading(true); setError("")
    try { setSources(await listReusableSignIns()) }
    catch { setError("Could not discover existing sign-ins. Retry, or connect a provider.") }
    finally { setLoading(false) }
  }
  useEffect(() => { void discover() }, [])
  const candidates = sources.filter(source => source.status === "ready").flatMap(source => source.connections.map(connection => ({ source, connection, value: JSON.stringify([source.id, connection.id]) })))
  const unavailable = sources.some(source => source.status === "unavailable")
  const skipped = sources.reduce((count, source) => count + source.unavailableCount, 0)
  const selected = candidates.find(candidate => candidate.value === choice)
  const providerNames = [...new Set(candidates.map(candidate => providerName(candidate.connection.integrationID)))]
  return <section className="settings-section">
    <h3>Saved OpenCode sign-ins</h3>
    <p className="settings-hint muted">Optionally copy one sign-in from another OpenCode installation into doc.haus. Environment credentials work directly without copying. Model access is checked separately.</p>
    {loading ? <p role="status">Looking for saved sign-ins…</p> : <>
      {candidates.length > 0 ? <>
        <p role="status">Saved sign-ins found for {providerNames.join(", ")}. Choose one to copy; model access has not been tested.</p>
        <label>Saved sign-in<select style={{ width: "100%", minWidth: 0, maxWidth: "100%" }} value={choice} disabled={pending} onChange={event => { setChoice(event.target.value); setMessage(""); setError("") }}>
          <option value="">Choose a saved sign-in</option>
          {candidates.map(({ source, connection, value }) => <option key={value} value={value}>{providerName(connection.integrationID)} · {connection.label} · {connection.authType === "oauth" ? "OAuth" : "API key"} ({source.label})</option>)}
        </select></label>
        <button className="primary" disabled={!selected || pending} onClick={async () => {
          if (!selected || pending) return
          setPending(true); setError(""); setMessage("")
          try {
            const copied = await reuseSignIn(selected.source.id, selected.connection.id)
            setChoice(""); setMessage("The selected sign-in was copied. Model access has not been tested.")
            try { await onReused(copied.integrationID) }
            catch { setError("The sign-in was copied, but the catalog could not refresh. Use Refresh provider catalog below.") }
          } catch { setError("The selected sign-in could not be copied or checked. Refresh the provider catalog before retrying, or connect a provider.") }
          finally { setPending(false) }
        }}>{pending ? "Copying selected sign-in…" : "Copy selected sign-in"}</button>
      </> : !error && !unavailable && !skipped && <p className="settings-hint muted">No saved OpenCode sign-ins were found. Connect a provider.</p>}
      {sources.filter(source => source.status === "unavailable").map(source => <p key={source.id} className="settings-warning">{candidates.length ? "Some other OpenCode stores could not be read. " : "OpenCode sign-in discovery is incomplete. "}{source.label} is unavailable or uses an unsupported format.{candidates.length ? " The saved sign-ins listed above are still available to choose." : " You can connect a provider."}</p>)}
      {skipped > 0 && <p className="settings-hint muted">{skipped} {candidates.length ? "other " : ""}saved {skipped === 1 ? "sign-in could" : "sign-ins could"} not be offered for reuse.{candidates.length ? " The sign-ins listed above are still available." : " You can connect those providers."}</p>}
      {error && <p role="alert" className="settings-warning">{error}</p>}
      {(error || unavailable || skipped > 0) && <button disabled={pending} onClick={() => void discover()}>Retry sign-in discovery</button>}
      {message && <p role="status">{message}</p>}
    </>}
  </section>
}

function providerName(id: string) {
  return ({ "opencode-go": "OpenCode Go", opencode: "OpenCode Zen", "zai-coding-plan": "Z.AI Coding Plan", zai: "Z.AI", openrouter: "OpenRouter", openai: "OpenAI", google: "Google AI", azure: "Azure" } as Record<string, string>)[id] ?? id
}
