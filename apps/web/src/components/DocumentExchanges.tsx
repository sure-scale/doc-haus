import { useEffect, useRef, useState } from "react"
import { correctReceivedBase, correctReceivedDocument, downloadForSending, exchangePayloadUrl, fetchDocumentExchanges, type DocumentExchange } from "../api/ingest/exchanges"
import { fetchDocumentVersions, type DocumentVersion, type DocumentVersions } from "../api/ingest/versions"
import { getMatter } from "../api/ingest/matters"
import { formatDateTime } from "../lib/format"
import useModalFocus from "./useModalFocus"
import { getPreferences } from "../api/ingest/preferences"
import "./DocumentExchanges.css"

export default function DocumentExchanges({ matterId, name, versions, refresh }: { matterId: string; name: string; versions: DocumentVersion[]; refresh?: () => void }) {
  const [events, setEvents] = useState<DocumentExchange[]>([])
  const [error, setError] = useState<string>()
  const [sending, setSending] = useState(false)
  const [correcting, setCorrecting] = useState<DocumentExchange>()
  const [linking, setLinking] = useState<DocumentExchange>()
  const [round, setRound] = useState(0)
  useEffect(() => {
    let cancelled = false
    setEvents([]); setError(undefined)
    fetchDocumentExchanges(matterId, name).then((value) => { if (!cancelled) { setEvents(value); setError(undefined) } }, (cause) => { if (!cancelled) setError(cause instanceof Error ? cause.message : String(cause)) })
    return () => { cancelled = true }
  }, [matterId, name, versions, round])
  const available = versions.filter((version) => !version.removed)
  const versionNumber = (id: string) => versions.find((version) => version.id === id)?.number
  return <section className="version-exchanges" aria-label="Send and receive history">
    <h3>Sent and received</h3>
    <p>Downloads for internal review do not count as sends.</p>
    <button type="button" disabled={!available.length} onClick={() => setSending(true)}>Record a send</button>
    {error && <p role="alert">{error}</p>}
    {!events.length && !error && <p>No sends or receipts recorded.</p>}
    <ol>{events.map((event) => <li key={event.id}>
      <strong>{event.kind === "sent" ? `Sent · Draft ${event.draftNumber}` : "Received"}</strong>
      <p>{event.linkedBranch ? `Linked branch · ${event.linkedBranch.sourceName} · Version ${event.linkedBranch.versionNumber}` : `Version ${versionNumber(event.versionId) ?? "no longer available"}`}{event.party ? ` · ${event.party}` : " · Party not recorded"}</p>
      <p>{formatDateTime(event.occurredAt)}</p>
      {(event.association?.baseVersionId ?? event.baseVersionId) && <p>{event.kind === "sent" ? "Blackline against" : "Returned from"} {event.linkedBranch ? `${event.linkedBranch.targetName} · ` : ""}version {event.linkedBranch?.baseVersionNumber ?? versionNumber(event.association?.baseVersionId ?? event.baseVersionId!) ?? "no longer available"}</p>}
      {event.match && <p>{event.match.reason}</p>}
      {event.source && <p>From {event.source.documentName ?? "Email source"}{event.source.sender ? ` · ${event.source.sender}` : ""}{event.source.date ? ` · ${event.source.date}` : ""}</p>}
      {event.baseCorrections?.map(correction => <p key={correction.requestId}>Base corrected to version {versionNumber(correction.toVersionId) ?? "no longer available"} · {formatDateTime(correction.recordedAt)}</p>)}
      {event.kind === "received" && !event.association && available.some(version => version.id === event.versionId) && <button type="button" onClick={() => setCorrecting(event)}>Correct base version</button>}
      {event.kind === "received" && (event.linkedBranch?.available || !event.association && available.some(version => version.id === event.versionId)) && <button type="button" onClick={() => setLinking(event)}>Correct linked document</button>}
      {event.payload && versions.some((version) => version.id === event.versionId && !version.removed) && (!event.baseVersionId || versions.some((version) => version.id === event.baseVersionId && !version.removed)) && <a href={exchangePayloadUrl(matterId, event.id)} download>Download exact sent copy</a>}
    </li>)}</ol>
    {sending && <SendDialog key={`${matterId}:${name}`} matterId={matterId} name={name} versions={available} onClose={() => setSending(false)} onRecorded={() => { setRound((value) => value + 1); refresh?.() }} />}
    {correcting && <CorrectBaseDialog key={correcting.id} matterId={matterId} exchange={correcting} versions={available} onClose={() => setCorrecting(undefined)} onRecorded={() => { setRound(value => value + 1); refresh?.() }} />}
    {linking && <LinkReturnDialog key={linking.id} matterId={matterId} exchange={linking} onClose={() => setLinking(undefined)} onRecorded={() => { setRound(value => value + 1); refresh?.() }} />}
  </section>
}

function LinkReturnDialog({ matterId, exchange, onClose, onRecorded }: { matterId: string; exchange: DocumentExchange; onClose: () => void; onRecorded: () => void }) {
  const dialog = useRef<HTMLDialogElement>(null), alive = useRef(true), requestId = useRef(crypto.randomUUID())
  const [names, setNames] = useState<string[]>([]), [name, setName] = useState(""), [versions, setVersions] = useState<DocumentVersions>(), [base, setBase] = useState("")
  const [busy, setBusy] = useState(false), [error, setError] = useState<string>()
  useModalFocus(dialog)
  useEffect(() => {
    alive.current = true; dialog.current?.showModal()
    getMatter(matterId).then(matter => { if (alive.current) setNames(matter.documents.filter(document => /\.docx$/i.test(document.name)).map(document => document.name)) }, cause => { if (alive.current) setError(cause instanceof Error ? cause.message : String(cause)) })
    return () => { alive.current = false; dialog.current?.close() }
  }, [matterId])
  useEffect(() => {
    let cancelled = false; setVersions(undefined); setBase("")
    if (name) fetchDocumentVersions(matterId, name).then(value => {
      if (cancelled) return
      setVersions(value); setBase(value.versions.find(version => !version.removed && version.id !== exchange.versionId)?.id ?? "")
    }, cause => { if (!cancelled) setError(cause instanceof Error ? cause.message : String(cause)) })
    return () => { cancelled = true }
  }, [matterId, name, exchange.versionId])
  async function save() {
    if (!versions?.document || !base) return
    setBusy(true); setError(undefined)
    try {
      await correctReceivedDocument(matterId, exchange.id, { targetDocumentId: versions.document.id, baseVersionId: base,
        expectedDocumentId: exchange.association?.documentId ?? exchange.documentId, expectedBaseVersionId: exchange.association?.baseVersionId ?? exchange.baseVersionId, requestId: requestId.current })
      if (!alive.current) return
      onRecorded(); onClose()
    } catch (cause) { if (alive.current) setError(cause instanceof Error ? cause.message : String(cause)) }
    finally { if (alive.current) setBusy(false) }
  }
  return <dialog ref={dialog} className="upload-review-dialog" aria-label="Correct linked returned document" onCancel={event => { event.preventDefault(); event.stopPropagation(); if (!busy) onClose() }}>
    <div className="upload-review-content"><h2>Correct linked document</h2><p>Link this received copy to the document and version it came from. It will appear there as a linked branch. Its original file, retained bytes and both working copies stay as saved.</p>
      <label>Document<select value={name} disabled={busy} onChange={event => { setName(event.target.value); requestId.current = crypto.randomUUID() }}><option value="">Choose a document</option>{names.map(item => <option key={item} value={item}>{item}</option>)}</select></label>
      {versions && <label>Base version<select value={base} disabled={busy} onChange={event => { setBase(event.target.value); requestId.current = crypto.randomUUID() }}>{versions.versions.filter(version => !version.removed && version.id !== exchange.versionId).map(version => <option key={version.id} value={version.id}>Version {version.number}</option>)}</select></label>}
      {error && <p role="alert">{error}</p>}
    </div><div className="upload-review-actions"><button type="button" disabled={busy} onClick={onClose}>Cancel</button><button type="button" className="primary" disabled={busy || !versions?.document || !base} onClick={save}>{busy ? "Linking…" : "Confirm document link"}</button></div>
  </dialog>
}

function CorrectBaseDialog({ matterId, exchange, versions, onClose, onRecorded }: { matterId: string; exchange: DocumentExchange; versions: DocumentVersion[]; onClose: () => void; onRecorded: () => void }) {
  const dialog = useRef<HTMLDialogElement>(null), alive = useRef(true), requestId = useRef(crypto.randomUUID())
  const [base, setBase] = useState(exchange.baseVersionId ?? ""), [busy, setBusy] = useState(false), [error, setError] = useState<string>()
  useModalFocus(dialog)
  useEffect(() => { alive.current = true; dialog.current?.showModal(); return () => { alive.current = false; dialog.current?.close() } }, [])
  async function save() {
    setBusy(true); setError(undefined)
    try {
      await correctReceivedBase(matterId, exchange.id, { baseVersionId: base, expectedBaseVersionId: exchange.baseVersionId, requestId: requestId.current })
      if (!alive.current) return
      onRecorded(); onClose()
    } catch (cause) { if (alive.current) setError(cause instanceof Error ? cause.message : String(cause)) }
    finally { if (alive.current) setBusy(false) }
  }
  return <dialog ref={dialog} className="upload-review-dialog" aria-label="Correct returned base" onCancel={event => { event.preventDefault(); event.stopPropagation(); if (!busy) onClose() }}>
    <div className="upload-review-content"><h2>Correct returned base</h2><p>Choose the version this returned file was made from. This changes its recorded relationship. The retained file bytes and the current working copy stay as saved.</p>
      <label>Base version<select value={base} disabled={busy} onChange={event => { setBase(event.target.value); requestId.current = crypto.randomUUID() }}>
        <option value="">Choose a base</option>{versions.filter(version => version.id !== exchange.versionId).map(version => <option key={version.id} value={version.id}>Version {version.number}</option>)}
      </select></label>{error && <p role="alert">{error}</p>}
    </div><div className="upload-review-actions"><button type="button" disabled={busy} onClick={onClose}>Cancel</button><button type="button" className="primary" disabled={busy || !base || base === exchange.baseVersionId} onClick={save}>{busy ? "Correcting…" : "Confirm base correction"}</button></div>
  </dialog>
}

export function SendDialog({ matterId, name, versions, onClose, onRecorded, initialVersion }: { matterId: string; name: string; versions: DocumentVersion[]; onClose: () => void; onRecorded: () => void; initialVersion?: string }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const [version, setVersion] = useState(initialVersion ?? versions.find((item) => item.current)?.id ?? versions[0]?.id ?? "")
  const [party, setParty] = useState("")
  const [marker, setMarker] = useState(true)
  const [preferencesReady, setPreferencesReady] = useState(false)
  const [base, setBase] = useState("")
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string>()
  const requestId = useRef(crypto.randomUUID())
  const alive = useRef(true)
  useModalFocus(dialog)
  useEffect(() => { alive.current = true; dialog.current?.showModal(); return () => { alive.current = false; dialog.current?.close() } }, [])
  useEffect(() => {
    let cancelled = false
    getPreferences().then(prefs => { if (!cancelled) { setMarker(prefs.documentMatchingId); setPreferencesReady(true) } }, () => { if (!cancelled) setError("The matching ID preference could not be loaded. Close and reopen this dialog to try again.") })
    return () => { cancelled = true }
  }, [])
  async function send() {
    setBusy(true); setError(undefined)
    try {
      const result = await downloadForSending(matterId, version, { requestId: requestId.current, party: party.trim() || undefined, marker, baseVersionId: base || undefined })
      if (!alive.current) return
      const url = URL.createObjectURL(new Blob([result.bytes as BlobPart], { type: "application/octet-stream" }))
      const link = document.createElement("a"); link.href = url; link.download = name; link.click(); URL.revokeObjectURL(url)
      onRecorded(); onClose()
    } catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)) }
    finally { setBusy(false) }
  }
  return <dialog ref={dialog} className="upload-review-dialog" aria-label="Record a send" onCancel={(event) => { event.preventDefault(); event.stopPropagation(); if (!busy) onClose() }}>
    <div className="upload-review-content">
      <h2>Record a send</h2><p>{name}</p>
      <p>Use this when sending the file to the other side. It records the exact copy and increments Draft N. Downloading it here does not email or share it.</p>
      <label>Version<select value={version} disabled={busy} onChange={(event) => { setVersion(event.target.value); setBase(""); requestId.current = crypto.randomUUID() }}>{versions.map((item) => <option key={item.id} value={item.id}>Version {item.number}{item.current ? " (current saved copy)" : ""}</option>)}</select></label>
      <p>Unsaved changes in Word are not included.</p>
      <label>Recipient or party (optional)<input value={party} disabled={busy} onChange={(event) => { setParty(event.target.value); requestId.current = crypto.randomUUID() }} /></label>
      {name.toLowerCase().endsWith(".docx") && <><label>Comparison copy<select value={base} disabled={busy} onChange={(event) => { setBase(event.target.value); requestId.current = crypto.randomUUID() }}><option value="">Saved file</option>{versions.filter((item) => item.id !== version).map((item) => <option key={item.id} value={item.id}>Blackline against version {item.number}</option>)}</select></label>
      <label><input type="checkbox" checked={marker} disabled={busy || !preferencesReady} onChange={(event) => { setMarker(event.target.checked); requestId.current = crypto.randomUUID() }} />Include an opaque matching ID</label><p>Helps recognize the document if it comes back renamed. It carries no access permission. Turn it off for a copy whose metadata must be scrubbed.</p></>}
      {error && <p role="alert">{error}</p>}
    </div><div className="upload-review-actions"><button type="button" disabled={busy} onClick={onClose}>Cancel</button><button type="button" className="primary" disabled={busy || !preferencesReady || !versions.some(item => item.id === version)} onClick={send}>{busy ? "Recording…" : "Record as sent and download"}</button></div>
  </dialog>
}
