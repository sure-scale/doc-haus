import { useEffect, useRef, useState } from "react"
import { identifyReturnedFile, inspectReturnedFile, saveReturnedFile, type ReturnInspection } from "../api/ingest/exchanges"
import type { DocumentReturnTieBreak } from "../../../../dochaus/lib/document-return-tie-break-contract"
import { fetchDocumentVersions, type DocumentVersions } from "../api/ingest/versions"
import type { CompareParams } from "../api/ingest/compare"
import BlacklineViewer from "./BlacklineViewer"
import useModalFocus from "./useModalFocus"

// A returned file is inspected before choosing its document and parent. The
// lawyer can keep it separate; uncertain similarity never silently links files.
type ReturnedDocumentProps = { matterId: string; documents: readonly { name: string }[]; onNew: (file: File) => void; onSaved: () => void }
export default function ReturnedDocument(props: ReturnedDocumentProps) {
  return <ReturnedDocumentContext key={props.matterId} {...props} />
}

function ReturnedDocumentContext({ matterId, documents, onNew, onSaved }: ReturnedDocumentProps) {
  const input = useRef<HTMLInputElement>(null)
  const [file, setFile] = useState<File>()
  const [comparison, setComparison] = useState<CompareParams>()
  const [round, setRound] = useState(0)
  return <>
    <button type="button" className="linklike" onClick={() => input.current?.click()}>Add a returned version</button>
    <input ref={input} type="file" accept=".docx" hidden onChange={(event) => { setFile(event.target.files?.[0]); setRound(value => value + 1); setComparison(undefined); event.target.value = "" }} />
    {file && <ReturnDialog documents={documents} key={round} matterId={matterId} file={file} onClose={() => setFile(undefined)} onNew={() => { onNew(file); setFile(undefined) }} onSaved={(params) => { onSaved(); setFile(undefined); setComparison(params) }} />}
    {comparison && <BlacklineViewer matterId={matterId} params={comparison} onClose={() => setComparison(undefined)} />}
  </>
}

export function ReturnDialog({ matterId, documents, file, onClose, onNew, onSaved, initialInspection }: { matterId: string; documents: readonly { name: string }[]; file: File; onClose: () => void; onNew: () => void; onSaved: (params: CompareParams) => void; initialInspection?: ReturnInspection }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const [inspection, setInspection] = useState<ReturnInspection>()
  const [base, setBase] = useState("")
  const [selectedName, setSelectedName] = useState("")
  const [versions, setVersions] = useState<DocumentVersions>()
  const [party, setParty] = useState("")
  const [emailName, setEmailName] = useState("")
  const [inspectionError, setInspectionError] = useState<string>()
  const [advisory, setAdvisory] = useState<DocumentReturnTieBreak>()
  const [identifying, setIdentifying] = useState(false)
  const [advisoryError, setAdvisoryError] = useState<string>()
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string>()
  const request = useRef(crypto.randomUUID())
  const alive = useRef(true)
  const selectionEdited = useRef(false), partyEdited = useRef(false)
  const adviceEpoch = useRef(0), adviceRead = useRef<AbortController | undefined>(undefined)
  const candidate = inspection?.candidates.find((item) => item.versionId === base)
  const chosen = versions?.versions.find((item) => item.id === base && !item.removed)
  useModalFocus(dialog)
  useEffect(() => { alive.current = true; dialog.current?.showModal(); return () => { alive.current = false; adviceEpoch.current++; adviceRead.current?.abort(); dialog.current?.close() } }, [])
  useEffect(() => {
    let cancelled = false
    adviceEpoch.current++; adviceRead.current?.abort(); setAdvisory(undefined); setAdvisoryError(undefined); setIdentifying(false)
    setInspection(undefined); setInspectionError(undefined); setError(undefined)
    ;(!emailName && initialInspection ? Promise.resolve(initialInspection) : inspectReturnedFile(matterId, file, emailName || undefined)).then((value) => {
      if (cancelled) return
      setInspection(value)
      if (!selectionEdited.current) { setBase(value.selected?.versionId ?? value.candidates[0]?.versionId ?? ""); setSelectedName(value.selected?.documentName ?? value.candidates[0]?.documentName ?? "") }
      if (!partyEdited.current) setParty(value.source?.sender ?? "")
    }, (cause) => { if (!cancelled) setInspectionError(cause instanceof Error ? cause.message : String(cause)) })
    return () => { cancelled = true }
  }, [matterId, file, initialInspection, emailName])
  useEffect(() => {
    let cancelled = false; setVersions(undefined); setError(undefined)
    if (selectedName) fetchDocumentVersions(matterId, selectedName).then((value) => { if (!cancelled) { setVersions(value); setBase((previous) => value.versions.some((item) => item.id === previous && !item.removed) ? previous : value.working_version ?? value.versions.find((item) => !item.removed)?.id ?? "") } }, (cause) => { if (!cancelled) setError(cause instanceof Error ? cause.message : String(cause)) })
    return () => { cancelled = true }
  }, [matterId, selectedName])
  const current = versions?.versions.find((item) => item.id === versions.working_version && !item.removed)
  async function identify() {
    if (!inspection || inspection.kind !== "needs-confirmation" || inspection.candidates.length < 2 || identifying || busy) return
    const epoch = ++adviceEpoch.current, controller = new AbortController()
    adviceRead.current?.abort(); adviceRead.current = controller
    setIdentifying(true); setAdvisory(undefined); setAdvisoryError(undefined)
    try {
      const result = await identifyReturnedFile(matterId, file, emailName || undefined, controller.signal)
      if (!alive.current || epoch !== adviceEpoch.current || controller.signal.aborted) return
      setInspection({ ...result.inspection, source: result.source ?? result.inspection.source })
      setAdvisory(result.advisory)
    } catch (cause) { if (alive.current && epoch === adviceEpoch.current && !controller.signal.aborted) setAdvisoryError(cause instanceof Error ? cause.message : String(cause)) }
    finally { if (alive.current && epoch === adviceEpoch.current) setIdentifying(false) }
  }
  async function save() {
    if (!inspection || inspectionError || !chosen || !current || busy) return
    setBusy(true); setError(undefined)
    try {
      const result = await saveReturnedFile(matterId, file, { baseVersionId: base, priorSha256: current.sha256, requestId: request.current, party: party.trim() || undefined, emailName: emailName || undefined })
      if (!alive.current) return
      onSaved({ left: result.comparison.left ?? result.name, right: result.comparison.right ?? result.name, leftVersion: result.comparison.leftVersion, rightVersion: result.comparison.rightVersion })
    } catch (cause) { if (alive.current) setError(cause instanceof Error ? cause.message : String(cause)) }
    finally { if (alive.current) setBusy(false) }
  }
  const duplicate = inspection?.kind === "duplicate" && inspection.candidates.some(item => item.documentName === selectedName && item.sha256 === inspection.sha256)
  const emails = documents.filter(item => /\.(eml|msg)$/i.test(item.name))
  const suggested = advisory?.choice === "candidate" ? inspection?.candidates.find(item => item.versionId === advisory.versionId && documents.some(document => document.name === item.documentName)) : undefined
  return <dialog ref={dialog} className="upload-review-dialog" aria-label="Review returned file" onCancel={(event) => { event.preventDefault(); event.stopPropagation(); if (!busy) onClose() }}>
    <div className="upload-review-content"><h2>Review returned file</h2><p>{file.name}</p>
      {emails.length > 0 && <><label>Email source (optional)<select value={emailName} disabled={busy} onChange={event => { adviceEpoch.current++; adviceRead.current?.abort(); setAdvisory(undefined); setInspection(undefined); setEmailName(event.target.value); request.current = crypto.randomUUID() }}><option value="">No email source</option>{emails.map(item => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label><p>Choose the saved email this exact attachment came from to retain its sender and date.</p></>}
      {!inspection && !inspectionError && <p role="status">Looking for the document and version this came from…</p>}
      {inspection && <><p>{inspection.reason}</p>{inspection.warnings.map((warning) => <p key={warning}>{warning}</p>)}
        {inspection.kind === "needs-confirmation" && inspection.candidates.length > 1 && <button type="button" disabled={busy || identifying} onClick={identify}>{identifying ? "Checking candidates…" : "Help identify this return"}</button>}
        {advisory && <section aria-label="Return identification advice"><h3>Suggested match</h3><p>{advisory.choice === "candidate" ? "Review this suggestion before choosing a base." : advisory.choice === "new-document" ? "The file may be a separate document." : "No candidate could be identified confidently."}</p><p>{advisory.reason}</p>
          <p>Model: {advisory.model.providerID}/{advisory.model.id}{advisory.model.variant ? ` · ${advisory.model.variant}` : ""}</p>
          {advisory.evidence.map((evidence, index) => <div key={index}><p>Saved version {inspection.candidates.find(item => item.versionId === evidence.versionId)?.versionNumber ?? evidence.versionId}</p><blockquote>Returned file: {evidence.incomingQuote}</blockquote><blockquote>Saved candidate: {evidence.candidateQuote}</blockquote></div>)}
          {advisory.limitations.map((limitation, index) => <p key={index}>{limitation}</p>)}
          {suggested && <button type="button" disabled={busy} onClick={() => { selectionEdited.current = true; setSelectedName(suggested.documentName); setBase(suggested.versionId); request.current = crypto.randomUUID() }}>Use suggested base</button>}
        </section>}
        {advisoryError && <p role="alert">{advisoryError}</p>}
        {emailName && <div aria-label="Recorded email source"><p>Email source: {inspection.source?.documentName ?? "Not recorded"}</p><p>Sender: {inspection.source?.sender ?? "Not recorded"}</p><p>Date: {inspection.source?.date ?? "Not recorded"}</p></div>}
        <label>Document<select value={selectedName} disabled={busy} onChange={(event) => { selectionEdited.current = true; setSelectedName(event.target.value); setBase(""); request.current = crypto.randomUUID() }}><option value="">Choose a document</option>{documents.filter((item) => item.name.toLowerCase().endsWith(".docx")).map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
        {versions && <label>Base version<select value={base} disabled={busy} onChange={(event) => { selectionEdited.current = true; setBase(event.target.value); request.current = crypto.randomUUID() }}>{versions.versions.filter((item) => !item.removed).map((item) => <option key={item.id} value={item.id}>Version {item.number}{item.current ? " (current)" : ""}</option>)}</select></label>}
        {candidate && <p>{candidate.reason}</p>}
        {duplicate ? <p>This copy is already kept. Nothing will be saved unless you deliberately add a separate document.</p> : chosen && <><p>The returned file will become the current copy. Its parent stays version {chosen.number}; existing internal versions are kept. The blackline and untracked-change check open against this selected base after saving.</p>
          <label>Received from (optional)<input value={party} disabled={busy} onChange={(event) => { partyEdited.current = true; setParty(event.target.value); request.current = crypto.randomUUID() }} /></label>
          {versions && !current && <p>The current file changed outside its saved version. Save or refresh it before importing a return.</p>}
        </>}
      </>}
      {error && <p role="alert">{error}</p>}
      {inspectionError && <p role="alert">{inspectionError}</p>}
    </div><div className="upload-review-actions"><button type="button" disabled={busy} onClick={onClose}>Cancel</button><button type="button" disabled={busy || !inspection} onClick={onNew}>Add as separate document</button>{inspection && chosen && !duplicate && <button type="button" className="primary" disabled={busy || !current} onClick={save}>{busy ? "Saving…" : "Save returned version and compare"}</button>}</div>
  </dialog>
}
