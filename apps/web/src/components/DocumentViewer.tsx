import { useEffect, useMemo, useRef, useState } from "react"
import { useDocxodus } from "docxodus/react"
import { CommentRenderMode } from "docxodus"
import {
  fetchRedlinedBytes,
  fetchRedlines,
  acceptRedline,
  rejectRedline,
  acceptAllRedlines,
  rejectAllRedlines,
  prepareDocumentChanges,
  type Redline,
} from "../api/ingest/redlines"
import { convertPdfToDocx, documentContentUrl } from "../api/ingest/documents"
import { ModalClose } from "./ModalHeader"
import { focusCitationPassage, type CitationFocus } from "./citation-target"

// In-browser redline viewer. Fetches the matter's document with every pending
// redline applied as native tracked changes and converts it to HTML with the WASM
// runtime — insertions render green, deletions struck red, the redline lawyers
// expect. The change list alongside it accepts or rejects each proposal, one block
// at a time or the whole document at once. Bytes convert client-side, so the
// document never leaves the browser for a third-party service.
export default function DocumentViewer({
  matterId,
  name,
  focusId,
  citation,
  onClose,
  onChanged,
  onConverted,
  onSaveTemplate,
}: {
  matterId: string
  name: string
  // A redline id to scroll to and highlight on open, set when the viewer was
  // opened from a chat redline preview's "View in document" link.
  focusId?: number
  // Source search offsets use a different text projection from this viewer.
  // Locate a unique rendered excerpt instead, and disclose when it cannot match.
  citation?: CitationFocus
  onClose: () => void
  onChanged?: () => void
  // Called with the new .docx name after a PDF is converted, so the parent can
  // swap the viewer onto the freshly indexed editable document.
  onConverted?: (name: string) => void
  // Hand this document to draft mode to turn into a reusable template (a chat
  // with the prompt prefilled). DOCX only — the redline pipeline and templates
  // are DOCX-only. Omitted on surfaces that do not offer the action.
  onSaveTemplate?: () => void
}) {
  // PDFs render natively in an <iframe>; the docxodus WASM path and the redline
  // pipeline are DOCX-only. A PDF carries no redlines until it is converted.
  const isPdf = name.toLowerCase().endsWith(".pdf")
  const [converting, setConverting] = useState(false)
  const { isReady, error: wasmError, convertToHtml } = useDocxodus("/wasm/")
  const [html, setHtml] = useState<string>()
  const [redlines, setRedlines] = useState<Redline[]>([])
  const [error, setError] = useState<string>()
  const [busy, setBusy] = useState<number | "all" | "prepare">()
  const [editing,setEditing] = useState(false)
  const [instructions,setInstructions] = useState("")
  const [preparation,setPreparation] = useState<string>()
  // Bumped after an accept/reject resolves so the load effect re-runs — the
  // document re-renders and the change list shrinks without duplicating fetch logic.
  const [reload, setReload] = useState(0)
  const focusRef = useRef<HTMLDivElement>(null)
  const documentRef = useRef<HTMLDivElement>(null)
  const [passage, setPassage] = useState<"matched" | "ambiguous" | "unavailable">()
  const moreRef = useRef<HTMLDetailsElement>(null)
  // Keep React from resetting the HTML when passage status or viewer controls
  // re-render. The citation effect owns only its temporary text highlights.
  const documentPreview = useMemo(() => html && <div className="docx-render" ref={documentRef} dangerouslySetInnerHTML={{ __html: html }} />, [html])

  useEffect(() => {
    if (!isReady || isPdf) return
    let cancelled = false
    setHtml(undefined)
    setError(undefined)
    Promise.all([fetchRedlinedBytes(matterId, name).catch(async (cause) => {
      if (!cancelled) setError(cause instanceof Error ? cause.message : String(cause))
      const current = await fetch(documentContentUrl(matterId,name))
      if (!current.ok) throw cause
      return new Uint8Array(await current.arrayBuffer())
    }), fetchRedlines(matterId, name)])
      .then(async ([bytes, changes]) => {
        const out = await convertToHtml(bytes, {
          renderTrackedChanges: true,
          showDeletedContent: true,
          renderMoveOperations: true,
          renderHeadersAndFooters: true,
          commentRenderMode: CommentRenderMode.Margin,
        })
        if (cancelled) return
        setHtml(out)
        setRedlines(changes)
      })
      .catch((e) => !cancelled && setError(e instanceof Error ? e.message : String(e)))
    return () => {
      cancelled = true
    }
  }, [isReady, matterId, name, reload])

  // Once the change list has rendered, scroll the proposal the chat link pointed
  // at into view and highlight it, so the lawyer lands on that exact change rather
  // than the top of a long list.
  useEffect(() => {
    if (focusId !== undefined) focusRef.current?.scrollIntoView({ block: "center" })
  }, [focusId, redlines])

  useEffect(() => {
    setPassage(undefined)
    if (!citation || !html || !documentRef.current) return
    const root = documentRef.current
    setPassage(focusCitationPassage(root, citation.excerpt))
    return () => {
      for (const mark of root.querySelectorAll("mark.citation-passage")) mark.replaceWith(...mark.childNodes)
      root.normalize()
    }
  }, [citation, html])

  // A <details> menu stays open until its summary is clicked again; also close it
  // on a press anywhere else in the viewer, as a menu is expected to.
  useEffect(() => {
    const dismiss = (event: PointerEvent) => {
      const menu = moreRef.current
      if (menu?.open && !menu.contains(event.target as Node)) menu.open = false
    }
    document.addEventListener("pointerdown", dismiss)
    return () => document.removeEventListener("pointerdown", dismiss)
  }, [])

  async function resolve(action: () => Promise<void>, key: number | "all") {
    setBusy(key)
    setError(undefined)
    try {
      await action()
      onChanged?.()
      setReload((n) => n + 1)
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause))
    } finally {
      setBusy(undefined)
    }
  }

  async function prepare() {
    setBusy("prepare")
    setError(undefined)
    setPreparation(undefined)
    try {
      const result = await prepareDocumentChanges(matterId,name,instructions)
      setPreparation(`${result.proposed} proposed change${result.proposed === 1 ? "" : "s"} ready for review.${result.open_items.length ? ` Open items: ${result.open_items.join(" ")}` : ""}`)
      onChanged?.()
      setReload((value) => value+1)
    } catch (cause) {setError(cause instanceof Error ? cause.message : String(cause))}
    finally {setBusy(undefined)}
  }

  // Download the redlined .docx — the document with pending changes as native Word
  // tracked changes — so it can be sent to counsel to accept or reject in Word.
  // With nothing pending this is simply the current clean document.
  async function download() {
    try {
    const bytes = await fetchRedlinedBytes(matterId, name)
    const url = URL.createObjectURL(
      new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
    )
    const a = document.createElement("a")
    a.href = url
    a.download = redlines.length ? `${name.replace(/\.docx$/i, "")} (tracked changes).docx` : name
    a.click()
    URL.revokeObjectURL(url)
    } catch(cause) {setError(cause instanceof Error ? cause.message : String(cause))}
  }

  // Convert this PDF into an editable .docx sibling, index it, and swap the viewer
  // onto the new document so it can be redlined like any other contract.
  async function convert() {
    setConverting(true)
    setError(undefined)
    try {
      const result = await convertPdfToDocx(matterId, name)
      onChanged?.()
      onConverted?.(result.name)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setConverting(false)
    }
  }

  if (isPdf)
    return (
      <div className="viewer-overlay" onClick={onClose}>
        <div className="viewer-panel document-reader" role="dialog" aria-modal="true" aria-label={name} onClick={(e) => e.stopPropagation()}>
          <div className="viewer-bar">
            <ViewerTitle name={name} kind="PDF document" />
            <div className="viewer-bar-actions">
              <button
                className="primary"
                onClick={convert}
                disabled={converting}
                title="Convert this PDF to an editable Word document so it can be redlined"
              >
                {converting ? "Converting..." : "Convert to DOCX"}
              </button>
              <span className="viewer-bar-divider" aria-hidden="true" />
              <ModalClose onClose={onClose} />
            </div>
          </div>
          <div className="viewer-body">
            <div className="viewer-doc">
              {citation && <p role="status">Source section {citation.section}. PDF passage navigation is unavailable; the source document is open.</p>}
              {error && <p className="muted">{error}</p>}
              <iframe className="pdf-render" title={name} src={documentContentUrl(matterId, name)} />
            </div>
          </div>
        </div>
      </div>
    )

  return (
    <div className="viewer-overlay" onClick={onClose}>
      <div className="viewer-panel document-reader" role="dialog" aria-modal="true" aria-label={name} onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <ViewerTitle name={name} kind="Word document" pending={redlines.length} />
          <div className="viewer-bar-actions">
            <button onClick={download} disabled={!html} title="Download as a Word file with tracked changes">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M12 4v11" />
                <path d="m7 10 5 5 5-5" />
                <path d="M5 20h14" />
              </svg>
              <span className="viewer-action-label">{redlines.length ? "Download redline" : "Download"}</span>
            </button>
            <button className="primary" onClick={() => setEditing((value) => !value)} disabled={busy !== undefined} aria-expanded={editing} aria-controls="document-edit-request">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M12 20h9" />
                <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z" />
              </svg>
              Request edits
            </button>
            {onSaveTemplate && (
              <details className="viewer-more-actions" ref={moreRef} onKeyDown={(event) => {
                if (event.key !== "Escape" || !event.currentTarget.open) return
                event.preventDefault()
                event.stopPropagation()
                event.currentTarget.open = false
                event.currentTarget.querySelector("summary")?.focus()
              }}>
                <summary aria-label="More actions" title="More actions">
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                    <circle cx="5" cy="12" r="1.8" />
                    <circle cx="12" cy="12" r="1.8" />
                    <circle cx="19" cy="12" r="1.8" />
                  </svg>
                </summary>
                <div className="viewer-more-menu">
                  <button onClick={onSaveTemplate}>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                      <rect x="8" y="8" width="13" height="13" rx="2" />
                      <path d="M4 16V5a2 2 0 0 1 2-2h11" />
                    </svg>
                    <span>
                      Save as template
                      <small>Reusable copy with client details removed</small>
                    </span>
                  </button>
                </div>
              </details>
            )}
            <span className="viewer-bar-divider" aria-hidden="true" />
            <ModalClose onClose={onClose} />
          </div>
        </div>
        <div className="viewer-body">
          <div className="viewer-doc">
            {citation && <p role="status">Source section {citation.section}. {passage === "matched"
              ? "The cited excerpt is highlighted."
              : passage === "ambiguous"
                ? "This excerpt appears more than once. The document is open; no exact passage was selected."
                : passage === "unavailable"
                  ? "The cited excerpt could not be located in the current preview. The document is open; no exact passage was selected."
                  : "Locating the cited excerpt…"}</p>}
            {editing && <div id="document-edit-request" className="change-card document-edit-request">
              <label htmlFor="document-edit-instructions">Describe the changes you want to review</label>
              <textarea id="document-edit-instructions" value={instructions} maxLength={24000} disabled={busy !== undefined}
                onChange={(event) => setInstructions(event.target.value)} rows={4} style={{width:"100%"}} />
              <button className="primary" onClick={prepare} disabled={busy !== undefined || !instructions.trim()}>{busy === "prepare" ? "Preparing changes…" : "Prepare changes"}</button>
              {preparation && <p role="status">{preparation}</p>}
            </div>}
            {wasmError && <p className="muted">Viewer failed to load: {wasmError.message}</p>}
            {error && <p className="muted">{error}</p>}
            {!wasmError && !error && !html && (
              <div className="docx-render docx-skeleton" aria-label="Rendering document">
                {Array.from({ length: 14 }, (_, i) => <div key={i} className="sk-line" />)}
              </div>
            )}
            {documentPreview}
          </div>
          {redlines.length > 0 && (
            <aside className="changes-panel">
              <div className="changes-head">
                <span className="changes-count">
                  <strong>{redlines.length}</strong> pending change{redlines.length === 1 ? "" : "s"}
                </span>
                <div className="changes-actions">
                  <button
                    className="btn-accept"
                    disabled={busy !== undefined || redlines.some((row) => row.stale)}
                    onClick={() => resolve(() => acceptAllRedlines(matterId, name), "all")}
                  >
                    Accept all
                  </button>
                  <button
                    className="btn-reject"
                    disabled={busy !== undefined}
                    onClick={() => resolve(() => rejectAllRedlines(matterId, name), "all")}
                  >
                    Reject all
                  </button>
                </div>
              </div>
              {redlines.map((r) => (
                <div
                  className={`change-card${r.id === focusId ? " focused" : ""}`}
                  key={r.id}
                  ref={r.id === focusId ? focusRef : undefined}
                >
                  <div className="change-author">{r.author}</div>
                  {r.old_text && <div className="change-old">{r.old_text}</div>}
                  <div className="change-new">{r.new_text}</div>
                  {r.review_comment && <div className="change-comment">{r.review_comment}</div>}
                  {r.stale_reason && <div className="change-comment">{r.stale_reason}</div>}
                  <div className="change-buttons">
                    <button
                      className="btn-accept"
                      disabled={busy !== undefined || r.stale}
                      onClick={() => resolve(() => acceptRedline(matterId, r.id), r.id)}
                    >
                      {busy === r.id ? "..." : "Accept"}
                    </button>
                    <button
                      className="btn-reject"
                      disabled={busy !== undefined}
                      onClick={() => resolve(() => rejectRedline(matterId, r.id), r.id)}
                    >
                      Reject
                    </button>
                  </div>
                </div>
              ))}
            </aside>
          )}
        </div>
      </div>
    </div>
  )
}

// File glyph, name and a quiet status line. The pending count stays visible on
// narrow screens, where the change list sits below the document.
function ViewerTitle({ name, kind, pending = 0 }: { name: string; kind: string; pending?: number }) {
  return (
    <div className="viewer-title">
      <span className="viewer-file-icon" aria-hidden="true">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
          <path d="M14 2v6h6M8 13h8M8 17h5" />
        </svg>
      </span>
      <span className="viewer-title-text">
        <span className="viewer-title-name" title={name}>{name}</span>
        <span className="viewer-title-meta">
          {kind}
          {pending > 0 && <span className="viewer-title-pending">{pending} pending change{pending === 1 ? "" : "s"}</span>}
        </span>
      </span>
    </div>
  )
}
