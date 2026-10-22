import { useContext, useEffect, useMemo, useRef, useState, type Ref } from "react"
import { Closing } from "./useExit"
import { SplitButton } from "./SplitButton"
import { useDocxodus, type ConversionOptions } from "docxodus/react"
import { CommentRenderMode } from "docxodus"
import {
  fetchRedlinedBytes,
  fetchRedlines,
  fetchPreparedChanges,
  acceptRedline,
  rejectRedline,
  acceptAllRedlines,
  rejectAllRedlines,
  prepareDocumentChanges,
  type AcceptResult,
  type PreparedChanges,
  type Redline,
} from "../api/ingest/redlines"
import { convertPdfToDocx, documentContentUrl, fetchDocumentText, matterDocumentKind } from "../api/ingest/documents"
import { ModalClose } from "./ModalHeader"
import { CompareIcon } from "./CompareDialog"
import { citationPassage, focusCitationPassage, type CitationFocus } from "./citation-target"
import { useToast } from "./Toast"
import useModalFocus from "./useModalFocus"
import backdropClose from "./backdropClose"

// How a Word file with tracked changes is painted: insertions green, deletions
// struck red, moves, headers and footers, comments in the margin. Shared with
// the blackline view so both read the same way.
export const DOCX_RENDER_OPTIONS: ConversionOptions = {
  renderTrackedChanges: true,
  showDeletedContent: true,
  renderMoveOperations: true,
  renderHeadersAndFooters: true,
  commentRenderMode: CommentRenderMode.Margin,
}

// A matter document's viewer. Emails and spreadsheets have no Word package to
// render or redline, so they open as the text ingest extracted from them; Word
// files and PDFs open in the redline viewer.
export default function DocumentViewer(props: Parameters<typeof WordDocumentViewer>[0]) {
  const kind = matterDocumentKind(props.name)
  if (kind === "eml" || kind === "xlsx")
    return <DocumentTextViewer matterId={props.matterId} name={props.name} kind={kind} citation={props.citation} onClose={props.onClose} />
  return <WordDocumentViewer {...props} />
}

// In-browser redline viewer. Fetches the matter's document with every pending
// redline applied as native tracked changes and converts it to HTML with the WASM
// runtime — insertions render green, deletions struck red, the redline lawyers
// expect. The change list alongside it accepts or rejects each proposal, one block
// at a time or the whole document at once. Bytes convert client-side, so the
// document never leaves the browser for a third-party service.
function WordDocumentViewer({
  matterId,
  name,
  focusId,
  citation,
  onClose,
  onChanged,
  onConverted,
  onSaveTemplate,
  onCompare,
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
  // Compare this Word file with another version of it. DOCX only.
  onCompare?: () => void
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
  // The last Prepare's batch and open items, stored by ingest and read with the
  // change list, so they survive closing the viewer. The count line is derived
  // from the refetched list; the open items stay until the next Prepare.
  const [preparation,setPreparation] = useState<PreparedChanges | null>(null)
  // Bumped after an accept/reject resolves so the load effect re-runs — the
  // document re-renders and the change list shrinks without duplicating fetch logic.
  const [reload, setReload] = useState(0)
  const focusRef = useRef<HTMLDivElement>(null)
  const modalRef = useRef<HTMLDivElement>(null)
  useModalFocus(modalRef)
  const closing = useContext(Closing)
  const documentRef = useRef<HTMLDivElement>(null)
  const [passage, setPassage] = useState<"matched" | "ambiguous" | "unavailable">()
  const moreRef = useRef<HTMLDetailsElement>(null)
  const toast = useToast()
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
    }), fetchRedlines(matterId, name), fetchPreparedChanges(matterId, name)])
      .then(async ([bytes, changes, prepared]) => {
        const out = await convertToHtml(bytes, DOCX_RENDER_OPTIONS)
        if (cancelled) return
        setHtml(out)
        setRedlines(changes)
        setPreparation(prepared)
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

  async function resolve(action: () => Promise<AcceptResult | void>, key: number | "all") {
    setBusy(key)
    setError(undefined)
    try {
      const result = await action()
      for (const notice of [notesNotUpdatedNotice(result, key === "all"), redlineNotUpdatedNotice(result), combinedNotUpdatedNotice(result)])
        if (notice) toast("warning", notice)
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
    try {
      await prepareDocumentChanges(matterId,name,instructions)
      onChanged?.()
      setReload((value) => value+1)
    } catch (cause) {setError(cause instanceof Error ? cause.message : String(cause))}
    finally {setBusy(undefined)}
  }

  // Download either the redlined .docx — the pending changes that can still be
  // accepted, as native Word tracked changes, to send to counsel — or the current
  // accepted Word file exactly as stored. Pending proposals never block the latter.
  async function download(tracked: boolean) {
    setError(undefined)
    try {
    const bytes = tracked ? await fetchRedlinedBytes(matterId, name) : await (async () => {
      const current = await fetch(documentContentUrl(matterId, name))
      if (!current.ok) throw new Error(`Could not download ${name} (${current.status})`)
      return new Uint8Array(await current.arrayBuffer())
    })()
    const url = URL.createObjectURL(
      new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
    )
    const a = document.createElement("a")
    a.href = url
    a.download = tracked ? `${name.replace(/\.docx$/i, "")} (tracked changes).docx` : name
    a.click()
    URL.revokeObjectURL(url)
    } catch(cause) {setError(cause instanceof Error ? cause.message : String(cause))}
  }
  const staleNotice = staleExportNotice(redlines)
  const prepared = preparation && preparedNotice(preparation.redline_ids, redlines)

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
      <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
        <div ref={modalRef} className="viewer-panel document-reader" role="dialog" aria-modal="true" aria-label={name} tabIndex={-1} onClick={(e) => e.stopPropagation()} onKeyDown={(e) => {
          if (e.key === "Escape" && !e.defaultPrevented) { e.stopPropagation(); onClose() }
        }}>
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
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
      <div ref={modalRef} className="viewer-panel document-reader" role="dialog" aria-modal="true" aria-label={name} tabIndex={-1} onClick={(e) => e.stopPropagation()} onKeyDown={(e) => {
        if (e.key === "Escape" && !e.defaultPrevented) { e.stopPropagation(); onClose() }
      }}>
        <div className="viewer-bar">
          <ViewerTitle name={name} kind="Word document" pending={redlines.length} />
          <div className="viewer-bar-actions">
            <DownloadActions redlines={redlines} preview={Boolean(html)} onDownload={download} />
            <button className="primary" onClick={() => setEditing((value) => !value)} disabled={busy !== undefined} aria-expanded={editing} aria-controls="document-edit-request">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M12 20h9" />
                <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z" />
              </svg>
              Request edits
            </button>
            {(onSaveTemplate || onCompare) && (
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
                  {onCompare && (
                    <button onClick={onCompare}>
                      <CompareIcon size={16} />
                      <span>
                        Compare with…
                        <small>Blackline against another version</small>
                      </span>
                    </button>
                  )}
                  {onSaveTemplate && (
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
                  )}
                </div>
              </details>
            )}
            <span className="viewer-bar-divider" aria-hidden="true" />
            <ModalClose onClose={onClose} />
          </div>
        </div>
        <div className="viewer-body">
          <div className="viewer-doc">
            {citation && <p role="status">{citationStatus(citation.section, passage)}</p>}
            {editing && <div id="document-edit-request" className="change-card document-edit-request">
              <label htmlFor="document-edit-instructions">Describe the changes you want to review</label>
              <textarea id="document-edit-instructions" value={instructions} maxLength={24000} disabled={busy !== undefined}
                onChange={(event) => setInstructions(event.target.value)} rows={4} style={{width:"100%"}} />
              <button className="primary" onClick={prepare} disabled={busy !== undefined || !instructions.trim()}>{busy === "prepare" ? "Preparing changes…" : "Prepare changes"}</button>
              {prepared && <p role="status">{prepared}</p>}
            </div>}
            {preparation && preparation.open_items.length > 0 && <div className="document-open-items">
              <strong id="document-open-items">Open items</strong>
              <ul aria-labelledby="document-open-items">{preparation.open_items.map((item, index) => <li key={index}>{item}</li>)}</ul>
            </div>}
            {wasmError && <p className="muted">Viewer failed to load: {wasmError.message}</p>}
            {error && <p className="muted">{error}</p>}
            {staleNotice && <p role="status" className="muted">{staleNotice}</p>}
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

// Where a citation landed in the open document, said in the status line.
export function citationStatus(section: string, passage?: "matched" | "ambiguous" | "unavailable") {
  return `Source section ${section}. ${passage === "matched"
    ? "The cited excerpt is highlighted."
    : passage === "ambiguous"
      ? "This excerpt appears more than once. The document is open; no exact passage was selected."
      : passage === "unavailable"
        ? "The cited excerpt could not be located in the current preview. The document is open; no exact passage was selected."
        : "Locating the cited excerpt…"}`
}

// An email or spreadsheet, read-only, as the text ingest extracted from it:
// the same text the assistant reads and cites, so a citation's excerpt is found
// in it directly and highlighted. Nothing here edits, redlines or converts.
function DocumentTextViewer({ matterId, name, kind, citation, onClose }: {
  matterId: string
  name: string
  kind: "eml" | "xlsx"
  citation?: CitationFocus
  onClose: () => void
}) {
  const [text, setText] = useState<string>()
  const [error, setError] = useState<string>()
  const markRef = useRef<HTMLElement>(null)
  const modalRef = useRef<HTMLDivElement>(null)
  useModalFocus(modalRef)
  const closing = useContext(Closing)
  const match = useMemo(() => (text !== undefined && citation ? citationPassage(text, citation.excerpt) : undefined), [text, citation])

  useEffect(() => {
    let cancelled = false
    fetchDocumentText(matterId, name).then(
      (value) => !cancelled && setText(value),
      (cause) => !cancelled && setError(cause instanceof Error ? cause.message : String(cause)),
    )
    return () => {
      cancelled = true
    }
  }, [matterId, name])

  useEffect(() => {
    markRef.current?.scrollIntoView({ block: "center" })
    markRef.current?.focus({ preventScroll: true })
  }, [match])

  return (
    <div className={`viewer-overlay${closing ? " is-closing" : ""}`} inert={closing} {...backdropClose(() => onClose())}>
      <div ref={modalRef} className="viewer-panel document-reader" role="dialog" aria-modal="true" aria-label={name} tabIndex={-1} onClick={(e) => e.stopPropagation()} onKeyDown={(e) => {
        if (e.key === "Escape" && !e.defaultPrevented) { e.stopPropagation(); onClose() }
      }}>
        <div className="viewer-bar">
          <ViewerTitle name={name} kind={`${kind === "eml" ? "Email" : "Excel spreadsheet"} · extracted text, read-only`} />
          <div className="viewer-bar-actions">
            <ModalClose onClose={onClose} />
          </div>
        </div>
        <div className="viewer-body">
          <div className="viewer-doc">
            {citation && <p role="status">{citationStatus(citation.section, error ? "unavailable" : match?.status)}</p>}
            {error && <p className="muted">{error}</p>}
            {!error && text === undefined && (
              <div className="docx-render docx-skeleton" aria-label="Reading the text">
                {Array.from({ length: 14 }, (_, i) => <div key={i} className="sk-line" />)}
              </div>
            )}
            {text !== undefined && <ExtractedText kind={kind} text={text} match={match} markRef={markRef} />}
          </div>
        </div>
      </div>
    </div>
  )
}

// The extracted text on the viewer's sheet, the matched excerpt marked in place.
export function ExtractedText({ kind, text, match, markRef }: {
  kind: "eml" | "xlsx"
  text: string
  match?: ReturnType<typeof citationPassage>
  markRef?: Ref<HTMLElement>
}) {
  if (!text.trim()) return <p className="muted">No text was read from this file.</p>
  if (match?.status !== "matched") return <pre className={`docx-render document-text is-${kind}`}>{text}</pre>
  return (
    <pre className={`docx-render document-text is-${kind}`}>
      {text.slice(0, match.start)}
      <mark className="citation-passage" tabIndex={-1} ref={markRef}>{text.slice(match.start, match.end)}</mark>
      {text.slice(match.end)}
    </pre>
  )
}

// The current accepted Word file is always downloadable. A redline download is
// offered only while some pending proposal can still render as tracked changes.
export function DownloadActions({ redlines, preview, onDownload }: {
  redlines: Redline[]
  preview: boolean
  onDownload: (tracked: boolean) => void
}) {
  // Labels are visually hidden on narrow screens, so the two downloads need
  // distinct glyphs: a plain arrow, and a document with a pen for tracked changes.
  const download = (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M12 4v11" />
      <path d="m7 10 5 5 5-5" />
      <path d="M5 20h14" />
    </svg>
  )
  const downloadTitle = "Download the current accepted Word file without pending changes"
  if (!redlines.some((row) => !row.stale)) {
    return (
      <button onClick={() => onDownload(false)} title={downloadTitle}>
        {download}
        <span className="viewer-action-label">Download</span>
      </button>
    )
  }
  return (
    <SplitButton
      label={<span className="viewer-action-label">Download</span>}
      icon={download}
      title={downloadTitle}
      onClick={() => onDownload(false)}
      toggleLabel="More download options"
      items={[{
        label: "Download redline",
        title: "Download as a Word file with the pending changes as tracked changes",
        disabled: !preview,
        onSelect: () => onDownload(true),
        icon: (
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M13 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h4" />
            <path d="M13 3v5h5v3" />
            <path d="M8 13h5M8 17h2" />
            <path d="M19 14.5a1.8 1.8 0 0 1 2.5 2.5L16 22.5l-3.5 1 1-3.5z" />
          </svg>
        ),
      }]}
    />
  )
}

// Stale proposals stay listed until the reviewer rejects them, but the tracked-
// changes preview and redline download leave them out: they target an earlier
// Word revision. Say so, and name the action that actually clears them.
export function staleExportNotice(redlines: Redline[]) {
  const stale = redlines.filter((row) => row.stale).length
  if (!stale) return
  const one = stale === 1
  return `${one ? "1 stale proposal is" : `${stale} stale proposals are`} left out of the tracked changes and the redline download. ` +
    `Reject ${one ? "it" : "them"} to clear ${one ? "it" : "them"}; prepare a fresh change list if the change is still needed.`
}

// Prepare's batch counted against the refetched change list: accepted, rejected
// and stale proposals drop out (the stale notice names the last), so the line
// goes once none is left to review. An empty batch says so.
export function preparedNotice(batch: number[], redlines: Redline[]) {
  if (!batch.length) return "No changes were proposed."
  const fresh = redlines.filter((row) => batch.includes(row.id) && !row.stale).length
  if (fresh) return `${fresh} proposed change${fresh === 1 ? "" : "s"} ready for review.`
}

// An accept into a drafting run's body also records the change in the run's
// drafting notes. When that could not happen the accept still stands; say so.
const notesSkipReasons: Record<string, string> = {
  notes_not_delivered: "the run has no delivered drafting notes",
  notes_missing: "the notes file is no longer in the matter",
  notes_not_deliverable: "the notes file is no longer a delivered document",
  pending_notes_proposals: "the notes have pending proposed changes",
  write_failed: "the notes file could not be saved",
}
export function notesNotUpdatedNotice(result: AcceptResult | void, all = false) {
  if (result?.notes_updated !== false) return
  const reason = result.notes_skip_reason && notesSkipReasons[result.notes_skip_reason]
  return `${all ? "Changes" : "Change"} accepted. The drafting notes were not updated${reason ? `: ${reason}` : ""}.`
}

// An accept into a drafting run's body also rebuilds the run's delivered redline
// from the accepted body. When that could not happen the accept still stands and
// the redline shows the body as it was; say so.
const redlineSkipReasons: Record<string, string> = {
  redline_not_delivered: "the run has no delivered redline",
  redline_missing: "the redline file is no longer in the matter",
  redline_not_deliverable: "the redline file is no longer a delivered document",
  pending_redline_proposals: "the redline has pending proposed changes",
  rebuild_failed: "it could not be rebuilt from the accepted document",
  write_failed: "the redline file could not be saved",
}
export function redlineNotUpdatedNotice(result: AcceptResult | void) {
  if (result?.redline_updated !== false) return
  const reason = result.redline_skip_reason && redlineSkipReasons[result.redline_skip_reason]
  return `The redline was not updated${reason ? `: ${reason}` : ""}.`
}

// An accept into a part of a saved combined file (the draft or the memo) also
// rebuilds that file. When that could not happen the accept still stands and the
// combined file holds the parts as they were; say so.
const combinedSkipReasons: Record<string, string> = {
  combined_missing: "the combined file is no longer in the matter",
  combined_not_deliverable: "the combined file is no longer a delivered document",
  pending_combined_proposals: "the combined file has pending proposed changes",
  combined_changed: "the combined file has changes of its own, which are kept",
  build_refused: "the memo and the draft can no longer be combined. Download them separately",
  rebuild_failed: "it could not be rebuilt from the accepted documents",
  write_failed: "the combined file could not be saved",
}
export function combinedNotUpdatedNotice(result: AcceptResult | void) {
  if (result?.combined_updated !== false) return
  const reason = result.combined_skip_reason && combinedSkipReasons[result.combined_skip_reason]
  return `The combined file was not updated${reason ? `: ${reason}` : ""}.`
}

// File glyph, name and a quiet status line. The pending count stays visible on
// narrow screens, where the change list sits below the document.
export function ViewerTitle({ name, kind, pending = 0 }: { name: string; kind: string; pending?: number }) {
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
