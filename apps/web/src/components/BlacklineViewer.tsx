import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react"
import { useDocxodus } from "docxodus/react"
import {
  compareUrl,
  fetchBlacklineBytes,
  fetchComparison,
  type ChangeHunk,
  type ChangeLocation,
  type CompareParams,
  type ComparisonJson,
} from "../api/ingest/compare"
import { humanizeError } from "../api/ingest/errors"
import { fileStem } from "./CompareDialog"
import { DOCX_RENDER_OPTIONS, ViewerTitle } from "./DocumentViewer"
import { ModalClose } from "./ModalHeader"

// The blackline of two Word versions, painted with the same render as the
// document viewer, beside a list of the edits the later version made without
// tracked changes and then every change. View and download only: the blackline
// is generated on demand and never saved into the matter.
export default function BlacklineViewer({ matterId, params, onClose }: { matterId: string; params: CompareParams; onClose: () => void }) {
  const { isReady, error: wasmError, convertToHtml } = useDocxodus("/wasm/")
  const panel = useRef<HTMLDivElement>(null)
  // The fetched blackline, painted on screen and saved by Download.
  const [blackline, setBlackline] = useState<Uint8Array>()
  const [html, setHtml] = useState<string>()
  const [renderError, setRenderError] = useState<string>()
  const [comparison, setComparison] = useState<ComparisonJson>()
  const [panelError, setPanelError] = useState<string>()
  // One string names the comparison, so the effects re-run only when it changes.
  const key = compareUrl(matterId, params)
  const name = comparison?.blacklineName ?? `Blackline - ${fileStem(params.left)} vs ${fileStem(params.right)}.docx`
  const documentPreview = useMemo(() => html && <div className="docx-render" dangerouslySetInnerHTML={{ __html: html }} />, [html])

  useEffect(() => {
    let cancelled = false
    setComparison(undefined)
    setPanelError(undefined)
    fetchComparison(matterId, params).then(
      (result) => !cancelled && setComparison(result),
      (cause) => !cancelled && setPanelError(humanizeError(cause, "Could not compare the documents.")),
    )
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key])

  // Move focus into the dialog and hand it back to whatever opened it.
  useLayoutEffect(() => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    panel.current?.querySelector<HTMLButtonElement>(".modal-close")?.focus({ preventScroll: true })
    return () => {
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [])

  // Fetched apart from the viewer runtime, so the blackline can be downloaded
  // even when it cannot be painted.
  useEffect(() => {
    let cancelled = false
    setBlackline(undefined)
    setHtml(undefined)
    setRenderError(undefined)
    fetchBlacklineBytes(matterId, params).then(
      (bytes) => !cancelled && setBlackline(bytes),
      (cause) => !cancelled && setRenderError(humanizeError(cause, "Could not show the blackline.")),
    )
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key])

  useEffect(() => {
    if (!isReady || !blackline) return
    let cancelled = false
    convertToHtml(blackline, DOCX_RENDER_OPTIONS).then(
      (out) => !cancelled && setHtml(out),
      (cause) => !cancelled && setRenderError(humanizeError(cause, "Could not show the blackline.")),
    )
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isReady, blackline])

  const failed = renderError ?? (wasmError && `Viewer failed to load: ${wasmError.message}`)
  return (
    <div className="viewer-overlay" onClick={onClose}>
      <div ref={panel} className="viewer-panel document-reader blackline-reader" role="dialog" aria-modal="true" aria-label={name} onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <ViewerTitle name={name} kind={`Blackline: ${params.left} → ${params.right}`} />
          <div className="viewer-bar-actions">
            <button onClick={() => blackline && saveDocx(blackline, name)} disabled={!blackline} title="Download the blackline as a Word file with tracked changes">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <path d="M12 4v11" />
                <path d="m7 10 5 5 5-5" />
                <path d="M5 20h14" />
              </svg>
              <span className="viewer-action-label">Download blackline</span>
            </button>
            <span className="viewer-bar-divider" aria-hidden="true" />
            <ModalClose onClose={onClose} />
          </div>
        </div>
        <div className="viewer-body">
          <div className="viewer-doc">
            {failed && <p className="muted">{failed}</p>}
            {!failed && !html && (
              <div className="docx-render docx-skeleton" aria-label="Comparing documents">
                {Array.from({ length: 14 }, (_, i) => <div key={i} className="sk-line" />)}
              </div>
            )}
            {documentPreview}
          </div>
          {comparison ? (
            <ComparePanel comparison={comparison} />
          ) : panelError ? (
            // One refusal (a PDF, a missing file) fails both requests; say it once.
            panelError !== renderError && (
              <aside className="changes-panel compare-panel">
                <p className="muted">{panelError}</p>
              </aside>
            )
          ) : (
            <aside className="changes-panel compare-panel" aria-busy="true">
              <p className="muted" role="status">
                Comparing…
              </p>
            </aside>
          )}
        </div>
      </div>
    </div>
  )
}

// Save the blackline already on screen rather than asking the route again: the
// file matches what the lawyer reviewed, and a refusal can never be saved as a
// .docx.
export function saveDocx(bytes: Uint8Array, name: string) {
  const url = URL.createObjectURL(
    new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
  )
  const a = document.createElement("a")
  a.href = url
  a.download = name
  a.click()
  URL.revokeObjectURL(url)
}

const KIND: Record<ChangeHunk["kind"], string> = {
  modified: "Changed",
  inserted: "Inserted",
  deleted: "Deleted",
  moved: "Moved",
}
const SCOPE: Record<ChangeLocation["scope"], string | undefined> = {
  body: undefined,
  header: "Header",
  footer: "Footer",
  footnote: "Footnote",
  endnote: "Endnote",
}
const plural = (count: number, one: string, many = `${one}s`) => `${count} ${count === 1 ? one : many}`

// Where a change sits: its part of the document, the nearest heading, then the
// paragraph's own clause number or, failing one, its ordinal in the version the
// text comes from. Paragraph 0 means the compare could not place it.
export function hunkLocation(location: ChangeLocation, comparison: Pick<ComparisonJson, "left" | "right">) {
  const parts = [
    SCOPE[location.scope],
    location.heading,
    location.number ?? (location.paragraph ? `paragraph ${location.paragraph} in ${comparison[location.side].label}` : undefined),
  ].filter(Boolean) as string[]
  const text = parts.join(" · ") || "Location not found"
  return text.charAt(0).toUpperCase() + text.slice(1)
}

// How the untracked-edit check reads: a warning only when the later version
// has tracked changes that would hide edits made outside them.
function untrackedStatus({ left, right, untracked }: ComparisonJson) {
  if (!right.trackedChanges.count)
    return untracked.length
      ? { tone: "info", text: `${right.label} has no tracked changes, so every difference below was made without tracking.` }
      : { tone: "ok", text: `${right.label} has no tracked changes, and its wording matches ${left.label}.` }
  if (!untracked.length) return { tone: "ok", text: "No edits outside the tracked changes." }
  const one = untracked.length === 1
  return {
    tone: "warning",
    text: `${plural(untracked.length, "edit")} in ${right.label} ${one ? "was" : "were"} made without tracked changes. Its tracked changes do not show ${one ? "it" : "them"}.`,
  }
}

// What the comparison did with the inputs and the Word file, in plain words.
function disclosures({ left, right, options, attribution, untracked }: ComparisonJson) {
  const lines: string[] = []
  if (attribution === "label") lines.push(`Untracked edits are listed here only; in the Word file every change shows under ${right.label}.`)
  else if (untracked.length) lines.push(`In the Word file, edits made without tracked changes show under “Not tracked (${right.label})”.`)
  for (const side of [left, right]) {
    const { count, authors } = side.trackedChanges
    if (count) lines.push(`${side.label} had ${plural(count, "tracked change")}${authors.length ? ` (by ${authors.join(", ")})` : ""}; compared as accepted.`)
  }
  for (const side of [left, right])
    if (side.pendingProposals)
      lines.push(`${side.label} has ${plural(side.pendingProposals, "pending doc.haus proposal")}; ${side.pendingProposals === 1 ? "it is" : "they are"} not included.`)
  if (options.ignoreFormatting) lines.push("Formatting-only changes are left out.")
  if (!options.detectMoves) lines.push("Moved text shows as deleted and inserted.")
  if (!options.headersFooters) lines.push("Headers and footers were not compared.")
  return lines
}

export function ComparePanel({ comparison }: { comparison: ComparisonJson }) {
  const { left, right, changes, untracked } = comparison
  const status = untrackedStatus(comparison)
  return (
    <aside className="changes-panel compare-panel" aria-label="Changes">
      <div className="changes-head">
        <span className="changes-count">
          <strong>{changes.length}</strong> change{changes.length === 1 ? "" : "s"}
        </span>
        <span className="compare-sides-line">
          {left.label} → {right.label}
        </span>
      </div>
      <section className={`compare-section compare-untracked is-${status.tone}`} aria-labelledby="compare-untracked-title">
        <h3 id="compare-untracked-title">
          Edits made without tracked changes
          {untracked.length > 0 && <span className="compare-count">{untracked.length}</span>}
        </h3>
        <p className="compare-status">{status.text}</p>
        {untracked.map((hunk) => (
          <HunkCard key={hunk.id} hunk={hunk} comparison={comparison} />
        ))}
      </section>
      <section className="compare-section" aria-labelledby="compare-all-title">
        <h3 id="compare-all-title">All changes</h3>
        {changes.length === 0 && <p className="compare-status">No differences found.</p>}
        {changes.map((hunk) => (
          <HunkCard key={hunk.id} hunk={hunk} comparison={comparison} marked />
        ))}
      </section>
      <div className="compare-disclosures">
        {disclosures(comparison).map((line) => (
          <p key={line}>{line}</p>
        ))}
      </div>
    </aside>
  )
}

function HunkCard({ hunk, comparison, marked }: { hunk: ChangeHunk; comparison: ComparisonJson; marked?: boolean }) {
  const before = hunk.before && !(hunk.kind === "moved" && hunk.before === hunk.after) ? hunk.before : undefined
  return (
    <div className={`change-card compare-hunk${marked && hunk.untracked ? " is-untracked" : ""}`}>
      <div className="change-author compare-hunk-meta">
        <span>{hunk.id}</span>
        <span>{KIND[hunk.kind]}</span>
        {marked && hunk.untracked && <span className="compare-untracked-tag">Not tracked</span>}
      </div>
      <div className="compare-hunk-location">{hunkLocation(hunk.location, comparison)}</div>
      {before && <ClippedText className="change-old" text={before} />}
      {hunk.after && <ClippedText className="change-new" text={hunk.after} />}
    </div>
  )
}

// Paragraph text can run long; show the start and let the lawyer expand it.
const CLIP = 320
function ClippedText({ className, text }: { className: string; text: string }) {
  const [open, setOpen] = useState(false)
  const long = text.length > CLIP
  return (
    <>
      <div className={className}>{long && !open ? `${text.slice(0, CLIP).trimEnd()}…` : text}</div>
      {long && (
        <button type="button" className="linklike compare-more" onClick={() => setOpen((value) => !value)}>
          {open ? "Show less" : "Show all"}
        </button>
      )}
    </>
  )
}
