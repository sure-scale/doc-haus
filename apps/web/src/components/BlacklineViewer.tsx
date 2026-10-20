import { Fragment, useEffect, useLayoutEffect, useMemo, useRef, useState, type MouseEvent } from "react"
import { RevisionPresentation } from "docxodus"
import { useDocxodus } from "docxodus/react"
import {
  compareUrl,
  fetchBlacklineBytes,
  fetchComparison,
  type ChangeHunk,
  type ChangeLocation,
  type CompareParams,
  type ComparisonJson,
  type ComparisonSide,
  type ProposalOutcome,
} from "../api/ingest/compare"
import { humanizeError } from "../api/ingest/errors"
import { fileStem } from "./CompareDialog"
import { DOCX_RENDER_OPTIONS, ViewerTitle } from "./DocumentViewer"
import { ModalClose } from "./ModalHeader"

// The blackline of two Word versions beside a list of the edits the later
// version made without tracked changes, what became of the earlier version's
// own tracked changes, and then every change. View and download only: the
// blackline is generated on demand and never saved into the matter.
//
// The page uses Word's markup (a change bar per changed paragraph, green
// double-lined moves, no fills), coloured as comparisons usually are: blue
// double-underlined insertions and red struck deletions. Edits made without
// tracked changes sit on an amber highlight, as their own author sets them
// apart in the Word file.
export const BLACKLINE_RENDER_OPTIONS = { ...DOCX_RENDER_OPTIONS, revisionPresentation: RevisionPresentation.Word }

// focus: a change (C- or U-id) to show once the page is painted.
export default function BlacklineViewer({ matterId, params, focus, onClose }: { matterId: string; params: CompareParams; focus?: string; onClose: () => void }) {
  const { isReady, error: wasmError, convertToHtml } = useDocxodus("/wasm/")
  const panel = useRef<HTMLDivElement>(null)
  const page = useRef<HTMLDivElement>(null)
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
    convertToHtml(blackline, BLACKLINE_RENDER_OPTIONS).then(
      (out) => !cancelled && setHtml(out),
      (cause) => !cancelled && setRenderError(humanizeError(cause, "Could not show the blackline.")),
    )
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isReady, blackline])

  // Mark the edits made without tracked changes in the page itself. They carry
  // their own author only when the blackline's re-labelling held.
  useEffect(() => {
    const root = page.current
    if (!root || !html || comparison?.attribution !== "authors") return
    const author = untrackedAuthor(comparison.right.label)
    for (const mark of root.querySelectorAll<HTMLElement>("ins[data-author], del[data-author]"))
      mark.classList.toggle("is-untracked", mark.dataset.author === author)
  }, [html, comparison])

  const showChange = (hunk: ChangeHunk) => {
    const root = page.current?.querySelector<HTMLElement>(".docx-render:not(.docx-skeleton)")
    if (root) focusChange(root, hunk)
  }

  // Opened from a chat answer's change: show it in the page and in the list.
  useEffect(() => {
    const hunk = focus && html && comparison && [...comparison.untracked, ...comparison.changes].find((change) => change.id === focus)
    if (!hunk) return
    document.getElementById(`compare-${hunk.id}`)?.scrollIntoView({ block: "nearest" })
    showChange(hunk)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focus, html, comparison])

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
          <div ref={page} className="viewer-doc">
            {failed && <p className="muted">{failed}</p>}
            {!failed && !html && (
              <div className="docx-render docx-skeleton" aria-label="Comparing documents">
                {Array.from({ length: 14 }, (_, i) => <div key={i} className="sk-line" />)}
              </div>
            )}
            {documentPreview}
          </div>
          {comparison ? (
            <ComparePanel comparison={comparison} onShow={html ? showChange : undefined} />
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
// The Word author of an edit made without tracked changes, as the blackline sets it.
const untrackedAuthor = (label: string) => `Not tracked (${label})`

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

// In summary order.
const OUTCOME: Record<ProposalOutcome["outcome"], string> = {
  pending: "still pending",
  countered: "reversed by their tracked change",
  accepted: "accepted",
  rejected: "rejected or changed",
}

// "3 tracked changes: 1 still pending, 1 accepted, 1 rejected or changed."
// Outcomes that did not happen are left out.
function proposalSummary(proposals: ProposalOutcome[]) {
  const parts = (Object.keys(OUTCOME) as ProposalOutcome["outcome"][]).flatMap((outcome) => {
    const count = proposals.filter((proposal) => proposal.outcome === outcome).length
    return count ? [`${count} ${OUTCOME[outcome]}`] : []
  })
  return `${plural(proposals.length, "tracked change")}: ${parts.join(", ")}.`
}

const hadRevisions = ({ label, trackedChanges: { count, authors } }: ComparisonSide) =>
  `${label} had ${plural(count, "tracked change")}${authors.length ? ` (by ${authors.join(", ")})` : ""}`

// What the comparison did with the inputs and the Word file, in plain words.
function disclosures({ left, right, options, attribution, untracked, proposals = [] }: ComparisonJson) {
  const lines: string[] = []
  if (attribution === "label") lines.push(`Untracked edits are listed here only; in the Word file every change shows under ${right.label}.`)
  else if (untracked.length) lines.push(`In the Word file, edits made without tracked changes show under “${untrackedAuthor(right.label)}”.`)
  // The blackline reads left with its own tracked changes accepted; the
  // section listing them, not this line, says what became of each.
  const leftCount = left.trackedChanges.count
  if (leftCount)
    lines.push(
      `${hadRevisions(left)}; the blackline compares its text with ${leftCount === 1 ? "that change" : "those changes"} accepted.` +
        (proposals.length ? ` “${trackedIn(left)}” above shows what became of each in ${right.label}.` : ""),
    )
  if (right.trackedChanges.count) lines.push(`${hadRevisions(right)}; compared as accepted.`)
  // Not "pending": that word belongs to the earlier version's tracked changes.
  for (const side of [left, right])
    if (side.pendingProposals)
      lines.push(`${side.label} has ${plural(side.pendingProposals, "change")} in doc.haus that ${side.pendingProposals === 1 ? "is" : "are"} not in the file yet, so ${side.pendingProposals === 1 ? "it is" : "they are"} not included.`)
  if (options.ignoreFormatting) lines.push("Formatting-only changes are left out.")
  if (!options.detectMoves) lines.push("Moved text shows as deleted and inserted.")
  if (!options.headersFooters) lines.push("Headers and footers were not compared.")
  return lines
}

// The earlier version's own tracked changes, named neutrally: the earlier file
// is not always ours.
const trackedIn = (side: ComparisonSide) => `Tracked changes in ${side.label}`

// onShow, once the blackline is painted, scrolls it to a change.
export function ComparePanel({ comparison, onShow }: { comparison: ComparisonJson; onShow?: (hunk: ChangeHunk) => void }) {
  const { left, right, changes, untracked, proposals = [] } = comparison
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
        <Legend comparison={comparison} />
      </div>
      <section className={`compare-section compare-untracked is-${status.tone}`} aria-labelledby="compare-untracked-title">
        <h3 id="compare-untracked-title">
          Edits made without tracked changes
          {untracked.length > 0 && <span className="compare-count">{untracked.length}</span>}
        </h3>
        <p className="compare-status">{status.text}</p>
        {untracked.map((hunk) => (
          <HunkCard
            key={hunk.id}
            hunk={hunk}
            comparison={comparison}
            onShow={onShow}
            removes={proposals.filter((proposal) => proposal.outcome === "rejected" && proposal.untracked === hunk.id).map((proposal) => proposal.id)}
          />
        ))}
      </section>
      {proposals.length > 0 && (
        <section className="compare-section compare-proposals" aria-labelledby="compare-proposals-title">
          <h3 id="compare-proposals-title">
            {trackedIn(left)}
            <span className="compare-count">{proposals.length}</span>
          </h3>
          <p className="compare-status">{proposalSummary(proposals)}</p>
          {proposals.map((proposal) => (
            <ProposalCard key={proposal.id} proposal={proposal} comparison={comparison} />
          ))}
        </section>
      )}
      <section className="compare-section" aria-labelledby="compare-all-title">
        <h3 id="compare-all-title">All changes</h3>
        {changes.length === 0 && <p className="compare-status">No differences found.</p>}
        {changes.map((hunk) => (
          <HunkCard key={hunk.id} hunk={hunk} comparison={comparison} marked onShow={onShow} />
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

// The marks in use, as a comparison's key would list them.
function Legend({ comparison: { options, attribution, untracked } }: { comparison: ComparisonJson }) {
  return (
    <div className="compare-legend" aria-label="Key">
      <span className="redline-ins">Inserted</span>
      <span className="redline-del">Deleted</span>
      {options.detectMoves && <span className="redline-move">Moved</span>}
      {attribution === "authors" && untracked.length > 0 && <span className="redline-untracked">Not tracked</span>}
    </div>
  )
}

// One change, shown as its own redline. An untracked edit names the earlier
// version's tracked changes it removed, listed below. With onShow, a click on
// the card (or its place, for the keyboard) shows the change in the blackline.
function HunkCard({ hunk, comparison, marked, removes = [], onShow }: { hunk: ChangeHunk; comparison: ComparisonJson; marked?: boolean; removes?: string[]; onShow?: (hunk: ChangeHunk) => void }) {
  const pieces = useMemo(() => hunkPieces(hunk), [hunk])
  const location = hunkLocation(hunk.location, comparison)
  // Not when the click lands on a control or ends a text selection.
  const show = onShow && ((event: MouseEvent) => {
    if (!(event.target as Element).closest("button") && !window.getSelection()?.toString()) onShow(hunk)
  })
  return (
    <div
      id={`compare-${hunk.id}`}
      className={`change-card compare-hunk${marked && hunk.untracked ? " is-untracked" : ""}${onShow ? " is-linked" : ""}`}
      onClick={show}
    >
      <div className="change-author compare-hunk-meta">
        <span>{hunk.id}</span>
        <span>{KIND[hunk.kind]}</span>
        {marked && hunk.untracked && <span className="compare-untracked-tag">Not tracked</span>}
        {removes.length > 0 && (
          <span className="compare-outcome-tag is-warn">
            {`Removes tracked ${removes.length === 1 ? "change" : "changes"} `}
            {removes.map((id, index) => (
              <Fragment key={id}>
                {index > 0 && ", "}
                <CompareRef id={id} section={trackedIn(comparison.left)} />
              </Fragment>
            ))}
          </span>
        )}
      </div>
      <div className="compare-hunk-location">
        {onShow ? (
          <button type="button" className="linklike compare-show" title="Show this change in the blackline" onClick={() => onShow(hunk)}>
            {location}
          </button>
        ) : (
          location
        )}
      </div>
      <Redline pieces={pieces} />
    </div>
  )
}

const OUTCOME_TAG: Record<Exclude<ProposalOutcome["outcome"], "rejected">, string> = {
  pending: "Still pending",
  countered: "Still pending — their tracked change reverses it",
  accepted: "Accepted",
}

// One tracked change pending in the earlier version and what became of it. A
// rejected one names the untracked edit that removed it; any other names an
// untracked edit in its paragraph. Both edits sit above. Not a blackline
// change itself, so it has no place in the page to show.
function ProposalCard({ proposal, comparison }: { proposal: ProposalOutcome; comparison: ComparisonJson }) {
  const rejected = proposal.outcome === "rejected"
  const ref = proposal.untracked
  const refButton = ref && <CompareRef id={ref} section="Edits made without tracked changes" />
  return (
    <div id={`compare-${proposal.id}`} className={`change-card compare-hunk compare-proposal${rejected ? " is-rejected" : ""}`}>
      <div className="change-author compare-hunk-meta">
        <span>{proposal.id}</span>
        <span>{proposal.kind === "inserted" ? "Insertion" : "Deletion"}</span>
        {proposal.author && <span>{proposal.author}</span>}
        {proposal.outcome === "rejected" ? (
          <span className="compare-outcome-tag is-warn">
            {`Rejected or changed — not shown as a tracked change${ref ? ", see " : ""}`}
            {refButton}
          </span>
        ) : (
          <>
            <span className="compare-outcome-tag">{OUTCOME_TAG[proposal.outcome]}</span>
            {ref && (
              <span className="compare-outcome-tag is-warn">
                {"Same paragraph has an untracked edit — see "}
                {refButton}
              </span>
            )}
          </>
        )}
      </div>
      <div className="compare-hunk-location">{hunkLocation(proposal.location, comparison)}</div>
      <Redline pieces={[{ kind: proposal.kind === "inserted" ? "ins" : "del", text: proposal.text }]} />
    </div>
  )
}

// A tag's link to another card in this panel.
function CompareRef({ id, section }: { id: string; section: string }) {
  return (
    <button
      type="button"
      className="linklike compare-ref"
      title={`Show ${id} under ${section}`}
      onClick={() => document.getElementById(`compare-${id}`)?.scrollIntoView({ block: "center", behavior: "smooth" })}
    >
      {id}
    </button>
  )
}

// A change's text as a redline: unchanged words plain, deletions struck,
// insertions underlined, moved text marked as moved.
export type Piece = { kind: "same" | "del" | "ins" | "move"; text: string }

export function hunkPieces(hunk: Pick<ChangeHunk, "kind" | "before" | "after">): Piece[] {
  if (hunk.kind === "moved" && hunk.before === hunk.after) return [{ kind: "move", text: hunk.after ?? "" }]
  if (!hunk.before) return [{ kind: "ins", text: hunk.after ?? "" }]
  if (!hunk.after) return [{ kind: "del", text: hunk.before }]
  return wordDiff(hunk.before, hunk.after)
}

// Words, runs of white space and single other characters, so a changed word
// reads whole.
const tokens = (text: string) => text.match(/\s+|[\p{L}\p{N}\p{M}]+|[^\s\p{L}\p{N}\p{M}]/gu) ?? []

// Word-level redline of one change. The common start and end are set aside
// first, so the table covers only the changed middle; a middle too large to
// align reads as one deletion and one insertion.
export function wordDiff(before: string, after: string): Piece[] {
  const a = tokens(before), b = tokens(after)
  let head = 0
  while (head < a.length && head < b.length && a[head] === b[head]) head++
  let tail = 0
  while (tail < a.length - head && tail < b.length - head && a[a.length - 1 - tail] === b[b.length - 1 - tail]) tail++
  const x = a.slice(head, a.length - tail), y = b.slice(head, b.length - tail)
  const middle: Piece[] = []
  if (x.length * y.length > 250_000) middle.push({ kind: "del", text: x.join("") }, { kind: "ins", text: y.join("") })
  else {
    // Longest common run of tokens from x[i] and y[j] onwards.
    const width = y.length + 1
    const common = new Uint32Array((x.length + 1) * width)
    for (let i = x.length - 1; i >= 0; i--)
      for (let j = y.length - 1; j >= 0; j--)
        common[i * width + j] = x[i] === y[j] ? common[(i + 1) * width + j + 1] + 1 : Math.max(common[(i + 1) * width + j], common[i * width + j + 1])
    let i = 0, j = 0
    while (i < x.length || j < y.length) {
      if (i < x.length && j < y.length && x[i] === y[j]) middle.push({ kind: "same", text: x[i++] }), j++
      else if (i < x.length && (j === y.length || common[(i + 1) * width + j] >= common[i * width + j + 1])) middle.push({ kind: "del", text: x[i++] })
      else middle.push({ kind: "ins", text: y[j++] })
    }
  }
  return merged([{ kind: "same", text: a.slice(0, head).join("") }, ...middle, { kind: "same", text: a.slice(a.length - tail).join("") }])
}

// Each run of changes reads as one insertion then one deletion, the order the
// blackline itself uses. White space alone between two changes joins the
// run, and goes to each side that continues after it.
function merged(pieces: Piece[]): Piece[] {
  const out: Piece[] = []
  let del = "", ins = "", gap = { del: "", ins: "" }
  const flush = () => {
    if (ins) out.push({ kind: "ins", text: ins })
    if (del) out.push({ kind: "del", text: del })
    del = ins = ""
    gap = { del: "", ins: "" }
  }
  pieces.forEach((piece, index) => {
    if (!piece.text) return
    if (piece.kind === "del") (del += (del ? gap.del : "") + piece.text), (gap.del = "")
    else if (piece.kind === "ins") (ins += (ins ? gap.ins : "") + piece.text), (gap.ins = "")
    else if ((del || ins) && !piece.text.trim() && pieces.slice(index + 1).find((next) => next.text)?.kind !== "same")
      (gap.del += piece.text), (gap.ins += piece.text)
    else {
      flush()
      const last = out.at(-1)
      if (last?.kind === "same") last.text += piece.text
      else out.push({ ...piece })
    }
  })
  flush()
  return out
}

// Unchanged words kept either side of a change, and the characters of an
// inserted or deleted run shown before "Show all".
const CONTEXT = 12
const CLIP = 320

// The redline as first shown: long unchanged text cut to the words nearest a
// change and long runs clipped. The same array when nothing was cut.
export function excerpt(pieces: Piece[]): Piece[] {
  let cut = false
  const out = pieces.map((piece, index): Piece => {
    if (piece.kind !== "same") {
      if (piece.text.length <= CLIP) return piece
      cut = true
      return { ...piece, text: `${piece.text.slice(0, CLIP).trimEnd()}…` }
    }
    // Each word with the white space around it.
    const words = piece.text.match(/\s*\S+\s*/g) ?? []
    const first = index === 0, last = index === pieces.length - 1
    if (first && words.length > CONTEXT + 3) return (cut = true), { ...piece, text: `… ${words.slice(-CONTEXT).join("")}` }
    if (last && words.length > CONTEXT + 3) return (cut = true), { ...piece, text: `${words.slice(0, CONTEXT).join("").trimEnd()}…` }
    if (!first && !last && words.length > 2 * CONTEXT + 3)
      return (cut = true), { ...piece, text: `${words.slice(0, CONTEXT).join("")}… ${words.slice(-CONTEXT).join("")}` }
    return piece
  })
  return cut ? out : pieces
}

export function Redline({ pieces }: { pieces: Piece[] }) {
  const [open, setOpen] = useState(false)
  const short = useMemo(() => excerpt(pieces), [pieces])
  return (
    <>
      <div className="compare-redline">
        {(open ? pieces : short).map((piece, index) =>
          piece.kind === "same" ? (
            <Fragment key={index}>{piece.text}</Fragment>
          ) : piece.kind === "del" ? (
            <del key={index}>{piece.text}</del>
          ) : (
            <ins key={index} className={piece.kind === "move" ? "is-move" : undefined}>
              {piece.text}
            </ins>
          ),
        )}
      </div>
      {short !== pieces && (
        <button type="button" className="linklike compare-more" onClick={() => setOpen((value) => !value)}>
          {open ? "Show less" : "Show all"}
        </button>
      )}
    </>
  )
}

// A paragraph of the painted blackline, read as each version has it: before
// without its insertions, after without its deletions.
export type BlockText = { before: string; after: string; marked: boolean }
const squash = (text: string) => text.replace(/\s+/g, " ").trim()
const words = (text: string) => text.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? []

// The paragraph a change sits in: the one whose text, read as the change's
// side, holds the change's first line. Repeated wording prefers a paragraph
// with a revision, then the one nearest the change's own place. An untracked
// edit is read without their tracked changes, so where those share its
// paragraph no text holds it whole; the revised paragraph sharing most of its
// words is taken instead.
export function changeBlock(blocks: BlockText[], hunk: ChangeHunk): number | undefined {
  const side = hunk.after ? "after" : "before"
  const line = squash((hunk[side] ?? "").split("\n").find((text) => text.trim()) ?? "")
  if (!line) return undefined
  const distance = (index: number) => Math.abs(index + 1 - hunk.location.paragraph)
  const holders = blocks.flatMap((block, index) => (block[side].includes(line) ? [index] : []))
  if (holders.length) return holders.sort((a, b) => Number(blocks[b].marked) - Number(blocks[a].marked) || distance(a) - distance(b))[0]
  const wanted = new Set(words(line))
  let best: number | undefined, score = 0.5
  blocks.forEach((block, index) => {
    if (!block.marked) return
    const have = new Set(words(block[side]))
    const shared = [...wanted].filter((word) => have.has(word)).length / wanted.size
    if (shared > score) (best = index), (score = shared)
  })
  return best
}

// Leaf blocks only: a list item holding paragraphs is read through them.
const BLOCKS = "p, li, h1, h2, h3, h4, h5, h6"

function blockText(block: Element, drop: string) {
  const copy = block.cloneNode(true) as Element
  for (const node of copy.querySelectorAll(drop)) node.remove()
  return squash(copy.textContent ?? "")
}

// Scroll the painted blackline to a change's paragraph, flash it and move
// focus there, as a click in Word's revisions list does.
function focusChange(root: HTMLElement, hunk: ChangeHunk) {
  const blocks = [...root.querySelectorAll<HTMLElement>(BLOCKS)].filter((block) => !block.querySelector(BLOCKS))
  const index = changeBlock(
    blocks.map((block) => ({ before: blockText(block, "ins"), after: blockText(block, "del"), marked: !!block.querySelector("ins, del") })),
    hunk,
  )
  if (index === undefined) return
  const block = blocks[index]
  for (const shown of root.querySelectorAll(".compare-focus")) shown.classList.remove("compare-focus")
  void block.offsetWidth // restarts the flash when the same change is shown again
  block.classList.add("compare-focus")
  block.tabIndex = -1
  block.scrollIntoView({ block: "center", behavior: "smooth" })
  block.focus({ preventScroll: true })
}
