import { Fragment, useMemo, type ReactNode } from "react"
import ReactMarkdown from "react-markdown"
import type { CompareParams } from "../../api/ingest/compare"
import { OUTCOME_TAG, Redline, hunkPieces } from "../BlacklineViewer"
import MarkdownContent from "../markdown/MarkdownContent"
import { compareParams } from "./CompareView"
import type { ListedChange, VersionCompare } from "./turn-model"

// The Compare assistant's answer to a Word comparison, read like a reviewing
// pane: each change it names is one row with its id (a link into the
// blackline), its redline and the assistant's note. The redline comes from the
// comparison itself, never from the model's text, and so does a proposal's
// outcome.

// The comparison an answer can draw from: the turn's only one, with its
// changes listed.
export function chatCompare(compares: VersionCompare[] | undefined) {
  return compares?.length === 1 && compares[0]!.hunks.length ? compares[0] : undefined
}

type Segment = { markdown: string } | { ids: string[]; note: string; source: string }

// An entry opens with its change ids in bold, "**C3** · …" or "- **C2, C4** · …",
// and runs to a blank line or the next block. P-ids are the earlier version's
// own tracked changes.
const ENTRY = /^ {0,3}(?:[-*+]\s+|\d+[.)]\s+)?(\*\*\s*[CUP]\d+\b[^*]*\*\*.*)$/
const BLOCK = /^\s*(?:[-*+]\s|\d+[.)]\s|#|\||>)/

export function answerSegments(text: string): Segment[] {
  const out: Segment[] = []
  let markdown: string[] = []
  let entry: string[] | undefined
  const close = () => {
    if (entry) {
      const [, bold = "", after = ""] = /^\*\*([^*]*)\*\*(.*)$/.exec(entry[0]!) ?? []
      const note = [after.replace(/^\s*[·:—–-]\s*/, ""), ...entry.slice(1)].join(" ").trim()
      out.push({ ids: [...new Set(bold.match(/\b[CUP]\d+\b/g))], note, source: entry.join("\n") })
    }
    entry = undefined
  }
  for (const line of text.split("\n")) {
    const opens = ENTRY.exec(line)
    if (opens) {
      close()
      if (markdown.length) out.push({ markdown: markdown.join("\n") })
      markdown = []
      entry = [opens[1]!]
    } else if (entry && line.trim() && !BLOCK.test(line)) entry.push(line.trim())
    else {
      close()
      markdown.push(line)
    }
  }
  close()
  if (markdown.length) out.push({ markdown: markdown.join("\n") })
  return out
}

// "<clause> — Material — Favors <party>: <sentence>", as the prompt asks; any
// other wording reads as the note alone.
const META = /^(?:(.*?)\s+[—–]\s+)?(Material|Minor)\s+[—–]\s+(Favou?rs\b[^:]*?)\s*:\s*([\s\S]*)$/

export default function CompareAnswer({ text, compare, onOpen }: { text: string; compare: VersionCompare; onOpen?: (params: CompareParams, change: string) => void }) {
  const segments = useMemo(() => answerSegments(text), [text])
  const listed = useMemo(() => new Map(compare.hunks.map((hunk) => [hunk.id, hunk])), [compare])
  const show = onOpen && ((id: string) => onOpen(compareParams(compare), id))
  return (
    <div className="markdown compare-answer">
      {segments.map((segment, i) => {
        if ("markdown" in segment) return <MarkdownContent key={i}>{segment.markdown}</MarkdownContent>
        const changes = segment.ids.flatMap((id) => listed.get(id) ?? [])
        if (!changes.length) return <MarkdownContent key={i}>{segment.source}</MarkdownContent>
        return <ChangeRow key={i} changes={changes} note={segment.note} onShow={show} />
      })}
    </div>
  )
}

function ChangeRow({ changes, note, onShow }: { changes: ListedChange[]; note: string; onShow?: (id: string) => void }) {
  const proposal = changes.length === 1 && changes[0]!.outcome ? changes[0]! : undefined
  const [, clause, level, favors, sentence] = (!proposal && META.exec(note)) || []
  return (
    <div className="change-row">
      {changes.map((change) => (
        <Fragment key={change.id}>
          <ChangeId id={change.id} onShow={onShow} />
          <ChangeRedline change={change} />
        </Fragment>
      ))}
      <div className="change-meta">
        {proposal ? (
          <ProposalOutcome proposal={proposal} onShow={onShow} />
        ) : level ? (
          <>
            {clause && <span className="change-clause">{clause}</span>}
            <span className={`change-level is-${level.toLowerCase()}`}>{level}</span>
            <span className="change-favors">{favors}</span>
            {sentence && <Note>{sentence}</Note>}
          </>
        ) : (
          note && <Note>{note}</Note>
        )}
      </div>
    </div>
  )
}

function ProposalOutcome({ proposal: { outcome, untracked }, onShow }: { proposal: ListedChange; onShow?: (id: string) => void }) {
  const rejected = outcome === "rejected"
  return (
    <span className="change-outcome">
      {rejected ? "Rejected or changed — not shown as a tracked change" : OUTCOME_TAG[outcome!]}
      {untracked && (
        <>
          {rejected ? ", see " : " — same paragraph has an untracked edit, see "}
          <ChangeId id={untracked} onShow={onShow} inline />
        </>
      )}
    </span>
  )
}

function ChangeId({ id, onShow, inline }: { id: string; onShow?: (id: string) => void; inline?: boolean }) {
  const className = inline ? "change-ref" : "change-id"
  return onShow ? (
    <button type="button" className={className} title={`Show ${id} in the blackline`} onClick={() => onShow(id)}>
      {id}
    </button>
  ) : (
    <span className={className}>{id}</span>
  )
}

function ChangeRedline({ change }: { change: ListedChange }) {
  const pieces = useMemo(() => hunkPieces(change), [change])
  return (
    <div className="change-redline">
      <Redline pieces={pieces} />
    </div>
  )
}

// The assistant's sentence, with its inline markdown but no paragraph of its own.
const inline = { p: ({ children }: { children?: ReactNode }) => <>{children}</> }
function Note({ children }: { children: string }) {
  return (
    <span className="change-note">
      <ReactMarkdown components={inline}>{children}</ReactMarkdown>
    </span>
  )
}
