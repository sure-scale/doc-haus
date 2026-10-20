import { useMemo } from "react"
import type { CompareParams } from "../../api/ingest/compare"
import { Redline, hunkPieces } from "../BlacklineViewer"
import MarkdownContent from "../markdown/MarkdownContent"
import { compareParams } from "./CompareView"
import type { ListedChange, VersionCompare } from "./turn-model"

// The Compare assistant's answer to a Word comparison: its own words, with each
// change an entry names drawn under it as a redline. The redline comes from the
// comparison itself, never from the model's text.

// The comparison an answer can draw from: the turn's only one, with its
// changes listed.
export function chatCompare(compares: VersionCompare[] | undefined) {
  return compares?.length === 1 && compares[0]!.hunks.length ? compares[0] : undefined
}

type Segment = { markdown: string } | { entry: string; ids: string[] }

// An entry opens with its change ids in bold, "**C3** · …" or "- **C2, C4** · …",
// and runs to a blank line or the next block.
const ENTRY = /^ {0,3}(?:[-*+]\s+|\d+[.)]\s+)?(\*\*\s*[CU]\d+\b[^*]*\*\*.*)$/
const BLOCK = /^\s*(?:[-*+]\s|\d+[.)]\s|#|\||>)/

export function answerSegments(text: string): Segment[] {
  const out: Segment[] = []
  let markdown: string[] = []
  let entry: string[] | undefined
  const close = () => {
    if (entry) out.push({ entry: entry.join("\n"), ids: [...new Set(entry[0]!.split("**")[1]!.match(/\b[CU]\d+\b/g))] })
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

export default function CompareAnswer({ text, compare, onOpen }: { text: string; compare: VersionCompare; onOpen?: (params: CompareParams, change: string) => void }) {
  const segments = useMemo(() => answerSegments(text), [text])
  const listed = useMemo(() => new Map(compare.hunks.map((hunk) => [hunk.id, hunk])), [compare])
  return (
    <div className="markdown compare-answer">
      {segments.map((segment, i) =>
        "markdown" in segment ? (
          <MarkdownContent key={i}>{segment.markdown}</MarkdownContent>
        ) : (
          <div key={i} className="compare-entry">
            <MarkdownContent>{segment.entry}</MarkdownContent>
            {segment.ids.map((id) => {
              const hunk = listed.get(id)
              return hunk && <EntryChange key={id} hunk={hunk} label={segment.ids.length > 1} onShow={onOpen && (() => onOpen(compareParams(compare), id))} />
            })}
          </div>
        ),
      )}
    </div>
  )
}

function EntryChange({ hunk, label, onShow }: { hunk: ListedChange; label: boolean; onShow?: () => void }) {
  const pieces = useMemo(() => hunkPieces(hunk), [hunk])
  return (
    <div className="compare-entry-change">
      {label && <span className="compare-entry-id">{hunk.id}</span>}
      <Redline pieces={pieces} />
      {onShow && (
        <button type="button" className="linklike compare-show" onClick={onShow} title={`Show ${hunk.id} in the blackline`}>
          Show in blackline
        </button>
      )}
    </div>
  )
}
