import { Fragment, useMemo, useState } from "react"
import { wordChanges } from "../../../../dochaus/lib/word-changes"

import { excerpt, type Piece } from "../../../../dochaus/lib/redline"

export { wordDiff, excerpt, type Piece } from "../../../../dochaus/lib/redline"

// A doc.haus proposal as its tracked change is written: the same changes the
// docx writer makes (wordChanges in dochaus/lib/word-changes.ts), each one's
// deleted words before its inserted ones, the words between them plain.
export function proposalPieces(before: string, after: string): Piece[] {
  let at = 0
  const pieces = wordChanges(before, after).flatMap(({ offset, replaced, changed }): Piece[] => {
    const same = before.slice(at, offset)
    at = offset + replaced.length
    return [{ kind: "same", text: same }, { kind: "del", text: replaced }, { kind: "ins", text: changed }]
  })
  return [...pieces, { kind: "same" as const, text: before.slice(at) }].filter((piece) => piece.text)
}

// What a screen reader hears before a marked run, as a change card's badge
// says it for a whole change.
const SAID: Record<Exclude<Piece["kind"], "same">, string> = { del: "Deleted: ", ins: "Inserted: ", move: "Moved: " }

// With `labelled`, each marked run also says what it is to a screen reader.
export function Redline({ pieces, labelled }: { pieces: Piece[]; labelled?: boolean }) {
  const [open, setOpen] = useState(false)
  const short = useMemo(() => excerpt(pieces), [pieces])
  return (
    <>
      <div className="compare-redline">
        {(open ? pieces : short).map((piece, index) => {
          if (piece.kind === "same") return <Fragment key={index}>{piece.text}</Fragment>
          const said = labelled && <span className="sr-only">{SAID[piece.kind]}</span>
          return piece.kind === "del" ? (
            <del key={index}>{said}{piece.text}</del>
          ) : (
            <ins key={index} className={piece.kind === "move" ? "is-move" : undefined}>
              {said}{piece.text}
            </ins>
          )
        })}
      </div>
      {short !== pieces && (
        <button type="button" className="linklike compare-more" onClick={() => setOpen((value) => !value)}>
          {open ? "Show less" : "Show all"}
        </button>
      )}
    </>
  )
}

// A doc.haus proposal in its card: one passage, only its changed words marked.
export function ProposalRedline({ before, after }: { before: string; after: string }) {
  const pieces = useMemo(() => proposalPieces(before, after), [before, after])
  return (
    <div className="proposal-redline">
      <Redline pieces={pieces} labelled />
    </div>
  )
}
