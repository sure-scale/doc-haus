import { Fragment, useMemo, useState } from "react"

// A change's text as a redline: unchanged words plain, deletions struck,
// insertions underlined, moved text marked as moved. Shared by the
// comparison's cards, a chat answer's changes and doc.haus's own proposals.
export type Piece = { kind: "same" | "del" | "ins" | "move"; text: string }

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

// A doc.haus proposal as Word shows a tracked change: each run's deleted
// words before its inserted ones. A lone insertion or deletion that the words
// around it repeat could stand at several places; it stands at the nearest
// whose both ends are sentence boundaries, as the tracked write places it
// (changedRange in dochaus/lib/docx-amendments.ts), so a new sentence is never
// shown inside the one whose words it repeats.
export function proposalPieces(before: string, after: string): Piece[] {
  const pieces = wordDiff(before, after)
  for (let i = 0; i + 1 < pieces.length; i++)
    if (pieces[i]!.kind === "ins" && pieces[i + 1]!.kind === "del") [pieces[i], pieces[i + 1]] = [pieces[i + 1]!, pieces[i]!]
  const changed = pieces.filter((piece) => piece.kind !== "same")
  if (changed.length !== 1) return pieces
  const at = pieces.indexOf(changed[0]!)
  const head = tokens(at ? pieces[at - 1]!.text : ""), run = tokens(changed[0]!.text)
  const text = [...head, ...run, ...tokens(pieces[at + 1]?.text ?? "")]
  const start = head.length, end = start + run.length
  const space = (token?: string) => token !== undefined && /^\s/.test(token)
  const stop = (token?: string) => token !== undefined && /^[.!?]$/.test(token)
  const boundary = (i: number) => i === 0 || i === text.length || (space(text[i - 1]) && stop(text[i - 2])) || (stop(text[i - 1]) && space(text[i]))
  if (boundary(start) && boundary(end)) return pieces
  // Moved k tokens earlier, the run still reads the same where the k tokens
  // before it match its own last k.
  for (let k = 1; k <= start && text[start - k] === text[end - k]; k++)
    if (boundary(start - k) && boundary(end - k))
      return [
        { kind: "same", text: text.slice(0, start - k).join("") },
        { kind: changed[0]!.kind, text: text.slice(start - k, end - k).join("") },
        { kind: "same", text: text.slice(end - k).join("") },
      ].filter((piece): piece is Piece => !!piece.text)
  return pieces
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
