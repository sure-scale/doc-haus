import { docxodus } from "./docxodus"

// Body, headers, footers, footnotes and endnotes; never comments. When a
// comparison excludes page furniture its blackline keeps the left's furniture.
const STORIES = 31
const HEADERS_FOOTERS = 6

export type RedlineMismatch = { view: "accepted" | "baseline"; scope: string; paragraph: number }
type Docxodus = Awaited<ReturnType<typeof docxodus>>

function storyText(dx: Docxodus, bytes: Uint8Array, scope: number) {
  const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
  try {
    return session.grep("^[\\s\\S]*$", { scope }).flatMap(match => {
      const paragraph = match.text.normalize("NFC").replace(/\s+/g, " ").trim()
      return paragraph ? [{ scope: match.enclosingAnchor.scope, text: paragraph }] : []
    })
  } finally { session.close() }
}

function textMismatch(view: RedlineMismatch["view"], actual: ReturnType<typeof storyText>, expected: ReturnType<typeof storyText>): RedlineMismatch | undefined {
  const ordinal = new Map<string, number>()
  for (let index = 0; index < Math.max(actual.length, expected.length); index++) {
    const row = expected[index] ?? actual[index]!
    const paragraph = (ordinal.get(row.scope) ?? 0) + 1
    ordinal.set(row.scope, paragraph)
    if (actual[index]?.scope !== expected[index]?.scope || actual[index]?.text !== expected[index]?.text)
      return { view, scope: row.scope, paragraph }
  }
}

/** One text contract for the preview, lean exporter and version comparison.
 * Native accept/reject handles nested revisions and moved paragraph marks.
 * Ignore empty paragraphs and whitespace cosmetics, preserving story/order. */
export async function redlineRoundTripMismatch(blackline: Uint8Array, left: Uint8Array, right: Uint8Array, headersFooters = true): Promise<RedlineMismatch | undefined> {
  const dx = await docxodus()
  const scope = headersFooters ? STORIES : STORIES & ~HEADERS_FOOTERS
  const text = (bytes: Uint8Array) => storyText(dx, bytes, scope)
  // The comparer accepts revisions already pending in its inputs. They remain
  // proposals in the source; the exporter separately discloses that projection.
  const expectedLeft = text(await dx.docxDiffAcceptRevisions(left))
  const expectedRight = text(await dx.docxDiffAcceptRevisions(right))
  const projections = [
    ["baseline", text(await dx.docxDiffRejectRevisions(blackline)), expectedLeft],
    ["accepted", text(await dx.docxDiffAcceptRevisions(blackline)), expectedRight],
  ] as const
  for (const [view, actual, expected] of projections) {
    const mismatch = textMismatch(view, actual, expected)
    if (mismatch) return mismatch
  }
}

export async function assertRedlineRoundTrip(blackline: Uint8Array, left: Uint8Array, right: Uint8Array, headersFooters = true) {
  const mismatch = await redlineRoundTripMismatch(blackline, left, right, headersFooters)
  if (mismatch) throw new Error(`Word redline export refused: its ${mismatch.view === "accepted" ? "accept-all" : "reject-all"} text differs at ${mismatch.scope} paragraph ${mismatch.paragraph}. The source and intended document have not been changed.`)
}

/** A layered export is the source edited in place, its own pending revisions
 * kept. The revisions it adds are the ids the source's native listing lacks.
 * Rejecting only those must give the source's text, and accepting only those
 * the baked text (the source with the same edits untracked); either way the
 * source's revisions remain, by id, author, type, text and date. */
export async function assertLayeredRoundTrip(layered: Uint8Array, source: Uint8Array, baked: Uint8Array) {
  const dx = await docxodus()
  const revisions = (bytes: Uint8Array) => {
    const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
    try {
      return session.listRevisions().map(({ id, author, type, text, date }) => ({ id, author, type, text, date }))
    } finally { session.close() }
  }
  const text = async (bytes: Uint8Array) => storyText(dx, await dx.docxDiffAcceptRevisions(bytes), STORIES)
  const theirs = revisions(source)
  const kept = new Set(theirs.map(row => row.id))
  const added = revisions(layered).filter(row => !kept.has(row.id)).map(row => row.id)
  for (const [view, expected] of [["baseline", source], ["accepted", baked]] as const) {
    const decision = view === "baseline" ? "rejecting" : "accepting"
    const session = dx.openDocxSession(layered, { persistAnchorIds: false })
    let decided: Uint8Array
    try {
      for (const id of added)
        if (!(view === "baseline" ? session.rejectRevision(id) : session.acceptRevision(id)).success)
          throw new Error(`Word redline export refused: ${decision} doc.haus's tracked change ${id} on its own failed. The source has not been changed.`)
      decided = session.save()
    } finally { session.close() }
    if (JSON.stringify(revisions(decided)) !== JSON.stringify(theirs))
      throw new Error(`Word redline export refused: ${decision} only doc.haus's changes alters the source's own tracked changes. The source has not been changed.`)
    const mismatch = textMismatch(view, await text(decided), await text(expected))
    if (mismatch)
      throw new Error(`Word redline export refused: ${decision} only doc.haus's changes leaves text that differs at ${mismatch.scope} paragraph ${mismatch.paragraph}. The source has not been changed.`)
  }
}
