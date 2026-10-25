import { docxodus } from "./docxodus"

// Body, headers, footers, footnotes and endnotes; never comments. When a
// comparison excludes page furniture its blackline keeps the left's furniture.
const STORIES = 31
const HEADERS_FOOTERS = 6

export type RedlineMismatch = { view: "accepted" | "baseline"; scope: string; paragraph: number }

/** One text contract for the preview, lean exporter and version comparison.
 * Native accept/reject handles nested revisions and moved paragraph marks.
 * Ignore empty paragraphs and whitespace cosmetics, preserving story/order. */
export async function redlineRoundTripMismatch(blackline: Uint8Array, left: Uint8Array, right: Uint8Array, headersFooters = true): Promise<RedlineMismatch | undefined> {
  const dx = await docxodus()
  const scope = headersFooters ? STORIES : STORIES & ~HEADERS_FOOTERS
  const text = (bytes: Uint8Array) => {
    const session = dx.openDocxSession(bytes, { persistAnchorIds: false })
    try {
      return session.grep("^[\\s\\S]*$", { scope }).flatMap(match => {
        const paragraph = match.text.normalize("NFC").replace(/\s+/g, " ").trim()
        return paragraph ? [{ scope: match.enclosingAnchor.scope, text: paragraph }] : []
      })
    } finally { session.close() }
  }
  // The comparer accepts revisions already pending in its inputs. They remain
  // proposals in the source; the exporter separately discloses that projection.
  const expectedLeft = text(await dx.docxDiffAcceptRevisions(left))
  const expectedRight = text(await dx.docxDiffAcceptRevisions(right))
  const projections = [
    ["baseline", text(await dx.docxDiffRejectRevisions(blackline)), expectedLeft],
    ["accepted", text(await dx.docxDiffAcceptRevisions(blackline)), expectedRight],
  ] as const
  for (const [view, actual, expected] of projections) {
    const ordinal = new Map<string, number>()
    for (let index = 0; index < Math.max(actual.length, expected.length); index++) {
      const row = expected[index] ?? actual[index]!
      const paragraph = (ordinal.get(row.scope) ?? 0) + 1
      ordinal.set(row.scope, paragraph)
      if (actual[index]?.scope !== expected[index]?.scope || actual[index]?.text !== expected[index]?.text)
        return { view, scope: row.scope, paragraph }
    }
  }
}

export async function assertRedlineRoundTrip(blackline: Uint8Array, left: Uint8Array, right: Uint8Array, headersFooters = true) {
  const mismatch = await redlineRoundTripMismatch(blackline, left, right, headersFooters)
  if (mismatch) throw new Error(`Word redline export refused: its ${mismatch.view === "accepted" ? "accept-all" : "reject-all"} text differs at ${mismatch.scope} paragraph ${mismatch.paragraph}. The source and intended document have not been changed.`)
}
