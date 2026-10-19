import { INGEST_URL } from "../../config"

// Compare two Word versions of a matter document: a blackline (native tracked
// changes, earlier → later) plus a check for edits the later version made
// without tracked changes. One read-only ingest route serves both the JSON and
// the .docx; nothing is written to the matter and no model is called.

export type CompareOptions = { ignoreFormatting: boolean; detectMoves: boolean; headersFooters: boolean }

// left = the earlier version (e.g. what we sent); right = the later one (e.g.
// the draft that came back). Labels default to each file's name without its
// extension; the right label is the author of the blackline's changes in Word.
export type CompareParams = {
  left: string
  right: string
  leftLabel?: string
  rightLabel?: string
} & Partial<CompareOptions>

export type ChangeLocation = {
  scope: "body" | "header" | "footer" | "footnote" | "endnote"
  // 1-based block ordinal within the scope, counted in `side`.
  paragraph: number
  side: "left" | "right"
  // Nearest preceding body heading, with its Word numbering when present.
  heading?: string
  // The paragraph's own Word auto-number when present.
  number?: string
}

export type ChangeHunk = {
  // "C1"… for blackline changes, "U1"… for untracked edits, in document order.
  id: string
  kind: "modified" | "inserted" | "deleted" | "moved"
  location: ChangeLocation
  before?: string
  after?: string
  untracked: boolean
}

// Word tracked changes already pending in an input file.
export type SideRevisions = { count: number; authors: string[] }

export type ComparisonSide = {
  name: string
  label: string
  trackedChanges: SideRevisions
  // Pending doc.haus proposals on the stored file; not part of the comparison.
  pendingProposals: number
}

export type ComparisonJson = {
  left: ComparisonSide
  right: ComparisonSide
  options: CompareOptions
  // "authors": edits made without tracking carry their own author in the Word
  // file; "label": every change carries the right label, as Word Compare does.
  attribution: "authors" | "label"
  changes: ChangeHunk[]
  untracked: ChangeHunk[]
  blacklineName: string
}

// Option flags default to on at the route; send each one so the URL says
// exactly what was compared. A blank label is left to the route's default.
export function compareUrl(
  matterId: string,
  params: CompareParams,
  { format = "docx", download = false }: { format?: "json" | "docx"; download?: boolean } = {},
): string {
  const flag = (value: boolean | undefined) => (value === false ? "0" : "1")
  const query: [string, string | undefined][] = [
    ["left", params.left],
    ["right", params.right],
    ["leftLabel", params.leftLabel?.trim() || undefined],
    ["rightLabel", params.rightLabel?.trim() || undefined],
    ["ignoreFormatting", flag(params.ignoreFormatting)],
    ["moves", flag(params.detectMoves)],
    ["headers", flag(params.headersFooters)],
    ["format", format],
    ["download", download ? "1" : undefined],
  ]
  const search = query.flatMap(([key, value]) => (value === undefined ? [] : [`${key}=${encodeURIComponent(value)}`])).join("&")
  return `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/compare?${search}`
}

export async function fetchComparison(matterId: string, params: CompareParams): Promise<ComparisonJson> {
  const res = await fetch(compareUrl(matterId, params, { format: "json" }))
  if (!res.ok) throw new Error(await compareError(res))
  return res.json()
}

export async function fetchBlacklineBytes(matterId: string, params: CompareParams): Promise<Uint8Array> {
  const res = await fetch(compareUrl(matterId, params))
  if (!res.ok) throw new Error(await compareError(res))
  return new Uint8Array(await res.arrayBuffer())
}

// The route answers a refusal with { error: <code>, message: <plain words> };
// show the words, never the code.
async function compareError(res: Response) {
  const body = (await res.json().catch(() => undefined)) as { message?: unknown } | undefined
  return typeof body?.message === "string" && body.message.trim() ? body.message : `Could not compare the documents (${res.status})`
}
