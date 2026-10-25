import type { DocumentMarkup, MarkupComment, MarkupLocation, MarkupRevision } from "../api/ingest/markup"
import { commentThreads } from "./MarkupPanel"

// Which painted paragraph each change and comment thread belongs to, for the
// change history card (#297). Docxodus paints each Word paragraph as one
// block in document order, which is how the reader counts body
// `locations[].paragraph`; comments match by their native id. Headers,
// footers and margin notes sit outside the section divs and never count.

export const PARAGRAPHS = "p, h1, h2, h3, h4, h5, h6"
// What opens the card: tracked changes, commented text, comment markers and
// whole inserted, deleted or reformatted paragraphs (a section's own change is
// painted on the paragraph that ends it).
export const TARGETS = [
  "ins", "del", ".rev-format-change", ".rev-para-mark-del", ".comment-highlight", "a.comment-marker",
  `:is(${PARAGRAPHS}):is(.rev-para-del, .rev-para-ins, .rev-para-format-change, .rev-section-format-change)`,
].join(", ")
// What the reader leaves out of a paragraph's preview, or the page adds to it.
const UNPAINTED = "del.rev-del:not(.rev-move-from), [data-list-marker], a.comment-marker, .rev-para-mark-del"

export type Thread = ReturnType<typeof commentThreads>[number]
// One painted paragraph with something to show: its first mark is the
// paragraph's one Tab stop.
export type Clause = {
  ordinal: number
  paragraph: Element
  mark: HTMLElement
  place: string
  revisions: MarkupRevision[]
  threads: Thread[]
}

export function paintedParagraphs(root: ParentNode): Element[] {
  return [...root.querySelectorAll("div[data-section-index]")].flatMap((section) => [...section.querySelectorAll(PARAGRAPHS)])
}

const letters = (text: string) => text.normalize("NFC").replace(/[^\p{L}\p{N}]/gu, "")

// The reader's preview is the paragraph's opening, "…" when cut. Only letters
// and digits are compared, so spacing, punctuation and pilcrows never matter.
export function contextMatches(context: string, painted: string) {
  const want = letters(context)
  const have = letters(painted)
  return context.endsWith("…") ? have.startsWith(want) : have === want
}

function paintedText(paragraph: Element) {
  const copy = paragraph.cloneNode(true) as Element
  for (const element of copy.querySelectorAll(UNPAINTED)) element.remove()
  return copy.textContent ?? ""
}

// Whether the page counts paragraphs as the reader does: every located body
// change's preview matches the paragraph painted at its ordinal. One mismatch
// (a text box the page paints inline, say) turns revisions off for the whole
// document, since a shifted count would put every later change on the wrong
// clause.
export function revisionsLineUp(paragraphs: Element[], revisions: MarkupRevision[]) {
  const texts = new Map<Element, string>()
  const textOf = (paragraph: Element) => texts.get(paragraph) ?? texts.set(paragraph, paintedText(paragraph)).get(paragraph)!
  return revisions.every((revision) => revision.locations.every((place) => {
    if (place.paragraph === null || !place.context) return true
    const paragraph = paragraphs[place.paragraph - 1]
    return paragraph !== undefined && contextMatches(place.context, textOf(paragraph))
  }))
}

// Comment ids painted in a paragraph: on a highlight or marker, or only in a
// root marker's `comment-ref-N` id.
function commentIds(paragraph: Element) {
  const ids = new Set<number>()
  const add = (value: string | null) => value && /^\d+$/.test(value) && ids.add(Number(value))
  for (const element of [paragraph, ...paragraph.querySelectorAll("[data-comment-id]")]) add(element.getAttribute("data-comment-id"))
  for (const marker of paragraph.querySelectorAll('a.comment-marker[id^="comment-ref-"]')) add(marker.id.slice("comment-ref-".length))
  return ids
}

const capital = (text: string) => text.charAt(0).toUpperCase() + text.slice(1)
const joined = (names: string[]) => (names.length > 1 ? `${names.slice(0, -1).join(", ")} and ${names.at(-1)}` : names[0] ?? "")
// The list number Word prints before a paragraph, when it names one (a bullet does not).
const listLabel = (place: Pick<MarkupLocation, "label">) => (place.label && /[\p{L}\p{N}]/u.test(place.label) ? place.label : undefined)

// A place as the lawyer sees it in Word: its list number ("clause 5",
// "Article 2"), else its paragraph's ordinal, as the "Their changes" panel says.
function placeName(place: Pick<MarkupLocation, "paragraph" | "label">) {
  const label = listLabel(place)
  if (label) return /^\p{L}{2,}\s/u.test(label) ? label : `clause ${label.replace(/\.$/, "")}`
  return place.paragraph !== null ? `paragraph ${place.paragraph}` : undefined
}

// Every painted paragraph with a change or comment thread on it, keyed by its
// element. Revisions are offered only when `revisions` allows them and the
// page lines up with the reader; comment threads match by id either way, and
// notes (which the page does not paint) are left out.
export function mapClauses(root: ParentNode, markup: Pick<DocumentMarkup, "revisions" | "comments">, revisions: boolean) {
  const paragraphs = paintedParagraphs(root)
  const body = revisions ? markup.revisions.filter((revision) => revision.scope === "body") : []
  const byOrdinal = new Map<number, MarkupRevision[]>()
  if (revisionsLineUp(paragraphs, body)) {
    for (const revision of body) {
      for (const ordinal of new Set(revision.locations.map((place) => place.paragraph))) {
        if (ordinal !== null) (byOrdinal.get(ordinal) ?? byOrdinal.set(ordinal, []).get(ordinal)!).push(revision)
      }
    }
  }
  const threads = commentThreads(markup.comments)
  const threadOf = new Map<number, Thread>()
  for (const thread of threads) for (const comment of [thread.root, ...thread.replies]) if (!comment.scope) threadOf.set(comment.id, thread)
  const clauses = new Map<Element, Clause>()
  paragraphs.forEach((paragraph, index) => {
    const ordinal = index + 1
    const own = byOrdinal.get(ordinal) ?? []
    const painted = new Set([...commentIds(paragraph)].flatMap((id) => threadOf.get(id) ?? []))
    const mine = threads.filter((thread) => painted.has(thread))
    if (!own.length && !mine.length) return
    const mark = paragraph.matches(TARGETS) ? paragraph : paragraph.querySelector(TARGETS)
    if (!(mark instanceof HTMLElement)) return
    const labelled = own.flatMap((revision) => revision.locations).find((place) => place.paragraph === ordinal && listLabel(place))
    clauses.set(paragraph, { ordinal, paragraph, mark, place: capital(labelled ? placeName(labelled)! : `paragraph ${ordinal}`), revisions: own, threads: mine })
  })
  return clauses
}

// A move seen from one of its places: where it came from or went to.
export function movePlace(revision: MarkupRevision, ordinal: number) {
  if (revision.kind !== "move") return
  const here = revision.locations.find((place) => place.paragraph === ordinal)
  const others = revision.locations.filter((place) => place.paragraph !== null && place.paragraph !== ordinal)
  if (!here || !others.length) return
  const facing = here.side ? others.filter((place) => place.side !== here.side) : others
  const where = joined([...new Set((facing.length ? facing : others).map((place) => placeName(place)!))])
  if (here.side === "to") return `Moved here from ${where}`
  if (here.side === "from") return `Moved to ${where}`
  return `Moved between here and ${where}`
}

export type TimelineItem =
  | { type: "revision"; revision: MarkupRevision }
  | { type: "comment"; comment: MarkupComment; parent?: MarkupComment; quote: string }
// One author at one time; `date` is null for every undated node, which come last.
export type TimelineNode = { author: string; date: string | null; items: TimelineItem[] }

const timeOf = (date: string | null) => {
  const time = date ? Date.parse(date) : Number.NaN
  return Number.isNaN(time) ? undefined : time
}

// The clause's changes and comments as how it got here: oldest first, ties
// and undated items in document order (changes as listed, then each thread's
// comments). Consecutive items by one author at one time are one node, so a
// replace by one reviewer reads as one step.
export function clauseTimeline(revisions: MarkupRevision[], threads: Thread[]): TimelineNode[] {
  const entries = [
    ...revisions.map((revision) => ({ author: revision.author, date: revision.date, item: { type: "revision", revision } as TimelineItem })),
    ...threads.flatMap(({ root, replies }) => {
      const byId = new Map([root, ...replies].map((comment) => [comment.id, comment]))
      return [root, ...replies].map((comment) => ({
        author: comment.author,
        date: comment.date,
        item: { type: "comment", comment, parent: comment.parent === null ? undefined : byId.get(comment.parent), quote: comment === root ? root.quote : "" } as TimelineItem,
      }))
    }),
  ].map((entry) => ({ ...entry, time: timeOf(entry.date) }))
  const dated = entries.filter((entry) => entry.time !== undefined).sort((a, b) => a.time! - b.time!)
  const nodes: (TimelineNode & { time?: number })[] = []
  for (const entry of [...dated, ...entries.filter((entry) => entry.time === undefined)]) {
    const last = nodes.at(-1)
    if (last && last.time === entry.time && last.author.trim() === entry.author.trim()) last.items.push(entry.item)
    else nodes.push({ author: entry.author, date: entry.time === undefined ? null : entry.date, time: entry.time, items: [entry.item] })
  }
  return nodes.map(({ author, date, items }) => ({ author, date, items }))
}
