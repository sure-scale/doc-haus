import { STRIP_RE } from "../../../../dochaus/lib/untrusted-text"
import type { DocumentMarkup } from "../api/ingest/markup"
import { formatDateTime } from "../lib/format"
import { authorTones, initialsOf, nameKey, toneClass, toneOf } from "./comment-authors"
import { flagNotes } from "./MarkupPanel"

// Word's comment cards for the comments Docxodus sets beside the page. Its
// notes are restyled in styles.css; this pass adds what CSS cannot: each
// author's initials in the colour the Tracked changes panel gives them, the
// date as the panel writes it, screening flags, a link between a card and
// its commented text, and plain words for the "[0]" markers. On a narrow page
// each card moves under its paragraph and back again when the page widens.
// While the Tracked changes panel lists the comments, as Word's Comments pane
// does, the cards give way to a comment button beside each commented
// paragraph, which opens its thread in the panel.
// Everything it adds or moves is undone by the function it returns.

export type CommentDecoration = {
  // Every author the panel colours (its revisions' and comments' authors, in
  // its order), so a name has the same colour in both. Without it, the
  // authors named in the rendered page.
  authors?: string[]
  // The lawyer's own name, shown in the brand colour.
  self?: string | null
  // Word's initials and screening notes by comment id, from the file's markup.
  initials?: Map<string, string | null>
  flags?: Map<string, string[]>
  // The panel is open beside the page: show a thread there, by its root
  // comment's id, when its button or its text is clicked.
  onReveal?: (id: number) => void
}

// The same threshold as the stacked layout's container query (design-documents.css).
const NARROW = 560
const BLOCKS = "p,li,td,th,h1,h2,h3,h4,h5,h6"
const SVG = "http://www.w3.org/2000/svg"
// The characters ingest strips from the file's comments, and flags: bidi
// controls, zero-width and tag characters. A card shows its text as the
// panel and the agents read it, so a right-to-left override cannot show
// "0001" as "1000" beside a flag saying it was removed.
const visible = (text: string) => text.replace(STRIP_RE, "")

// What the panel knows about the comments in the working file. An older
// version's comments may carry other ids, so it gets none of this.
export function commentDecoration(markup: DocumentMarkup | undefined, self: string | null | undefined): CommentDecoration {
  if (!markup) return { self }
  const body = markup.comments.filter((comment) => !comment.scope)
  return {
    authors: [...markup.revisions.map((revision) => revision.author), ...markup.comments.map((comment) => comment.author)],
    self,
    initials: new Map(body.map((comment) => [String(comment.id), comment.initials])),
    flags: new Map(body.flatMap((comment) => {
      const notes = flagNotes(comment.findings)
      return notes.length ? [[String(comment.id), notes] as const] : []
    })),
  }
}

function icon(document: Document, paths: string[]) {
  const svg = document.createElementNS(SVG, "svg")
  for (const [name, value] of Object.entries({ class: "markup-icon", width: "14", height: "14", viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", "stroke-width": "2", "stroke-linecap": "round", "stroke-linejoin": "round", "aria-hidden": "true" }))
    svg.setAttribute(name, value)
  for (const d of paths) svg.appendChild(document.createElementNS(SVG, "path")).setAttribute("d", d)
  return svg
}

const active = new WeakMap<HTMLElement, () => void>()

export function decorateComments(root: HTMLElement, decoration: CommentDecoration = {}): () => void {
  active.get(root)?.()
  const column = root.querySelector<HTMLElement>(".comment-margin-column")
  const notes = column ? [...column.querySelectorAll<HTMLElement>(":scope > .comment-margin-note")] : []
  if (!column || !notes.length) return () => {}
  const document = root.ownerDocument
  const pane = decoration.onReveal
  const content = root.querySelector(".comment-margin-content") ?? root
  const all = [...column.querySelectorAll<HTMLElement>(".comment-margin-note")]
  const anchors = [...root.querySelectorAll<HTMLElement>(".comment-highlight[data-comment-id]")]
  const markers = [...root.querySelectorAll<HTMLElement>("a.comment-marker")]
  const added: Element[] = []
  const pins: HTMLElement[] = []
  const pinned: Element[] = []
  const element = <K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text?: string) => {
    const node = document.createElement(tag)
    node.className = className
    if (text !== undefined) node.textContent = text
    added.push(node)
    return node
  }
  // Each card's own text (author, date, comment) without them; the file's
  // text is kept to put back.
  const stripped: [Text, string][] = []
  for (const note of all) {
    const walker = document.createTreeWalker(note, NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
      const text = visible(node.data)
      if (text === node.data) continue
      stripped.push([node, node.data])
      node.data = text
    }
  }
  const authorOf = (note: HTMLElement) => visible(note.dataset.author ?? "").trim()

  // Colours exactly as the panel's: its list of names when given, else every
  // author the page names (tracked changes and comments alike).
  const names = decoration.authors ?? [...root.querySelectorAll<HTMLElement>("[data-author]")].map((node) => visible(node.dataset.author ?? ""))
  const tones = authorTones(names)
  const self = decoration.self?.trim()
  const toneFor = (author: string) => toneOf(tones, author, self ? (name) => nameKey(name) === nameKey(self) : undefined)
  const toneVariable = (author: string) => {
    const tone = toneFor(author)
    return tone === "self" ? "var(--accent)" : `var(--author-${tone + 1})`
  }
  const rootOf = (note: Element) => (note.parentElement?.closest<HTMLElement>(".comment-margin-note") ?? note) as HTMLElement
  const byId = new Map(all.map((note) => [note.dataset.commentId ?? "", note]))
  const named = (note: HTMLElement) => {
    const author = authorOf(note)
    return `${note.classList.contains("comment-resolved") ? "Resolved comment" : "Comment"}${author ? ` by ${author}` : ""}`
  }
  // The paragraph a thread's text is in: its highlight's, else its marker's.
  // A bookmark is painted as <a id="its name">, so one the file names
  // "comment-ref-N" is not the marker; only Docxodus's marker has the class.
  const blockOf = (note: HTMLElement) => {
    const id = CSS.escape(note.dataset.commentId ?? "")
    const hit = content.querySelector(`.comment-highlight[data-comment-id="${id}"]`) ?? content.querySelector(`a.comment-marker#comment-ref-${id}`)
    return hit?.closest(BLOCKS) ?? undefined
  }

  for (const note of all) {
    const id = note.dataset.commentId ?? ""
    const author = authorOf(note)
    const reply = note.classList.contains("comment-margin-reply")
    const resolved = note.classList.contains("comment-resolved")
    note.style.setProperty("--comment-tone", toneVariable(author))
    const header = note.querySelector<HTMLElement>(":scope > .comment-margin-note-header")
    if (header) {
      const avatar = element("span", `markup-avatar ${toneClass(toneFor(author))}${reply ? " is-small" : ""}`, initialsOf(author, decoration.initials?.get(id)))
      avatar.setAttribute("aria-hidden", "true")
      header.prepend(avatar)
      const when = formatDateTime(note.dataset.date)
      if (when) {
        const time = element("time", "doc-comment-time", when)
        time.dateTime = note.dataset.date!
        ;(header.querySelector(":scope > .comment-margin-date") ?? header.querySelector(":scope > .comment-margin-author"))?.after(time)
      }
      if (resolved && !header.querySelector(":scope > .comment-resolved-badge")) {
        const badge = element("span", "comment-resolved-badge", "Resolved")
        ;(header.querySelector(":scope > .doc-comment-time") ?? header.lastElementChild)?.after(badge)
      }
    }
    const flags = decoration.flags?.get(id)
    if (flags?.length) {
      const flag = element("div", "markup-flag doc-comment-flag")
      const text = document.createElement("span")
      for (const line of flags) text.appendChild(document.createElement("span")).textContent = line
      flag.append(icon(document, ["M12 3 2 21h20L12 3z", "M12 10v5M12 18h.01"]), text)
      const bodyText = note.querySelector(":scope > .comment-margin-note-body")
      if (bodyText) bodyText.after(flag)
      else note.append(flag)
    }
    if (!reply) {
      note.tabIndex = 0
      note.setAttribute("role", "note")
      const by = author ? ` by ${author}` : ""
      note.setAttribute("aria-label", `Comment${by}${resolved ? ", resolved" : ""}`)
    }
  }

  // The commented text takes its thread's colour, for the tint and underline.
  for (const anchor of anchors) {
    const note = byId.get(anchor.dataset.commentId ?? "")
    if (note) anchor.style.setProperty("--comment-tone", toneVariable(authorOf(rootOf(note))))
  }

  // "[0]" means nothing to a lawyer; the marker is kept, out of sight, as
  // words a screen reader reads at that point in the text.
  for (const marker of markers) {
    const id = marker.dataset.commentId ?? marker.id.replace(/^comment-ref-/, "")
    const note = byId.get(id)
    const author = visible(note?.dataset.author ?? marker.title.replace(/^Comment by /, "")).trim()
    const kind = marker.dataset.parentId !== undefined || note?.classList.contains("comment-margin-reply") ? "reply" : note?.classList.contains("comment-resolved") ? "resolved comment" : "comment"
    marker.dataset.markerText = marker.textContent ?? ""
    marker.textContent = `(${kind}${author ? ` by ${author}` : ""})`
    marker.classList.add("sr-only")
  }

  // With the panel open the cards are hidden (styles.css) and each thread has
  // a button in the right margin beside its paragraph, as Word marks a comment
  // while its Comments pane is open. It holds no text, so the text walks of
  // the markers and citations are unchanged.
  if (pane) {
    root.dataset.comments = "pane"
    const count = new Map<Element, number>()
    for (const note of notes) {
      const block = blockOf(note)
      if (!block) continue
      const pin = element("button", "doc-comment-pin")
      pin.type = "button"
      pin.dataset.commentId = note.dataset.commentId
      pin.title = named(note)
      pin.setAttribute("aria-label", named(note))
      pin.style.setProperty("--comment-tone", note.style.getPropertyValue("--comment-tone"))
      pin.style.setProperty("--pin-index", String(count.get(block) ?? 0))
      if (note.classList.contains("comment-resolved")) pin.classList.add("is-resolved")
      pin.append(icon(document, ["M5 4h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H9l-5 4V5a1 1 0 0 1 1-1z"]))
      if (!count.has(block)) {
        block.classList.add("doc-comment-block")
        pinned.push(block)
      }
      count.set(block, (count.get(block) ?? 0) + 1)
      block.append(pin)
      pins.push(pin)
    }
  }

  // A card, its button and its text light up together while any of them, or
  // the thread's card in the panel, is hovered or focused; a reply lights its
  // thread's.
  const host = root.closest<HTMLElement>(".viewer-body") ?? root
  const threadOf = (target: EventTarget | null) => {
    if (!(target instanceof Element) || !host.contains(target)) return undefined
    const own = target.closest<HTMLElement>(".doc-comment-pin, .markup-thread[data-comment-id]")
    if (own) return own.dataset.commentId
    if (!root.contains(target)) return undefined
    const note = target.closest(".comment-margin-note")
    if (note) return rootOf(note).dataset.commentId
    const id = target.closest<HTMLElement>(".comment-highlight[data-comment-id]")?.dataset.commentId
    const owner = id === undefined ? undefined : byId.get(id)
    return owner ? rootOf(owner).dataset.commentId : id
  }
  let hovered: string | undefined
  let focused: string | undefined
  let lit: string | undefined
  const light = () => {
    const next = hovered ?? focused
    if (next === lit) return
    for (const node of [...all, ...anchors, ...pins]) node.classList.remove("is-linked")
    lit = next
    if (next === undefined) return
    for (const node of [...anchors, ...pins]) if (node.dataset.commentId === next) node.classList.add("is-linked")
    byId.get(next)?.classList.add("is-linked")
  }
  const over = (event: Event) => { hovered = threadOf(event.target); light() }
  const out = (event: Event) => {
    const next = (event as MouseEvent).relatedTarget
    if (!(next instanceof Node) || !host.contains(next)) { hovered = undefined; light() }
  }
  const focusIn = (event: Event) => { focused = threadOf(event.target); light() }
  const focusOut = () => { focused = undefined; light() }
  // A click on a thread's button, or on its text (not a selection), opens it in the panel.
  const click = (event: Event) => {
    if (!pane || !(event.target instanceof Element)) return
    const pin = event.target.closest<HTMLElement>(".doc-comment-pin")
    const anchor = pin ? undefined : event.target.closest(".comment-highlight[data-comment-id]")
    if (anchor && document.getSelection()?.isCollapsed === false) return
    // Text outside every comment has no thread, not thread 0.
    const id = pin ? pin.dataset.commentId : anchor ? threadOf(anchor) : undefined
    if (id) pane(Number(id))
  }
  host.addEventListener("mouseover", over)
  host.addEventListener("mouseout", out)
  host.addEventListener("focusin", focusIn)
  host.addEventListener("focusout", focusOut)
  root.addEventListener("click", click)

  // On a narrow page each card moves (one copy, read once in reading order)
  // to just after the paragraph it comments on; several on one paragraph
  // share a stack. A cell or list item keeps its cards inside it. With the
  // panel open the cards are hidden and stay put.
  const stacks: HTMLElement[] = []
  let stacked = false
  const stack = () => {
    const placed = new Map<Element, HTMLElement>()
    for (const note of notes) {
      const block = blockOf(note)
      if (!block) continue
      let group = placed.get(block)
      if (!group) {
        group = document.createElement("div")
        group.className = "doc-comment-stack"
        if (block.matches("li,td,th")) block.append(group)
        else block.after(group)
        placed.set(block, group)
        stacks.push(group)
      }
      group.append(note)
    }
  }
  const unstack = () => {
    column.append(...notes)
    for (const group of stacks.splice(0)) group.remove()
  }
  const fit = (width: number) => {
    const narrow = !pane && width <= NARROW
    if (narrow === stacked) return
    stacked = narrow
    if (narrow) stack()
    else unstack()
  }
  const style = getComputedStyle(root)
  fit(root.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight))
  const observer = new ResizeObserver((entries) => {
    const entry = entries.at(-1)
    if (entry) fit(entry.contentRect.width)
  })
  observer.observe(root)

  const undo = () => {
    observer.disconnect()
    host.removeEventListener("mouseover", over)
    host.removeEventListener("mouseout", out)
    host.removeEventListener("focusin", focusIn)
    host.removeEventListener("focusout", focusOut)
    root.removeEventListener("click", click)
    if (stacked) unstack()
    for (const node of added.splice(0)) node.remove()
    for (const [node, text] of stripped.splice(0)) node.data = text
    for (const block of pinned.splice(0)) block.classList.remove("doc-comment-block")
    delete root.dataset.comments
    for (const node of [...all, ...anchors]) {
      node.style.removeProperty("--comment-tone")
      node.classList.remove("is-linked")
    }
    for (const note of notes) {
      note.removeAttribute("tabindex")
      note.removeAttribute("role")
      note.removeAttribute("aria-label")
    }
    for (const marker of markers) {
      if (marker.dataset.markerText !== undefined) marker.textContent = marker.dataset.markerText
      delete marker.dataset.markerText
      marker.classList.remove("sr-only")
    }
    if (active.get(root) === undo) active.delete(root)
  }
  active.set(root, undo)
  return undo
}
