import { Fragment, useEffect, useId, useLayoutEffect, useRef, useState, type RefObject } from "react"
import type { DocumentMarkup, MarkupFinding, MarkupRevision } from "../api/ingest/markup"
import { formatDateTime } from "../lib/format"
import { changeLabel, flagNotes } from "./MarkupPanel"
import { clauseTimeline, CONTAINERS, mapClauses, markAt, movePlace, PARAGRAPHS, TARGETS, type Clause, type TimelineItem } from "./change-timeline"
import "./ChangeTimeline.css"

// A change history card on marked-up text in the document viewer (#297): who
// changed the clause under the pointer or keyboard focus, when, and what it
// said, oldest first, from the markup list "Their changes" already reads.
// Mouse and pen open it after a short rest and keep it open while on the
// clause's marks or the card; focus on a clause's first mark opens it at once
// and Enter moves into it; a tap opens it with a Close button. A click on text
// never opens, pins or closes it, and a drag never opens it, so selecting text
// is unaffected. Every name, change, quote and comment is the other side's
// text and renders as plain text.

type Point = { x: number; y: number }
type Opened = { clause: Clause; anchor: Element; point?: Point; touch: boolean }

// How long the pointer rests before the card opens, switches or closes.
const REST = 300
// Space between the card and its line, and between the card and the window edge.
const GAP = 8
const MARGIN = 8
const SIDES = ["below", "above", "right", "left"] as const
// What Tab reaches, as the viewer's focus trap (useModalFocus) counts it.
const FOCUSABLE = 'button:not(:disabled), a[href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), summary, iframe, [tabindex]:not([tabindex="-1"])'
const HINT = "Change history available. Press Enter to read it."

const authorName = (name: string) => name.trim() || "No author named"
const plural = (count: number, one: string) => `${count} ${one}${count === 1 ? "" : "s"}`

// A whole paragraph's text ends in the pilcrow its mark travels with; a
// paragraph mark has no text of its own.
function changeText(revision: MarkupRevision) {
  if (revision.family === "paragraph_mark") return ""
  return revision.kind !== "format" && revision.text.endsWith("¶") ? revision.text.slice(0, -1) : revision.text
}

// The anchor's line boxes. A whole paragraph is one block box, so the lines of
// its text are used instead; a changed table, row or cell points at its first
// paragraph, where the reader places the change.
function lineBoxes(anchor: Element) {
  const target = anchor.matches(CONTAINERS) ? anchor.querySelector(PARAGRAPHS) ?? anchor : anchor
  const boxes = [...target.getClientRects()]
  if (!target.matches(PARAGRAPHS)) return boxes
  const range = document.createRange()
  range.selectNodeContents(target)
  const lines = [...range.getClientRects()].filter((rect) => rect.width > 0 && rect.height > 0)
  return lines.length ? lines : boxes
}

// Whether any of the line shows inside every scrolling box around the page
// (the viewer's document pane, or its body on a narrow screen).
function inView(line: DOMRect, clips: Element[]) {
  return clips.every((clip) => {
    const box = clip.getBoundingClientRect()
    return line.bottom > box.top && line.top < box.bottom && line.right > box.left && line.left < box.right
  })
}

// The line under the pointer, else the first (always the first for keyboard).
function lineAt(boxes: DOMRect[], point?: Point) {
  if (!point) return 0
  const under = boxes.findIndex((rect) => point.x >= rect.left && point.x <= rect.right && point.y >= rect.top && point.y <= rect.bottom)
  return under >= 0 ? under : Math.max(0, boxes.findIndex((rect) => point.y >= rect.top && point.y <= rect.bottom))
}

// Below, above, right, then left of the line: the first side with room for
// the whole card, else the roomier of below and above with the card's height
// capped so its list scrolls. Then it shifts to stay inside the window, and
// vertically inside the pane (`top`..`bottom`) so it never covers the viewer's
// toolbar. The app zooms the body (`--ui-scale`), so window measures are
// divided by it.
function placeCard(card: HTMLElement, list: HTMLElement | null, line: DOMRect, top: number, bottom: number) {
  const scale = Number(getComputedStyle(document.body).zoom) || 1
  const box = card.getBoundingClientRect()
  const width = box.width
  const height = box.height + (list ? (list.scrollHeight - list.clientHeight) * scale : 0)
  const room = {
    below: bottom - MARGIN - line.bottom - GAP,
    above: line.top - GAP - MARGIN - top,
    right: innerWidth - MARGIN - line.right - GAP,
    left: line.left - GAP - MARGIN,
  }
  const fits = (side: (typeof SIDES)[number]) => side === "below" || side === "above"
    ? height <= room[side] && width <= innerWidth - 2 * MARGIN
    : width <= room[side] && height <= bottom - top - 2 * MARGIN
  const fitting = SIDES.find(fits)
  const side = fitting ?? (room.below >= room.above ? "below" : "above")
  const cap = fitting ? undefined : Math.max(room[side], 0)
  const shown = cap === undefined ? height : Math.min(height, cap)
  const left = side === "right" ? line.right + GAP : side === "left" ? line.left - GAP - width : line.left
  const y = side === "below" ? line.bottom + GAP : side === "above" ? line.top - GAP - shown : line.top
  const within = (value: number, size: number, start: number, end: number) => Math.max(start + MARGIN, Math.min(value, end - MARGIN - size))
  card.style.maxHeight = cap === undefined ? "" : `${cap / scale}px`
  card.style.left = `${within(left, width, 0, innerWidth) / scale}px`
  card.style.top = `${within(y, shown, top, bottom) / scale}px`
  card.dataset.placement = side
  card.style.visibility = "visible"
}

export default function ChangeTimeline({ root, html, markup, revisions }: {
  // The rendered page (`.docx-render`); the card is its sibling, so the
  // page's inline-size container cannot trap the card's fixed position.
  root: RefObject<HTMLElement | null>
  html?: string
  // The working file's fresh markup reading; undefined offers nothing.
  markup?: DocumentMarkup
  // False when the page is not the raw working file (live doc.haus proposals,
  // or source changes accepted out of view): its paragraph count cannot be
  // trusted, so only comment threads are offered.
  revisions: boolean
}) {
  const [opened, setOpened] = useState<Opened>()
  const card = useRef<HTMLDivElement>(null)
  const list = useRef<HTMLDivElement>(null)
  // The page's handlers, for the card's own pointer, focus, Tab and Close events.
  const controls = useRef<{ hover: (inside: boolean) => void; blur: (next: Element | null) => void; tab: (event: KeyboardEvent) => void; close: () => void }>(undefined)
  // Enter on a closed clause's mark: focus the list once the card is placed.
  const focusList = useRef(false)
  const id = useId()
  const titleId = `${id}-title`
  // Read with each clause's mark, since browsers drop aria-expanded on ins and del.
  const hintId = `${id}-hint`
  const placeId = `${id}-place`

  // Map the page once per rendering and reading, give each clause one Tab
  // stop, and follow the pointer, focus and taps through delegated listeners.
  useEffect(() => {
    const page = root.current
    if (!html || !markup || !page) return
    const clauses = mapClauses(page, markup, revisions)
    if (!clauses.size) return
    const byMark = new Map<Element, Clause>([...clauses.values()].map((clause) => [clause.mark, clause]))
    const saved: [HTMLElement, string, string | null][] = []
    const described = new Map<Element, string>()
    for (const clause of clauses.values()) {
      const before = clause.mark.getAttribute("aria-describedby")
      described.set(clause.mark, before ? `${before} ${hintId}` : hintId)
      for (const [name, value] of [["tabindex", "0"], ["aria-haspopup", "dialog"], ["aria-expanded", "false"], ["aria-describedby", described.get(clause.mark)!]] as const) {
        saved.push([clause.mark, name, clause.mark.getAttribute(name)])
        clause.mark.setAttribute(name, value)
      }
      clause.mark.classList.add("change-timeline-mark")
    }
    const marksOf = (clause: Clause) => new Set([clause.mark, ...(clause.paragraph.matches(TARGETS) ? [clause.paragraph] : []), ...clause.paragraph.querySelectorAll(TARGETS)])

    let current: Opened | undefined
    // Whether the keyboard opened the card (or went back to its mark), so
    // focus on the mark holds it; a click that focuses a mark does not.
    let keyboard = false
    // What the mouse or pen is on: a clause's mark, the card, or neither;
    // `dragging` while a button is down, as when selecting text.
    const hover: { clause?: Clause; mark?: Element; point?: Point; card: boolean; dragging: boolean } = { card: false, dragging: false }
    let opening: ReturnType<typeof setTimeout> | undefined
    let closing: ReturnType<typeof setTimeout> | undefined
    let pending: Clause | undefined
    let pressedBy = ""
    // Focus going back to a mark after Escape must not reopen its card.
    let returning = false

    const stop = () => {
      clearTimeout(opening)
      clearTimeout(closing)
      opening = closing = pending = undefined
    }
    function show(next?: Opened, byKeyboard = false) {
      if (current) {
        current.clause.mark.setAttribute("aria-expanded", "false")
        current.clause.mark.setAttribute("aria-describedby", described.get(current.clause.mark)!)
        current.clause.mark.removeAttribute("aria-controls")
        for (const element of marksOf(current.clause)) element.classList.remove("is-timeline-active")
      }
      if (next) {
        next.clause.mark.setAttribute("aria-expanded", "true")
        next.clause.mark.setAttribute("aria-describedby", `${described.get(next.clause.mark)} ${placeId}`)
        next.clause.mark.setAttribute("aria-controls", id)
        for (const element of marksOf(next.clause)) element.classList.add("is-timeline-active")
      } else hover.card = false
      current = next
      keyboard = !!next && byKeyboard
      setOpened(next)
    }
    const focusHeld = () => {
      const active = document.activeElement
      return !!current && !!active && ((keyboard && active === current.clause.mark) || !!card.current?.contains(active))
    }
    const held = () => !!current && (current.touch || hover.card || hover.clause === current.clause || focusHeld())

    // Hover: a rest on a mark opens its clause (or switches to it); leaving
    // both the clause's marks and the card closes it a moment later, which is
    // the grace for crossing the gap between them. Focus wins over hover.
    function sync() {
      if (hover.card || (current && hover.clause === current.clause)) return stop()
      if (hover.clause && !hover.dragging && !focusHeld()) {
        clearTimeout(closing)
        closing = undefined
        if (pending === hover.clause) return
        clearTimeout(opening)
        pending = hover.clause
        opening = setTimeout(() => {
          pending = undefined
          if (hover.clause && hover.mark?.isConnected) show({ clause: hover.clause, anchor: hover.mark, point: hover.point, touch: false })
        }, REST)
        return
      }
      clearTimeout(opening)
      opening = pending = undefined
      if (!current || closing) return
      closing = setTimeout(() => {
        closing = undefined
        if (current && !held()) show(undefined)
      }, REST)
    }

    // The clause a mark belongs to, for a mark in a painted paragraph with
    // something to show.
    const hit = (node: EventTarget | null) => {
      if (!(node instanceof Element) || !page.contains(node)) return
      const found = markAt(node)
      const clause = found && page.contains(found.mark) ? clauses.get(found.paragraph) : undefined
      return clause ? { clause, mark: found!.mark } : undefined
    }
    const point = (event: PointerEvent) => ({ x: event.clientX, y: event.clientY })
    const over = (event: PointerEvent) => {
      if (event.pointerType === "touch") return
      const found = hit(event.target)
      hover.clause = found?.clause
      hover.mark = found?.mark
      hover.point = point(event)
      hover.dragging = event.buttons !== 0
      sync()
    }
    const out = (event: PointerEvent) => {
      if (event.pointerType === "touch") return
      const found = hit(event.relatedTarget)
      hover.clause = found?.clause
      hover.mark = found?.mark
      hover.dragging = event.buttons !== 0
      sync()
    }
    const move = (event: PointerEvent) => {
      if (event.pointerType !== "touch") hover.point = point(event)
    }
    // A press cancels a pending open (it may start a selection); any press
    // outside the card and every mark closes it.
    const press = (event: PointerEvent) => {
      pressedBy = event.pointerType
      clearTimeout(opening)
      opening = pending = undefined
      if (!current || card.current?.contains(event.target as Node) || hit(event.target)) return
      stop()
      show(undefined)
    }
    // Touch has no hover: a tap on a mark opens its card, unless it selected text.
    const tap = (event: MouseEvent) => {
      const found = pressedBy === "touch" ? hit(event.target) : undefined
      if (!found || document.getSelection()?.isCollapsed === false) return
      stop()
      show({ clause: found.clause, anchor: found.mark, point: { x: event.clientX, y: event.clientY }, touch: true })
    }
    // Keyboard focus opens at once; a mouse press that focuses the mark does not.
    const visible = (element: Element) => {
      try {
        return element.matches(":focus-visible")
      } catch {
        return true
      }
    }
    const focused = (event: FocusEvent) => {
      const clause = byMark.get(event.target as Element)
      if (!clause || returning || !visible(clause.mark)) return
      stop()
      show({ clause, anchor: clause.mark, touch: false }, true)
    }
    // Focus leaving the mark and the card closes it, unless hover or a tap holds it.
    const blurred = (next: Element | null) => {
      if (!current || current.touch || hover.card || hover.clause === current.clause) return
      if (next && (next === current.clause.mark || card.current?.contains(next))) return
      stop()
      show(undefined)
    }
    const left = (event: FocusEvent) => {
      if (current && event.target === current.clause.mark) blurred(event.relatedTarget as Element | null)
    }
    // Enter or Space on a mark moves into the card, so a long list can be scrolled.
    const keyed = (event: KeyboardEvent) => {
      const clause = byMark.get(event.target as Element)
      if (!clause || (event.key !== "Enter" && event.key !== " ")) return
      event.preventDefault()
      stop()
      if (current?.clause === clause && list.current) return list.current.focus()
      focusList.current = true
      show({ clause, anchor: clause.mark, touch: false }, true)
    }
    // Focus back on the open clause's mark, which keeps its card open and in place.
    const toMark = () => {
      if (!current) return
      keyboard = true
      returning = true
      current.clause.mark.focus()
      returning = false
    }
    // Shift+Tab from the card's first stop goes back to the clause's mark. Tab
    // from its last stop carries on from the mark, as if the card were there,
    // so the next clause's mark is next; when nothing between the mark and the
    // card takes focus, Tab carries on past the card.
    const tab = (event: KeyboardEvent) => {
      const box = card.current
      if (!current || !box) return
      const shown = (element: HTMLElement) => !element.closest("[inert]") && element.getClientRects().length > 0 && getComputedStyle(element).visibility !== "hidden"
      const stops = [...box.querySelectorAll<HTMLElement>(FOCUSABLE)].filter(shown)
      if (document.activeElement !== (event.shiftKey ? stops[0] : stops.at(-1))) return
      if (event.shiftKey) {
        event.preventDefault()
        return toMark()
      }
      const scope = box.parentElement?.closest('[aria-modal="true"]') ?? document.body
      const order = [...scope.querySelectorAll<HTMLElement>(FOCUSABLE)].filter((element) => element === current!.clause.mark || shown(element))
      const next = order[order.indexOf(current.clause.mark) + 1]
      if (next && !box.contains(next)) toMark()
    }
    // Closing the card (Escape, Close, or its line scrolled away) sends focus
    // inside it back to the clause's mark, so focus is never lost.
    const dismiss = () => {
      if (!current) return
      const mark = current.clause.mark
      const inside = !!card.current?.contains(document.activeElement)
      stop()
      show(undefined)
      if (!inside) return
      returning = true
      mark.focus({ preventScroll: true })
      returning = false
    }
    // Escape closes the card before the viewer sees it.
    const escape = (event: KeyboardEvent) => {
      if (!current || event.key !== "Escape" || event.defaultPrevented) return
      event.preventDefault()
      event.stopPropagation()
      dismiss()
    }
    controls.current = {
      hover: (inside) => {
        hover.card = inside
        sync()
      },
      blur: blurred,
      tab,
      close: dismiss,
    }

    page.addEventListener("pointerover", over)
    page.addEventListener("pointerout", out)
    page.addEventListener("pointermove", move, { passive: true })
    page.addEventListener("click", tap)
    page.addEventListener("focusin", focused)
    page.addEventListener("focusout", left)
    page.addEventListener("keydown", keyed)
    document.addEventListener("pointerdown", press, true)
    document.addEventListener("keydown", escape, true)
    return () => {
      page.removeEventListener("pointerover", over)
      page.removeEventListener("pointerout", out)
      page.removeEventListener("pointermove", move)
      page.removeEventListener("click", tap)
      page.removeEventListener("focusin", focused)
      page.removeEventListener("focusout", left)
      page.removeEventListener("keydown", keyed)
      document.removeEventListener("pointerdown", press, true)
      document.removeEventListener("keydown", escape, true)
      stop()
      if (current) show(undefined)
      controls.current = undefined
      for (const [element, name, value] of saved.reverse()) {
        if (value === null) element.removeAttribute(name)
        else element.setAttribute(name, value)
      }
      for (const clause of clauses.values()) clause.mark.classList.remove("change-timeline-mark")
    }
  }, [root, html, markup, revisions, id, hintId, placeId])

  // Placed before it paints, and again on resize and on any scroll but the
  // card's own. The line it points at stays the one first chosen.
  useLayoutEffect(() => {
    const element = card.current
    if (!opened || !element) return
    let line: number | undefined
    let frame = 0
    // The scrolling boxes the page sits in: a card whose line scrolls out of
    // them closes rather than stay clamped over the viewer's toolbar.
    const clips: Element[] = []
    for (let box = opened.anchor.parentElement; box && box !== document.body; box = box.parentElement) {
      const style = getComputedStyle(box)
      if (style.overflowX !== "visible" || style.overflowY !== "visible") clips.push(box)
    }
    const place = (moved: boolean) => {
      frame = 0
      if (!opened.anchor.isConnected) return controls.current?.close()
      const boxes = lineBoxes(opened.anchor)
      line ??= lineAt(boxes, opened.point)
      const at = boxes[line] ?? boxes[0] ?? opened.anchor.getBoundingClientRect()
      if (moved && !inView(at, clips)) return controls.current?.close()
      // The pane: the scrolling boxes outside the page (a table's own scroll
      // box would cramp the card), within the window.
      const pane = clips.filter((clip) => !root.current?.contains(clip)).map((clip) => clip.getBoundingClientRect())
      placeCard(element, list.current, at, Math.max(0, ...pane.map((box) => box.top)), Math.min(innerHeight, ...pane.map((box) => box.bottom)))
    }
    place(false)
    if (focusList.current) {
      focusList.current = false
      list.current?.focus({ preventScroll: true })
    }
    const later = (event: Event) => {
      if (event.target instanceof Node && element.contains(event.target)) return
      frame ||= requestAnimationFrame(() => place(true))
    }
    window.addEventListener("resize", later)
    window.addEventListener("scroll", later, true)
    return () => {
      cancelAnimationFrame(frame)
      window.removeEventListener("resize", later)
      window.removeEventListener("scroll", later, true)
    }
  }, [opened])

  const hint = markup && <span id={hintId} hidden>{HINT}</span>
  if (!opened) return hint || null
  const { clause } = opened
  const nodes = clauseTimeline(clause.revisions, clause.threads)
  const undated = nodes.findIndex((node) => node.date === null)
  const counts = [
    clause.revisions.length > 0 && plural(clause.revisions.length, "change"),
    clause.threads.length > 0 && plural(clause.threads.length, "comment thread"),
  ].filter(Boolean).join(" · ")
  return (
    <>
      {hint}
      <div ref={card} id={id} className="change-timeline" role="dialog" aria-labelledby={titleId}
        onPointerEnter={(event) => {
          if (event.pointerType !== "touch") controls.current?.hover(true)
        }}
        onPointerLeave={(event) => {
          if (event.pointerType !== "touch") controls.current?.hover(false)
        }}
        onBlur={(event) => controls.current?.blur(event.relatedTarget)}
        onKeyDown={(event) => {
          if (event.key === "Tab" && !event.altKey && !event.ctrlKey && !event.metaKey) controls.current?.tab(event.nativeEvent)
        }}>
        <div className="change-timeline-head">
          <div>
            <h2 className="change-timeline-title" id={titleId}>Change history</h2>
            <p className="change-timeline-place" id={placeId}>{clause.place} · {counts}</p>
          </div>
          {opened.touch && <button type="button" className="change-timeline-close" onClick={() => controls.current?.close()}>Close</button>}
        </div>
        <div ref={list} key={clause.ordinal} className="change-timeline-body" tabIndex={0} aria-label="Changes and comments, oldest first">
          <ol className="change-timeline-list">
            {nodes.map((node, index) => (
              <Fragment key={index}>
                {index === undated && <li className="change-timeline-undated">Date not recorded</li>}
                <li className="change-timeline-node">
                  <div className="change-timeline-meta">
                    <span className="change-timeline-author">{authorName(node.author)}</span>
                    {node.date && <time dateTime={node.date}>{formatDateTime(node.date)}</time>}
                  </div>
                  <ul>
                    {node.items.map((item) => (
                      <Action key={item.type === "revision" ? item.revision.id : `comment-${item.comment.id}`} item={item} ordinal={clause.ordinal} />
                    ))}
                  </ul>
                </li>
              </Fragment>
            ))}
          </ol>
        </div>
      </div>
    </>
  )
}

// One change or comment: what it is, its text, and what screening found.
function Action({ item, ordinal }: { item: TimelineItem; ordinal: number }) {
  if (item.type === "revision") {
    const { revision } = item
    const text = changeText(revision)
    const moved = movePlace(revision, ordinal)
    return (
      <li className="change-timeline-action" data-kind={revision.kind}>
        <span className="change-timeline-kind">{changeLabel(revision)}</span>
        {text && <div className={revision.kind === "delete" ? "change-old" : revision.kind === "insert" ? "change-new" : "change-timeline-text"}>{text}</div>}
        {moved && <p className="change-timeline-note">{moved}</p>}
        <Flags findings={revision.findings} />
      </li>
    )
  }
  const { comment, parent, quote } = item
  return (
    <li className="change-timeline-action" data-kind="comment">
      <span className="change-timeline-kind">{parent ? `Reply to ${authorName(parent.author)}` : comment.parent !== null ? "Reply" : "Comment"}</span>
      {comment.resolved === true && <span className="version-tag is-viewing change-timeline-tag">Resolved</span>}
      {quote && <blockquote className="change-timeline-quote">{quote}</blockquote>}
      <p className="change-timeline-comment">{comment.text}</p>
      <Flags findings={comment.findings} />
    </li>
  )
}

function Flags({ findings }: { findings: MarkupFinding[] }) {
  return <>{flagNotes(findings).map((note) => <p key={note} className="change-timeline-flag">{note}</p>)}</>
}
