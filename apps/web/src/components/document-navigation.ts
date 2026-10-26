// Read the rendered document without wrapping or replacing its text. Word runs,
// citation marks and comment links keep their nodes and event handlers.
const ignored = "style, script, noscript, [hidden], .comment-margin-note, .comment-marker, .doc-comment-pin"
const blocks = "p, li, td, th, h1, h2, h3, h4, h5, h6, pre"

function visible(element: Element) {
  const style = element.ownerDocument.defaultView?.getComputedStyle(element)
  return Boolean(element.getClientRects().length) && style?.visibility !== "hidden" && style?.visibility !== "collapse"
}

export function documentMatches(root: HTMLElement, query: string): Range[] {
  const needle = query.trim()
  if (!needle) return []
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT)
  const nodes: { node: Text; start: number; end: number }[] = []
  let text = ""
  let block: Element | null | undefined
  const visibility = new Map<Element, boolean>()
  while (walker.nextNode()) {
    if (walker.currentNode.nodeType === Node.ELEMENT_NODE) {
      const element = walker.currentNode as Element
      // Word hard breaks render as <br> inside the same paragraph. They
      // separate words even though there is no text node for the break.
      if (element.tagName === "BR" && !element.closest(ignored) && visible(element)) text += "\n"
      continue
    }
    const node = walker.currentNode as Text
    // Docxodus follows a hard break with a standalone direction mark. It has
    // no visible characters and must not interrupt whitespace matching.
    if (/^[\u200e\u200f]+$/.test(node.data)) continue
    const parent = node.parentElement
    if (!parent || parent.closest(ignored)) continue
    if (!visibility.has(parent)) visibility.set(parent, visible(parent))
    if (!visibility.get(parent)) continue
    const nextBlock = parent.closest(blocks)
    if (nodes.length && block !== nextBlock) text += "\n"
    nodes.push({ node, start: text.length, end: text.length + node.length })
    text += node.data
    block = nextBlock
  }
  const literal = needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+")
  const matches: Range[] = []
  let cursor = 0
  for (const match of text.matchAll(new RegExp(literal, "giu"))) {
    const start = match.index, end = start + match[0].length
    while (nodes[cursor] && nodes[cursor].end <= start) cursor++
    const first = nodes[cursor]
    let last = cursor
    while (nodes[last] && nodes[last].end < end) last++
    if (!first || !nodes[last]) continue
    const range = root.ownerDocument.createRange()
    range.setStart(first.node, start - first.start)
    range.setEnd(nodes[last].node, end - nodes[last].start)
    matches.push(range)
  }
  return matches
}

export type DocumentHeading = { element: HTMLElement; level: number; text: string }

export function documentHeadings(root: HTMLElement): DocumentHeading[] {
  return [...root.querySelectorAll<HTMLElement>("h1, h2, h3, h4, h5, h6, [role=heading][aria-level]")].flatMap((element) => {
    if (element.closest(ignored) || !visible(element)) return []
    // Docxodus renders Word outline levels as h1–h6 and native numbering as
    // text within them. Read only the heading, leaving margin comments out.
    const clone = element.cloneNode(true) as HTMLElement
    for (const ignoredNode of clone.querySelectorAll(ignored)) ignoredNode.remove()
    const text = clone.textContent?.replace(/\s+/g, " ").trim()
    const level = Number(element.getAttribute("aria-level") ?? element.tagName.slice(1))
    return text ? [{ element, level: Math.min(6, Math.max(1, level || 1)), text }] : []
  })
}

export function focusDocumentHeading(heading: HTMLElement) {
  const previous = heading.getAttribute("tabindex")
  heading.tabIndex = -1
  heading.scrollIntoView({ block: "start", behavior: "instant" })
  heading.focus({ preventScroll: true })
  heading.addEventListener("blur", () => {
    if (previous === null) heading.removeAttribute("tabindex")
    else heading.setAttribute("tabindex", previous)
  }, { once: true })
}

export function scrollDocumentMatch(range: Range, root: HTMLElement) {
  // Scroll the text's rectangle, rather than the top of a potentially long
  // paragraph. The scroll container changes on narrow screens.
  const rect = range.getBoundingClientRect()
  const page = root.closest<HTMLElement>(".viewer-doc")
  const body = root.closest<HTMLElement>(".viewer-body")
  const scroll = page && getComputedStyle(page).overflowY === "auto" ? page : body
  if (!scroll || !rect.height) return
  const viewport = scroll.getBoundingClientRect()
  scroll.scrollBy({ top: rect.top - viewport.top - viewport.height / 2 + rect.height / 2, behavior: "instant" })
}
