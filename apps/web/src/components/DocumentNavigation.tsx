import { useEffect, useId, useRef, useState, type KeyboardEvent, type RefObject } from "react"
import { documentHeadings, documentMatches, focusDocumentHeading, scrollDocumentMatch, type DocumentHeading } from "./document-navigation"
import "./DocumentNavigation.css"

// Shared by the Word preview and extracted-text viewer. Native PDFs keep their
// browser reader's navigation; their text is outside this document's DOM.
export function useDocumentNavigation(root: RefObject<HTMLElement | null>, content: string | undefined, withOutline = true) {
  const id = `document-navigation-${useId().replace(/[^a-z0-9-]/gi, "")}`
  const [findOpen, setFindOpen] = useState(false)
  const [outlineOpen, setOutlineOpen] = useState(false)
  const [query, setQuery] = useState("")
  const [matches, setMatches] = useState<Range[]>([])
  const [current, setCurrent] = useState(0)
  const [headings, setHeadings] = useState<DocumentHeading[]>([])
  const [selected, setSelected] = useState<HTMLElement>()
  const [scrollIntent, setScrollIntent] = useState(0)
  const scrollRequested = useRef(false)
  const findToggle = useRef<HTMLButtonElement>(null)
  const outlineToggle = useRef<HTMLButtonElement>(null)
  const input = useRef<HTMLInputElement>(null)
  const findName = `${id}-find`, currentName = `${id}-current`

  useEffect(() => {
    if (findOpen) {
      scrollRequested.current = true
      input.current?.focus()
    }
  }, [findOpen])

  useEffect(() => {
    setCurrent(0)
    scrollRequested.current = true
  }, [query, content])

  useEffect(() => {
    const element = root.current
    if (!element || content === undefined) {
      setMatches([])
      setHeadings([])
      return
    }
    let frame = 0
    function read() {
      if (!element) return
      const found = findOpen ? documentMatches(element, query) : []
      setMatches(found)
      setCurrent((index) => Math.max(0, Math.min(index, found.length - 1)))
      if (withOutline) setHeadings(documentHeadings(element))
    }
    read()
    // Citations and the existing comment/marker decorators can split text nodes
    // after conversion. Re-read their final DOM without interfering with them.
    const observer = new MutationObserver(() => {
      cancelAnimationFrame(frame)
      frame = requestAnimationFrame(read)
    })
    observer.observe(element, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["class", "style", "hidden", "data-comments"] })
    return () => { observer.disconnect(); cancelAnimationFrame(frame) }
  }, [root, content, findOpen, query, withOutline])

  useEffect(() => {
    if (!findOpen || !matches.length) return
    const range = matches[current]
    if (!range || !root.current) return
    const registry = typeof CSS !== "undefined" && "highlights" in CSS ? CSS.highlights : undefined
    if (registry && typeof Highlight !== "undefined") {
      const all = new Highlight(...matches)
      const active = new Highlight(range)
      active.priority = 1
      registry.set(findName, all)
      registry.set(currentName, active)
      if (scrollRequested.current) scrollDocumentMatch(range, root.current)
      scrollRequested.current = false
      return () => { registry.delete(findName); registry.delete(currentName) }
    }
    // Older browsers can still navigate using their native selection. Keep
    // focus in the Find field and clear only the selection this effect owns.
    const selection = root.current.ownerDocument.getSelection()
    selection?.removeAllRanges()
    selection?.addRange(range)
    if (scrollRequested.current) scrollDocumentMatch(range, root.current)
    scrollRequested.current = false
    return () => {
      if (selection?.rangeCount && selection.getRangeAt(0) === range) selection.removeAllRanges()
    }
  }, [matches, current, findOpen, root, findName, currentName, scrollIntent])

  const closeFind = () => { setFindOpen(false); findToggle.current?.focus() }
  const closeOutline = () => { setOutlineOpen(false); outlineToggle.current?.focus() }
  const move = (step: number) => {
    scrollRequested.current = true
    setScrollIntent((value) => value + 1)
    setCurrent((index) => matches.length ? (index + step + matches.length) % matches.length : 0)
  }
  function onKeyDown(event: KeyboardEvent) {
    if (event.nativeEvent.isComposing && event.key === "Enter") return
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "f" && content !== undefined) {
      event.preventDefault(); event.stopPropagation()
      setFindOpen(true)
      input.current?.focus()
      input.current?.select()
      return
    }
    if (findOpen && (event.key === "F3" || ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "g") || (event.key === "Enter" && event.target === input.current))) {
      event.preventDefault(); event.stopPropagation(); move(event.shiftKey ? -1 : 1)
    } else if (event.key === "Escape" && (findOpen || outlineOpen)) {
      event.preventDefault(); event.stopPropagation()
      if (findOpen) closeFind()
      else closeOutline()
    }
  }

  return {
    onKeyDown,
    bar: <div className="document-navigation-bar">
      <div className="document-navigation-toggles">
        <button type="button" ref={findToggle} aria-expanded={findOpen} aria-controls={`${id}-search`} disabled={content === undefined} onClick={() => findOpen ? closeFind() : setFindOpen(true)}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><circle cx="10.5" cy="10.5" r="6.5" /><path d="m16 16 5 5" /></svg>
          Find
        </button>
        {withOutline && <button type="button" ref={outlineToggle} aria-expanded={outlineOpen} aria-controls={`${id}-outline`} disabled={content === undefined} onClick={() => outlineOpen ? closeOutline() : setOutlineOpen(true)}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h13M3 6h1M3 12h1M3 18h1" /></svg>
          Outline
        </button>}
      </div>
      {findOpen && <div id={`${id}-search`} className="document-find" role="search" aria-label="Find in document">
        <label className="sr-only" htmlFor={`${id}-query`}>Find in document</label>
        <input id={`${id}-query`} ref={input} value={query} placeholder="Find in document" type="search" autoComplete="off" onChange={(event) => setQuery(event.target.value)} />
        <span className="document-find-count" role="status" aria-live="polite" aria-atomic="true">{!query.trim() ? "Enter text" : !matches.length ? "No matches" : `${current + 1} of ${matches.length}`}</span>
        <button type="button" aria-label="Previous match" title="Previous match (Shift+Enter)" disabled={!matches.length} onClick={() => move(-1)}>↑</button>
        <button type="button" aria-label="Next match" title="Next match (Enter)" disabled={!matches.length} onClick={() => move(1)}>↓</button>
        <button type="button" aria-label="Close Find" onClick={closeFind}>×</button>
      </div>}
      <style>{`::highlight(${findName}) { background-color: var(--md-sys-color-tertiary-container); color: var(--md-sys-color-on-tertiary-container); } ::highlight(${currentName}) { background-color: var(--md-sys-color-primary); color: var(--md-sys-color-on-primary); text-decoration: underline; }`}</style>
    </div>,
    outline: outlineOpen && <nav id={`${id}-outline`} className="document-outline" aria-label="Document outline">
      <div className="document-outline-head"><strong>Outline</strong><button type="button" aria-label="Close outline" onClick={closeOutline}>×</button></div>
      {headings.length ? <ol>{headings.map(({ element, level, text }, index) => <li key={index} style={{ paddingInlineStart: `${(level - 1) * 12}px` }}>
        <button type="button" aria-current={selected === element ? "location" : undefined} onClick={() => {
          setSelected(element)
          if (window.matchMedia("(max-width: 720px)").matches) setOutlineOpen(false)
          focusDocumentHeading(element)
        }}>{text}</button>
      </li>)}</ol> : <p className="muted">No headings in this document.</p>}
    </nav>,
  }
}
