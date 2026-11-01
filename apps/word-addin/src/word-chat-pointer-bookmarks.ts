const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  XMLNS = "http://www.w3.org/2000/xmlns/"
const POINTER = /^_dh_p[0-9a-f]{12}$/
const bookmark = (e: Element) => e.namespaceURI === W && (e.localName === "bookmarkStart" || e.localName === "bookmarkEnd")

/** doc.haus marks each paragraph it renders with one hidden pointer bookmark that
 * wraps the paragraph's runs. Return that pair when the paragraph's own children
 * hold exactly it (start first, end last, same numeric id, nothing else), an empty
 * list when the paragraph holds no bookmark, and undefined for any other bookmark shape. */
export function pointerBookmarks(children: Element[]): Element[] | undefined {
  const marks = children.filter(bookmark)
  if (!marks.length) return []
  const [start, end] = marks,
    first = children.find((e) => !(e.namespaceURI === W && e.localName === "pPr"))
  const ok = (e: Element, name: boolean) =>
    Array.from(e.attributes).every(
      (a) => a.namespaceURI === XMLNS || (a.namespaceURI === W && (a.localName === "id" || (name && a.localName === "name"))),
    ) && /^\d+$/.test(e.getAttributeNS(W, "id") ?? "")
  return marks.length === 2 &&
    start!.localName === "bookmarkStart" &&
    end!.localName === "bookmarkEnd" &&
    POINTER.test(start!.getAttributeNS(W, "name") ?? "") &&
    ok(start!, true) &&
    ok(end!, false) &&
    start!.getAttributeNS(W, "id") === end!.getAttributeNS(W, "id") &&
    first === start &&
    children.at(-1) === end
    ? marks
    : undefined
}
