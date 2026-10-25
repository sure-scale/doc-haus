// The converter writes a Word hyperlink's target into the rendered page as the
// file gives it, so a target such as javascript: would run in the app when the
// lawyer clicks the link. After a viewer renders a document, and before anyone
// can click it, a link keeps its target only when it is a place in the document
// (#bookmark) or the browser reads it as a web or mail address. Any other link
// keeps its text and loses its target. The browser's own URL parser decides, so
// the target is judged exactly as a click would read it.
const LINK_PROTOCOLS = new Set(["http:", "https:", "mailto:"])

function linkProtocol(href: string) {
  try {
    return new URL(href, document.baseURI).protocol
  } catch {
    return undefined
  }
}

export function removeUnsafeLinkTargets(root: ParentNode) {
  for (const link of root.querySelectorAll("[href]")) {
    const href = link.getAttribute("href")!
    if (href.startsWith("#")) continue
    if (!LINK_PROTOCOLS.has(linkProtocol(href) ?? "")) link.removeAttribute("href")
  }
}
