import { useEffect, useRef, useState, type RefObject } from "react"
import type { ClauseHistoryRequest } from "../../../../dochaus/lib/clause-history-contract"
import type { DocumentVersions } from "../api/ingest/versions"
import { ClauseHistory } from "./ClauseHistory"
import { paintedParagraphs, PARAGRAPHS } from "./change-timeline"

/** Capture the whole paragraph, not just the highlighted excerpt. The server
 * verifies this wording against the exact saved bytes; rendered ordinals and
 * content-derived HTML ids never establish cross-version identity. */
export function selectedHistoryParagraph(root: HTMLElement, selection: Selection | null): string | undefined {
  if (!selection?.rangeCount) return
  const range = selection.getRangeAt(0)
  const element = (node: Node) => node instanceof root.ownerDocument.defaultView!.Element ? node as Element : node.parentElement
  const start = element(range.startContainer)?.closest(PARAGRAPHS)
  const end = element(range.endContainer)?.closest(PARAGRAPHS)
  if (!start || start !== end || !paintedParagraphs(root).includes(start)) return
  const copy = start.cloneNode(true) as Element
  for (const decoration of copy.querySelectorAll("del, .rev-move-from, [data-list-marker], a.comment-marker, .rev-para-mark-del, .rev-para-mark-ins")) decoration.remove()
  // These are typed Word characters represented by DOM layout elements.
  // Do not collapse whitespace or normalize punctuation to force a match.
  for (const tab of copy.querySelectorAll("[data-docx-tab]")) tab.replaceWith(root.ownerDocument.createTextNode("\t"))
  for (const line of copy.querySelectorAll("br")) {
    // Docxodus emits one plain LRM immediately after its <br>, outside the
    // spans containing actual run text. Preserve directional marks in runs.
    const marker = line.nextSibling
    if (marker?.nodeType === 3 && marker.textContent?.startsWith("\u200e")) marker.textContent = marker.textContent.slice(1)
    line.replaceWith(root.ownerDocument.createTextNode("\n"))
  }
  const text = copy.textContent ?? ""
  return text.trim() ? text : undefined
}

export function DocumentClauseHistory({ matterId, name, root, content, versions, shownVersion }: {
  matterId: string; name: string; root: RefObject<HTMLDivElement | null>; content?: string; versions?: DocumentVersions; shownVersion?: string
}) {
  const [paragraph, setParagraph] = useState<string>()
  const [request, setRequest] = useState<ClauseHistoryRequest>()
  const trigger = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    setParagraph(undefined); setRequest(undefined)
    const node = root.current
    if (!node || !content) return
    const read = () => {
      const selection = node.ownerDocument.getSelection()
      // Clicking History may collapse the browser selection into its button.
      // Keep the last captured clause until another document selection or load.
      if (selection?.anchorNode && node.contains(selection.anchorNode)) setParagraph(selectedHistoryParagraph(node, selection))
    }
    node.ownerDocument.addEventListener("selectionchange", read)
    node.addEventListener("pointerup", read)
    node.addEventListener("keyup", read)
    return () => { node.ownerDocument.removeEventListener("selectionchange", read); node.removeEventListener("pointerup", read); node.removeEventListener("keyup", read) }
  }, [matterId, name, root, content, shownVersion, versions?.working_version])
  const anchorVersion = shownVersion ?? versions?.working_version
  const current = versions?.versions.find(version => version.current)
  return <div className="document-clause-history">
    <div className="document-clause-history-control">
      <button ref={trigger} type="button" disabled={!paragraph || !anchorVersion || !current} aria-expanded={Boolean(request)}
        title={!anchorVersion ? "Refresh the document's saved versions before opening history" : "Select text in one paragraph to inspect its earlier wording"}
        onClick={() => { if (paragraph && anchorVersion && current) setRequest({ anchorVersion, paragraph: { scope: "body", text: paragraph }, expectedCurrentHash: current.sha256 }) }}>Clause history</button>
      <span>Select text in one paragraph to see its saved history.</span>
    </div>
    {request && <ClauseHistory matterId={matterId} name={name} request={request} onClose={() => { setRequest(undefined); trigger.current?.focus() }} />}
  </div>
}
