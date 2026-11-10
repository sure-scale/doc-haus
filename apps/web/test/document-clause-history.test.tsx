import { expect, test } from "bun:test"
import { Window } from "happy-dom"
import { act, createElement, createRef } from "react"
import { Document, Packer, Paragraph } from "docx"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { DocumentClauseHistory, selectedHistoryParagraph } from "../src/components/DocumentClauseHistory"
import { docxodus } from "../../../dochaus/lib/docxodus"
import type { ComparisonParagraph } from "../../../dochaus/lib/version-compare-contract"
import type { DocumentVersions } from "../src/api/ingest/versions"

function selection(page: Window, start: Node, end = start, from = 0, to = end.nodeType === 3 ? end.textContent?.length ?? 0 : end.childNodes.length) {
  const range = page.document.createRange(); range.setStart(start as never, from); range.setEnd(end as never, to)
  const current = page.document.getSelection()!; current.removeAllRanges(); current.addRange(range)
  return current as unknown as Selection
}

test("a partial or collapsed selection captures only one whole body paragraph with literal Unicode and whitespace", async () => {
  const page = new Window(), root = page.document.createElement("div")
  root.innerHTML = '<header><p>Header wording.</p></header><div data-section-index="0"><h2>1. Terms</h2><p>Literal 💬 A  B. A B is different.</p><table><tbody><tr><td><p>Cell wording.</p></td></tr></tbody></table></div><p>Margin note.</p>'
  page.document.body.append(root)
  try {
    const paragraph = root.querySelector("div[data-section-index] > p")!, text = paragraph.firstChild!
    for (const [from, to] of [[3,3], [3,10]]) expect(selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, text as unknown as Node, undefined, from, to))).toBe("Literal 💬 A  B. A B is different.")
    expect(selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, root.querySelector("header p")!.firstChild! as unknown as Node))).toBeUndefined()
    expect(selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, root.lastChild!.firstChild! as unknown as Node))).toBeUndefined()
    expect(selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, text as unknown as Node, root.querySelector("td p")!.firstChild! as unknown as Node))).toBeUndefined()
    expect(selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, root.querySelector("td p")!.firstChild! as unknown as Node))).toBe("Cell wording.")
    expect(root.querySelector("div[data-section-index] > p")!.textContent).toBe("Literal 💬 A  B. A B is different.")
  } finally { await page.happyDOM.abort() }
})

test("Current wording excludes tracked deletions and rendered markers and retains tabs and line breaks", async () => {
  const page = new Window(), root = page.document.createElement("div")
  root.innerHTML = '<div data-section-index="0"><p><span data-list-marker>3.</span>Alpha<span data-docx-tab>layout tab</span>Beta<del class="rev-del">Old</del><br>\u200e<ins>New</ins><span>\u200e literal</span><span class="rev-move-from">Moved away</span><a class="comment-marker">[1]</a><span class="rev-para-mark-ins">¶</span><span class="rev-para-mark-del">¶</span></p></div>'
  page.document.body.append(root)
  const before = root.innerHTML
  try {
    expect(selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, root.querySelector("ins")!.firstChild! as unknown as Node))).toBe("Alpha\tBeta\nNew\u200e literal")
    expect(root.innerHTML).toBe(before)
  } finally { await page.happyDOM.abort() }
})

test("actual Docxodus rendered Current wording matches the saved paragraph inventory", async () => {
  const parts = unzipSync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph("Placeholder"), new Paragraph("Other clause.")] }] })))
  const paragraph = '<w:p><w:r><w:t xml:space="preserve">Alpha 💬</w:t><w:tab/><w:t>Beta</w:t><w:br/></w:r><w:del w:id="1" w:author="Reviewer"><w:r><w:delText>Old</w:delText></w:r></w:del><w:ins w:id="2" w:author="Reviewer"><w:r><w:t>Current</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>'
  parts["word/document.xml"] = strToU8(strFromU8(parts["word/document.xml"]!).replace(/<w:p[\s>][\s\S]*?<\/w:p>/, paragraph))
  const bytes = zipSync(parts), engine = await docxodus()
  // Exercise the actual backend projection at runtime. Its source graph is
  // checked by dochaus; browser checks consume only its shared result type.
  const { readComparisonParagraphs } = await import(new URL("../../../dochaus/lib/version-compare.ts", import.meta.url).href) as { readComparisonParagraphs(bytes: Uint8Array): Promise<ComparisonParagraph[]> }
  const [html, saved] = await Promise.all([engine.convertDocxToHtml(bytes, { renderTrackedChanges: true, showDeletedContent: true, renderMoveOperations: true }), readComparisonParagraphs(bytes)])
  const page = new Window(), root = page.document.createElement("div"); root.innerHTML = html; page.document.body.append(root)
  try {
    const paragraph = root.querySelector('div[data-section-index] p')!
    expect(Boolean(paragraph)).toBe(true)
    const current = selectedHistoryParagraph(root as unknown as HTMLElement, selection(page, paragraph.querySelector("ins")?.firstChild as unknown as Node ?? paragraph.firstChild! as unknown as Node))
    expect(current).toBe("Alpha 💬\tBeta\nCurrent end.")
    expect(saved.find(item => item.location.scope === "body")?.text).toBe(current)
  } finally { await page.happyDOM.abort() }
})

test("the web adapter pins full text to the shown saved version, preserves the viewer, and restores History focus", async () => {
  const browser = new Window({ url: "http://localhost:5275" })
  const names = ["window", "document", "HTMLElement", "Element", "Node", "Event", "IS_REACT_ACT_ENVIRONMENT"] as const
  const savedGlobals = names.map(name => [name, Object.getOwnPropertyDescriptor(globalThis, name)] as const)
  Object.defineProperties(globalThis, Object.fromEntries(names.map(name => [name, { configurable: true, writable: true, value: name === "IS_REACT_ACT_ENVIRONMENT" ? true : browser[name as keyof Window] }])))
  const fetchBefore = globalThis.fetch, requests: { url: string; init?: RequestInit }[] = []
  globalThis.fetch = (async (url, init) => { requests.push({ url: String(url), init }); return Response.json({ document: { id: "doc", name: "Agreement.docx", currentVersion: "v3" }, anchorVersion: "v1", currentVersion: "v3", anchor: { status: "verified", paragraphs: [] }, rounds: [], wordingView: "current-text", limitations: [] }) }) as typeof fetch
  const { createRoot } = await import("react-dom/client")
  const container = document.createElement("div"), viewer = document.createElement("div"); document.body.append(viewer, container)
  viewer.innerHTML = '<div data-section-index="0"><p>Full first clause.</p><p>Another clause.</p></div>'
  const root = createRoot(container), ref = createRef<HTMLDivElement>(); ref.current = viewer
  const versions: DocumentVersions = { name: "Agreement.docx", document: { id: "doc", name: "Agreement.docx", current_version: "v3" }, working_version: "v3", events: [], versions: [{ id: "v3", number: 3, sha256: "a".repeat(64), current: true } as DocumentVersions["versions"][number]] }
  try {
    await act(async () => { root.render(createElement(DocumentClauseHistory, { matterId: "matter", name: "Agreement.docx", root: ref, content: "rendered content", versions, shownVersion: "v1" })) })
    const trigger = container.querySelector<HTMLButtonElement>("button")!
    expect(trigger.disabled).toBe(true)
    await act(async () => { selection(browser, viewer.querySelector("p")!.firstChild!, undefined, 3, 7); document.dispatchEvent(new Event("selectionchange")) })
    const before = viewer.innerHTML, selected = document.getSelection()?.toString()
    expect(trigger.disabled).toBe(false)
    await act(async () => { trigger.click() })
    expect(JSON.parse(requests[0]!.init!.body as string)).toEqual({ anchorVersion: "v1", paragraph: { scope: "body", text: "Full first clause." }, expectedCurrentHash: "a".repeat(64) })
    expect(viewer.innerHTML).toBe(before); expect(document.getSelection()?.toString()).toBe(selected)
    await act(async () => { container.querySelector(".clause-history")!.dispatchEvent(new browser.KeyboardEvent("keydown", { key: "Escape", bubbles: true }) as unknown as KeyboardEvent) })
    expect(Boolean(container.querySelector(".clause-history"))).toBe(false)
    expect(document.activeElement === trigger).toBe(true)
    await act(async () => { root.render(createElement(DocumentClauseHistory, { matterId: "other", name: "Other.docx", root: ref, content: "new document", versions })) })
    expect(container.querySelector<HTMLButtonElement>("button")!.disabled).toBe(true)
    expect(requests).toHaveLength(1)
  } finally {
    await act(async () => root.unmount()); globalThis.fetch = fetchBefore; await browser.happyDOM.close()
    for (const [name, descriptor] of savedGlobals) { if (descriptor) Object.defineProperty(globalThis, name, descriptor); else delete (globalThis as Record<string, unknown>)[name] }
  }
})
