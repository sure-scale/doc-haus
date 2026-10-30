import { expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { Window } from "happy-dom"
import { DocumentProposalReview, DownloadActions } from "../src/components/DocumentViewer"
import type { Redline } from "../src/api/ingest/redlines"

type Insertion = NonNullable<Redline["paragraph_insert"]>
const reference = "The  existing paragraph says <a>literal</a> & keeps every word."
const text = `New <b>literal</b> | &  repeated spaces ${"X".repeat(600)}`
const location = { body_paragraph_number: 8, headings: ["Service terms", "Payment"], previous_paragraph: "The  preceding paragraph.", next_paragraph: "The following paragraph." } satisfies Insertion["location"]
const insertion: Insertion = { kind: "paragraph_insert", reference_text: reference, position: "after", text, location }
const proposal = (changes: Partial<Redline> = {}): Redline => ({
  id: 91, doc_name: "Synthetic insertion.docx", scope: "paragraph_insert", paragraph_insert: insertion,
  old_text: "Legacy old summary must not display", new_text: "Legacy joined summary must not display", find_text: "Compatibility summary",
  author: "Synthetic Reviewing Lawyer", review_comment: "Insert the requested standalone paragraph.", created_at: 1, ...changes,
})
function render(value: Redline | Redline[]) {
  const page = new Window()
  page.document.body.innerHTML = renderToStaticMarkup(createElement(DocumentProposalReview, {
    redlines: Array.isArray(value) ? value : [value], wordLocked: false, onAccept() {}, onReject() {}, onAcceptAll() {}, onRejectAll() {},
  }))
  const button = (name: string) => [...page.document.querySelectorAll("button")].find(button => button.textContent === name)!
  return { page, root: page.document.body, button, close: () => page.happyDOM.abort() }
}
function acceptance(view: ReturnType<typeof render>, enabled: boolean) {
  expect(view.button("Accept").disabled).toBe(!enabled); expect(view.button("Accept all").disabled).toBe(!enabled)
  expect(view.button("Reject").disabled).toBe(false); expect(view.button("Reject all").disabled).toBe(false)
}

test("a paragraph insertion shows full exact escaped reference/new text in before or after order without replacing the reference", () => {
  for (const position of ["before", "after"] as const) {
    const original = proposal({ paragraph_insert: { ...insertion, position } }), snapshot = JSON.stringify(original), view = render(original)
    try {
      expect(view.root.querySelector(".paragraph-insertion-title")!.textContent).toBe(`Insert a paragraph ${position} the reference`)
      expect([...view.root.querySelectorAll(".paragraph-insertion-block")].map(block => block.getAttribute("aria-label"))).toEqual(position === "before"
        ? ["New paragraph", "Reference paragraph, unchanged"] : ["Reference paragraph, unchanged", "New paragraph"])
      expect(view.root.querySelector(".paragraph-insertion-reference .paragraph-insertion-text")!.textContent).toBe(reference)
      expect(view.root.querySelector(".paragraph-insertion-new .paragraph-insertion-text")!.textContent).toBe(text)
      expect(view.root.querySelector("a,b,del,ins,.proposal-redline,.table-row-proposal,.paragraph-spacing-proposal")).toBeNull()
      expect(view.root.textContent).toContain(original.review_comment)
      expect(view.root.textContent).not.toContain(original.old_text); expect(view.root.textContent).not.toContain(original.new_text)
      acceptance(view, true); expect(JSON.stringify(original)).toBe(snapshot)
    } finally { view.close() }
  }
})

test("the same frozen insertion becomes non-acceptable when stale, keeps original context and leaves rejection/download available", () => {
  const original = proposal(), fresh = render(original)
  try { acceptance(fresh, true); expect(fresh.root.textContent).toContain("Reference paragraph · unchanged") }
  finally { fresh.close() }
  const stale = proposal({ stale: true, stale_reason: "This reference belongs to an earlier Word file. Prepare a fresh change list." }), view = render(stale)
  try {
    expect(view.root.querySelector(".paragraph-insertion-title")!.textContent).toBe("Insert a paragraph after the original reference")
    expect(view.root.querySelector(".paragraph-insertion-unavailable")!.textContent).toBe("Reference location needs refreshing.")
    expect(view.root.querySelector(".paragraph-insertion-reference")!.getAttribute("aria-label")).toBe("Original reference paragraph")
    expect(view.root.querySelector(".paragraph-insertion-reference .paragraph-insertion-text")!.textContent).toBe(reference)
    expect(view.root.querySelector(".paragraph-insertion-new .paragraph-insertion-text")!.textContent).toBe(text)
    expect(view.root.querySelector(".paragraph-insertion-location summary")!.textContent).toBe("Original body paragraph 8")
    expect([...view.root.querySelectorAll(".paragraph-insertion-location dt")].map(node => node.textContent)).toEqual(["Original heading path", "Original previous paragraph", "Original next paragraph"])
    expect([...view.root.querySelectorAll(".paragraph-insertion-source-text")].map(node => node.textContent)).toEqual([...location.headings, location.previous_paragraph, location.next_paragraph])
    expect(view.root.textContent).toContain(stale.stale_reason!); acceptance(view, false)
    const page = new Window(); page.document.body.innerHTML = renderToStaticMarkup(createElement(DownloadActions, { redlines: [stale], preview: false, onDownload() {} }))
    expect(page.document.querySelector("button")!.textContent).toBe("Download"); expect(page.document.querySelector("button")!.disabled).toBe(false)
    page.happyDOM.abort()
  } finally { view.close() }
})

test("whitespace reference content remains exact with a separate annotation, while a matching stored operation stays immutable", () => {
  const viewValue: Insertion = { ...insertion, reference_text: "  ", text: "An ordinary new paragraph." }
  const original = proposal({ paragraph_insert: viewValue, target_json: JSON.stringify({ version: 5, ...viewValue, source_sha256: "a".repeat(64), prepared_role: "source", target_handle: "opaque-source-reference" }) })
  const snapshot = JSON.stringify(original), view = render(original)
  try {
    expect(view.root.querySelector(".paragraph-insertion-reference .paragraph-insertion-text")!.textContent).toBe("  ")
    expect(view.root.querySelector(".paragraph-insertion-reference .paragraph-insertion-note")!.textContent).toBe("Whitespace only")
    expect(view.root.textContent).not.toContain("opaque-source-reference")
    acceptance(view, true); expect(JSON.stringify(original)).toBe(snapshot)
  } finally { view.close() }
})

test("missing, unsupported, disguised v5 and conflicting frozen insertion data never become ordinary diffs or enable either Accept", () => {
  for (const original of [
    proposal({ paragraph_insert: undefined }),
    proposal({ scope: "phrase", paragraph_insert: undefined, target_json: JSON.stringify({ version: 5, kind: "unknown" }) }),
    proposal({ scope: "clause", paragraph_insert: undefined, target_json: JSON.stringify({ kind: "paragraph_insert" }) }),
    proposal({ paragraph_insert: { ...insertion, position: "middle" } as unknown as Insertion }),
    proposal({ paragraph_insert: { ...insertion, text: "Two\nparagraphs" } }),
    proposal({ paragraph_insert: { ...insertion, style: "Heading1" } as unknown as Insertion }),
    proposal({ paragraph_insert: { ...insertion, location: undefined } as unknown as Insertion }),
    proposal({ paragraph_insert: { ...insertion, location: { ...location, body_paragraph_number: 0 } } }),
    proposal({ paragraph_insert: { ...insertion, location: { ...location, next_paragraph: 3 } } as unknown as Insertion }),
    proposal({ target_json: JSON.stringify({ version: 5, ...insertion, text: "Different stored words." }) }),
    proposal({ target_json: JSON.stringify({ version: 5, ...insertion, location: { ...location, body_paragraph_number: 9 } }) }),
    proposal({ target_json: JSON.stringify({ version: 5, ...insertion, location: { ...location, headings: ["A different section"] } }) }),
  ]) {
    const view = render(original)
    try {
      expect(view.root.querySelector(".paragraph-insertion-unavailable")!.textContent).toContain("Paragraph insertion preview unavailable")
      expect(view.root.querySelector(".proposal-redline,.table-row-proposal,.paragraph-spacing-proposal,ins,del")).toBeNull()
      expect(view.root.textContent).not.toContain(original.old_text); expect(view.root.textContent).not.toContain(original.new_text); acceptance(view, false)
    } finally { view.close() }
  }
})

test("identical references with identical surrounding words remain distinguishable by verified body paragraph number", () => {
  const first = proposal({ id: 92, paragraph_insert: { ...insertion, location: { ...location, body_paragraph_number: 8 } } })
  const second = proposal({ id: 93, paragraph_insert: { ...insertion, location: { ...location, body_paragraph_number: 18 } } })
  const originals = [first, second], snapshot = JSON.stringify(originals), view = render(originals)
  try {
    const cards = [...view.root.querySelectorAll(".paragraph-insertion-proposal")]
    expect(cards.length).toBe(2)
    expect(cards.map(card => card.querySelector("summary")!.textContent)).toEqual(["Body paragraph 8", "Body paragraph 18"])
    for (const card of cards) {
      expect(card.querySelector("details")!.open).toBe(true)
      expect(card.querySelector(".paragraph-insertion-reference .paragraph-insertion-text")!.textContent).toBe(reference)
      expect([...card.querySelectorAll(".paragraph-insertion-source-text")].map(node => node.textContent)).toEqual([...location.headings, location.previous_paragraph, location.next_paragraph])
    }
    expect([...view.root.querySelectorAll("button")].filter(button => button.textContent === "Accept").every(button => !button.disabled)).toBe(true)
    expect(view.button("Accept all").disabled).toBe(false)
    expect(JSON.stringify(originals)).toBe(snapshot)
  } finally { view.close() }
})

test("source context retains exact escaped whitespace text and distinguishes absent adjacent paragraphs from empty ones", () => {
  const original = proposal({ paragraph_insert: { ...insertion, location: { body_paragraph_number: 1, headings: ["Terms <a>literal</a>  & conditions"], previous_paragraph: null, next_paragraph: "  " } } })
  const snapshot = JSON.stringify(original), view = render(original)
  try {
    const context = view.root.querySelector(".paragraph-insertion-location")!
    expect([...context.querySelectorAll(".paragraph-insertion-source-text")].map(node => node.textContent)).toEqual(["Terms <a>literal</a>  & conditions", "  "])
    expect(context.querySelector(".paragraph-insertion-context-note")!.textContent).toBe("No adjacent body paragraph")
    expect(context.querySelector(".paragraph-insertion-note")!.textContent).toBe("Whitespace only")
    expect(context.querySelector("a")).toBeNull()
    acceptance(view, true); expect(JSON.stringify(original)).toBe(snapshot)
  } finally { view.close() }
  const withoutHeading = render(proposal({ paragraph_insert: { ...insertion, location: { body_paragraph_number: 1, headings: [], previous_paragraph: "", next_paragraph: null } } }))
  try {
    expect(withoutHeading.root.textContent).toContain("No heading")
    expect(withoutHeading.root.querySelector(".paragraph-insertion-source-text")!.textContent).toBe("")
    expect(withoutHeading.root.querySelector(".paragraph-insertion-note")!.textContent).toBe("Empty paragraph")
    acceptance(withoutHeading, true)
  } finally { withoutHeading.close() }
})
