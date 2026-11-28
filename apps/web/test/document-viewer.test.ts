import { afterEach, describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { Window } from "happy-dom"
import { FILLED_VALUE_FILL } from "../../../dochaus/lib/marker-patterns"
import { fetchDocumentText } from "../src/api/ingest/documents"
import { acceptAllRedlines, acceptRedline, type Redline } from "../src/api/ingest/redlines"
import { citationPassage } from "../src/components/citation-target"
import { decorateFilledValues } from "../src/components/document-markers"
import {
  acceptedEditsHint,
  combinedNotUpdatedNotice,
  ExtractedText,
  notesNotUpdatedNotice,
  pdfCitationStatus,
  preparedNotice,
  redlineNotUpdatedNotice,
  requestEditsLine,
  staleExportNotice,
} from "../src/components/DocumentViewer"

const proposal = (id: number, stale?: true): Redline => ({
  id,
  doc_name: "working.docx",
  scope: "phrase",
  find_text: "thirty days",
  old_text: "thirty days",
  new_text: "forty-five days",
  author: "Synthetic Reviewing Lawyer",
  review_comment: "",
  created_at: id,
  ...(stale ? { stale, stale_reason: "Targets an earlier Word revision." } : {}),
})

test("PDF citation status quotes preserve literal source characters and cap prefix length", () => {
  const excerpt = "# Token `A|B` uses \\*literal\\* and **stars**.\n---\nKeep  spaces."
  const shown = renderToStaticMarkup(pdfCitationStatus({ section: "Tokens", excerpt }, { status: "matched", page: 2 }))
  expect(shown).toContain(`<q>${excerpt}</q>`)
  const long = "A|B ".repeat(30)
  const cut = renderToStaticMarkup(pdfCitationStatus({ section: "Tokens", excerpt: long }, { status: "matched", page: 2 }))
  const quote = cut.match(/<q>([\s\S]*)<\/q>/)?.[1]
  expect(quote?.endsWith("…")).toBe(true)
  expect(long.startsWith(quote!.slice(0, -1))).toBe(true)
  expect(quote!.length).toBeLessThanOrEqual(101)
})

test("highlights exactly the spans a batch preview shaded as filled values", () => {
  const browser = new Window()
  try {
    browser.document.body.innerHTML = `<style>.docx-a { background: #${FILLED_VALUE_FILL}; } .docx-b { background: #FFF2CC; }</style>
      <div id="page"><p>Dear <span class="docx-a">Example Contact</span>, re <span class="docx-b">[insert date]</span> and <span>plain text</span>.</p></div>`
    const page = browser.document.getElementById("page") as unknown as HTMLElement
    expect(decorateFilledValues(page)).toBe(1)
    const marked = [...page.querySelectorAll(".doc-filled")]
    expect(marked.map((span) => span.textContent)).toEqual(["Example Contact"])
  } finally {
    browser.happyDOM.abort()
  }
})

describe("proposal status notices", () => {
  test("stale notice appears only when stale proposals exist and reflects their count", () => {
    expect(staleExportNotice([proposal(2)])).toBeUndefined()
    expect(staleExportNotice([proposal(1, true), proposal(2)])).toContain("1 stale proposal")
    expect(staleExportNotice([proposal(1, true), proposal(3, true)])).toContain("2 stale proposals")
  })

  test("attribution hints stay silent when unnecessary and never attribute edits to doc.haus itself", () => {
    expect(requestEditsLine(0, "Pat Lawyer")).toBeUndefined()
    expect(requestEditsLine(1, "Pat Lawyer")).toContain("Pat Lawyer")
    expect(requestEditsLine(3, "doc.haus")).not.toContain("doc.haus")
    expect(requestEditsLine(3, null)).toBeDefined()

    const by = (...authors: string[]) => authors.map((author) => ({ author }))
    expect(acceptedEditsHint(by("Morgan Okafor"), "Pat Lawyer")).toContain("Pat Lawyer")
    expect(acceptedEditsHint(by(" pat lawyer", "Pat Lawyer"), "Pat Lawyer")).toBeUndefined()
    expect(acceptedEditsHint(by("Morgan Okafor"), "doc.haus")).not.toContain("doc.haus")
    expect(acceptedEditsHint(by("doc.haus"), null)).toBeUndefined()
    expect(acceptedEditsHint(by("Morgan Okafor"), undefined)).toBeUndefined()
  })

  test("preparedNotice counts only fresh pending proposals from the batch", () => {
    expect(preparedNotice([1, 2, 3], [proposal(1), proposal(2), proposal(3)])).toContain("3 ")
    expect(preparedNotice([2, 3], [proposal(1, true), proposal(2), proposal(3)])).toContain("2 ")
    expect(preparedNotice([1, 2, 3], [proposal(2), proposal(3, true)])).toContain("1 ")
    expect(preparedNotice([1, 2, 3], [proposal(3, true)])).toBeUndefined()
    expect(preparedNotice([1, 2], [])).toBeUndefined()
    expect(preparedNotice([4], [proposal(3, true)])).toBeUndefined()
    expect(preparedNotice([], [proposal(3)])).toBeDefined()
  })

  test("skipped redline, notes, and combined update notices appear only when updated is false", () => {
    for (const reason of ["pending_redline_proposals", "rebuild_failed", "unknown_reason"]) {
      expect(redlineNotUpdatedNotice({ redline_updated: false, redline_skip_reason: reason })).toBeDefined()
    }
    expect(redlineNotUpdatedNotice({ redline_updated: true, notes_updated: false })).toBeUndefined()
    expect(redlineNotUpdatedNotice({})).toBeUndefined()
    expect(redlineNotUpdatedNotice(undefined)).toBeUndefined()

    for (const reason of ["pending_notes_proposals", "write_failed", "unknown_reason"]) {
      expect(notesNotUpdatedNotice({ notes_updated: false, notes_skip_reason: reason })).toBeDefined()
    }
    expect(notesNotUpdatedNotice({ notes_updated: true })).toBeUndefined()
    expect(notesNotUpdatedNotice({})).toBeUndefined()
    expect(notesNotUpdatedNotice(undefined)).toBeUndefined()

    for (const reason of ["build_refused", "combined_changed", "unknown_reason"]) {
      expect(combinedNotUpdatedNotice({ combined_updated: false, combined_skip_reason: reason })).toBeDefined()
    }
    expect(combinedNotUpdatedNotice({ combined_updated: true })).toBeUndefined()
    expect(combinedNotUpdatedNotice({ notes_updated: false })).toBeUndefined()
    expect(combinedNotUpdatedNotice(undefined)).toBeUndefined()
  })
})

describe("accept and extracted-text API contracts", () => {
  const originalFetch = globalThis.fetch
  afterEach(() => { globalThis.fetch = originalFetch })

  test("acceptRedline and acceptAllRedlines preserve combined-file rebuild fields only when reported", async () => {
    const answer = (body: unknown) => { globalThis.fetch = (async () => Response.json(body)) as unknown as typeof fetch }
    answer({ ok: true, notes_updated: true, combined_updated: false, combined_skip_reason: "build_refused" })
    expect(await acceptRedline("m", 1)).toEqual({ notes_updated: true, combined_updated: false, combined_skip_reason: "build_refused" })
    answer({ ok: true, accepted: 2, combined_updated: true })
    expect(await acceptAllRedlines("m", "Supply note - clean.docx")).toEqual({ combined_updated: true })
    answer({ ok: true, notes_updated: true, redline_updated: true })
    expect(await acceptRedline("m", 1)).not.toHaveProperty("combined_updated")
  })

  test("ExtractedText escapes angle brackets and marks only matched passages", () => {
    const text = "Subject: Supply terms\nFrom: Synthetic Buyer <buyer@example.invalid>\n\nPayment is due\n on delivery.\nThanks"
    const match = citationPassage(text, "Payment is due on delivery.")
    const html = renderToStaticMarkup(createElement(ExtractedText, { kind: "eml", text, match }))
    expect(html).toContain("&lt;buyer@example.invalid&gt;")
    expect(html).not.toContain("<buyer@example.invalid>")
    expect(html).toContain("<mark")
    expect(html).toContain("Payment is due\n on delivery.</mark>")
    for (const unmatched of [citationPassage(text, "Payment is due in advance."), undefined]) {
      expect(renderToStaticMarkup(createElement(ExtractedText, { kind: "xlsx", text, match: unmatched }))).not.toContain("<mark")
    }
  })

  test("fetchDocumentText encodes the filename and surfaces HTTP errors", async () => {
    const requests: string[] = []
    globalThis.fetch = (async (url: string) => {
      requests.push(String(url))
      return requests.length === 1 ? Response.json({ text: "Synthetic body", injection: null }) : Response.json({ error: "Not found" }, { status: 404 })
    }) as unknown as typeof fetch
    expect(await fetchDocumentText("m", "Client instructions.eml")).toBe("Synthetic body")
    expect(requests[0]).toEndWith("/matters/m/documents/text?name=Client%20instructions.eml")
    await expect(fetchDocumentText("m", "gone.xlsx")).rejects.toThrow("Not found")
  })
})
