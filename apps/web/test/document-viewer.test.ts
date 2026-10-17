import { describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { DownloadActions, staleExportNotice } from "../src/components/DocumentViewer"
import type { Redline } from "../src/api/ingest/redlines"

const proposal = (id: number, stale?: true): Redline => ({
  id, doc_name: "working.docx", scope: "phrase", find_text: "thirty days", old_text: "thirty days", new_text: "forty-five days",
  author: "Synthetic Reviewing Lawyer", review_comment: "", created_at: id,
  ...(stale ? { stale, stale_reason: "This proposal targets an earlier Word revision. Prepare a fresh change list." } : {}),
})
const buttons = (redlines: Redline[], preview = true) =>
  [...renderToStaticMarkup(createElement(DownloadActions, { redlines, preview, onDownload: () => {} }))
    .matchAll(/<button([^>]*)>([\s\S]*?)<\/button>/g)].map(([, attributes, content]) => ({
    label: content.replace(/<[^>]+>/g, "").trim(), disabled: attributes.includes("disabled"),
  }))

describe("document viewer downloads", () => {
  test("the current Word file stays downloadable whatever is pending", () => {
    expect(buttons([])).toEqual([{ label: "Download", disabled: false }])
    expect(buttons([proposal(1, true)])).toEqual([{ label: "Download", disabled: false }])
    expect(buttons([proposal(1, true), proposal(2)])).toEqual([
      { label: "Download redline", disabled: false }, { label: "Download", disabled: false },
    ])
    expect(buttons([proposal(2)], false)).toEqual([
      { label: "Download redline", disabled: true }, { label: "Download", disabled: false },
    ])
  })

  test("stale proposals are named as excluded with the action that clears them", () => {
    expect(staleExportNotice([proposal(2)])).toBeUndefined()
    expect(staleExportNotice([proposal(1, true), proposal(2)])).toBe(
      "1 stale proposal is left out of the tracked changes and the redline download. Reject it to clear it; prepare a fresh change list if the change is still needed.")
    expect(staleExportNotice([proposal(1, true), proposal(3, true)])).toStartWith("2 stale proposals are left out")
    expect(staleExportNotice([proposal(1, true), proposal(3, true)])).toContain("Reject them to clear them")
  })
})
