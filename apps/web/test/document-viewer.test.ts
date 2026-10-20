import { afterEach, describe, expect, test } from "bun:test"
import { createElement } from "react"
import { renderToStaticMarkup } from "react-dom/server"
import { combinedNotUpdatedNotice, DownloadActions, notesNotUpdatedNotice, preparedNotice, redlineNotUpdatedNotice, staleExportNotice } from "../src/components/DocumentViewer"
import { acceptAllRedlines, acceptRedline, type Redline } from "../src/api/ingest/redlines"

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

describe("the Prepare line", () => {
  test("counts only the batch's proposals still pending and fresh, and goes when none is left", () => {
    expect(preparedNotice([1, 2, 3], [proposal(1), proposal(2), proposal(3)])).toBe("3 proposed changes ready for review.")
    // An older stale proposal outside the batch is the stale notice's to report.
    expect(preparedNotice([2, 3], [proposal(1, true), proposal(2), proposal(3)])).toBe("2 proposed changes ready for review.")
    // Card 1 accepted, card 3 stale since: one is left to review.
    expect(preparedNotice([1, 2, 3], [proposal(2), proposal(3, true)])).toBe("1 proposed change ready for review.")
    expect(preparedNotice([1, 2, 3], [proposal(3, true)])).toBeUndefined()
    expect(preparedNotice([1, 2], [])).toBeUndefined()
    // Before the list refetches, a new batch is not yet counted.
    expect(preparedNotice([4], [proposal(3, true)])).toBeUndefined()
    expect(preparedNotice([], [proposal(3)])).toBe("No changes were proposed.")
  })
})

describe("the delivered redline after an accept", () => {
  test("a redline left as it was is named in plain words; anything else says nothing", () => {
    expect(redlineNotUpdatedNotice({ redline_updated: false, redline_skip_reason: "pending_redline_proposals" })).toBe(
      "The redline was not updated: the redline has pending proposed changes.")
    expect(redlineNotUpdatedNotice({ redline_updated: false, redline_skip_reason: "rebuild_failed" })).toBe(
      "The redline was not updated: it could not be rebuilt from the accepted document.")
    expect(redlineNotUpdatedNotice({ redline_updated: false, redline_skip_reason: "unknown_reason" })).toBe("The redline was not updated.")
    expect(redlineNotUpdatedNotice({ redline_updated: true, notes_updated: false })).toBeUndefined()
    expect(redlineNotUpdatedNotice({})).toBeUndefined()
    expect(redlineNotUpdatedNotice(undefined)).toBeUndefined()
  })
})

describe("drafting notes after an accept", () => {
  test("a skipped notes update is named in plain words; anything else says nothing", () => {
    expect(notesNotUpdatedNotice({ notes_updated: false, notes_skip_reason: "pending_notes_proposals" })).toBe(
      "Change accepted. The drafting notes were not updated: the notes have pending proposed changes.")
    expect(notesNotUpdatedNotice({ notes_updated: false, notes_skip_reason: "write_failed" }, true)).toBe(
      "Changes accepted. The drafting notes were not updated: the notes file could not be saved.")
    expect(notesNotUpdatedNotice({ notes_updated: false, notes_skip_reason: "unknown_reason" })).toBe(
      "Change accepted. The drafting notes were not updated.")
    expect(notesNotUpdatedNotice({ notes_updated: true })).toBeUndefined()
    expect(notesNotUpdatedNotice({})).toBeUndefined()
    expect(notesNotUpdatedNotice(undefined)).toBeUndefined()
  })
})

describe("the saved combined file after an accept", () => {
  const originalFetch = globalThis.fetch
  afterEach(() => { globalThis.fetch = originalFetch })

  test("an accept reports whether the combined file was rebuilt, and an accept without one reports nothing", async () => {
    const answer = (body: unknown) => { globalThis.fetch = (async () => Response.json(body)) as unknown as typeof fetch }
    answer({ ok: true, notes_updated: true, combined_updated: false, combined_skip_reason: "build_refused" })
    expect(await acceptRedline("m", 1)).toEqual({ notes_updated: true, combined_updated: false, combined_skip_reason: "build_refused" })
    answer({ ok: true, accepted: 2, combined_updated: true })
    expect(await acceptAllRedlines("m", "Supply note - clean.docx")).toEqual({ combined_updated: true })
    answer({ ok: true, notes_updated: true, redline_updated: true })
    expect(await acceptRedline("m", 1)).not.toHaveProperty("combined_updated")
  })

  test("a combined file left as it was is named in plain words; anything else says nothing", () => {
    expect(combinedNotUpdatedNotice({ combined_updated: false, combined_skip_reason: "build_refused" })).toBe(
      "The combined file was not updated: the memo and the draft can no longer be combined. Download them separately.")
    expect(combinedNotUpdatedNotice({ combined_updated: false, combined_skip_reason: "combined_changed" })).toBe(
      "The combined file was not updated: the combined file has changes of its own, which are kept.")
    expect(combinedNotUpdatedNotice({ combined_updated: false, combined_skip_reason: "unknown_reason" })).toBe("The combined file was not updated.")
    expect(combinedNotUpdatedNotice({ combined_updated: true })).toBeUndefined()
    expect(combinedNotUpdatedNotice({ notes_updated: false })).toBeUndefined()
    expect(combinedNotUpdatedNotice(undefined)).toBeUndefined()
  })
})
