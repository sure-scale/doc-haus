import { expect, test } from "bun:test"
import type { WordConnectionState } from "../../../dochaus/lib/word-addin-contract"
import { savePresentation } from "./pane-presentation"
import type { PaneState } from "./sync"

const current: WordConnectionState = { version: { base: { id: "v2", number: 2 }, current: { id: "v2", number: 2 }, state: "current" }, lease: { state: "available", expiresAt: null } }
const state = (input: Partial<PaneState> = {}): PaneState => ({ status: "not-saved", busy: false, canSave: true, content: "unverified", connection: current, message: "Ready to save.", ...input })

test("matching server versions and restored receipts do not imply that fresh Word content is saved", () => {
  const receipt = { base: "v2", number: 2, fingerprint: "a".repeat(64), at: 1_700_000_000_000, overlap: false }
  for (const acknowledgement of [undefined, receipt]) {
    const view = savePresentation(state({ acknowledgement }))
    expect(view.badge).toBe("Ready")
    expect(view.message).toBe("Version 2 in doc.haus")
    expect(view.primary).toBe("save")
    expect(view.mismatch).toBeUndefined()
    expect(view.help).toBeUndefined()
  }
  expect(savePresentation(state({ connection: undefined })).message).toContain("Save to confirm this copy")
})

test("only observed changes get Changes to save; verified saves get the close action and one receipt line", () => {
  expect(savePresentation(state({ content: "changed" }))).toMatchObject({ badge: "Changes to save", primary: "save" })
  const receipt = { base: "v2", number: 2, fingerprint: "a".repeat(64), at: 1_700_000_000_000, overlap: false }
  const view = savePresentation(state({ status: "saved", content: "confirmed", acknowledgement: receipt }))
  expect(view.badge).toBe("Saved")
  expect(view.primary).toBe("close")
  expect(view.message).toContain("Version 2 received at")
  expect(view.mismatch).toBeUndefined()
  expect(savePresentation(state({ status: "saving", message: "Saving in Word…" }))).toMatchObject({ badge: "Saving…", primary: "save", message: "Saving in Word…" })
  expect(savePresentation(state({ problem: true, content: "confirmed", acknowledgement: receipt, message: "doc.haus did not confirm this save." }))).toMatchObject({ badge: "Needs attention", primary: "save", message: "doc.haus did not confirm this save." })
})

test("blocked copies show contextual recovery and only a real version mismatch gets the extra version line", () => {
  const stale = savePresentation(state({ canSave: false, connection: { ...current, version: { ...current.version, base: { id: "v1", number: 1 }, state: "stale" } } }))
  expect(stale.badge).toBe("Saving paused")
  expect(stale.mismatch).toBe("This copy: version 1. In doc.haus: version 2.")
  expect(stale.help).toContain("Save As")
  const other = savePresentation(state({ canSave: false, connection: { ...current, lease: { state: "other", generation: "other-owner", expiresAt: 90_000 } } }))
  expect(other.badge).toBe("Saving paused")
  expect(other.mismatch).toBeUndefined()
  expect(other.help).toContain("take over")
})
