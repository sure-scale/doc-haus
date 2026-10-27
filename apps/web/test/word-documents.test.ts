import { afterEach, describe, expect, test } from "bun:test"
import { fetchWordConfig, fetchWordDocuments, linkedWordDocument, recoverWordEditing, revokeWordDocument, type WordOpenDocument } from "../src/api/ingest/word"
import { wordSetupReason, wordVersionsChanged, wordWarning } from "../src/components/WordDocuments"
import { versionOrigin } from "../src/components/DocumentVersions"

const originalFetch = globalThis.fetch
const stubFetch = (respond: () => Promise<Response>): typeof fetch => Object.assign(respond, { preconnect: originalFetch.preconnect })
afterEach(() => { globalThis.fetch = originalFetch })
const entry = (patch: Partial<WordOpenDocument> = {}): WordOpenDocument => ({ name: "Agreement.docx", open: true, currentVersion: "v-1", overlap: false, ...patch })

describe("Word connection", () => {
  test("editing recovery compares the displayed version and lease before changing ownership", async () => {
    let sent: { url: string; init?: RequestInit } | undefined
    globalThis.fetch = (async (input, init) => {
      sent = { url: String(input), init }
      return Response.json({ ok: true })
    }) as typeof fetch
    await recoverWordEditing("m", entry({ lease: { state: "other", generation: "lease-a", expiresAt: Date.now() + 1000 } }), "takeover")
    expect(sent).toMatchObject({ url: expect.stringContaining("/matters/m/documents/word/recover?name=Agreement.docx"), init: {
      method: "POST", body: JSON.stringify({ action: "takeover", expectedCurrent: "v-1", expectedGeneration: "lease-a" }),
    } })
  })
  test("a linked copy sends the browser origin through normal ingest and returns the DOCX without reading its grant", async () => {
    let sent: { url: string; init?: RequestInit } | undefined
    const bytes = new Uint8Array([80, 75, 3, 4])
    globalThis.fetch = (async (input, init) => {
      sent = { url: String(input), init }
      return new Response(bytes, { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } })
    }) as typeof fetch
    const copy = await linkedWordDocument("matter / one", "Agreement #1 & schedule.docx", "http://localhost:5173")
    expect(sent).toMatchObject({
      url: expect.stringContaining("/matters/matter%20%2F%20one/documents/word?name=Agreement%20%231%20%26%20schedule.docx"),
      init: { method: "POST", headers: { "Content-Type": "application/json" }, body: '{"origin":"http://localhost:5173"}' },
    })
    expect(new Uint8Array(await copy.arrayBuffer())).toEqual(bytes)
  })

  test("status keeps inactive linked copies available for disconnect, and disconnect needs no document token", async () => {
    const documents = [entry({ open: false, overlap: true })]
    const calls: Array<{ url: string; init?: RequestInit }> = []
    globalThis.fetch = (async (input, init) => {
      calls.push({ url: String(input), init })
      return init?.method === "DELETE" ? new Response(null, { status: 204 }) : Response.json({ documents })
    }) as typeof fetch
    expect(await fetchWordDocuments("m")).toEqual(documents)
    await revokeWordDocument("m", "Agreement.docx")
    expect(calls).toEqual([
      { url: expect.stringContaining("/matters/m/documents/word"), init: undefined },
      { url: expect.stringContaining("/matters/m/documents/word?name=Agreement.docx"), init: { method: "DELETE" } },
    ])
  })

  test("setup and write failures retain the service's actionable explanation", async () => {
    globalThis.fetch = stubFetch(async () => Response.json({ available: false, origin: null, manifestUrl: null, reason: "A trusted HTTPS connection has not been set up." }))
    const config = await fetchWordConfig()
    expect(wordSetupReason(config)).toBe("A trusted HTTPS connection has not been set up.")
    globalThis.fetch = stubFetch(async () => Response.json({ error: "Word setup must be completed before downloading a linked copy." }, { status: 503 }))
    await expect(linkedWordDocument("m", "Agreement.docx", "http://localhost")).rejects.toThrow("Word setup must be completed")
    await expect(revokeWordDocument("m", "Agreement.docx")).rejects.toThrow("Word setup must be completed")
  })
})

describe("Word status semantics", () => {
  test("an active editing lease explains blocked changes even after heartbeat contact stops", () => {
    const locked = entry({ open: false, lease: { state: "other", generation: "lease-a", expiresAt: Date.now() + 1000 } })
    expect(wordWarning([locked], "Agreement.docx")).toContain("Changes here are locked")
    expect(wordWarning([locked], "Schedule.docx")).toBeUndefined()
  })
  test("only received version identities trigger a document refresh; heartbeat and overlap changes alone do not", () => {
    expect(wordVersionsChanged([entry()], [entry({ open: false, overlap: true })])).toBe(false)
    expect(wordVersionsChanged([entry()], [entry({ currentVersion: "v-2" })])).toBe(true)
    // A first linked copy may save before the next status poll sees its grant.
    expect(wordVersionsChanged([], [entry({ currentVersion: "v-2" })])).toBe(true)
    expect(wordVersionsChanged([], [entry({ currentVersion: null })])).toBe(false)
  })

  test("any active Word document warns matter-wide, while inactive links leave no open warning", () => {
    expect(wordWarning([entry({ open: false, overlap: true })])).toBeUndefined()
    const warning = wordWarning([entry(), entry({ name: "Schedule.docx" })], "Another.docx")
    expect(warning).toStartWith("Agreement.docx, Schedule.docx are open in Word.")
    expect(warning).toContain("Use Save to doc.haus in Word and close the document before changing it here")
    expect(warning).toContain("compared with the current version")
    expect(wordWarning([entry(), entry({ name: "Schedule.docx" })], "Schedule.docx")).toStartWith("Schedule.docx is open in Word.")
  })

  test("the explicit Word writer has a distinct saved-from-Word version label", () => {
    expect(versionOrigin({ id: "v-2", parent: "v-1", origin: "word" })).toBe("Saved from Word")
  })
})
