import { expect, test } from "bun:test"
import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { deskRequestHandler } from "./desk-client"
import { DeskRequestError } from "./desk-host"
import { isDeskRequest } from "./desk-protocol"

const link: WordDocumentLink = { version: 1, apiOrigin: "https://localhost:4612", origin: "http://localhost:5274", matter: "synthetic-matter", matterName: "Synthetic matter", document: "synthetic-document", documentName: "Agreement.docx", baseVersion: "version-one", token: "synthetic-capability" }

test("desk reads use only scoped HTTPS endpoints and do not carry a save lease or a capability URL", async () => {
  let calls = 0
  const request = deskRequestHandler(link, async (input, init) => {
    calls++
    expect(String(input)).toBe(`${link.apiOrigin}/word/desk/source`)
    expect(String(input)).not.toContain(link.token)
    expect(init?.method).toBe("POST")
    expect(init?.credentials).toBe("omit")
    expect(init?.redirect).toBe("error")
    const headers = new Headers(init?.headers)
    expect(headers.get("Authorization")).toBe(`Bearer ${link.token}`)
    expect(headers.get("x-dochaus-word-version")).toBe(WORD_ADDIN_VERSION)
    expect(headers.has("x-dochaus-word-lease")).toBe(false)
    expect(headers.has("x-dochaus-word-session")).toBe(false)
    expect(JSON.parse(init?.body as string)).toEqual({ sourceId: "1", expectedHash: "a".repeat(64) })
    return Response.json({ text: "Synthetic saved source" })
  })
  expect(await request({ action: "source", sourceId: "1", expectedHash: "a".repeat(64) })).toEqual({ text: "Synthetic saved source" })
  expect(calls).toBe(1)
  expect(() => deskRequestHandler({ ...link, apiOrigin: "http://untrusted.example" })).toThrow("HTTPS")
})

test("a changed source is disclosed with the server refusal", async () => {
  let changed = false
  const request = deskRequestHandler(link, async () => changed
    ? Response.json({ code: "desk_source_changed", error: "This source changed. Refresh the source list and choose it again." }, { status: 409 })
    : Response.json({ text: "Original saved source" }))
  expect(await request({ action: "source", sourceId: "1", expectedHash: "a".repeat(64) })).toEqual({ text: "Original saved source" })
  changed = true
  const error = await request({ action: "source", sourceId: "1", expectedHash: "a".repeat(64) }).catch((error: unknown) => error)
  expect(error).toBeInstanceOf(DeskRequestError)
  expect((error as Error).message).toContain("This source changed")
})

test("Word handoffs inject the linked matter origin and validate the native target before returning it", async () => {
  const uri = `doc.haus://word/open?source=${encodeURIComponent(`${link.apiOrigin}/word/open/${"a".repeat(43)}`)}`
  const id = "00000000-0000-0000-0000-000000000001"
  const input = { action: "open" as const, sourceId: "1", expectedHash: "a".repeat(64) }
  const request = deskRequestHandler(link, async (url, init) => {
    expect(String(url)).toBe(`${link.apiOrigin}/word/desk/open`)
    expect(JSON.parse(init!.body as string)).toEqual({ sourceId: "1", expectedHash: input.expectedHash, origin: link.origin })
    expect(new Headers(init!.headers).has("x-dochaus-word-lease")).toBe(false)
    return Response.json({ uri, id })
  })
  expect(await request(input)).toEqual({ uri, id })
  expect(isDeskRequest(input)).toBe(true)
  expect(isDeskRequest({ ...input, origin: "https://untrusted.example" })).toBe(false)
  for (const invalid of ["https://untrusted.example", uri.replace(encodeURIComponent(link.apiOrigin), encodeURIComponent("https://untrusted.example")), `${uri}&extra=1`]) {
    const refuse = deskRequestHandler(link, async () => Response.json({ uri: invalid, id }))
    await expect(refuse(input)).rejects.toThrow("Word connection could not be prepared")
  }
})

test("document lists carry only the validated linked matter's browser URL for other file formats", async () => {
  const request = deskRequestHandler(link, async () => Response.json({ sources: [] }))
  expect(await request({ action: "sources" })).toEqual({ sources: [], matterUrl: `${link.origin}/matter/${link.matter}?view=documents` })
  expect(() => deskRequestHandler({ ...link, origin: "https://user:secret@untrusted.example" })).toThrow("browser origin")
})

test("document desk cannot forward attachment actions to another capability route", async () => {
  let calls = 0
  const request = deskRequestHandler(link, async () => { calls++; return Response.json({}) })
  await expect(request({ action: "chat-materials" })).rejects.toThrow("Open Attach from the chat")
  await expect(request({ action: "chat-upload", name: "Notes.txt", data: "YWJj", mode: "request" })).rejects.toThrow("Open Attach from the chat")
  expect(calls).toBe(0)
})
