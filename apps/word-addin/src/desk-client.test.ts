import { expect, test } from "bun:test"
import { WORD_ADDIN_VERSION, type WordDocumentLink } from "../../../dochaus/lib/word-addin-contract"
import { deskRequestHandler } from "./desk-client"
import { DeskRequestError } from "./desk-host"

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
