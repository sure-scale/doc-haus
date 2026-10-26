import { expect, test } from "bun:test"
import { WordClient, httpsOrigin, readLink } from "./client"
import { renderManifest } from "./manifest"
import { WORD_ADDIN_ID, WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"

const link = {
  version: 1 as const,
  apiOrigin: "https://localhost:4612",
  origin: "http://localhost:4444",
  matter: "synthetic",
  matterName: "Synthetic matter",
  document: "sample.docx",
  documentName: "Sample",
  token: "synthetic-test-capability",
  baseVersion: "original",
}

test("only same-origin HTTPS receivers receive a document capability", () => {
  expect(readLink(link, link.apiOrigin)).toEqual(link)
  expect(readLink(JSON.stringify(link), link.apiOrigin)).toEqual(link)
  expect(readLink({ ...link, apiOrigin: "https://other.example" }, link.apiOrigin)).toBeNull()
  for (const origin of [
    "http://localhost:4612",
    "https://user:pass@localhost",
    "https://localhost/path",
    "https://localhost/?q=x",
  ]) {
    expect(() => httpsOrigin(origin)).toThrow()
  }
})

test("save uses bearer-only POST, disallows redirects, and requires a version acknowledgement", async () => {
  let init: RequestInit | undefined
  let url = ""
  const client = new WordClient(link, "session", async (input, options) => {
    url = String(input)
    init = options
    return Response.json({ version: { id: "v1", number: 1, parent: "original", created: true }, overlap: false })
  })
  const value = await client.save(new Uint8Array([1]), "original")
  expect(value.version.id).toBe("v1")
  expect(url).toBe("https://localhost:4612/word/saves")
  expect(url).not.toContain(link.token)
  expect(init?.redirect).toBe("error")
  expect(init?.credentials).toBe("omit")
  expect(init?.referrerPolicy).toBe("no-referrer")
  expect(new Headers(init?.headers).get("x-dochaus-word-base-version")).toBe("original")
  const unconfirmed = new WordClient(link, "session", async () => Response.json({ ok: true }))
  await expect(unconfirmed.save(new Uint8Array([1]), "original")).rejects.toThrow("did not confirm")
})

test("manifest uses the shared identity and designated auto-open pane", () => {
  const manifest = renderManifest(link.apiOrigin)
  expect(manifest).toContain(`<Id>${WORD_ADDIN_ID}</Id>`)
  expect(manifest).toContain(`<Version>${WORD_ADDIN_VERSION}</Version>`)
  expect(manifest).toContain("<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>")
  expect(manifest).toContain("https://localhost:4612/word-addin/pane.html")
  expect(manifest.match(/xsi:type="ShowTaskpane"/g)).toHaveLength(1)
  expect(manifest).not.toContain('xsi:type="ExecuteFunction"')
  // A function runtime must not boot the pane's document writer as well.
  expect(manifest).not.toContain("<FunctionFile")
  expect(() => renderManifest("http://localhost:4612")).toThrow()
})

test("save deadline covers a success body that stalls after headers", async () => {
  const server = Bun.serve({
    port: 0,
    fetch: () =>
      new Response(
        new ReadableStream({
          start(controller) {
            controller.enqueue(new TextEncoder().encode('{"version":'))
          },
        }),
        { headers: { "Content-Type": "application/json" } },
      ),
  })
  try {
    const client = new WordClient(link, "session", (_url, init) => fetch(server.url, init), 25)
    await expect(client.save(new Uint8Array([1]), "original")).rejects.toThrow("did not confirm")
  } finally {
    server.stop(true)
  }
})

test("heartbeat and Done release unused success bodies", async () => {
  let cancelled = 0
  const client = new WordClient(
    link,
    "session",
    async () =>
      new Response(
        new ReadableStream({
          cancel() {
            cancelled++
          },
        }),
      ),
  )
  await client.heartbeat()
  await client.done()
  expect(cancelled).toBe(2)
})

test("default transport invokes native fetch with the global receiver", async () => {
  const previous = globalThis.fetch
  let called = false
  globalThis.fetch = function (this: unknown) {
    if (this !== globalThis) throw new TypeError("Illegal invocation")
    called = true
    return Promise.resolve(
      Response.json({
        version: { id: "v1", number: 1, parent: "original", created: true },
        overlap: false,
      }),
    )
  } as unknown as typeof fetch
  try {
    const client = new WordClient(link, "session")
    expect((await client.save(new Uint8Array([1]), "original")).version.id).toBe("v1")
    expect(called).toBe(true)
  } finally {
    globalThis.fetch = previous
  }
})
