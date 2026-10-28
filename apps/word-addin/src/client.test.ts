import { expect, test } from "bun:test"
import { ConnectionError, WordClient, httpsOrigin, isCurrentPaneVersion, readLink } from "./client"
import { renderManifest } from "./manifest"
import { wordBootstrapUrl } from "./bootstrap-contract"
import { WORD_ADDIN_ID, WORD_ADDIN_VERSION, type WordConnectionState } from "../../../dochaus/lib/word-addin-contract"

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

const connection = (id = "v1"): WordConnectionState => ({
  version: { base: { id, number: 1 }, current: { id, number: 1 }, state: "current" },
  lease: { state: "owned", generation: "generation", expiresAt: 90_000 },
})

test("latest-copy handoff stays scoped to the original HTTPS receiver and never changes the old base", async () => {
  let request: RequestInit | undefined
  const uri = `doc.haus://word/open?source=${encodeURIComponent(`${link.apiOrigin}/word/open/${"a".repeat(43)}`)}`
  const client = new WordClient(link, "session", async (url, init) => {
    expect(String(url)).toBe(`${link.apiOrigin}/word/latest`)
    request = init
    return Response.json({ uri, id: "ticket" })
  })
  expect(await client.latest("old-base")).toBe(uri)
  expect(JSON.parse(request!.body as string)).toEqual({ origin: link.origin })
  expect(new Headers(request!.headers).get("x-dochaus-word-base-version")).toBe("old-base")
  expect(link.baseVersion).toBe("original")
  for (const bad of ["https://attacker.example", uri.replace(encodeURIComponent(link.apiOrigin), encodeURIComponent("https://attacker.example")), `${uri}&extra=1`]) {
    const invalid = new WordClient(link, "session", async () => Response.json({ uri: bad }))
    await expect(invalid.latest("old-base")).rejects.toThrow("could not be opened")
  }
})

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
    return Response.json({
      version: { id: "v1", number: 1, parent: "original", created: true },
      overlap: false,
      connection: connection(),
    })
  })
  const value = await client.save(new Uint8Array([1]), "original", "generation")
  expect(value.version.id).toBe("v1")
  expect(url).toBe("https://localhost:4612/word/saves")
  expect(url).not.toContain(link.token)
  expect(init?.redirect).toBe("error")
  expect(init?.credentials).toBe("omit")
  expect(init?.referrerPolicy).toBe("no-referrer")
  expect(new Headers(init?.headers).get("x-dochaus-word-base-version")).toBe("original")
  expect(new Headers(init?.headers).get("x-dochaus-word-version")).toBe(WORD_ADDIN_VERSION)
  expect(new Headers(init?.headers).get("x-dochaus-word-lease")).toBe("generation")
  const unconfirmed = new WordClient(link, "session", async () => Response.json({ ok: true }))
  await expect(unconfirmed.save(new Uint8Array([1]), "original", "generation")).rejects.toThrow("did not confirm")
})

test("manifest uses the shared identity and designated auto-open pane", () => {
  const manifest = renderManifest(link.apiOrigin)
  expect(manifest).toContain(`<Id>${WORD_ADDIN_ID}</Id>`)
  expect(manifest).toContain(`<Version>${WORD_ADDIN_VERSION}</Version>`)
  expect(manifest).toContain("<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>")
  expect(manifest).toContain(`<AppDomain>https://doc.haus</AppDomain>`)
  expect(manifest).toContain(`<AppDomain>${link.apiOrigin}</AppDomain>`)
  expect(manifest).toContain(wordBootstrapUrl(link.apiOrigin))
  expect(manifest).toContain('DefaultValue="https://doc.haus/word-addin/assets/icon.png"')
  expect(
    manifest.match(new RegExp(`bootstrap\\?version=${WORD_ADDIN_VERSION.replaceAll(".", "\\.")}`, "g")),
  ).toHaveLength(2)
  expect(manifest.match(/xsi:type="ShowTaskpane"/g)).toHaveLength(1)
  expect(manifest).toContain('id="Button.Label" DefaultValue="doc.haus"')
  expect(manifest).not.toContain('xsi:type="ExecuteFunction"')
  // A function runtime must not boot the pane's document writer as well.
  expect(manifest).not.toContain("<FunctionFile")
  expect(() => renderManifest("http://localhost:4612")).toThrow()
})

test("an outdated client receives clear update guidance without retrying", async () => {
  let requests = 0
  const client = new WordClient(link, "session", async () => {
    requests++
    return new Response(null, { status: 426 })
  })
  await expect(client.save(new Uint8Array([1]), "original", "generation")).rejects.toThrow("out of date")
  expect(requests).toBe(1)
})

test("only the exact single current manifest URL version can start a pane", () => {
  const url = `${link.apiOrigin}/word-addin/pane.html`
  expect(isCurrentPaneVersion(`${url}?version=${WORD_ADDIN_VERSION}`)).toBe(true)
  for (const value of [
    url,
    `${url}?version=1.0.0.0`,
    `${url}?version=`,
    `${url}?version=${WORD_ADDIN_VERSION}&version=${WORD_ADDIN_VERSION}`,
    "invalid URL",
  ])
    expect(isCurrentPaneVersion(value)).toBe(false)
})

test("save deadline covers a success body that stalls after headers", async () => {
  const acknowledgement = {
    version: { id: "v1", number: 1, parent: "original", created: true },
    overlap: false,
    connection: connection(),
  }
  const fast = new WordClient(link, "session", async () => Response.json(acknowledgement), 25)
  expect(await fast.save(new Uint8Array([1]), "original", "generation")).toEqual(acknowledgement)

  const events: string[] = []
  const client = new WordClient(link, "session", async (_url, init) => {
    const signal = init!.signal!
    const response = new Response(new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode('{"version":'))
        signal.addEventListener("abort", () => {
          events.push("abort")
          controller.error(signal.reason)
        }, { once: true })
      },
    }), { headers: { "Content-Type": "application/json" } })
    const json = response.json.bind(response)
    response.json = () => {
      events.push("body")
      return json()
    }
    events.push("headers")
    return response
  }, 25)
  try {
    await client.save(new Uint8Array([1]), "original", "generation")
    throw new Error("Expected the body deadline to expire")
  } catch (error) {
    expect(error).toBeInstanceOf(ConnectionError)
    expect((error as ConnectionError).message).toContain("did not confirm")
    expect((error as ConnectionError).retryable).toBe(true)
  }
  expect(events).toEqual(["headers", "body", "abort"])
})

test("heartbeat and Done read authoritative state and carry the current base and fence", async () => {
  const requests: RequestInit[] = []
  const client = new WordClient(link, "session", async (_url, init) => {
    requests.push(init!)
    return Response.json(connection())
  })
  expect(await client.heartbeat("v1", "generation")).toEqual(connection())
  expect(await client.done("v1", "generation")).toEqual(connection())
  for (const request of requests) {
    expect(new Headers(request.headers).get("x-dochaus-word-base-version")).toBe("v1")
    expect(new Headers(request.headers).get("x-dochaus-word-lease")).toBe("generation")
    expect(new Headers(request.headers).get("x-dochaus-word-session")).toBe("session")
  }
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
        connection: connection(),
      }),
    )
  } as unknown as typeof fetch
  try {
    const client = new WordClient(link, "session")
    expect((await client.save(new Uint8Array([1]), "original", "generation")).version.id).toBe("v1")
    expect(called).toBe(true)
  } finally {
    globalThis.fetch = previous
  }
})

test("lease commands send conditional takeover observations and the release fence", async () => {
  const requests: RequestInit[] = []
  const client = new WordClient(link, "session", async (_url, init) => {
    requests.push(init!)
    return Response.json(connection("original"))
  })
  await client.lease("inspect", "original")
  await client.lease("claim", "original")
  await client.lease("takeover", "original", undefined, { current: "observed", generation: "observed-fence" })
  await client.lease("release", "original", "generation")
  expect(requests.map((request) => JSON.parse(String(request.body)))).toEqual([
    { action: "inspect", baseVersion: "original" },
    { action: "claim", baseVersion: "original" },
    { action: "takeover", baseVersion: "original", expectedCurrent: "observed", expectedGeneration: "observed-fence" },
    { action: "release", baseVersion: "original" },
  ])
  expect(new Headers(requests[3]!.headers).get("x-dochaus-word-lease")).toBe("generation")
})

test("conflict responses preserve authoritative version and ownership without retry permission", async () => {
  for (const status of [409, 410, 423]) {
    const state = connection("original")
    state.version.state = "stale"
    state.version.current = { id: "newer", number: 2 }
    state.lease.state = "other"
    const client = new WordClient(link, "session", async () => Response.json(state, { status }))
    try {
      await client.save(new Uint8Array([1]), "original", "generation")
      throw new Error("Expected refusal")
    } catch (error) {
      expect(error).toBeInstanceOf(ConnectionError)
      expect((error as ConnectionError).retryable).toBe(false)
      expect((error as ConnectionError).connection).toEqual(state)
    }
  }
})

test("malformed or inconsistent save ownership cannot become an acknowledgement", async () => {
  for (const state of [
    undefined,
    { version: {}, lease: {} },
    connection("different"),
    { ...connection(), lease: { state: "owned", generation: "different", expiresAt: 90_000 } },
  ]) {
    const client = new WordClient(link, "session", async () =>
      Response.json({
        version: { id: "v1", number: 1, parent: "original", created: true },
        overlap: false,
        connection: state,
      }),
    )
    await expect(client.save(new Uint8Array([1]), "original", "generation")).rejects.toThrow("did not confirm")
  }
})

test("inline latest content is scoped, bounded and integrity checked without changing the current base", async () => {
  const bytes = new Uint8Array([80, 75, 3, 4]), encoded = btoa(String.fromCharCode(...bytes))
  const digest = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)), byte => byte.toString(16).padStart(2, "0")).join("")
  const payload = { version: { id: "new-base", number: 2 }, bytes: encoded, sha256: digest }
  const client = new WordClient(link, "session", async (url, init) => {
    expect(String(url)).toBe(`${link.apiOrigin}/word/latest`)
    expect(JSON.parse(init!.body as string)).toEqual({ origin: link.origin, mode: "inline" })
    expect(new Headers(init!.headers).get("x-dochaus-word-base-version")).toBe("old-base")
    return Response.json(payload)
  })
  expect(await client.latestContent("old-base")).toEqual({ version: payload.version, bytes })
  expect(link.baseVersion).toBe("original")
  for (const bad of [{ ...payload, sha256: "0".repeat(64) }, { ...payload, bytes: "!bad" }, { ...payload, version: { id: "", number: 2 } }, { ...payload, bytes: "A".repeat(6_000_000) }]) {
    await expect(new WordClient(link, "session", async () => Response.json(bad)).latestContent("old-base")).rejects.toThrow("could not be verified")
  }
})
