import { expect, test } from "bun:test"
import { listReusableSignIns, reuseSignIn } from "../src/api/ingest/credential-reuse"

test("reuse discovery and explicit selection use the cookie-free cross-origin ingest contract", async () => {
  const original = globalThis.fetch
  const requests: Request[] = []
  const connection = { id: "selected-connection", integrationID: "opencode-go", label: "Go account", authType: "key" as const }
  const source = { id: "native-source", label: "OpenCode", status: "ready" as const, unavailableCount: 0, connections: [connection] }
  globalThis.fetch = (async (input, init) => {
    const request = new Request(input, init)
    requests.push(request)
    // Ingest permits the web origin without credentialed CORS. Native engine
    // authentication belongs to its server proxy, not a browser cookie.
    expect(init?.credentials).toBe("omit")
    expect(request.headers.has("cookie")).toBe(false)
    expect(request.headers.has("authorization")).toBe(false)
    return Response.json(request.method === "POST" ? { connection } : { sources: [source] }, {
      headers: { "Access-Control-Allow-Origin": "http://localhost:5174" },
    })
  }) as typeof fetch
  try {
    const sources = await listReusableSignIns()
    expect(sources).toEqual([source])
    expect(requests.map(request => request.method)).toEqual(["GET"])
    const selected = sources[0]!.connections[0]!
    expect(await reuseSignIn(sources[0]!.id, selected.id)).toEqual(connection)
    expect(await requests[1]!.json()).toEqual({ sourceId: "native-source", connectionId: "selected-connection" })
    expect(requests.map(request => new URL(request.url).pathname)).toEqual(["/auth/reuse", "/auth/reuse"])
  } finally { globalThis.fetch = original }
})
