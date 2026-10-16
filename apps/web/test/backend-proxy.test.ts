import { expect, test } from "bun:test"
import { createServer } from "vite"
import { backendProxy } from "../backend-proxy"

test("web proxy preserves native paths, query, directory headers, binary bodies, downloads and SSE", async () => {
  const upstream = Bun.serve({ port: 0, async fetch(request) {
    const url = new URL(request.url)
    if (url.pathname === "/api/events") return new Response(new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode("data: first\n\n"))
        setTimeout(() => { controller.enqueue(new TextEncoder().encode("data: second\n\n")); controller.close() }, 200)
      },
    }), { headers: { "content-type": "text/event-stream" } })
    if (url.pathname === "/download") return new Response(new Uint8Array([0, 255, 3]), {
      headers: { "content-disposition": 'attachment; filename="sample.bin"', "content-type": "application/octet-stream" },
    })
    return Response.json({ path: url.pathname, query: url.search, directory: request.headers.get("x-opencode-directory"),
      contentType: request.headers.get("content-type"), method: request.method, bytes: [...new Uint8Array(await request.arrayBuffer())] })
  } })
  const web = await createServer({ configFile: false, appType: "custom", server: {
    host: "127.0.0.1", port: 0, proxy: backendProxy({ DOC_HAUS_ENGINE_PORT: String(upstream.port), DOC_HAUS_INGEST_PORT: String(upstream.port) }),
  } })
  try {
    await web.listen()
    const address = web.httpServer?.address()
    if (!address || typeof address === "string") throw new Error("Proxy fixture failed to listen")
    const origin = `http://127.0.0.1:${address.port}`
    for (const prefix of ["engine", "ingest"]) {
      const result = await fetch(`${origin}/${prefix}/api/upload?name=a%20b&limit=2`, {
        method: "POST", headers: { "x-opencode-directory": "/synthetic/matter", "content-type": "application/octet-stream" }, body: new Uint8Array([0, 255, 1]),
      })
      expect(await result.json()).toEqual({ path: "/api/upload", query: "?name=a%20b&limit=2", directory: "/synthetic/matter", contentType: "application/octet-stream", method: "POST", bytes: [0, 255, 1] })
      const download = await fetch(`${origin}/${prefix}/download`)
      expect(download.headers.get("content-disposition")).toBe('attachment; filename="sample.bin"')
      expect([...new Uint8Array(await download.arrayBuffer())]).toEqual([0, 255, 3])
      const stream = await fetch(`${origin}/${prefix}/api/events`)
      expect(stream.headers.get("content-type")).toContain("text/event-stream")
      const reader = stream.body!.getReader()
      const first = await reader.read()
      expect(new TextDecoder().decode(first.value)).toBe("data: first\n\n")
      const second = await reader.read()
      expect(new TextDecoder().decode(second.value)).toBe("data: second\n\n")
      expect((await reader.read()).done).toBe(true)
    }
    expect((await fetch(`${origin}/engineer/api/upload`)).status).toBe(404)
  } finally { await web.close(); upstream.stop(true) }
}, 15_000)
