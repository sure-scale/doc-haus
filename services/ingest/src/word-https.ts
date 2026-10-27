import path from "node:path"
import { fileURLToPath } from "node:url"
import { renderManifest } from "../../../apps/word-addin/src/manifest"
import { wordHttpsConfiguration, type WordHttpsConfiguration } from "./word-configuration"

type Fetch = (request: Request) => Response | Promise<Response>
const assets = fileURLToPath(new URL("../../../apps/word-addin/dist/", import.meta.url))
const tokenRoutes = new Set(["/word/lease", "/word/saves", "/word/heartbeat", "/word/done",
  "/word/desk/sources", "/word/desk/source", "/word/desk/versions", "/word/desk/compare"])
let startupFailure: string | undefined

/** A recoverable bridge failure must never make the ordinary product unavailable. */
export function wordHttpsFailure() { return startupFailure }

/** This HTTPS listener exposes static pane files and scoped document capabilities.
 * It never forwards the ordinary matter, document-read or administration API. */
export function wordHttpsHandler(fetch: Fetch, configuration: WordHttpsConfiguration, directory = assets): Fetch {
  return async request => {
    const url = new URL(request.url)
    if (tokenRoutes.has(url.pathname)) {
      if (request.method !== "POST" && request.method !== "OPTIONS") return new Response("Method not allowed", { status: 405 })
      return fetch(request)
    }
    if (!url.pathname.startsWith("/word-addin/") || !["GET", "HEAD"].includes(request.method))
      return new Response("Not found", { status: 404 })
    const headers = { "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff", "Referrer-Policy": "no-referrer" }
    if (url.pathname === "/word-addin/manifest.xml")
      return new Response(request.method === "HEAD" ? null : renderManifest(configuration.origin), { headers: { ...headers, "Content-Type": "application/xml; charset=utf-8", "Content-Disposition": 'attachment; filename="doc-haus-word.xml"' } })
    let relative: string
    try { relative = decodeURIComponent(url.pathname.slice("/word-addin/".length)) || "pane.html" }
    catch { return new Response("Invalid path", { status: 400 }) }
    const absolute = path.resolve(directory, relative)
    if (relative.includes("\0") || relative.includes("\\") || !absolute.startsWith(path.resolve(directory) + path.sep))
      return new Response("Not found", { status: 404 })
    const file = Bun.file(absolute)
    if (!await file.exists()) return new Response("The Word pane is not built. Restart doc.haus after installing dependencies.", { status: 503 })
    return new Response(request.method === "HEAD" ? null : file, { headers: { ...headers, "Content-Type": file.type } })
  }
}

export function startWordHttps(fetch: Fetch) {
  startupFailure = undefined
  try {
    const configuration = wordHttpsConfiguration()
    if (!configuration || configuration.external) return
    const server = Bun.serve({
      hostname: configuration.host,
      port: configuration.port,
      tls: { cert: Bun.file(configuration.certificate!), key: Bun.file(configuration.key!) },
      maxRequestBodySize: 128 * 1024 * 1024,
      fetch: wordHttpsHandler(fetch, configuration),
    })
    console.log(`doc.haus Word add-in listening on ${configuration.origin}`)
    return server
  } catch {
    // Native TLS errors can contain the private key/certificate paths. Keep the
    // log and API reason useful without retaining those diagnostics.
    startupFailure = "The Word HTTPS connection could not start. Check Word setup and that its configured port is free, then restart doc.haus."
    console.warn(`doc.haus: ${startupFailure}`)
  }
}
