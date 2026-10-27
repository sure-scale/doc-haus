import path from "node:path"
import { createHash } from "node:crypto"
import { existsSync, readFileSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { renderManifest } from "../../../apps/word-addin/src/manifest"
import { WORD_ADDIN_VERSION } from "../../../dochaus/lib/word-addin-contract"
import { WORD_BOOTSTRAP_ORIGIN, WORD_HEALTH_PATH } from "../../../apps/word-addin/src/bootstrap-contract"
import { wordHttpsConfiguration, type WordHttpsConfiguration } from "./word-configuration"
import { wordOpenPath, wordOpenResultPath } from "./word-open-tickets"

type Fetch = (request: Request) => Response | Promise<Response>
const assets = fileURLToPath(new URL("../../../apps/word-addin/dist/", import.meta.url))
const tokenRoutes = new Set(["/word/lease", "/word/saves", "/word/heartbeat", "/word/done",
  "/word/desk/sources", "/word/desk/source", "/word/desk/versions", "/word/desk/compare"])
let startupFailure: string | undefined
let activeServer: ReturnType<typeof Bun.serve> | undefined
let activeIdentity: string | undefined
let activeAssets = assets
const paneFiles = ["pane.html", "pane.js", "pane.css", "desk.html", "desk.js", "desk.css", "assets/icon.png"]

export function wordPaneBuilt(directory = assets) {
  return paneFiles.every(file => existsSync(path.join(directory, file)))
}

function configurationIdentity(configuration: WordHttpsConfiguration) {
  return JSON.stringify([configuration.origin, configuration.host, configuration.port,
    configuration.external ? "external" : createHash("sha256").update(readFileSync(configuration.certificate!)).update(readFileSync(configuration.key!)).digest("hex")])
}

export function wordHttpsRunning(configuration: WordHttpsConfiguration) {
  try { return Boolean(activeServer && activeIdentity === configurationIdentity(configuration)) }
  catch { return false }
}

export function wordHttpsAvailable(configuration: WordHttpsConfiguration) {
  return configuration.external || (wordHttpsRunning(configuration) && wordPaneBuilt(activeAssets))
}

export function stopWordHttps() {
  activeServer?.stop(true)
  activeServer = undefined
  activeIdentity = undefined
}

/** A recoverable bridge failure must never make the ordinary product unavailable. */
export function wordHttpsFailure() { return startupFailure }

/** This HTTPS listener exposes static pane files and scoped document capabilities.
 * It never forwards the ordinary matter, document-read or administration API. */
export function wordHttpsHandler(fetch: Fetch, configuration: WordHttpsConfiguration, directory = assets): Fetch {
  return async request => {
    const url = new URL(request.url)
    if (wordOpenPath(url.pathname) || wordOpenResultPath(url.pathname)) {
      if (request.headers.has("Origin")) return new Response("Origin refused", { status: 403 })
      if (request.method !== (wordOpenResultPath(url.pathname) ? "POST" : "GET")) return new Response("Method not allowed", { status: 405 })
      return fetch(request)
    }
    if (url.pathname === WORD_HEALTH_PATH) {
      // Only the public holding page may read this across origins. This does
      // not widen the bearer-only Word commands or expose the ordinary API.
      const origin = request.headers.get("Origin")
      if (origin && origin !== WORD_BOOTSTRAP_ORIGIN && origin !== configuration.origin)
        return new Response("Origin refused", { status: 403 })
      const headers = new Headers({ "Cache-Control": "no-store", "Referrer-Policy": "no-referrer", "Vary": "Origin" })
      if (origin) headers.set("Access-Control-Allow-Origin", origin)
      if (!["GET", "HEAD", "OPTIONS"].includes(request.method))
        return new Response("Method not allowed", { status: 405, headers })
      if (request.method === "OPTIONS") {
        headers.set("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS")
        return new Response(null, { status: 204, headers })
      }
      const ready = wordPaneBuilt(directory)
      headers.set("Content-Type", "application/json; charset=utf-8")
      return new Response(request.method === "HEAD" ? null : JSON.stringify({ ready, version: WORD_ADDIN_VERSION }),
        { status: ready ? 200 : 503, headers })
    }
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

export function startWordHttps(fetch: Fetch, environment: Record<string, string | undefined> = process.env, directory = assets) {
  startupFailure = undefined
  try {
    const configuration = wordHttpsConfiguration(environment)
    if (!configuration || configuration.external) { stopWordHttps(); return }
    const identity = configurationIdentity(configuration)
    if (activeServer && activeIdentity === identity && activeAssets === directory) return activeServer
    stopWordHttps()
    const server = Bun.serve({
      hostname: configuration.host,
      port: configuration.port,
      tls: { cert: Bun.file(configuration.certificate!), key: Bun.file(configuration.key!) },
      maxRequestBodySize: 128 * 1024 * 1024,
      fetch: wordHttpsHandler(fetch, configuration, directory),
    })
    activeServer = server
    activeIdentity = identity
    activeAssets = directory
    console.log(`doc.haus Word add-in listening on ${configuration.origin}`)
    return server
  } catch {
    // Native TLS errors can contain the private key/certificate paths. Keep the
    // log and API reason useful without retaining those diagnostics.
    startupFailure = "The Word HTTPS connection could not start. Check Word setup and that its configured port is free, then choose Repair in Settings → Add-ins."
    console.warn(`doc.haus: ${startupFailure}`)
  }
}
