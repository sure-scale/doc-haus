import { existsSync, readFileSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { setupWordAddin, validateWordCertificate, wordAddinRegistered, wordCertificateTrusted, wordInstalled, type WordSetupOptions } from "../../../scripts/setup-word-addin"
import { wordConfigurationFile, wordHttpsConfiguration } from "./word-configuration"
import { startWordHttps, wordHttpsFailure, wordHttpsRunning, wordPaneBuilt } from "./word-https"
import { wordBootstrapUrl } from "../../../apps/word-addin/src/bootstrap-contract"

export type WordSetupStatus = {
  platform: "macos" | "windows" | "unsupported"
  wordInstalled: boolean
  canSetup: boolean
  configured: boolean
  registered: boolean
  trusted: boolean
  running: boolean
  available: boolean
  status: "ready" | "not_configured" | "needs_repair" | "unsupported" | "word_missing" | "external"
  origin: string | null
  manifestUrl: string | null
  reason?: string
  activatedWithoutRestart: boolean
  restartWord: boolean
}

type Options = WordSetupOptions & {
  fetch: (request: Request) => Response | Promise<Response>
  assets?: string
}
const loopback = (hostname: string) => ["localhost", "127.0.0.1", "[::1]", "::1"].includes(hostname)

/** UI setup is deliberately limited to the local installation. Reverse-proxy
 * HTTPS, certificate paths and arbitrary process arguments are never accepted. */
export function localWordSetupRequest(request: Request) {
  const url = new URL(request.url)
  if (!loopback(url.hostname)) return false
  const origin = request.headers.get("origin")
  if (origin) {
    try {
      const browser = new URL(origin)
      if (!loopback(browser.hostname) || !["http:", "https:"].includes(browser.protocol) || browser.origin !== origin) return false
    } catch { return false }
  }
  const site = request.headers.get("sec-fetch-site")
  return !site || ["same-origin", "same-site", "none"].includes(site) || Boolean(origin)
}

export function wordSetupRoutes(options: Options) {
  const environment = options.environment ?? process.env
  const native = { ...options, environment, nativeConfirmation: true }
  // Request Host and Origin are caller-controlled. A network-facing ingest
  // bind cannot prove a native setup request came from this computer.
  const localIngest = loopback(environment.INGEST_HOST ?? "127.0.0.1")
  let pending: Promise<WordSetupStatus> | undefined
  let restartWord = false

  function externalConfigured() {
    if (environment.DOCHAUS_WORD_EXTERNAL_HTTPS === "1") return true
    try {
      const saved = JSON.parse(readFileSync(wordConfigurationFile(environment), "utf8"))
      return !environment.DOCHAUS_WORD_ORIGIN && saved?.external === true
    } catch { return false }
  }

  async function status(): Promise<WordSetupStatus> {
    const platform = (options.platform ?? process.platform) === "darwin" ? "macos" : (options.platform ?? process.platform) === "win32" ? "windows" : "unsupported"
    const installed = await wordInstalled(native).catch(() => false)
    let configuration: ReturnType<typeof wordHttpsConfiguration>
    let invalid = false
    try { configuration = wordHttpsConfiguration(environment) } catch { invalid = true }
    const external = externalConfigured()
    const origin = configuration?.origin ?? null
    let configuredOrigin: string | undefined = environment.DOCHAUS_WORD_ORIGIN
    if (!configuredOrigin) {
      try {
        const saved = JSON.parse(readFileSync(wordConfigurationFile(environment), "utf8"))
        if (typeof saved?.origin === "string") configuredOrigin = saved.origin
      } catch { /* Malformed local settings can be repaired. */ }
    }
    let localOrigin = true
    try { if (configuredOrigin) localOrigin = loopback(new URL(configuredOrigin).hostname) } catch { /* Explicit setup validates or repairs the origin. */ }
    const local = localOrigin && (!environment.DOCHAUS_WORD_HOST || loopback(environment.DOCHAUS_WORD_HOST)) &&
      (!configuration || (loopback(new URL(configuration.origin).hostname) && loopback(configuration.host)))
    const canSetup = localIngest && platform !== "unsupported" && installed && !external && local
    let registered = false, trusted = false, running = false
    if (configuration && !external) {
      try {
        validateWordCertificate(configuration.certificate!, configuration.key!, new URL(configuration.origin).hostname)
        trusted = await wordCertificateTrusted(configuration.certificate!, new URL(configuration.origin).hostname, native)
      } catch { /* Invalid TLS material stays unavailable and can be repaired. */ }
      const manifest = path.join(path.dirname(wordConfigurationFile(environment)), "word-addin", "manifest.xml")
      registered = await wordAddinRegistered(configuration.origin, manifest, native)
      running = wordHttpsRunning(configuration) && wordPaneBuilt(options.assets)
    } else if (configuration && external) {
      // Configuration alone does not prove that the firm's HTTPS proxy serves
      // this installation. Keep normal TLS verification and refuse redirects.
      const available = await Promise.all(["manifest.xml", "pane.html"].map(async file => {
        try {
          const response = await fetch(`${configuration.origin}/word-addin/${file}`, { redirect: "error", signal: AbortSignal.timeout(4000) })
          if (!response.ok) return false
          const content = await response.text()
          return file === "manifest.xml" ? content.includes(wordBootstrapUrl(configuration.origin).replaceAll("&", "&amp;")) : content.includes("office.js")
        } catch { return false }
      }))
      running = trusted = available.every(Boolean)
    }
    const available = external ? running : Boolean(configuration && registered && trusted && running)
    const base = { platform, wordInstalled: installed, canSetup, configured: Boolean(configuration), registered, trusted, running, available,
      origin, manifestUrl: origin ? `${origin}/word-addin/manifest.xml` : null, activatedWithoutRestart: !external && running, restartWord } as const
    if (external) return { ...base, status: "external", reason: available ? "This Word connection is managed by your administrator." : "Your administrator needs to check the Word HTTPS connection." }
    if (available) return { ...base, status: "ready" }
    if (platform === "unsupported") return { ...base, status: "unsupported", reason: "Local Word setup needs Microsoft Word on macOS or Windows." }
    if (!installed) return { ...base, status: "word_missing", reason: "Install Microsoft Word on this computer, then set it up here." }
    if (!localIngest) return { ...base, status: "needs_repair", reason: "Set up Word in a local doc.haus installation on this computer." }
    if (!local) return { ...base, status: "needs_repair", reason: "Your administrator needs to check this Word connection." }
    if (invalid || existsSync(wordConfigurationFile(environment))) return { ...base, status: "needs_repair",
      reason: wordHttpsFailure() ?? (!wordPaneBuilt(options.assets) ? "The Word files are missing. Update doc.haus, then set up Word again." : "The Word connection needs repair. Set it up again here.") }
    return { ...base, status: "not_configured", reason: "Set up Word to open and save matter documents from Microsoft Word." }
  }

  async function setup() {
    const current = await status()
    if (!current.canSetup) return current
    const result = await setupWordAddin([], native)
    restartWord = result?.registrationChanged ?? false
    startWordHttps(options.fetch, environment, options.assets)
    return status()
  }

  const app = new Hono()
  app.get("/word/setup", async c => {
    c.header("Cache-Control", "no-store")
    const result = await status()
    return c.json(localWordSetupRequest(c.req.raw) ? result : { ...result, canSetup: false })
  })
  app.post("/word/setup", async c => {
    c.header("Cache-Control", "no-store")
    if (!localIngest || !localWordSetupRequest(c.req.raw)) return c.json({ error: "Set up Word from doc.haus on this computer." }, 403)
    if (c.req.header("content-type")?.split(";")[0].trim().toLowerCase() !== "application/json") return c.json({ error: "Choose Set up Word in doc.haus." }, 415)
    // There are no setup options to accept from the browser. A small explicit
    // JSON body also prevents ordinary cross-site form submissions.
    const text = await setupBody(c.req.raw)
    let input: unknown
    try { input = text === undefined ? undefined : JSON.parse(text) } catch { /* Invalid input is refused below. */ }
    if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).length) return c.json({ error: "Choose Set up Word in doc.haus." }, 400)
    try {
      if (!pending) pending = setup().finally(() => { pending = undefined })
      const result = await pending
      return c.json(result, result.available ? 200 : 409)
    } catch {
      // OS errors may contain certificate/key paths. Only this fixed message
      // crosses the API boundary; read status again to inspect actual progress.
      return c.json({ error: "Word setup did not finish. Allow the system confirmation, then try again.", ...(await status()) }, 503)
    }
  })
  return app
}

async function setupBody(request: Request) {
  const reader = request.body?.getReader()
  if (!reader) return
  const timer = setTimeout(() => { void reader.cancel().catch(() => {}) }, 5000)
  const chunks: Uint8Array[] = []
  let length = 0
  try {
    while (true) {
      const chunk = await reader.read()
      if (chunk.done) return new TextDecoder().decode(Buffer.concat(chunks))
      length += chunk.value.length
      if (length > 128) { await reader.cancel(); return }
      chunks.push(chunk.value)
    }
  } catch { return }
  finally { clearTimeout(timer); reader.releaseLock() }
}
