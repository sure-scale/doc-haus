import { Hono } from "hono"
import { patchEngineConfig, validateSettingsPatch } from "./engine-config"

type Selection = { sessionID: string; primary: string; fast: string }
type Options = {
  engine?: string
  password?: () => string | undefined
  transport?: (url: string, init: RequestInit) => Promise<Response>
  writeSettings?: (patch: Record<string, unknown>) => unknown
}

// The engine owns candidate credentials; ingest remains the only settings writer.
// Never forward engine bodies or thrown errors: either can contain account data.
export function providerSetupRoutes(options: Options = {}) {
  const app = new Hono()
  let saving = false
  app.post("/engine/provider-setup", async c => {
    c.header("Cache-Control", "no-store")
    const origin = c.req.header("origin")
    if (origin && !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) return c.json({ error: "Origin is not allowed." }, 403)
    if (c.req.header("content-type")?.split(";")[0].trim().toLowerCase() !== "application/json") return c.json({ error: "Save AI setup using JSON." }, 415)
    const input = await readSelection(c.req.raw)
    if (!input) return c.json({ error: "Select a tested connection and both models." }, 400)
    // Keep the credential/config transition ordered across concurrent Save clicks.
    if (saving) return c.json({ error: "Another AI setup is being saved. Try again shortly." }, 409)
    saving = true
    const engine = options.engine ?? process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
    const request = (action: "prepare" | "save" | "rollback" | "close") => {
      const password = options.password ? options.password() : process.env.OPENCODE_SERVER_PASSWORD
      return (options.transport ?? fetch)(`${engine}/api/dochaus/provider-candidate/${action === "close" ? `session?sessionID=${encodeURIComponent(input.sessionID)}` : action}`, {
        method: action === "close" ? "DELETE" : "POST",
        signal: AbortSignal.timeout(10_000),
        headers: { ...(password ? { Authorization: `Basic ${btoa(`opencode:${password}`)}` } : {}), ...(action === "close" ? {} : { "Content-Type": "application/json" }) },
        ...(action === "close" ? {} : { body: JSON.stringify(action === "rollback" ? { sessionID: input.sessionID } : input) }),
      })
    }
    const rollback = async () => { try { return (await request("rollback")).ok } catch { return false } }
    try {
      let patch: Record<string, unknown>
      try {
        const prepared = await request("prepare")
        if (!prepared.ok) return c.json({ error: "The selected connection and models must pass their tests before saving.", code: "setup_not_ready" }, 400)
        const result: unknown = await prepared.json()
        if (!result || typeof result !== "object" || !("patch" in result)) throw new Error()
        patch = validateSettingsPatch(result.patch)
        const primary = typeof patch.model === "string" ? patch.model : ""
        const fast = typeof patch.small_model === "string" ? patch.small_model : ""
        if (Object.keys(patch).some(key => !["model", "small_model", "provider"].includes(key)) ||
            primary.slice(primary.indexOf("/") + 1) !== input.primary || fast.slice(fast.indexOf("/") + 1) !== input.fast ||
            primary.split("/")[0] !== fast.split("/")[0]) throw new Error()
      } catch {
        return c.json({ error: "Could not validate AI setup. Your current setup is unchanged.", code: "setup_validation_failed" }, 400)
      }
      try {
        const saved = await request("save")
        if (!saved.ok) throw new Error()
      } catch {
        const restored = await rollback()
        return c.json({ error: restored ? "Could not save the connection. Your current setup is unchanged." : "Could not save the connection or confirm the previous connection was restored. Check AI setup before continuing.", code: restored ? "connection_save_failed" : "setup_partial_state", partialState: !restored }, 503)
      }
      try {
        // Synchronous validated read/merge/atomic rename preserves other settings.
        const write = options.writeSettings ?? patchEngineConfig
        write(patch)
      } catch {
        const restored = await rollback()
        return c.json({ error: restored ? "Could not save AI settings. The previous connection was restored." : "AI settings could not be saved and the previous connection could not be restored. Check AI setup before continuing.", code: restored ? "settings_save_failed" : "setup_partial_state", partialState: !restored }, 503)
      }
      // Settings and credentials are saved. A cleanup failure must not invite a
      // replay of this completed save; the candidate expires in the engine.
      let sessionClosePending = false
      try { sessionClosePending = !(await request("close")).ok } catch { sessionClosePending = true }
      return c.json({ ok: true, ...(sessionClosePending ? { sessionClosePending: true } : {}) })
    } finally { saving = false }
  })
  return app
}

async function readSelection(request: Request): Promise<Selection | undefined> {
  const reader = request.body?.getReader()
  if (!reader) return
  let expired = false
  const timer = setTimeout(() => { expired = true; void reader.cancel().catch(() => {}) }, 5000)
  const chunks: Uint8Array[] = []
  let length = 0
  try {
    while (true) {
      const chunk = await reader.read()
      if (expired) return
      if (chunk.done) break
      length += chunk.value.length
      if (length > 2048) { await reader.cancel(); return }
      chunks.push(chunk.value)
    }
    const value: unknown = JSON.parse(new TextDecoder().decode(Buffer.concat(chunks)))
    if (!value || typeof value !== "object" || Array.isArray(value) || Object.keys(value).sort().join(",") !== "fast,primary,sessionID") return
    const input = value as Selection
    if (typeof input.sessionID !== "string" || !/^[a-zA-Z0-9._-]{1,128}$/.test(input.sessionID) ||
        [input.primary, input.fast].some(model => typeof model !== "string" || !model || model.length > 256 || /[\s\0]/.test(model))) return
    return input
  } catch { return }
  finally { clearTimeout(timer); reader.releaseLock() }
}
