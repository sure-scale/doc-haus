import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs"
import path from "node:path"
import { Hono } from "hono"
import { readEngineConfig } from "./engine-config"

export const SETUP_STEPS = ["models", "research", "details"] as const
export type SetupStep = typeof SETUP_STEPS[number]
export type SetupProgress = { version: 1; completed: SetupStep[]; skipped: SetupStep[]; next: SetupStep | "done"; ready: boolean }
type Progress = Omit<SetupProgress, "ready" | "next">
type CatalogModel = { providerID: string; id: string; enabled: boolean; variants?: { id: string }[] }

export function validModelPair(config: { model?: unknown; small_model?: unknown }, models: CatalogModel[]) {
  if (typeof config.model !== "string" || typeof config.small_model !== "string") return false
  function available(reference: string) {
    // Native model references retain an optional, explicitly selected #variant.
    const [base, variant, extra] = reference.split("#")
    if (extra !== undefined || variant === "") return undefined
    return models.find(model => model.enabled && `${model.providerID}/${model.id}` === base &&
      (variant === undefined || model.variants?.some(candidate => candidate.id === variant)))
  }
  const primary = available(config.model)
  const fast = available(config.small_model)
  return Boolean(primary && fast && primary.providerID === fast.providerID)
}

export function setupRoutes(options: { directory: string; config?: () => { model?: unknown; small_model?: unknown }; models?: () => Promise<CatalogModel[]> }) {
  const file = path.join(options.directory, "setup.json")
  const config = options.config ?? readEngineConfig
  const models = options.models ?? (async () => {
    const engine = process.env.OPENCODE_URL ?? `http://127.0.0.1:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}`
    const password = process.env.OPENCODE_SERVER_PASSWORD
    const unavailable = "Could not check the engine model catalog. Retry when the engine is ready."
    // Saving preferences reloads research consent protection, and the engine
    // refuses catalog reads until it is active again. Wait out that short reload
    // inside one 8 s budget, well within the web client's 12 s timeout.
    const deadline = Date.now() + 8000
    while (true) {
      const res = await fetch(`${engine}/api/model`, {
        signal: AbortSignal.timeout(Math.max(deadline - Date.now(), 1)),
        ...(password ? { headers: { Authorization: `Basic ${btoa(`opencode:${password}`)}` } } : {}),
      }).catch(() => { throw new Error(unavailable) })
      if (res.ok) {
        const result = await res.json() as { data: CatalogModel[] }
        if (!Array.isArray(result.data)) throw new Error("The engine returned an invalid model catalog.")
        return result.data
      }
      const readiness = res.status === 503 ? await res.json().catch(() => undefined) as { legalState?: string; consentState?: string } | undefined : undefined
      const consent = readiness?.legalState === "active" ? readiness.consentState : undefined
      if (consent === "pending" && Date.now() + 200 < deadline) { await Bun.sleep(200); continue }
      if (consent === "pending") throw new Error("Research protection is still loading. Retry in a moment.")
      if (consent) throw new Error("Research protection could not load, so the engine refuses setup checks. Restart doc.haus, then retry.")
      throw new Error(unavailable)
    }
  })
  function write(progress: Progress) {
    mkdirSync(options.directory, { recursive: true })
    const temporary = `${file}.${crypto.randomUUID()}.tmp`
    try { writeFileSync(temporary, JSON.stringify(progress, null, 2) + "\n", { mode: 0o600, flag: "wx" }); renameSync(temporary, file) }
    finally { rmSync(temporary, { force: true }) }
  }
  async function read(): Promise<SetupProgress> {
    const ready = validModelPair(config(), await models())
    const progress: Progress = existsSync(file) ? JSON.parse(readFileSync(file, "utf8")) : {
      version: 1, completed: ready ? ["models"] : [], skipped: ready ? ["research", "details"] : [],
    }
    if (!existsSync(file)) write(progress)
    if (progress.version !== 1 || !Array.isArray(progress.completed) || !Array.isArray(progress.skipped))
      throw new Error("Setup progress could not be read. Open Settings to check your configuration.")
    return summary(progress, ready)
  }
  function summary(progress: Progress, ready: boolean): SetupProgress {
    const completed = progress.completed.filter(step => SETUP_STEPS.includes(step) && (step !== "models" || ready))
    const skipped = progress.skipped.filter(step => step !== "models" && SETUP_STEPS.includes(step) && !completed.includes(step))
    return { version: 1, completed, skipped, next: SETUP_STEPS.find(step => !completed.includes(step) && !skipped.includes(step)) ?? "done", ready }
  }
  const routes = new Hono()
  routes.get("/setup", async c => {
    try { return c.json(await read()) }
    catch { return c.json({ error: "Could not check setup. Make sure the engine is running, then retry." }, 503) }
  })
  routes.patch("/setup", async c => {
    try {
      const input: unknown = await c.req.json()
      if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(key => !["step", "status"].includes(key))) throw new Error("Choose a setup step and status.")
      const value = input as { step: SetupStep; status: "completed" | "skipped" }
      if (!SETUP_STEPS.includes(value.step) || !["completed", "skipped"].includes(value.status)) throw new Error("Choose a setup step and status.")
      if (value.step === "models" && value.status === "skipped") throw new Error("Save both models before continuing.")
      const current = await read()
      if (value.step === "models" && !current.ready) throw new Error("Save an available primary and fast model from the same provider before continuing.")
      if (value.step !== "models" && !current.completed.includes("models")) throw new Error("Complete model setup first.")
      const progress: Progress = { version: 1, completed: current.completed.filter(step => step !== value.step), skipped: current.skipped.filter(step => step !== value.step) }
      progress[value.status].push(value.step)
      write(progress)
      // A second catalog read could land in a new consent reload and report a saved step as failed.
      return c.json(summary(progress, current.ready))
    } catch (error) { return c.json({ error: error instanceof Error ? error.message : "Could not save setup progress." }, 400) }
  })
  return routes
}
