import { Hono } from "hono"
import path from "node:path"
import { mkdirSync } from "node:fs"
import { workspaceRoot } from "./matter"
import { DASHBOARD_STATE_MAX_BYTES, DashboardStateError, parseDashboardState, readDashboardState, writeDashboardState } from "./dashboard-state"

// Dashboard sessions live outside matter folders, like the library builders.
// Ingest owns the directory; the browser uses it to scope the native client.
export function dashboardRoutes() {
  const app = new Hono()
  app.get("/dashboard/directory", (c) => {
    const dir = path.join(workspaceRoot(), ".dashboard")
    mkdirSync(dir, { recursive: true })
    return c.json({ dir })
  })
  app.get("/dashboard/tabs", (c) => {
    c.header("Cache-Control", "no-store")
    try { return c.json(readDashboardState()) }
    catch { return c.json({ error: "Could not read saved dashboard tabs. Retry." }, 503) }
  })
  app.put("/dashboard/tabs", async (c) => {
    c.header("Cache-Control", "no-store")
    let input: unknown
    try {
      const body = await c.req.text()
      if (new TextEncoder().encode(body).byteLength > DASHBOARD_STATE_MAX_BYTES) return c.json({ error: "Dashboard settings are too large." }, 413)
      input = JSON.parse(body)
    } catch { return c.json({ error: "Choose valid dashboard settings." }, 400) }
    try {
      if (typeof input !== "object" || input === null || Array.isArray(input) || Object.keys(input).some(key => !["state", "revision"].includes(key)))
        throw new DashboardStateError("Choose dashboard settings and their saved revision.")
      const value = input as { state: unknown; revision: number }
      if (!Number.isSafeInteger(value.revision) || value.revision < 0 || value.revision >= Number.MAX_SAFE_INTEGER) throw new DashboardStateError("Choose a valid saved dashboard revision.")
      const state = parseDashboardState(value.state)
      const result = writeDashboardState(state, value.revision)
      return result.saved ? c.json(result.snapshot) : c.json({ error: "Dashboard tabs changed elsewhere. Reload the saved tabs before trying again.", ...result.snapshot }, 409)
    } catch (error) {
      if (error instanceof DashboardStateError) return c.json({ error: error.message }, 400)
      return c.json({ error: "Could not save dashboard tabs. Your changes have not been confirmed. Retry." }, 503)
    }
  })
  return app
}
