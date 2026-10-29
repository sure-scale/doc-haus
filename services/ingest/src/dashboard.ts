import { Hono } from "hono"
import path from "node:path"
import { mkdirSync } from "node:fs"
import { workspaceRoot } from "./matter"

// Dashboard sessions live outside matter folders, like the library builders.
// Ingest owns the directory; the browser uses it to scope the native client.
export function dashboardRoutes() {
  const app = new Hono()
  app.get("/dashboard/directory", (c) => {
    const dir = path.join(workspaceRoot(), ".dashboard")
    mkdirSync(dir, { recursive: true })
    return c.json({ dir })
  })
  return app
}
