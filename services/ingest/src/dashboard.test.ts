import { expect, test } from "bun:test"
import { mkdtempSync, existsSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { dashboardRoutes } from "./dashboard"

test("dashboard sessions use the current workspace's own non-matter directory", async () => {
  const root = mkdtempSync(path.join(tmpdir(), "dochaus-dashboard-route-"))
  const previous = process.env.WORKSPACE_ROOT
  try {
    process.env.WORKSPACE_ROOT = root
    const app = dashboardRoutes()
    const res = await app.request("/dashboard/directory")
    expect(res.status).toBe(200)
    expect(await res.json()).toEqual({ dir: path.join(root, ".dashboard") })
    expect(existsSync(path.join(root, ".dashboard"))).toBe(true)
    expect(existsSync(path.join(root, ".dashboard", "matter.json"))).toBe(false)
  } finally {
    if (previous === undefined) delete process.env.WORKSPACE_ROOT
    else process.env.WORKSPACE_ROOT = previous
    rmSync(root, { recursive: true, force: true })
  }
})
