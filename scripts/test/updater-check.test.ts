import { afterEach, expect, test } from "bun:test"
import { mkdtemp, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { RepoUpdater } from "../updater"
import { updateRoutes } from "../../services/ingest/src/updater"

const roots: string[] = []
afterEach(async () => { for (const root of roots.splice(0)) await rm(root, { recursive: true, force: true }) })

async function fixture() {
  const root = await mkdtemp(path.join(os.tmpdir(), "dochaus-update-check-"))
  roots.push(root)
  const current = { version: 1, name: "Current", notes: "Current release" }
  await Bun.write(path.join(root, "release.json"), JSON.stringify(current))
  class MockUpdater extends RepoUpdater {
    calls: string[][] = []
    configured = true
    connected = true
    override async git(...args: string[]) {
      this.calls.push(args)
      if (args[0] === "symbolic-ref") return "local-branch"
      if (args[0] === "config") {
        if (!this.configured) throw new Error("Git could not complete the update.")
        return args[2].endsWith(".remote") ? "private" : "refs/heads/dochaus"
      }
      if (!this.connected) throw new Error("Git could not complete the update. Check your repository access and connection.")
      if (args[0] === "fetch") return ""
      if (["rev-parse", "log"].includes(args[0])) return "a".repeat(40)
      if (args[0] === "show") return JSON.stringify(current)
      throw new Error(`Unexpected command: ${args[0]}`)
    }
  }
  const updater = new MockUpdater(root, path.join(root, "runtime"))
  const app = updateRoutes(updater, async () => false, async () => { throw new Error("Restart must not run") }, () => { throw new Error("Signal must not run") })
  return { updater, app }
}

test("HTTP 200 check error identifies missing tracking configuration without fetching; explicit retry refreshes cache", async () => {
  const { updater, app } = await fixture()
  expect(await (await app.request("/updates")).json()).toMatchObject({ available: false })
  expect(updater.calls.some(call => call[0] === "fetch")).toBe(true)
  updater.calls = []
  updater.configured = false
  const failed = await app.request("/updates?refresh=1")
  expect(failed.status).toBe(200)
  expect(await failed.json()).toMatchObject({ available: false, error: expect.stringContaining("no update branch configured") })
  expect(updater.calls.some(call => call[0] === "fetch")).toBe(false)
  updater.configured = true
  expect(await (await app.request("/updates")).json()).toHaveProperty("error")
  expect(await (await app.request("/updates?refresh=1")).json()).not.toHaveProperty("error")
})

test("disconnected update checks recover on explicit retry and unavailable checks reject apply", async () => {
  const { updater, app } = await fixture()
  updater.connected = false
  expect(await (await app.request("/updates")).json()).toMatchObject({ available: false, error: expect.stringContaining("connection") })
  const apply = await app.request("/updates/apply", { method: "POST", headers: { origin: "http://localhost:5173", "x-dochaus-update": "1", "content-type": "application/json" }, body: JSON.stringify({ target: "a".repeat(40) }) })
  expect(apply.status).toBe(409)
  expect(await apply.json()).toHaveProperty("error", "The available update changed. Check for updates again.")
  updater.connected = true
  expect(await (await app.request("/updates?refresh=1")).json()).not.toHaveProperty("error")
  expect(updater.calls.some(call => ["merge", "reset", "push"].includes(call[0]))).toBe(false)
})
