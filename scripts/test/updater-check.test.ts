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
    connected = true
    override async git(...args: string[]) {
      this.calls.push(args)
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

test("checks fetch the canonical public release branch without requiring checkout tracking; retry refreshes errors", async () => {
  const { updater, app } = await fixture()
  expect(await (await app.request("/updates")).json()).toMatchObject({ available: false })
  expect(updater.calls.find(call => call[0] === "fetch")).toEqual(["fetch", "--no-tags", "--", "origin", "refs/heads/dochaus:refs/remotes/origin/dochaus"])
  expect(updater.calls.find(call => call[0] === "rev-parse")).toEqual(["rev-parse", "refs/remotes/origin/dochaus"])
  expect(updater.calls.some(call => ["symbolic-ref", "config"].includes(call[0]))).toBe(false)
  updater.calls = []
  updater.connected = false
  const failed = await app.request("/updates?refresh=1")
  expect(failed.status).toBe(200)
  expect(await failed.json()).toMatchObject({ available: false, error: expect.stringContaining("connection") })
  expect(updater.calls.find(call => call[0] === "fetch")).toEqual(["fetch", "--no-tags", "--", "origin", "refs/heads/dochaus:refs/remotes/origin/dochaus"])
  updater.connected = true
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
