import { afterEach, expect, test } from "bun:test"
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import type { Plugin } from "@opencode/plugin"
import researchMcp from "../plugin/research-mcp"
import { researchConsentStatus } from "../lib/research-consent-status"

const previous = process.env.WORKSPACE_ROOT
let cleanup: Plugin.Cleanup | void
let root: string
afterEach(async () => {
  await cleanup?.()
  cleanup = undefined
  if (previous === undefined) delete process.env.WORKSPACE_ROOT
  else process.env.WORKSPACE_ROOT = previous
  if (root) rmSync(root, { recursive: true, force: true })
})

async function fixture(preferences: Record<string, unknown> = {}) {
  root = mkdtempSync(path.join(tmpdir(), "dochaus-research-mcp-"))
  process.env.WORKSPACE_ROOT = root
  const servers = new Map<string, { type: "local" | "remote"; [key: string]: unknown }>()
  let transform: Parameters<Plugin.Context["mcp"]["transform"]>[0]
  const editor = {
    list: () => Array.from(servers), get: (name: string) => servers.get(name), update: () => {},
    set: (name: string, config: { type: "local" | "remote"; [key: string]: unknown }) => servers.set(name, config),
    remove: (name: string) => servers.delete(name),
  }
  const mcp = {
    transform: async (callback: typeof transform) => {
      transform = callback
      return { dispose: async () => {} }
    },
    reload: async () => { transform(editor) },
  }
  cleanup = await researchMcp.setup({ mcp, location: { directory: root } } as unknown as Plugin.Context)
  const save = (input: Record<string, unknown>) => {
    writeFileSync(path.join(root, ".preferences", "preferences.json"), JSON.stringify(input))
  }
  save(preferences)
  await mcp.reload()
  return { servers, save, reload: mcp.reload }
}

test("shipped MCPs and old preferences never opt into CourtListener", async () => {
  const config = JSON.parse(readFileSync(path.resolve(import.meta.dir, "../opencode.json"), "utf8"))
  expect(config.mcp.servers.courtlistener).toBeUndefined()
  const { servers } = await fixture({ caseLawApiKey: "synthetic-token" })
  expect(servers.has("courtlistener")).toBe(false)
})

test("opt-in requires a token and uses Token authentication without OAuth", async () => {
  const { servers, save, reload } = await fixture({ courtlistenerMcpEnabled: true })
  expect(servers.has("courtlistener")).toBe(false)
  save({ courtlistenerMcpEnabled: true, caseLawApiKey: " synthetic-token " })
  await reload()
  expect(servers.get("courtlistener")).toMatchObject({
    url: "https://mcp.courtlistener.com/", oauth: false,
    headers: { Authorization: "Token synthetic-token" },
  })
})

test("saved settings update warm matters, rotate tokens and disconnect on opt-out", async () => {
  const { servers, save } = await fixture()
  save({ courtlistenerMcpEnabled: true, caseLawApiKey: "first-synthetic-token" })
  await until(() => servers.has("courtlistener"))
  save({ courtlistenerMcpEnabled: true, caseLawApiKey: "second-synthetic-token" })
  await until(() => JSON.stringify(servers.get("courtlistener")).includes("second-synthetic-token"))
  save({ courtlistenerMcpEnabled: false, caseLawApiKey: "second-synthetic-token" })
  await until(() => !servers.has("courtlistener"))
}, 30_000)

test("account and native refreshes retain local Python and remove every unapproved remote", async () => {
  const { servers, save, reload } = await fixture({ courtlistenerMcpEnabled: true, caseLawApiKey: "synthetic-token" })
  servers.set("python", { type: "local", command: ["neutral-python-fixture"] })
  servers.set("account-remote", { type: "remote", url: "https://unapproved.test/" })
  servers.set("courtlistener-alias", { type: "remote", url: "https://mcp.courtlistener.com/" })
  servers.set("courtlistener", { type: "remote", url: "https://unapproved.test/" })
  await reload()
  expect([...servers.keys()].sort()).toEqual(["courtlistener", "python"])
  expect(servers.get("courtlistener")?.url).toBe("https://mcp.courtlistener.com/")
  save({ courtlistenerMcpEnabled: false, caseLawApiKey: "synthetic-token" })
  servers.set("account-change", { type: "remote", url: "https://new-account.test/" })
  await reload()
  expect([...servers.keys()]).toEqual(["python"])
  expect(researchConsentStatus(root)).toBe("active")
})

test("saved consent does not override a disabled CourtListener configuration", async () => {
  const { servers, reload } = await fixture({ courtlistenerMcpEnabled: true, caseLawApiKey: "synthetic-token" })
  servers.set("courtlistener", { type: "remote", url: "https://mcp.courtlistener.com/", disabled: true })
  await reload()
  expect(servers.has("courtlistener")).toBe(false)
})

test("malformed and unreadable preferences revoke consent without throwing", async () => {
  const { servers, reload } = await fixture({ courtlistenerMcpEnabled: true, caseLawApiKey: "synthetic-token" })
  const file = path.join(root, ".preferences", "preferences.json")
  writeFileSync(file, "{malformed")
  await reload()
  expect(servers.has("courtlistener")).toBe(false)
  rmSync(file)
  mkdirSync(file)
  servers.set("account-remote", { type: "remote", url: "https://unapproved.test/" })
  await reload()
  expect(servers.size).toBe(0)
})

test("native transform registration failure cannot claim activation or independently filter native state", async () => {
  root = mkdtempSync(path.join(tmpdir(), "dochaus-research-registration-failure-"))
  process.env.WORKSPACE_ROOT = root
  const native = new Map([["unapproved-remote", { type: "remote", url: "https://unapproved.test/" }]])
  // Native registration failed before our callback runs. The additive plugin
  // cannot remove those entries; the engine must reject readiness and requests.
  const mcp = { transform: async () => { throw new Error("synthetic registration failure") } }
  cleanup = await researchMcp.setup({ mcp, location: { directory: root } } as unknown as Plugin.Context)
  expect(researchConsentStatus(root)).toBe("failed")
  expect(native.has("unapproved-remote")).toBe(true)
  expect(researchConsentStatus(root + "-missing")).toBe("pending")
})

test("a failing native editor does not throw from the transform and reports failed activation", async () => {
  root = mkdtempSync(path.join(tmpdir(), "dochaus-research-editor-failure-"))
  process.env.WORKSPACE_ROOT = root
  let transform: Parameters<Plugin.Context["mcp"]["transform"]>[0]
  const mcp = {
    transform: async (callback: typeof transform) => { transform = callback; return { dispose: async () => {} } },
    reload: async () => { transform({ list: () => { throw new Error("synthetic native editor failure") } } as never) },
  }
  cleanup = await researchMcp.setup({ mcp, location: { directory: root } } as unknown as Plugin.Context)
  await expect(mcp.reload()).resolves.toBeUndefined()
  expect(researchConsentStatus(root)).toBe("failed")
})

// The plugin reloads on a file watch; on macOS a watch event can take a few
// seconds to arrive.
async function until(ready: () => boolean) {
  const deadline = Date.now() + 10_000
  while (!ready()) {
    if (Date.now() > deadline) throw new Error("Saved connector preferences did not apply")
    await Bun.sleep(20)
  }
}
