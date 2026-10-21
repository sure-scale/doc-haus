import { expect, test } from "bun:test"
import { connectionMessage, createProviderHealthMonitor, providerConnection, testModelConnection } from "../src/api/opencode/provider-health"

const model = { providerID: "fixture", id: "primary" }
const config = async () => ({ model: "fixture/primary", small_model: "fixture/primary" })

test("startup verifies the saved model once when primary and fast are identical", async () => {
  const calls: unknown[] = []
  const monitor = createProviderHealthMonitor(config, async (...args) => { calls.push(args); return { ok: true } })
  const first = monitor.check()
  expect(monitor.check()).toBe(first)
  await first
  expect(calls).toEqual([["fixture", "primary", false]])
  expect(monitor.getSnapshot()).toEqual([{ providerID: "fixture", modelID: "primary", status: "ready" }])
})

test("an expired saved key is rejected and checking after replacement clears it", async () => {
  let expired = true
  const forces: boolean[] = []
  const monitor = createProviderHealthMonitor(config, async (_provider, _model, force) => {
    forces.push(force)
    return expired ? { ok: false, kind: "auth", expired: true } : { ok: true }
  })
  await monitor.check()
  expect(monitor.getSnapshot()[0]?.status).toBe("auth")
  expect(connectionMessage(monitor.getSnapshot()[0]!)).toContain("Create a new key")
  expired = false
  await monitor.check(true)
  expect(monitor.getSnapshot()[0]?.status).toBe("ready")
  expect(forces).toEqual([false, true])
})

test("distinct primary and fast access failures both remain visible; auth takes precedence", async () => {
  const monitor = createProviderHealthMonitor(async () => ({ model: "fixture/primary", small_model: "fixture/fast" }),
    async (_provider, id) => id === "fast" ? { ok: false, kind: "auth" } : { ok: true })
  await monitor.check()
  expect(monitor.getSnapshot()).toHaveLength(2)
  expect(providerConnection(monitor.getSnapshot(), "fixture")?.status).toBe("auth")
})

test("typed runtime auth failure invalidates verified access for the actual failing model", async () => {
  const monitor = createProviderHealthMonitor(config, async () => ({ ok: true }))
  await monitor.check()
  monitor.rejected(model, { type: "provider.rate-limit", message: "temporary" })
  expect(monitor.getSnapshot()[0]?.status).toBe("ready")
  monitor.rejected(model, { type: "provider.auth", message: "This API key has expired." })
  expect(monitor.getSnapshot()[0]).toEqual({ providerID: "fixture", modelID: "primary", status: "auth", expired: true })
})

test("a late failure from before reconnection cannot overwrite the new successful check", async () => {
  const monitor = createProviderHealthMonitor(config, async () => ({ ok: true }))
  await monitor.check()
  const oldVersion = monitor.version()
  await monitor.check(true)
  monitor.rejected(model, { type: "provider.auth", message: "expired key" }, oldVersion)
  expect(monitor.getSnapshot()[0]?.status).toBe("ready")
})

test("a runtime failure wins over an older in-flight success", async () => {
  let resolve: ((value: { ok: true }) => void) | undefined
  const monitor = createProviderHealthMonitor(config, () => new Promise(done => { resolve = done }))
  const checking = monitor.check()
  await Promise.resolve()
  monitor.rejected(model, { type: "provider.auth", message: "invalid credentials" })
  resolve!({ ok: true })
  await checking
  expect(monitor.getSnapshot()[0]?.status).toBe("auth")
})

test("configuration outage is an unverifiable connection, never rejected credentials", async () => {
  const monitor = createProviderHealthMonitor(async () => { throw new Error("offline") }, async () => ({ ok: true }))
  await monitor.check()
  expect(monitor.getSnapshot()[0]?.status).toBe("unavailable")
  expect(connectionMessage(monitor.getSnapshot()[0]!)).toContain("network")
})

test("managed probe distinguishes typed provider rejection from local engine authentication", async () => {
  const saved = globalThis.fetch
  try {
    globalThis.fetch = (async () => Response.json({ error: "Authentication required" }, { status: 401 })) as unknown as typeof fetch
    expect(await testModelConnection("fixture", "primary", false)).toEqual({ ok: false, kind: "unavailable" })
    globalThis.fetch = (async () => Response.json({ ok: false, kind: "auth", expired: true, error: "private diagnostic must not be copied" }, { status: 503 })) as unknown as typeof fetch
    expect(await testModelConnection("fixture", "primary", true)).toEqual({ ok: false, kind: "auth", expired: true })
    globalThis.fetch = (async () => Response.json({ ok: false, kind: "unavailable" }, { status: 503 })) as unknown as typeof fetch
    expect(await testModelConnection("fixture", "primary", false)).toEqual({ ok: false, kind: "unavailable" })
  } finally { globalThis.fetch = saved }
})


test("connection checks use the underlying catalog model once across generation variants", async () => {
  const calls: string[] = []
  const monitor = createProviderHealthMonitor(async () => ({ model: "fixture/primary#high", small_model: "fixture/primary#low" }),
    async (provider, id) => { calls.push(`${provider}/${id}`); return { ok: true } })
  await monitor.check()
  expect(calls).toEqual(["fixture/primary"])
})


test("rechecking also verifies a failed session model after the saved default changed", async () => {
  const calls: string[] = []
  let oldModelWorks = false
  const monitor = createProviderHealthMonitor(config, async (_provider, id) => {
    calls.push(id)
    return id === "old-session-model" && !oldModelWorks ? { ok: false, kind: "auth" } : { ok: true }
  })
  await monitor.check()
  monitor.rejected({ providerID: "fixture", id: "old-session-model" }, { type: "provider.auth", message: "Invalid key" })
  await monitor.check(true)
  expect(calls).toEqual(["primary", "old-session-model", "primary"])
  expect(providerConnection(monitor.getSnapshot(), "fixture")?.status).toBe("auth")
  oldModelWorks = true
  await monitor.check(true)
  expect(monitor.getSnapshot().every(item => item.status === "ready")).toBe(true)
})
