import { expect, test } from "bun:test"
import { WordSync, type Transport } from "./sync"
import type { WordConnectionState } from "../../../dochaus/lib/word-addin-contract"

function fixture(owner: "available" | "owned" | "other" = "available") {
  const events: string[] = []
  let version = "base-one", saved = false
  const connection = (): WordConnectionState => ({ version: { base: { id: version, number: 1 }, current: { id: version, number: 1 }, state: "current" }, lease: owner === "available" ? { state: "available", expiresAt: null } : { state: owner, generation: "generation-one", expiresAt: 999999 } })
  const transport: Transport = {
    async lease(action) { events.push(action); if (action === "claim" && owner === "available") owner = "owned"; return connection() },
    async save() { events.push("upload"); version = "saved-two"; return { version: { id: version, number: 2, parent: "base-one", created: true }, overlap: false, connection: connection() } },
    async heartbeat() { return connection() }, async done() { return connection() },
  }
  const sync = new WordSync("base-one", { saved: async () => saved, save: async () => { events.push("native-save"); saved = true }, export: async () => { events.push("export"); return new Uint8Array([1]) } }, transport, { read: () => null, write() {} }, () => {}, async () => "synthetic-fingerprint", async () => {})
  return { sync, events }
}
function deferred() { let resolve!: () => void; const promise = new Promise<void>(done => { resolve = done }); return { promise, resolve } }

test("reviewed local editing claims available ownership and leaves canonical history untouched", async () => {
  const { sync, events } = fixture()
  await sync.start(); events.length = 0
  const identity = await sync.localEdit(async value => { events.push("local-edit"); return value })
  expect(identity).toEqual({ base: "base-one", generation: "generation-one" })
  expect(events).toEqual(["inspect", "claim", "local-edit"])
  expect(sync.snapshot().content).toBe("changed"); expect(sync.snapshot().status).toBe("not-saved"); expect(sync.snapshot().busy).toBe(false)
})

test("Save waits for deferred native application; no export or upload races the reviewed edit", async () => {
  const { sync, events } = fixture("owned"), entered = deferred(), release = deferred()
  await sync.start(); events.length = 0
  const edit = sync.localEdit(async () => { events.push("edit-start"); entered.resolve(); await release.promise; events.push("edit-end") })
  await entered.promise
  const save = sync.save()
  expect(sync.snapshot().busy).toBe(true); expect(await sync.tick()).toBe(false)
  expect(events).toEqual(["inspect", "edit-start"])
  release.resolve(); await edit; await save
  expect(events.indexOf("export")).toBeGreaterThan(events.indexOf("edit-end"))
  expect(events.indexOf("upload")).toBeGreaterThan(events.indexOf("edit-end"))
})

test("another owner refuses Apply without a takeover, native edit or canonical write", async () => {
  const { sync, events } = fixture("other")
  await sync.start(); events.length = 0
  await expect(sync.localEdit(async () => { events.push("local-edit") })).rejects.toThrow("Another editing session")
  expect(events).toEqual(["inspect"]); expect(sync.snapshot().problem).toBe(true)
})
