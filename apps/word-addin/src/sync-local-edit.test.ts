import { expect, test } from "bun:test"
import { WordSync, type Acknowledgement, type Transport } from "./sync"
import { ConnectionError } from "./client"
import type { WordConnectionState } from "../../../dochaus/lib/word-addin-contract"

function fixture(owner: "available" | "owned" | "other" = "available", checkpoint: Acknowledgement | null = null) {
  const events: string[] = []
  let version = "base-one", saved = false, bytes = 1
  let operation: { sequence: number; operationId: string | null; state: "idle" | "pending" | "confirmed" | "not-dispatched" } = { sequence: 0, operationId: null, state: "idle" }
  const connection = (): WordConnectionState => ({ version: { base: { id: version, number: 1 }, current: { id: version, number: 1 }, state: "current" }, nativeOperation: { sequence: operation.sequence, state: operation.state }, lease: owner === "available" ? { state: "available", expiresAt: null } : { state: owner, generation: "generation-one", expiresAt: 999999 } })
  const transport: Transport = {
    async lease(action) { events.push(action); if (action === "claim" && owner === "available") owner = "owned"; return connection() },
    async nativeOperation(_base, _generation, input) {
      events.push(input.action)
      if (input.action === "begin") operation = { sequence: input.sequence + 1, operationId: input.operationId, state: "pending" }
      if (input.action === "settle") operation = { sequence: input.sequence, operationId: input.operationId, state: input.outcome }
      return { operation, connection: connection() }
    },
    async save() { events.push("upload"); version = "saved-two"; return { version: { id: version, number: 2, parent: "base-one", created: true }, overlap: false, connection: connection() } },
    async heartbeat() { return connection() }, async done() { return connection() },
  }
  const host = { saved: async () => saved, save: async () => { events.push("native-save"); saved = true }, export: async () => { events.push("export"); return new Uint8Array([bytes]) } }
  const checkpoints = { read: () => checkpoint, write(value: Acknowledgement) { checkpoint = value } }
  const reopen = () => new WordSync("base-one", host, transport, checkpoints, () => {}, async value => String(value[0]), async () => {})
  const sync = reopen()
  return { sync, reopen, transport, events, get saved() { return saved }, set saved(value: boolean) { saved = value }, get bytes() { return bytes }, set bytes(value: number) { bytes = value } }
}
function deferred() { let resolve!: () => void; const promise = new Promise<void>(done => { resolve = done }); return { promise, resolve } }

test("reviewed local editing claims available ownership and leaves canonical history untouched", async () => {
  const { sync, events } = fixture()
  await sync.start(); events.length = 0
  const identity = await sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async value => { events.push("local-edit"); return value })
  expect(identity).toEqual({ base: "base-one", generation: "generation-one" })
  expect(events).toEqual(["inspect", "claim", "begin", "local-edit", "settle"])
  expect(sync.snapshot().content).toBe("changed"); expect(sync.snapshot().status).toBe("not-saved"); expect(sync.snapshot().busy).toBe(false)
})

test("Save waits for deferred native application; no export or upload races the reviewed edit", async () => {
  const { sync, events } = fixture("owned"), entered = deferred(), release = deferred()
  await sync.start(); events.length = 0
  const edit = sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { events.push("edit-start"); entered.resolve(); await release.promise; events.push("edit-end") })
  await entered.promise
  const save = sync.save()
  expect(sync.snapshot().busy).toBe(true); expect(await sync.tick()).toBe(false)
  expect(events).toEqual(["inspect", "begin", "edit-start"])
  release.resolve(); await edit; await save
  expect(events.indexOf("export")).toBeGreaterThan(events.indexOf("edit-end"))
  expect(events.indexOf("upload")).toBeGreaterThan(events.indexOf("edit-end"))
})

test("another owner refuses Apply without a takeover, native edit or canonical write", async () => {
  const { sync, events } = fixture("other")
  await sync.start(); events.length = 0
  await expect(sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { events.push("local-edit") })).rejects.toThrow("Another editing session")
  expect(events).toEqual(["inspect"]); expect(sync.snapshot().problem).toBe(true)
})

test("Undo reporting natively saved after a reviewed edit cannot publish until deliberate Save", async () => {
  const f = fixture("owned")
  await f.sync.start()
  await f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { f.bytes = 2; f.saved = false })
  f.events.length = 0
  await f.sync.tick()
  expect(f.events.some(event => ["export", "upload", "native-save"].includes(event))).toBe(false)
  // Native Undo restores the original package and Word's saved flag, even
  // though this fresh pane has never acknowledged a canonical Save.
  f.bytes = 1; f.saved = true
  await f.sync.tick()
  await f.sync.refresh()
  expect(f.events.some(event => ["export", "upload", "native-save"].includes(event))).toBe(false)
  expect(f.sync.snapshot()).toMatchObject({ status: "not-saved", content: "changed", acknowledgement: undefined })
  expect(await f.sync.save()).toBe(true)
  expect(f.events.filter(event => event === "upload")).toHaveLength(1)
  expect(f.events.filter(event => event === "native-save")).toHaveLength(1)
  expect(f.sync.snapshot().status).toBe("saved")
  await f.sync.tick()
  await f.sync.refresh()
  expect(f.events.filter(event => event === "upload")).toHaveLength(1)
})

test("an uncertain reviewed edit is fenced before its callback can mutate and throw", async () => {
  const f = fixture("owned", { base: "base-one", fingerprint: "1", number: 1, at: 0, overlap: false })
  // No start() is required for localEdit's guarded identity check. A later
  // refresh must not probe an unchecked saved checkpoint through an export.
  await expect(f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { f.bytes = 2; f.saved = true; throw new Error("Lost native response") })).rejects.toThrow("Lost native response")
  f.events.length = 0
  await f.sync.tick()
  await f.sync.refresh()
  expect(f.events.some(event => ["export", "upload", "native-save"].includes(event))).toBe(false)
  expect(f.sync.snapshot()).toMatchObject({ status: "not-saved", content: "changed" })
  expect(await f.sync.done()).toBe(false)
  expect(f.events.filter(event => event === "upload")).toHaveLength(0)
  expect(f.events.filter(event => event === "native-save")).toHaveLength(0)
})


test("a lost begin acknowledgement is read back before the only native dispatch", async () => {
  const f = fixture("owned"), command = f.transport.nativeOperation!, calls: string[] = []
  f.transport.nativeOperation = async (...args) => {
    const result = await command(...args)
    if (args[2].action === "begin") throw new ConnectionError("Lost begin response", true)
    return result
  }
  await f.sync.start()
  await f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { calls.push("native"); f.bytes = 2 })
  expect(calls).toEqual(["native"])
  expect(f.events.filter(event => ["begin", "read", "settle"].includes(event))).toEqual(["begin", "read", "settle"])
})

test("lost begin with unavailable read never dispatches or saves", async () => {
  const f = fixture("owned"), command = f.transport.nativeOperation!
  f.transport.nativeOperation = async (...args) => {
    if (args[2].action === "read") throw new ConnectionError("Offline", true)
    const result = await command(...args)
    if (args[2].action === "begin") throw new ConnectionError("Lost begin response", true)
    return result
  }
  await f.sync.start()
  let writes = 0
  await expect(f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { writes++ })).rejects.toThrow()
  expect(writes).toBe(0); expect(await f.sync.save()).toBe(false)
  expect(f.events).not.toContain("native-save"); expect(f.events).not.toContain("upload")
})

test("unknown admission with an idle server gives conditional recovery advice", async () => {
  const f = fixture("owned"), command = f.transport.nativeOperation!
  f.transport.nativeOperation = async (...args) => {
    if (args[2].action === "begin" || args[2].action === "read") throw new ConnectionError("Offline", true)
    return command(...args)
  }
  await f.sync.start()
  let writes = 0
  await expect(f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { writes++ })).rejects.toThrow()
  expect(writes).toBe(0)
  expect((await f.transport.lease("inspect", "base-one")).nativeOperation?.state).toBe("idle")
  expect(f.sync.snapshot().message).toContain("if it shows an unconfirmed Word change")
  expect(f.sync.snapshot().message).toContain("Check the document status")
  expect(await f.sync.save()).toBe(false)
})

test("lost settlement acknowledgement reads terminal status without replay", async () => {
  const f = fixture("owned"), command = f.transport.nativeOperation!, writes: number[] = []
  f.transport.nativeOperation = async (...args) => {
    const result = await command(...args)
    if (args[2].action === "settle") throw new ConnectionError("Lost settle response", true)
    return result
  }
  await f.sync.start(); await f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { writes.push(1); f.bytes = 2 })
  expect(writes).toEqual([1]); expect(f.events.filter(event => event === "read")).toHaveLength(1)
})

test("pane reconstruction cannot reapply or save an indeterminate native operation", async () => {
  const f = fixture("owned")
  await f.sync.start()
  await expect(f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { f.bytes = 2; f.saved = true; throw Object.assign(new Error("Lost Word response"), { mutationState: "uncertain" }) })).rejects.toThrow()
  expect(f.sync.snapshot().message).toContain("Create saved copy")
  expect(f.sync.snapshot().message).toContain("will not include unsent Word edits")
  const fresh = f.reopen(); await fresh.start(); f.events.length = 0
  expect(fresh.snapshot().message).toContain("Create saved copy")
  expect(fresh.snapshot().message).toContain("Keep this Word file")
  await fresh.tick(); expect(await fresh.save()).toBe(false); expect(await fresh.done()).toBe(false)
  let writes = 0
  await expect(fresh.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { writes++ })).rejects.toThrow()
  expect(writes).toBe(0); expect(f.events).not.toContain("native-save"); expect(f.events).not.toContain("upload")
})

test("closing the pane while begin is in flight prevents native dispatch", async () => {
  const f = fixture("owned"), command = f.transport.nativeOperation!, entered = deferred(), release = deferred()
  f.transport.nativeOperation = async (...args) => {
    const result = await command(...args)
    if (args[2].action === "begin") { entered.resolve(); await release.promise }
    return result
  }
  await f.sync.start(); let writes = 0
  const promise = f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { writes++ })
  await entered.promise; f.sync.dispose(); release.resolve()
  await expect(promise).rejects.toThrow(); expect(writes).toBe(0); expect(f.events).toContain("settle")
})

test("the reviewed conversation and turn bind the amendment before dispatch", async () => {
  const f = fixture("owned"), command = f.transport.nativeOperation!, received: unknown[] = []
  f.transport.nativeOperation = async (...args) => {
    if (args[2].action === "begin") received.push(args[2])
    return command(...args)
  }
  await f.sync.start()
  await f.sync.localEdit({ conversationId: "conversation-42", turnId: "turn-19" }, async () => {})
  expect(received).toHaveLength(1)
  expect(received[0]).toMatchObject({ action: "begin", intent: "amendment", sequence: 0, proposal: { conversationId: "conversation-42", turnId: "turn-19" } })
  expect((received[0] as { operationId: string }).operationId).toBeString()
})

test("a missing proposal binding cannot begin or dispatch a Word edit", async () => {
  const f = fixture("owned"); await f.sync.start(); f.events.length = 0
  let writes = 0
  await expect(f.sync.localEdit({ conversationId: "", turnId: "turn" }, async () => { writes++ })).rejects.toThrow("Review the proposal")
  expect(writes).toBe(0); expect(f.events).not.toContain("begin")
})
