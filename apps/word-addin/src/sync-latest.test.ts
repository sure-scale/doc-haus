import { expect, test } from "bun:test"
import type { WordConnectionState, WordLeaseAction, WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"
import { ConnectionError, type LatestWordContent } from "./client"
import { DocumentStateError, type WordHost } from "./office"
import { WordSync, browserCheckpoints, type Acknowledgement, type Transport } from "./sync"
import { documentFingerprint } from "./fingerprint"
import { strToU8, zipSync } from "fflate"

function fixture(owner: "owned" | "other" | "available" = "owned", fingerprint = async (value: Uint8Array) => String(value[0])) {
  const events: string[] = []
  let current = "v2"
  let bytes = 1
  let generation = "original-owner"
  let operation: { sequence: number; operationId: string | null; state: "idle" | "pending" | "confirmed" | "not-dispatched" } = { sequence: 0, operationId: null, state: "idle" }
  let checkpoint: Acknowledgement | null = { base: "v1", fingerprint: "1", number: 1, at: 0, overlap: false }
  const state = (base: string): WordConnectionState => ({
    version: { base: { id: base, number: base === "v1" ? 1 : 2 }, current: { id: current, number: 2 }, state: base === current ? "current" : "stale" },
    nativeOperation: { sequence: operation.sequence, state: operation.state },
    lease: { state: owner, expiresAt: owner === "available" ? null : 100_000, ...(owner === "available" ? {} : { generation }) },
  })
  const transport: Transport = {
    lease: async (action: WordLeaseAction, base, fence) => {
      events.push(`${action}:${base}`)
      if (action === "claim") { owner = "owned"; generation = "new-owner" }
      if (action === "release") {
        if (owner !== "owned" || fence !== generation) throw new ConnectionError("Lease changed", false, state(base))
        owner = "available"
      }
      return state(base)
    },
    nativeOperation: async (base, fence, input) => {
      events.push(`${input.action}:${base}`)
      if (owner !== "owned" || generation !== fence) throw new ConnectionError("Lease changed", false, state(base))
      if (input.action === "begin") operation = { sequence: input.sequence + 1, operationId: input.operationId, state: "pending" }
      if (input.action === "settle") operation = { sequence: input.sequence, operationId: input.operationId, state: input.outcome }
      return { operation, connection: state(base) }
    },
    save: async (_bytes, base) => {
      events.push(`upload:${base}`)
      current = "v3"
      return { version: { id: current, number: 3, parent: base, created: true }, overlap: false, connection: state(current) } satisfies WordSaveResponse
    },
    heartbeat: async base => { events.push(`heartbeat:${base}`); return state(base) },
    done: async base => { events.push(`done:${base}`); return state(base) },
  }
  const host: WordHost = {
    saved: async () => true,
    save: async () => { events.push("native-save") },
    export: async () => { events.push("export"); return new Uint8Array([bytes]) },
    replace: async (value, base) => { events.push(`replace:${base}`); bytes = value[0]! },
  }
  const sync = new WordSync("v1", host, transport, {
    read: () => checkpoint,
    write: value => { checkpoint = value },
    clear: () => { events.push("clear-checkpoint"); checkpoint = null },
  }, () => {}, fingerprint, async () => {}, () => 50_000)
  const load = async (base: string): Promise<LatestWordContent> => {
    events.push(`load:${base}`)
    return { version: { id: "v2", number: 2 }, bytes: new Uint8Array([2]) }
  }
  return {
    sync, host, transport, events, load,
    get bytes() { return bytes }, set bytes(value: number) { bytes = value },
    get current() { return current }, set current(value: string) { current = value },
    get checkpoint() { return checkpoint },
    set owner(value: typeof owner) { owner = value },
    set generation(value: string) { generation = value },
  }
}

test("refresh replaces the current buffer, adopts only its version and releases only the old own lease without publishing", async () => {
  const f = fixture()
  expect(await f.sync.loadLatest(f.load, version => { expect(version).toEqual({ id: "v2", number: 2 }); return true })).toBe(true)
  expect(f.events).toEqual(["inspect:v1", "export", "load:v1", "inspect:v2", "export", "inspect:v1", "begin:v1", "export", "replace:v2", "settle:v1", "clear-checkpoint", "release:v1", "inspect:v2"])
  expect(f.bytes).toBe(2)
  expect(f.checkpoint).toBeNull()
  expect(f.sync.snapshot()).toMatchObject({ status: "not-saved", content: "changed", connection: { version: { state: "current", base: { id: "v2" } }, lease: { state: "available" } } })
  expect(f.sync.snapshot().acknowledgement).toBeUndefined()
  expect(f.sync.snapshot().message).toContain("Latest version loaded")
  expect(f.sync.snapshot().message).not.toContain("changed again")
  await f.sync.tick()
  expect(f.events.some(value => value.startsWith("upload:") || value === "native-save" || value.startsWith("claim:"))).toBe(false)
  expect(await f.sync.save()).toBe(true)
  expect(f.events).toContain("upload:v2")
  expect(f.sync.snapshot().status).toBe("saved")
})

test("refresh leaves another editor's lease intact and does not steal it", async () => {
  const f = fixture("other")
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(false)
  expect(f.sync.snapshot().connection?.lease.state).toBe("other")
  expect(f.events.some(value => /^(release|claim|takeover|upload|begin|replace):/.test(value))).toBe(false)
  expect(f.bytes).toBe(1)
})

test("native AutoSave after refresh cannot publish even if lease release failed", async () => {
  const f = fixture()
  const lease = f.transport.lease
  f.transport.lease = async (...args) => {
    if (args[0] === "release") { f.events.push("release-failed"); throw new ConnectionError("Offline", true) }
    return lease(...args)
  }
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(true)
  await f.sync.tick()
  await f.sync.refresh()
  expect(f.sync.snapshot().status).toBe("not-saved")
  expect(f.events.some(value => value.startsWith("upload:") || value === "native-save")).toBe(false)
})

test("confirmation is required even for a natively saved file, and cancelling keeps all local content and lease", async () => {
  const f = fixture()
  let confirms = 0
  expect(await f.sync.loadLatest(f.load, () => { confirms++; return false })).toBe(false)
  expect(confirms).toBe(1)
  expect(f.bytes).toBe(1)
  expect(f.events).toEqual(["inspect:v1", "export", "load:v1"])
  expect(f.checkpoint?.base).toBe("v1")
})

test("changes to the captured Word package during confirmation refuse import causally", async () => {
  for (const changed of [false, true]) {
    const f = fixture()
    expect(await f.sync.loadLatest(f.load, () => { if (changed) f.bytes = 9; return true })).toBe(!changed)
    expect(f.bytes).toBe(changed ? 9 : 2)
    expect(f.events.includes("replace:v2")).toBe(!changed)
    if (changed) expect(f.sync.snapshot().message).toContain("changed while loading")
  }
})

test("fresh package checks catch real body and header formatting changes with unchanged text", async () => {
  for (const part of ["word/document.xml", "word/header1.xml"]) {
    for (const changed of [false, true]) {
      const f = fixture("owned", documentFingerprint)
      let bold = false
      const root = part.includes("header") ? "hdr" : "document"
      f.host.export = async () => zipSync({ [part]: strToU8(`<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t>Same clause text.</w:t></w:r></w:p></w:${root}>`) })
      expect(await f.sync.loadLatest(f.load, () => { bold = changed; return true })).toBe(!changed)
      expect(f.events.includes("replace:v2")).toBe(!changed)
    }
  }
})

test("a canonical version advanced during confirmation refuses import before any mutation", async () => {
  const f = fixture()
  expect(await f.sync.loadLatest(f.load, () => { f.current = "v3"; return true })).toBe(false)
  expect(f.bytes).toBe(1)
  expect(f.events).toEqual(["inspect:v1", "export", "load:v1", "inspect:v2"])
  expect(f.sync.snapshot().message).toContain("latest version changed")
})

test("an owner race during native import cannot release the replacement editor", async () => {
  const f = fixture()
  const replace = f.host.replace!
  f.host.replace = async (...args) => { await replace(...args); f.owner = "other"; f.generation = "other-owner" }
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(false)
  expect(f.events.some(value => /^(claim|takeover|upload):/.test(value))).toBe(false)
})

test("unsupported hosts and already current copies never download or mutate", async () => {
  for (const unsupported of [false, true]) {
    const f = fixture()
    if (unsupported) f.host.replace = undefined
    else f.current = "v1"
    expect(await f.sync.loadLatest(f.load, () => { throw new Error("No confirmation expected") })).toBe(false)
    expect(f.bytes).toBe(1)
    expect(f.events.some(value => /^(load|replace|claim|release|upload):/.test(value))).toBe(false)
  }
})

test("uncertain native/settings failure keeps old base, blocks repeat import, claim, Apply and all saving", async () => {
  const f = fixture()
  f.host.replace = async () => { f.events.push("partial-import"); f.bytes = 2; throw new DocumentStateError("Word could not preserve the link") }
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(false)
  expect(f.sync.snapshot().connection?.version.base?.id).toBe("v1")
  expect(f.checkpoint?.base).toBe("v1")
  f.current = "v1" // Even an unexpected server recovery cannot grant writing.
  f.owner = "available"
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(false)
  expect(await f.sync.save()).toBe(false)
  await f.sync.startEditing()
  await f.sync.refresh()
  await expect(f.sync.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { throw new Error("Must not apply") })).rejects.toBeInstanceOf(DocumentStateError)
  expect(f.events.filter(value => value === "partial-import")).toHaveLength(1)
  expect(f.events.some(value => /^(claim|release|takeover|upload):/.test(value) || value === "native-save")).toBe(false)
  expect(f.sync.snapshot().canSave).toBe(false)
})

test("an outstanding Save with an uncertain response refuses refresh without retrying or discarding its bytes", async () => {
  const f = fixture()
  f.current = "v1"
  f.transport.save = async () => { f.events.push("uncertain-save"); throw new ConnectionError("Offline", true) }
  expect(await f.sync.save()).toBe(false)
  f.current = "v2"
  const earlier = f.events.length
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(false)
  expect(f.events.slice(earlier)).toEqual([])
  expect(f.sync.snapshot().message).toContain("pending Save")
  expect(f.bytes).toBe(1)
})

test("after refreshed link persistence, later Save acknowledgements use the namespace the reopened copy will read", async () => {
  const values = new Map<string, string>()
  const storage = { getItem: (key: string) => values.get(key) ?? null, setItem: (key: string, value: string) => values.set(key, value), removeItem: (key: string) => values.delete(key) } as unknown as Storage
  const link = { version: 1 as const, origin: "http://localhost:4000", apiOrigin: "https://localhost:4500", matter: "synthetic", matterName: "Synthetic matter", document: "document", documentName: "Draft.docx", baseVersion: "v1", token: "synthetic-capability" }
  const initial = await browserCheckpoints(link, storage)
  const ack = { base: "v3", fingerprint: "a".repeat(64), number: 3, at: 0, overlap: false }
  initial.write({ ...ack, base: "v1" })
  await initial.rebase!("v2")
  expect(initial.read()).toBeNull()
  initial.write(ack)
  expect((await browserCheckpoints({ ...link, baseVersion: "v2" }, storage)).read()).toEqual(ack)
  expect((await browserCheckpoints(link, storage)).read()).toBeNull()
})


test("a failed post-import status check still reports replacement so stale chat context is cleared", async () => {
  const f = fixture("owned")
  const lease = f.transport.lease
  let inspections = 0
  f.transport.lease = async (...args) => {
    if (args[0] === "inspect" && ++inspections === 4) throw new ConnectionError("Offline", true)
    return lease(...args)
  }
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(true)
  expect(f.bytes).toBe(2)
  expect(f.sync.snapshot().status).toBe("not-saved")
  expect(f.sync.snapshot().acknowledgement).toBeUndefined()
  expect(f.events.some(value => value.startsWith("upload:") || value === "native-save")).toBe(false)
})

test("a failed post-BEGIN read settles as not dispatched without replacing Word", async () => {
  const f = fixture(); const exportCurrent = f.host.export; let exports = 0
  f.host.export = async () => { if (++exports === 3) throw new Error("Word export unavailable"); return exportCurrent() }
  expect(await f.sync.loadLatest(f.load, () => true)).toBe(false)
  expect(f.events).toContain("begin:v1"); expect(f.events).toContain("settle:v1")
  expect(f.events).not.toContain("replace:v2"); expect(f.bytes).toBe(1)
})

test("closing the pane during the post-BEGIN read never replaces Word", async () => {
  const f = fixture(); const exportCurrent = f.host.export; let exports = 0
  let release!: () => void, entered!: () => void
  const wait = new Promise<void>(resolve => { release = resolve }), ready = new Promise<void>(resolve => { entered = resolve })
  f.host.export = async () => { if (++exports === 3) { entered(); await wait }; return exportCurrent() }
  const loading = f.sync.loadLatest(f.load, () => true)
  await ready; f.sync.dispose(); release()
  expect(await loading).toBe(false); expect(f.events).toContain("settle:v1")
  expect(f.events).not.toContain("replace:v2"); expect(f.bytes).toBe(1)
})
