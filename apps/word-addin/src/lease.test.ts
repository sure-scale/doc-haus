import { expect, test } from "bun:test"
import type { WordConnectionState, WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"
import { ConnectionError } from "./client"
import { WordSync, type Acknowledgement, type Checkpoints, type Transport } from "./sync"

const fingerprint = async (bytes: Uint8Array) => String(bytes[0])
const immediate = async () => {}
const checkpoints = (value: Acknowledgement | null = null): Checkpoints => ({ read: () => value, write: () => {} })

function server() {
  let current = "v1"
  let number = 1
  let owner: string | undefined
  let generation = ""
  let serial = 0
  let loseResponses = 0
  let offlineRequests = 0
  let last: { base: string; byte: number; generation: string } | undefined
  const events: string[] = []
  const writes: { base: string; byte: number; generation: string }[] = []
  const state = (session: string, base: string): WordConnectionState => ({
    version: {
      base: { id: base, number: base === current ? number : 1 },
      current: { id: current, number },
      state: base === current ? "current" : "stale",
    },
    lease: {
      state: !owner ? "available" : session === owner ? "owned" : "other",
      ...(owner ? { generation } : {}),
      expiresAt: owner ? 90_000 : null,
    },
  })
  const refuse = (session: string, base: string): never => {
    throw new ConnectionError("Refused", false, state(session, base))
  }
  const transport = (session: string): Transport => ({
    lease: async (action, base, fence, expected) => {
      events.push(`${session}:${action}:${base}`)
      if (action === "inspect") return state(session, base)
      if (action === "release") {
        if (session !== owner || fence !== generation) return refuse(session, base)
        owner = undefined
        return state(session, base)
      }
      if (base !== current) return refuse(session, base)
      if (action === "takeover") {
        if (expected?.current !== current || expected.generation !== generation) return refuse(session, base)
      } else if (owner && owner !== session) return refuse(session, base)
      if (owner !== session) generation = `fence-${++serial}`
      owner = session
      return state(session, base)
    },
    save: async (bytes, base, fence) => {
      events.push(`${session}:upload:${base}:${fence}`)
      if (offlineRequests > 0) {
        offlineRequests--
        throw new ConnectionError("Request never reached server", true)
      }
      if (session !== owner || fence !== generation) return refuse(session, base)
      const retry = last?.base === base && last.byte === bytes[0] && last.generation === fence
      if (!retry) {
        if (base !== current) return refuse(session, base)
        last = { base, byte: bytes[0]!, generation: fence }
        writes.push(last)
        current = `v${++number}`
      }
      if (loseResponses > 0) {
        loseResponses--
        throw new ConnectionError("Response lost", true)
      }
      return {
        version: { id: current, number, parent: base, created: !retry },
        overlap: false,
        connection: state(session, current),
      } satisfies WordSaveResponse
    },
    heartbeat: async (base, fence) => {
      events.push(`${session}:heartbeat:${base}:${fence}`)
      if (session !== owner || fence !== generation || base !== current) return refuse(session, base)
      return state(session, base)
    },
    done: async (base, fence) => {
      events.push(`${session}:done:${base}:${fence}`)
      if (session !== owner || fence !== generation || base !== current) return refuse(session, base)
      owner = undefined
      return state(session, base)
    },
  })
  return {
    transport,
    events,
    writes,
    state,
    expire: () => {
      owner = undefined
    },
    lose: (count: number) => {
      loseResponses = count
    },
    offline: (count: number) => {
      offlineRequests = count
    },
    advance: () => {
      current = `v${++number}`
    },
  }
}

function pane(api: Transport, options: { base?: string; checkpoint?: Acknowledgement; initialByte?: number } = {}) {
  let byte = options.initialByte ?? 1
  let saved = true
  const hostEvents: string[] = []
  const sync = new WordSync(
    options.base ?? "v1",
    {
      saved: async () => saved,
      save: async () => {
        hostEvents.push("save")
        saved = true
      },
      export: async () => {
        hostEvents.push("export")
        return new Uint8Array([byte])
      },
      close: async (base) => {
        hostEvents.push(`close:${base}`)
      },
    },
    api,
    checkpoints(options.checkpoint),
    () => {},
    fingerprint,
    immediate,
  )
  return {
    sync,
    hostEvents,
    edit: (value: number, nativeSaved = true) => {
      byte = value
      saved = nativeSaved
    },
  }
}

test("two initial panes only inspect; an explicit owner saves and syncs while the other never adopts its parent", async () => {
  const api = server()
  const first = pane(api.transport("first"))
  const second = pane(api.transport("second"))
  await first.sync.start()
  await second.sync.start()
  await first.sync.tick()
  await second.sync.tick()
  expect(api.events.every((event) => event.includes(":inspect:"))).toBe(true)
  expect(api.writes).toEqual([])
  expect(first.hostEvents).toEqual([])
  expect(first.sync.snapshot().canSave).toBe(true)
  expect(first.sync.snapshot().message).toContain("Ready to save")
  expect(await first.sync.refresh()).toBe(true) // Deliberate Start saving here.
  expect(api.events.indexOf("first:claim:v1")).toBeLessThan(api.events.indexOf("first:upload:v1:fence-1"))
  expect(first.sync.snapshot().status).toBe("saved")
  first.edit(2)
  await first.sync.tick()
  expect(api.writes.map((write) => write.base)).toEqual(["v1", "v2"])
  await second.sync.tick()
  expect(second.sync.snapshot().connection?.version.state).toBe("stale")
  expect(await second.sync.save()).toBe(false)
  expect(await second.sync.done()).toBe(false)
  expect(second.hostEvents).toEqual([])
  expect(second.sync.snapshot().acknowledgement).toBeUndefined()
  expect(api.writes).toHaveLength(2)
})

test("matching checkpoints recover their acknowledged parent; a different forwarded copy keeps the embedded base", async () => {
  const api = server()
  api.advance()
  const previous: Acknowledgement = { base: "v2", number: 2, fingerprint: "1", at: 1, overlap: false }
  for (const matches of [true, false]) {
    const view = pane(api.transport(matches ? "match" : "fork"), { checkpoint: previous, initialByte: matches ? 1 : 9 })
    view.edit(matches ? 1 : 9, false)
    await view.sync.start()
    expect(view.sync.snapshot().status).toBe("not-saved")
    expect(api.writes).toEqual([])
    expect(view.sync.snapshot().connection?.version.state).toBe(matches ? "current" : "stale")
    expect(view.sync.snapshot().acknowledgement?.base).toBe(matches ? "v2" : undefined)
    view.edit(3, false)
    expect(await view.sync.done()).toBe(matches)
    if (matches) {
      expect(api.writes.map((write) => write.base)).toEqual(["v2"])
      expect(view.hostEvents.at(-1)).toBe("close:v3")
      api.writes.length = 0
    } else {
      expect(view.hostEvents.filter((event) => event === "save" || event.startsWith("close"))).toEqual([])
    }
  }
})

test("takeover uses the displayed version and generation; a raced observation cannot save or close", async () => {
  const api = server()
  const owner = pane(api.transport("owner"))
  const standby = pane(api.transport("standby"))
  await owner.sync.start()
  await owner.sync.takeover() // No displayed other owner; cannot manufacture ownership.
  expect(api.events.some((event) => event.includes(":takeover:"))).toBe(false)
  await owner.sync.refresh()
  const current = owner.sync.snapshot().acknowledgement!.base
  // A fresh copy of the current version can deliberately take over, never a stale copy.
  const currentStandby = pane(api.transport("standby"), { base: current })
  await currentStandby.sync.start()
  expect(currentStandby.sync.snapshot().canSave).toBe(false)
  expect(await currentStandby.sync.takeover()).toBe(true)
  expect(api.writes).toHaveLength(1)
  expect(currentStandby.hostEvents).toEqual([])
  await owner.sync.tick()
  expect(owner.sync.snapshot().canSave).toBe(false)
  expect(await owner.sync.save()).toBe(false)
  // A second displayed takeover observation becomes invalid when the version moves.
  const racing = pane(api.transport("racing"), { base: current })
  await racing.sync.start()
  api.advance()
  expect(await racing.sync.takeover()).toBe(false)
  expect(racing.sync.snapshot().connection?.version.state).toBe("stale")
  expect(racing.hostEvents).toEqual([])
  await standby.sync.start()
  expect(standby.sync.snapshot().connection?.version.state).toBe("stale")
})

test("release and expiry never silently reacquire; Check status does not unexpectedly claim", async () => {
  const api = server()
  const view = pane(api.transport("pane"))
  await view.sync.start()
  await view.sync.refresh()
  const writes = api.writes.length
  await view.sync.release()
  await view.sync.tick()
  expect(view.sync.snapshot().canSave).toBe(true)
  expect(api.events.filter((event) => event.includes(":claim:"))).toHaveLength(1)
  expect(api.writes).toHaveLength(writes)
  await view.sync.refresh()
  expect(view.sync.snapshot().canSave).toBe(true)
  api.expire()
  await view.sync.tick()
  expect(view.sync.snapshot().canSave).toBe(true)
  expect(api.events.filter((event) => event.includes(":claim:"))).toHaveLength(2)
  const checking = pane(api.transport("checking"), { base: view.sync.snapshot().acknowledgement!.base })
  await api.transport("other").lease("claim", view.sync.snapshot().acknowledgement!.base)
  await checking.sync.start()
  api.expire()
  await checking.sync.refresh() // Label was Check status; newly available is not an implicit claim.
  expect(checking.sync.snapshot().connection?.lease.state).toBe("available")
  expect(api.events.some((event) => event.startsWith("checking:claim"))).toBe(false)
})

test("lost save ACK retries the original fence and parent before inspecting; fenced retries never claim", async () => {
  for (const expired of [false, true]) {
    const api = server()
    const view = pane(api.transport("pane"))
    await view.sync.start()
    api.lose(2)
    expect(await view.sync.save()).toBe(false)
    expect(api.writes).toHaveLength(1)
    expect(view.sync.snapshot().acknowledgement).toBeUndefined()
    view.edit(2, false)
    if (expired) api.expire()
    expect(await view.sync.save()).toBe(!expired)
    if (expired) {
      expect(api.writes).toHaveLength(1)
      expect(view.sync.snapshot().canSave).toBe(false)
      expect(view.hostEvents.filter((event) => event === "save")).toHaveLength(1)
    } else {
      expect(api.writes.map(({ base, generation }) => ({ base, generation }))).toEqual([
        { base: "v1", generation: "fence-1" },
        { base: "v2", generation: "fence-1" },
      ])
      expect(view.sync.snapshot().acknowledgement?.fingerprint).toBe("2")
    }
    expect(api.events.filter((event) => event.includes(":claim:"))).toHaveLength(1)
  }
})

test("Done sends the acknowledged base and fence before close; its released fence cannot resume background writes", async () => {
  const api = server()
  const view = pane(api.transport("pane"))
  await view.sync.start()
  expect(await view.sync.done()).toBe(true)
  expect(api.events).toContain("pane:done:v2:fence-1")
  expect(view.hostEvents.at(-1)).toBe("close:v2")
  expect(view.sync.snapshot().connection?.lease.state).toBe("available")
  view.edit(2, false)
  await view.sync.tick()
  expect(view.sync.snapshot().status).toBe("not-saved")
  expect(view.sync.snapshot().canSave).toBe(true)
  expect(api.writes).toHaveLength(1)
})

test("explicit recovery of fenced pending bytes requires an unchanged base and a successful new claim", async () => {
  for (const takeover of [false, true]) {
    const api = server()
    const view = pane(api.transport("pane"))
    await view.sync.start()
    api.offline(2)
    expect(await view.sync.save()).toBe(false)
    expect(api.writes).toEqual([])
    api.expire()
    view.edit(2, false)
    if (takeover) {
      await api.transport("other").lease("claim", "v1")
      await view.sync.refresh() // Old pending generation is refused; display the current owner.
      expect(view.sync.snapshot().connection?.lease.state).toBe("other")
      expect(await view.sync.takeover()).toBe(true)
      expect(api.writes).toEqual([])
    }
    expect(await view.sync.save()).toBe(true)
    expect(api.writes).toEqual([{ byte: 2, base: "v1", generation: takeover ? "fence-3" : "fence-2" }])
    expect(view.sync.snapshot().acknowledgement?.fingerprint).toBe("2")
    expect(view.hostEvents.filter((event) => event === "save")).toHaveLength(2)
  }
})
