import { describe, expect, test } from "bun:test"
import { ConnectionError } from "./client"
import { DocumentStateError, type WordHost } from "./office"
import {
  WordSync as CheckedWordSync,
  browserCheckpoints,
  type Acknowledgement,
  type Checkpoints,
  type PaneState,
  type Transport,
} from "./sync"
import type { WordConnectionState, WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"

// These existing save/close regressions use a small protocol fixture.
// Ownership races and startup are exercised separately with a stateful server.
const connection = (id: string): WordConnectionState => ({
  version: { base: { id, number: 1 }, current: { id, number: 1 }, state: "current" },
  lease: { state: "owned", generation: "test-generation", expiresAt: 90_000 },
})
class WordSync extends CheckedWordSync {
  constructor(
    base: string,
    host: WordHost,
    transport: {
      save(bytes: Uint8Array, base: string): Promise<WordSaveResponse>
      heartbeat(): Promise<void>
      done(): Promise<void>
    },
    checkpoints: Checkpoints,
    render: (state: PaneState) => void,
    fingerprint?: (bytes: Uint8Array) => Promise<string>,
    wait?: (ms: number) => Promise<void>,
    now?: () => number,
  ) {
    let owned = true
    const state = (base: string): WordConnectionState =>
      owned ? connection(base) : { ...connection(base), lease: { state: "available", expiresAt: null } }
    super(
      base,
      host,
      {
        lease: async (action, base) => {
          if (action === "claim") owned = true
          return state(base)
        },
        save: transport.save,
        heartbeat: async (base) => {
          await transport.heartbeat()
          return connection(base)
        },
        done: async (base) => {
          await transport.done()
          owned = false
          return state(base)
        },
      },
      checkpoints,
      render,
      fingerprint,
      wait,
      now,
    )
  }
}
const acknowledgement = (id: string, number = 1): WordSaveResponse => ({
  version: { id, number, parent: null, created: true },
  overlap: false,
  connection: connection(id),
})
const checkpoints = (initial: Acknowledgement | null = null): Checkpoints => ({
  read: () => initial,
  write: () => {},
})
const fingerprint = async (bytes: Uint8Array) => String(bytes[0])
const immediate = async () => {}

test("confirmed save persists its base for reopen and never silently adopts a newer server version", async () => {
  let diskBase = "v1", uploads = 0, version = "v1"
  const observed: string[] = []
  const host: WordHost = {saved:async()=>true,save:immediate,export:async()=>new Uint8Array([1]),persistAcknowledgedBase:async (base,previous,hash)=>{expect(previous).toBe("v1");expect(hash).toBe("1");diskBase=base}}
  const transport: Transport = {
    lease:async (_action,base)=>{observed.push(base); return base === version ? connection(base) : {...connection(base),version:{base:{id:base,number:1},current:{id:version,number:2},state:"stale"}}},
    heartbeat:async base=>connection(base),done:async base=>connection(base),save:async()=>{uploads++;version="v2";return acknowledgement(version,2)},
  }
  const sync = new CheckedWordSync("v1",host,transport,checkpoints(),()=>{},fingerprint,immediate)
  expect(await sync.save()).toBe(true); expect(diskBase).toBe("v2"); expect(uploads).toBe(1)
  const reopened = new CheckedWordSync(diskBase,host,transport,checkpoints(),()=>{},fingerprint,immediate)
  await reopened.start(); expect(observed.at(-1)).toBe("v2"); expect(reopened.snapshot().connection?.version.state).toBe("current")
  version="v3"
  const stale = new CheckedWordSync(diskBase,host,transport,checkpoints(),()=>{},fingerprint,immediate)
  await stale.start(); expect(stale.snapshot().connection?.version.state).toBe("stale"); expect(stale.snapshot().canSave).toBe(false); expect(uploads).toBe(1)
})

test("post-ack metadata failure retries without another upload and concurrent content stays unconfirmed", async () => {
  for (const scenario of ["failure","edit"] as const) {
    let bytes=1,uploads=0,persistCalls=0
    const host:WordHost={saved:async()=>true,save:immediate,export:async()=>new Uint8Array([bytes]),persistAcknowledgedBase:async()=>{persistCalls++;if(scenario==="failure" && persistCalls===1)throw Error("metadata failed");if(scenario==="edit" && persistCalls===1)bytes=2}}
    const sync=new WordSync("v1",host,{heartbeat:immediate,done:immediate,save:async()=>{uploads++;return acknowledgement(`v${uploads+1}`)}},checkpoints(),()=>{},fingerprint,immediate)
    expect(await sync.save()).toBe(false); expect(sync.snapshot().status).not.toBe("saved"); expect(uploads).toBe(1)
    if (scenario==="failure") {expect(await sync.save()).toBe(true);expect(uploads).toBe(1);expect(persistCalls).toBe(2)}
    else {expect(sync.snapshot().content).toBe("changed");expect(await sync.save()).toBe(true);expect(uploads).toBe(2)}
  }
})

test("a post-ack edit persists the pending base before saving the newer buffer", async () => {
  for (const failedMetadata of [false,true]) {
    let bytes=1, uploads=0, link="v1", persisted="v1", calls=0, clock=0
    const writes: Array<{base:string;previous:string;fingerprint:string}> = []
    const host:WordHost={saved:async()=>true,save:immediate,export:async()=>new Uint8Array([bytes]),persistAcknowledgedBase:async(base,previous,hash)=>{
      calls++; writes.push({base,previous,fingerprint:hash})
      if (![previous,base].includes(link)) throw Error("link base mismatch")
      link=base
      if (failedMetadata && calls===1) throw Error("settings or native save failed")
      persisted=base
    }}
    const sync=new WordSync("v1",host,{heartbeat:immediate,done:immediate,save:async()=>{
      uploads++; if (!failedMetadata && uploads===1) bytes=2
      return acknowledgement(`v${uploads+1}`)
    }},checkpoints(),()=>{},fingerprint,immediate,()=>clock)
    expect(await sync.save()).toBe(false); expect(uploads).toBe(1); expect(sync.snapshot().status).not.toBe("saved")
    if (failedMetadata) bytes=2
    clock=31_000; await sync.tick(); expect(sync.snapshot().status).not.toBe("saved"); expect(uploads).toBe(1)
    expect(await sync.save()).toBe(true); expect(uploads).toBe(2); expect(link).toBe("v3"); expect(persisted).toBe("v3")
    expect(writes.at(-2)).toEqual({base:"v2",previous:"v1",fingerprint:"2"})
    expect(writes.at(-1)).toEqual({base:"v3",previous:"v2",fingerprint:"2"})
    expect(sync.snapshot().acknowledgement?.base).toBe("v3")
  }
})

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((value) => {
    resolve = value
  })
  return { promise, resolve }
}

describe("confirmed Word saves", () => {
  test("acknowledgement gates Saved and Done; all commands remain serialized", async () => {
    const started = deferred<void>()
    const uploaded = deferred<WordSaveResponse>()
    const closeStarted = deferred<boolean>()
    const closeReleased = deferred<void>()
    const events: string[] = []
    let bytes = 1
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: async () => {
          events.push("word-save")
        },
        export: async () => {
          events.push("export")
          return new Uint8Array([bytes])
        },
        close: async (base) => {
          events.push(`close:${base}`)
          closeStarted.resolve(true)
          await closeReleased.promise
        },
      },
      {
        heartbeat: immediate,
        save: async () => {
          events.push("upload")
          started.resolve()
          return uploaded.promise
        },
        done: async () => {
          events.push("done")
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    const saving = sync.save()
    await started.promise
    expect(sync.snapshot().status).toBe("saving")
    expect(sync.snapshot().acknowledgement).toBeUndefined()
    const completing = sync.done()
    expect(events).toEqual(["word-save", "export", "upload"])
    uploaded.resolve(acknowledgement("v1"))
    expect(await saving).toBe(true)
    expect(await Promise.race([closeStarted.promise, completing.then(() => false)])).toBe(true)
    expect(await sync.tick()).toBe(false)
    closeReleased.resolve()
    expect(await completing).toBe(true)
    expect(events).toEqual([
      "word-save",
      "export",
      "upload",
      "export",
      "word-save",
      "export",
      "upload",
      "export",
      "done",
      "export",
      "close:v1",
    ])
    expect(sync.snapshot().status).toBe("done")
    const completedEvents = [...events]
    await sync.tick()
    expect(events).toEqual(completedEvents)
  })

  test("a lost response retries exactly the same bytes and base, then advances parent", async () => {
    const requests: { bytes: number; base: string }[] = []
    let value = 1
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([value]),
      },
      {
        heartbeat: immediate,
        done: immediate,
        save: async (bytes, base) => {
          requests.push({ bytes: bytes[0]!, base })
          if (requests.length <= 2) throw new ConnectionError("Offline", true)
          return acknowledgement(`v${requests.length}`)
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    expect(await sync.save()).toBe(false)
    expect(requests).toEqual([
      { bytes: 1, base: "base" },
      { bytes: 1, base: "base" },
    ])
    expect(sync.snapshot().status).toBe("not-saved")
    value = 2
    expect(await sync.save()).toBe(true)
    expect(requests.slice(2)).toEqual([
      { bytes: 1, base: "base" },
      { bytes: 2, base: "v3" },
    ])
    expect(sync.snapshot().acknowledgement?.base).toBe("v4")
  })

  test("background sync never exports unsaved changes, including changes during export", async () => {
    let saved = false
    let exports = 0
    let uploads = 0
    let race = false
    const sync = new WordSync(
      "base",
      {
        saved: async () => saved,
        save: async () => { saved = true },
        export: async () => {
          exports++
          if (race) saved = false
          return new Uint8Array([1])
        },
      },
      {
        heartbeat: immediate,
        done: immediate,
        save: async () => {
          uploads++
          return acknowledgement("v1")
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    await sync.start()
    expect(exports).toBe(0)
    expect(await sync.save()).toBe(true)
    exports = 0; uploads = 0; race = true
    await sync.refresh()
    expect(exports).toBe(1)
    expect(uploads).toBe(0)
    expect(sync.snapshot().status).toBe("not-saved")
  })

  test("explicitly enabled sync captures changed same-minute saved content without timestamps", async () => {
    let bytes = 1
    const bases: string[] = []
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([bytes]),
      },
      {
        heartbeat: immediate,
        done: immediate,
        save: async (_bytes, base) => {
          bases.push(base)
          return acknowledgement(`v${bases.length}`)
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    await sync.start()
    expect(bases).toEqual([])
    await sync.refresh()
    await sync.tick()
    expect(bases).toEqual([])
    expect(await sync.save()).toBe(true)
    expect(bases).toEqual(["base"])
    bytes = 2
    await sync.tick()
    expect(bases).toEqual(["base", "v1"])
  })

  test("reopen reuses acknowledged base only when this saved file matches", async () => {
    const previous: Acknowledgement = { base: "ack", fingerprint: "1", number: 2, at: 1, overlap: false }
    for (const bytes of [1, 2]) {
      const requests: string[] = []
      const sync = new WordSync(
        "original",
        {
          saved: async () => true,
          save: immediate,
          export: async () => new Uint8Array([bytes]),
        },
        {
          heartbeat: immediate,
          done: immediate,
          save: async (_bytes, base) => {
            requests.push(base)
            return acknowledgement("new")
          },
        },
        checkpoints(previous),
        () => {},
        fingerprint,
        immediate,
      )
      await sync.start()
      expect(requests).toEqual([])
      expect(sync.snapshot().acknowledgement?.base).toBe(bytes === 1 ? "ack" : undefined)
      await sync.save()
      expect(requests).toEqual([bytes === 1 ? "ack" : "original"])
    }
  })

  test("initial dirty export restores only an exact checkpoint before later editing and Done", async () => {
    for (const matches of [true, false]) {
      const previous: Acknowledgement = { base: "ack-v4", fingerprint: "1", number: 4, at: 1, overlap: false }
      let bytes = matches ? 1 : 9
      let saved = false
      const bases: string[] = []
      const statuses: string[] = []
      const events: string[] = []
      const sync = new WordSync(
        "original-v1",
        {
          saved: async () => saved,
          save: async () => {
            saved = true
          },
          export: async () => {
            events.push("export")
            return new Uint8Array([bytes])
          },
        },
        {
          heartbeat: async () => {
            events.push("heartbeat")
          },
          done: immediate,
          save: async (_bytes, base) => {
            bases.push(base)
            return acknowledgement("new-v5", 5)
          },
        },
        checkpoints(previous),
        (state) => {
          statuses.push(state.status)
        },
        fingerprint,
        immediate,
      )
      await sync.start()
      expect(events).toEqual(["export"])
      expect(bases).toEqual([])
      expect(statuses).not.toContain("saved")
      expect(sync.snapshot().status).toBe("not-saved")
      expect(sync.snapshot().acknowledgement?.base).toBe(matches ? "ack-v4" : undefined)
      bytes = 2
      await sync.tick()
      expect(bases).toEqual([])
      expect(sync.snapshot().status).toBe("not-saved")
      expect(await sync.done()).toBe(true)
      expect(bases).toEqual([matches ? "ack-v4" : "original-v1"])
      expect(sync.snapshot().acknowledgement?.fingerprint).toBe("2")
    }
  })

  test("failed Done leaves heartbeat live and does not report completion", async () => {
    let heartbeats = 0
    let closes = 0
    let time = 0
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([1]),
        close: async () => {
          closes++
        },
      },
      {
        heartbeat: async () => {
          heartbeats++
        },
        save: async () => acknowledgement("v1"),
        done: async () => {
          throw new ConnectionError("Offline", true)
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
      () => time,
    )
    await sync.start()
    expect(await sync.done()).toBe(false)
    expect(sync.snapshot().status).toBe("not-saved")
    expect(closes).toBe(0)
    time = 35_000
    await sync.tick()
    expect(heartbeats).toBe(2)
  })

  test("Done cannot discard new changes made while the saved snapshot is uploading", async () => {
    let bytes = 1
    let done = 0
    let closes = 0
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([bytes]),
        close: async () => {
          closes++
        },
      },
      {
        heartbeat: immediate,
        save: async () => {
          bytes = 2
          return acknowledgement("v1")
        },
        done: async () => {
          done++
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    expect(await sync.done()).toBe(false)
    expect(done).toBe(0)
    expect(closes).toBe(0)
    expect(sync.snapshot().status).toBe("not-saved")
  })

  test("an explicit save requires a fresh server acknowledgement even for unchanged content", async () => {
    let available = true
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([1]),
      },
      {
        heartbeat: immediate,
        done: immediate,
        save: async () => {
          if (!available) throw new ConnectionError("Offline", true)
          return acknowledgement("v1")
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    expect(await sync.save()).toBe(true)
    expect(sync.snapshot().status).toBe("saved")
    available = false
    expect(await sync.save()).toBe(false)
    expect(sync.snapshot().status).toBe("not-saved")
  })

  test("Save and background capture do not confirm an old snapshot when Word saved newer edits during upload", async () => {
    for (const explicit of [true, false]) {
      let bytes = 1
      let race = false
      let uploads = 0
      const published: string[] = []
      const sync = new WordSync(
        "base",
        {
          saved: async () => true,
          save: immediate,
          export: async () => new Uint8Array([bytes]),
        },
        {
          heartbeat: immediate,
          done: immediate,
          save: async () => {
            uploads++
            if (race) bytes = 3
            return acknowledgement(race ? "v2" : "v1")
          },
        },
        checkpoints(),
        (state) => {
          published.push(state.status)
        },
        fingerprint,
        immediate,
      )
      expect(await sync.save()).toBe(true)
      published.length = 0; bytes = 2; race = true
      if (explicit) expect(await sync.save()).toBe(false)
      else await sync.refresh()
      expect(sync.snapshot().acknowledgement?.fingerprint).toBe("2")
      expect(sync.snapshot().status).toBe("not-saved")
      expect(published.slice(published.indexOf("saving"))).not.toContain("saved")
      if (explicit) {
        await sync.tick()
        expect(uploads).toBe(2)
      }
    }
  })

  test("retrying old pending bytes cannot publish Saved while the newer document is still being exported", async () => {
    let bytes = 1
    let available = false
    let exports = 0
    const statuses: string[] = []
    const secondExportStarted = deferred<void>()
    const secondExportReleased = deferred<void>()
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => {
          exports++
          if (exports === 2) {
            secondExportStarted.resolve()
            await secondExportReleased.promise
          }
          return new Uint8Array([bytes])
        },
      },
      {
        heartbeat: immediate,
        done: immediate,
        save: async () => {
          if (!available) throw new ConnectionError("Offline", true)
          return acknowledgement(`v${bytes}`)
        },
      },
      checkpoints(),
      (state) => {
        statuses.push(state.status)
      },
      fingerprint,
      immediate,
    )
    expect(await sync.save()).toBe(false)
    expect(sync.snapshot().connectionFailed).toBe(true)
    statuses.length = 0
    bytes = 2
    available = true
    const retried = sync.save()
    await secondExportStarted.promise
    expect(sync.snapshot().status).toBe("saving")
    expect(statuses).not.toContain("saved")
    secondExportReleased.resolve()
    expect(await retried).toBe(true)
    expect(sync.snapshot().acknowledgement?.fingerprint).toBe("2")
    expect(sync.snapshot().status).toBe("saved")
    expect(sync.snapshot().connectionFailed).toBeUndefined()
  })

  test("Done does not close or reclaim its released lease if edits arrive during the request", async () => {
    for (const nativeSaved of [false, true]) {
      let bytes = 1
      let saved = true
      let heartbeats = 0
      let closes = 0
      const statuses: string[] = []
      const sync = new WordSync(
        "base",
        {
          saved: async () => saved,
          save: immediate,
          export: async () => new Uint8Array([bytes]),
          close: async () => {
            closes++
          },
        },
        {
          heartbeat: async () => {
            heartbeats++
          },
          save: async () => acknowledgement("v1"),
          done: async () => {
            bytes = 2
            saved = nativeSaved
            statuses.length = 0
          },
        },
        checkpoints(),
        (state) => {
          statuses.push(state.status)
        },
        fingerprint,
        immediate,
      )
      expect(await sync.done()).toBe(false)
      expect(heartbeats).toBe(0)
      expect(closes).toBe(0)
      expect(sync.snapshot().status).toBe("not-saved")
      expect(statuses).not.toContain("done")
      expect(statuses).not.toContain("saved")
    }
  })

  test("Done never closes if the document upload was not acknowledged", async () => {
    let closes = 0
    let completions = 0
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([1]),
        close: async () => {
          closes++
        },
      },
      {
        heartbeat: immediate,
        save: async () => {
          throw new ConnectionError("Offline", true)
        },
        done: async () => {
          completions++
        },
      },
      checkpoints(),
      () => {},
      fingerprint,
      immediate,
    )
    expect(await sync.done()).toBe(false)
    expect(sync.snapshot().status).toBe("not-saved")
    expect(closes).toBe(0)
    expect(completions).toBe(0)
  })

  test("unsupported or rejected close keeps a confirmed save and provides a manual close fallback", async () => {
    for (const supported of [false, true]) {
      const sync = new WordSync(
        "base",
        {
          saved: async () => true,
          save: immediate,
          export: async () => new Uint8Array([1]),
          ...(supported
            ? {
                close: async () => {
                  throw new Error("private host path")
                },
              }
            : {}),
        },
        { heartbeat: immediate, save: async () => acknowledgement("v1"), done: immediate },
        checkpoints(),
        () => {},
        fingerprint,
        immediate,
      )
      expect(await sync.done()).toBe(true)
      expect(sync.snapshot().status).toBe("done")
      expect(sync.snapshot().acknowledgement?.base).toBe("v1")
      expect(sync.snapshot().message).toContain("Close it in Word")
      expect(sync.snapshot().message).not.toContain("private")
      expect(sync.snapshot().busy).toBe(false)
      await sync.checkStatus()
      expect(sync.snapshot().status).toBe("done")
      expect(sync.snapshot().message).toContain("Close it in Word")
    }
  })

  test("close-time edits or an unverifiable context do not retain Saved or reclaim a released lease", async () => {
    for (const outcome of ["unsaved", "native-saved", "declined", "unavailable"]) {
      let bytes = 1
      let saved = true
      let heartbeats = 0
      let closing = false
      const sync = new WordSync(
        "base",
        {
          saved: async () => {
            if (closing && outcome === "unavailable") throw new Error("Word context unavailable")
            return saved
          },
          save: immediate,
          export: async () => new Uint8Array([bytes]),
          close: async () => {
            closing = true
            bytes = 2
            saved = outcome === "native-saved"
            if (outcome === "declined") throw new DocumentStateError("New changes")
            throw new Error("Close failed")
          },
        },
        {
          heartbeat: async () => {
            heartbeats++
          },
          save: async () => acknowledgement("v1"),
          done: immediate,
        },
        checkpoints(),
        () => {},
        fingerprint,
        immediate,
      )
      expect(await sync.done()).toBe(false)
      expect(sync.snapshot().status).toBe("not-saved")
      expect(sync.snapshot().message).toContain("last save is in doc.haus")
      expect(heartbeats).toBe(0)
    }
  })

  test("editing after Done requires an explicit claim before new changes can save", async () => {
    for (const nativeSaved of [false, true]) {
      let bytes = 1
      let saved = true
      let time = 0
      let heartbeats = 0
      const sync = new WordSync(
        "base",
        {
          saved: async () => saved,
          save: async () => {
            saved = true
          },
          export: async () => new Uint8Array([bytes]),
        },
        {
          heartbeat: async () => {
            heartbeats++
          },
          done: immediate,
          save: async () => acknowledgement(`v${bytes}`),
        },
        checkpoints(),
        () => {},
        fingerprint,
        immediate,
        () => time,
      )
      expect(await sync.done()).toBe(true)
      expect(sync.snapshot().status).toBe("done")
      bytes = 2
      saved = nativeSaved
      time = nativeSaved ? 20_000 : 5_000
      await sync.tick()
      expect(heartbeats).toBe(0)
      expect(sync.snapshot().status).toBe("not-saved")
      expect(sync.snapshot().busy).toBe(false)
      expect(await sync.save()).toBe(true)
      expect(sync.snapshot().acknowledgement?.fingerprint).toBe("2")
    }
  })
})

test("browser checkpoint key separates original hand-outs and stores no capability", async () => {
  const values = new Map<string, string>()
  const storage = {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => {
      values.set(key, value)
    },
  } as Storage
  const link = {
    version: 1 as const,
    apiOrigin: "https://localhost:4612",
    origin: "http://localhost:4444",
    matter: "synthetic",
    matterName: "Synthetic matter",
    document: "sample.docx",
    documentName: "Sample",
    token: "synthetic-test-capability",
    baseVersion: "original",
  }
  const store = await browserCheckpoints(link, storage)
  const value: Acknowledgement = { base: "ack", fingerprint: "a".repeat(64), number: 1, at: 123, overlap: false }
  store.write(value)
  expect(store.read()).toEqual(value)
  expect(JSON.stringify([...values])).not.toContain(link.token)
  expect((await browserCheckpoints({ ...link, baseVersion: "other" }, storage)).read()).toBeNull()
})

test("unchanged background polling stays quiet while a queued explicit save stays busy and serialized", async () => {
  const pollStarted = deferred<void>(), pollReleased = deferred<void>()
  const uploadStarted = deferred<void>(), uploaded = deferred<WordSaveResponse>()
  const states: PaneState[] = []
  let saves = 0
  const sync = new WordSync("base", {
    saved: async () => true,
    save: async () => { saves++ },
    export: async () => new Uint8Array([1]),
  }, {
    heartbeat: async () => { pollStarted.resolve(); await pollReleased.promise },
    save: async () => {
      if (saves === 1) return acknowledgement("v1")
      uploadStarted.resolve()
      return uploaded.promise
    },
    done: immediate,
  }, checkpoints(), state => states.push(state), fingerprint, immediate)
  await sync.save()
  states.length = 0
  const polling = sync.tick()
  try {
    await pollStarted.promise
    expect(states.every(state => !state.busy)).toBe(true)
    expect(sync.snapshot().busy).toBe(false)
    expect(await sync.tick()).toBe(false)
    const saving = sync.save()
    const queued = states.length - 1
    expect(sync.snapshot().busy).toBe(true)
    expect(saves).toBe(1)
    pollReleased.resolve()
    await uploadStarted.promise
    expect(await polling).toBe(true)
    expect(states.slice(queued).every(state => state.busy)).toBe(true)
    expect(sync.snapshot().status).toBe("saving")
    uploaded.resolve(acknowledgement("v2"))
    expect(await saving).toBe(true)
    expect(sync.snapshot()).toMatchObject({ busy: false, status: "saved", content: "confirmed", acknowledgement: { base: "v2" } })
  } finally { pollReleased.resolve(); uploaded.resolve(acknowledgement("v2")); await polling }
})

test("a real background upload still shows Saving and remains busy until confirmation", async () => {
  const uploadStarted = deferred<void>(), uploaded = deferred<WordSaveResponse>()
  let bytes = 1
  const sync = new WordSync("base", {
    saved: async () => true, save: immediate, export: async () => new Uint8Array([bytes]),
  }, {
    heartbeat: immediate,
    save: async () => { if (bytes === 1) return acknowledgement("v1"); uploadStarted.resolve(); return uploaded.promise },
    done: immediate,
  }, checkpoints(), () => {}, fingerprint, immediate)
  await sync.save()
  bytes = 2
  const polling = sync.tick()
  try {
    await uploadStarted.promise
    expect(sync.snapshot()).toMatchObject({ busy: true, status: "saving" })
    expect(await sync.tick()).toBe(false)
    uploaded.resolve(acknowledgement("v2"))
    expect(await polling).toBe(true)
    expect(sync.snapshot()).toMatchObject({ busy: false, status: "saved", content: "confirmed", acknowledgement: { base: "v2" } })
  } finally { uploaded.resolve(acknowledgement("v2")); await polling }
})


test("rehydrated reviewed edits require explicit Save across lease ownership states", async () => {
  for (const outcome of ["success", "uncertain"] as const) for (const next of ["owned", "available", "other"] as const) {
    let owner: "owned" | "available" | "other" = "owned", bytes = 1, version = "base", uploads = 0, nativeSaves = 0
    let operation: { sequence: number; operationId: string | null; state: "idle" | "pending" | "confirmed" | "not-dispatched" } = { sequence: 0, operationId: null, state: "idle" }
    const current = (): WordConnectionState => ({ nativeOperation: { sequence: operation.sequence, state: operation.state }, version: { base: {id: version, number: 1}, current: {id: version, number: 1}, state: "current" }, lease: owner === "available" && operation.state !== "pending" ? { state: "available", expiresAt: null } : { state: owner === "owned" ? "owned" : "other", generation: "g1", expiresAt: 90000 } })
    const host = { saved: async () => true, save: async () => { nativeSaves++ }, export: async () => new Uint8Array([bytes]) }
    const transport = {
      lease: async (action: string) => { if ((action === "claim" || action === "takeover") && operation.state !== "pending") owner = "owned"; return current() },
      nativeOperation: async (_base: string, _generation: string, input: { action: "read" } | { action: "begin"; sequence: number; operationId: string } | { action: "settle"; sequence: number; operationId: string; outcome: "confirmed" | "not-dispatched" }) => {
        if (input.action === "begin") operation = { sequence: input.sequence + 1, operationId: input.operationId, state: "pending" }
        if (input.action === "settle") operation = { sequence: input.sequence, operationId: input.operationId, state: input.outcome }
        return { operation, connection: current() }
      },
      heartbeat: async () => current(), done: async () => current(),
      save: async () => { uploads++; version = `v${uploads}`; return { ...acknowledgement(version), connection: current() } },
    }
    const make = () => new CheckedWordSync("base", host, transport, checkpoints(), () => {}, fingerprint, immediate)
    const first = make()
    await first.start()
    const edit = first.localEdit({ conversationId: "conversation", turnId: "turn" }, async () => { bytes = 2; if (outcome === "uncertain") throw new Error("lost response") })
    if (outcome === "uncertain") await expect(edit).rejects.toThrow("lost response")
    else await edit
    await first.tick()
    expect(uploads).toBe(0)
    owner = next
    const reopened = make()
    await reopened.checkStatus()
    if (outcome === "success" && next === "available") expect(await reopened.startEditing()).toBe(true)
    if (outcome === "success" && next === "other") expect(await reopened.takeover()).toBe(true)
    await reopened.tick()
    expect(uploads).toBe(0)
    expect(nativeSaves).toBe(0)
    if (outcome === "uncertain") {
      expect(reopened.snapshot()).toMatchObject({status: "not-saved", canSave: false})
      expect(await reopened.save()).toBe(false)
      expect(uploads).toBe(0); expect(nativeSaves).toBe(0)
      continue
    }
    expect(reopened.snapshot()).toMatchObject({status: "not-saved", canSave: true})
    expect(await reopened.save()).toBe(true)
    expect(uploads).toBe(1)
    expect(nativeSaves).toBe(1)
    bytes = 3
    await reopened.tick()
    expect(uploads).toBe(2)
  }
})

test("a new pane stays neutral and never publishes an unverified saved buffer", async () => {
  let uploads = 0
  const sync = new WordSync("base", {saved: async () => true, save: immediate, export: async () => new Uint8Array([1])},
    {heartbeat: immediate, done: immediate, save: async () => { uploads++; return acknowledgement("v1") }}, checkpoints(), () => {}, fingerprint, immediate)
  await sync.start()
  await sync.tick()
  expect(uploads).toBe(0)
  expect(sync.snapshot()).toMatchObject({status: "not-saved", content: "unverified", canSave: true, message: "Ready to save. Choose Save to send this document."})
})

test("a failed native Save after retrying pending bytes cannot publish newer content in the background", async () => {
  let bytes = 1, uploads = 0, nativeSaves = 0, rejectUpload = false, failNative = false, clock = 0
  const sync = new WordSync("base", {
    saved: async () => true,
    save: async () => { nativeSaves++; if (failNative) throw new Error("native save failed") },
    export: async () => new Uint8Array([bytes]),
  }, {
    heartbeat: immediate, done: immediate,
    save: async () => {
      uploads++
      if (rejectUpload) throw new ConnectionError("offline")
      return acknowledgement(`v${uploads}`)
    },
  }, checkpoints(), () => {}, fingerprint, immediate, () => clock)
  expect(await sync.save()).toBe(true)
  bytes = 2; rejectUpload = true
  expect(await sync.save()).toBe(false)
  rejectUpload = false; failNative = true
  expect(await sync.save()).toBe(false)
  const afterPendingAcknowledged = uploads
  expect(sync.snapshot().acknowledgement?.fingerprint).toBe("2")
  bytes = 3; clock += 60_000
  await sync.tick()
  expect(uploads).toBe(afterPendingAcknowledged)
  expect(sync.snapshot().status).toBe("not-saved")
  failNative = false
  expect(await sync.save()).toBe(true)
  expect(uploads).toBe(afterPendingAcknowledged + 1)
  expect(nativeSaves).toBe(4)
})
