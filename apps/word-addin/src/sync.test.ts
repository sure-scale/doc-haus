import { describe, expect, test } from "bun:test"
import { ConnectionError } from "./client"
import { WordSync, browserCheckpoints, type Acknowledgement, type Checkpoints } from "./sync"
import type { WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"

const acknowledgement = (id: string, number = 1): WordSaveResponse => ({
  version: { id, number, parent: null, created: true },
  overlap: false,
})
const checkpoints = (initial: Acknowledgement | null = null): Checkpoints => ({
  read: () => initial,
  write: () => {},
})
const fingerprint = async (bytes: Uint8Array) => String(bytes[0])
const immediate = async () => {}

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
    const sync = new WordSync(
      "base",
      {
        saved: async () => saved,
        save: immediate,
        export: async () => {
          exports++
          saved = false
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
    saved = true
    await sync.tick()
    expect(exports).toBe(1)
    expect(uploads).toBe(0)
    expect(sync.snapshot().status).toBe("not-saved")
  })

  test("initial saved state is captured and changed same-minute content is detected without timestamps", async () => {
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
    await sync.tick()
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
      expect(requests).toEqual(bytes === 1 ? [] : ["original"])
      expect(sync.snapshot().acknowledgement?.base).toBe(bytes === 1 ? "ack" : "new")
    }
  })

  test("failed Done leaves heartbeat live and does not report completion", async () => {
    let heartbeats = 0
    let time = 0
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([1]),
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
    time = 35_000
    await sync.tick()
    expect(heartbeats).toBe(3)
  })

  test("Done cannot discard new changes made while the saved snapshot is uploading", async () => {
    let bytes = 1
    let done = 0
    const sync = new WordSync(
      "base",
      {
        saved: async () => true,
        save: immediate,
        export: async () => new Uint8Array([bytes]),
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
    expect(await sync.start()).toBe(true)
    expect(sync.snapshot().status).toBe("saved")
    available = false
    expect(await sync.save()).toBe(false)
    expect(sync.snapshot().status).toBe("not-saved")
  })

  test("Save and background capture do not confirm an old snapshot when Word saved newer edits during upload", async () => {
    for (const explicit of [true, false]) {
      let bytes = 1
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
            bytes = 2
            return acknowledgement("v1")
          },
        },
        checkpoints(),
        (state) => {
          published.push(state.status)
        },
        fingerprint,
        immediate,
      )
      if (explicit) expect(await sync.save()).toBe(false)
      else await sync.start()
      expect(sync.snapshot().acknowledgement?.fingerprint).toBe("1")
      expect(sync.snapshot().status).toBe("not-saved")
      expect(published).not.toContain("saved")
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
  })

  test("Done restores heartbeat if edits arrive during the done request, including native-saved edits", async () => {
    for (const nativeSaved of [false, true]) {
      let bytes = 1
      let saved = true
      let heartbeats = 0
      const statuses: string[] = []
      const sync = new WordSync(
        "base",
        {
          saved: async () => saved,
          save: immediate,
          export: async () => new Uint8Array([bytes]),
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
      expect(heartbeats).toBe(1)
      expect(sync.snapshot().status).toBe("not-saved")
      expect(statuses).not.toContain("done")
      expect(statuses).not.toContain("saved")
    }
  })

  test("editing after Done resumes the heartbeat and saving controls for unsaved and native-saved changes", async () => {
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
      expect(heartbeats).toBe(1)
      expect(sync.snapshot().status).toBe(nativeSaved ? "saved" : "not-saved")
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
