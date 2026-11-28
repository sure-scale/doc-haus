import { expect, test } from "bun:test"
import { zipSync, strToU8, strFromU8, unzipSync } from "fflate"
import { documentFingerprint } from "./fingerprint"
import { connectOffice, DocumentStateError, exportFile, type OfficeFile } from "./office"
import { WordSync, type Transport } from "./sync"
import type { WordConnectionState } from "../../../dochaus/lib/word-addin-contract"

test("sliced export assembles the file and always closes its handle on success or failure", async () => {
  for (const fails of [false, true]) {
    let closed = 0
    const file: OfficeFile = {
      size: 3,
      sliceCount: 2,
      getSliceAsync: (index, callback) =>
        callback({ status: fails ? "failed" : "succeeded", value: { data: index === 0 ? [1, 2] : [3] } }),
      closeAsync: (callback) => {
        closed++
        callback({ status: "succeeded", value: undefined })
      },
    }
    if (fails) await expect(exportFile(file)).rejects.toThrow("could not read")
    else expect(await exportFile(file)).toEqual(new Uint8Array([1, 2, 3]))
    expect(closed).toBe(1)
  }
})

const linkedDocument = {
  version: 1,
  apiOrigin: "https://localhost:4612",
  origin: "http://localhost:4444",
  matter: "synthetic-matter",
  matterName: "Synthetic matter",
  document: "draft.docx",
  documentName: "Draft",
  baseVersion: "original-v1",
  token: "synthetic-capability",
  extension: { retained: true },
}

function installCloseHost(supported: boolean, encoded = false, saveMode: "saved" | "dirty" | "failed" = "saved") {
  const events: string[] = []
  const exported = zipSync({"word/document.xml":strToU8("<document/>")})
  let nativeSaved = true
  let raw: unknown = encoded ? JSON.stringify(linkedDocument) : { ...linkedDocument }
  let callback: ((result: { status: string; value: void }) => void) | undefined
  let started!: () => void
  const saving = new Promise<void>((resolve) => {
    started = resolve
  })
  const context = {
    document: {
      get saved() {
        return nativeSaved
      },
      set saved(value: boolean) {
        nativeSaved = value
      },
      load: (property: string) => {
        events.push(`load:${property}`)
      },
      save: (behavior?: string) => {
        events.push(`save:${behavior}`)
        if (saveMode === "failed") throw new Error("Synthetic native save failure")
        nativeSaved = saveMode === "saved"
      },
      close: (behavior: string) => {
        expect(nativeSaved).toBe(true)
        events.push(`close:${behavior}`)
      },
    },
    sync: async () => {
      events.push("sync")
    },
  }
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value: {
      location: { origin: linkedDocument.apiOrigin },
      Office: {
        onReady: async () => ({ host: "Word" }),
        context: {
          document: {
            getFileAsync: (_type: unknown, _options: unknown, callback: (value: unknown) => void) => callback({status:"succeeded",value:{size:exported.length,sliceCount:1,getSliceAsync:(_index:number, done:(value:unknown)=>void) => done({status:"succeeded",value:{data:[...exported]}}),closeAsync:(done:(value:unknown)=>void)=>done({status:"succeeded",value:undefined})}}),
            settings: {
              get: () => raw,
              set: (key: string, value: unknown) => {
                expect(key).toBe("dochaus.link")
                events.push("settings-set")
                raw = value
                nativeSaved = false
              },
              saveAsync: (value: typeof callback) => {
                events.push("settings-save")
                callback = value
                started()
              },
            },
          },
          requirements: {
            isSetSupported: (name: string, version: string) =>
              name === "WordApi" && (version === "1.1" || (version === "1.5" && supported)),
          },
        },
        FileType: {Compressed:"compressed"},
      },
      Word: { run: async (action: (value: typeof context) => Promise<unknown>) => action(context) },
    },
  })
  return {
    events,
    exported,
    context,
    saving,
    get raw() {
      return raw
    },
    set raw(value: unknown) {
      raw = value
    },
    finish: (status = "succeeded") => callback!({ status, value: undefined }),
  }
}

test("Done persists only its acknowledged base before Save-close, preserving link encoding and WordApi 1.1", async () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    const unsupported = installCloseHost(false)
    const oldHost = (await connectOffice()).host
    expect(oldHost.close).toBeUndefined()
    expect(await oldHost.saved()).toBe(true)
    expect(unsupported.events).toEqual(["load:saved", "sync"])
    for (const encoded of [false, true]) {
      const fixture = installCloseHost(true, encoded)
      const { host } = await connectOffice()
      const closing = host.close!("ack-v2")
      await fixture.saving
      expect(fixture.events).toEqual(["load:saved", "sync", "settings-set", "settings-save"])
      expect(fixture.context.document.saved).toBe(false)
      expect(typeof fixture.raw).toBe(encoded ? "string" : "object")
      expect(encoded ? JSON.parse(fixture.raw as string) : fixture.raw).toEqual({
        ...linkedDocument,
        baseVersion: "ack-v2",
      })
      fixture.finish()
      await closing
      expect(fixture.events).toEqual([
        "load:saved",
        "sync",
        "settings-set",
        "settings-save",
        "save:Save",
        "sync",
        "load:saved",
        "sync",
        "close:Save",
        "sync",
      ])
      fixture.events.length = 0
      fixture.context.document.saved = false
      await expect(host.close!("ack-v3")).rejects.toBeInstanceOf(DocumentStateError)
      expect(fixture.events).toEqual(["load:saved", "sync"])
    }
  } finally {
    if (originalWindow) Object.defineProperty(globalThis, "window", originalWindow)
    else delete (globalThis as { window?: unknown }).window
  }
})

test("acknowledged base persists in the same linked file and refuses identity or persistence failures", async () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    for (const encoded of [false,true]) {
      const f = installCloseHost(true, encoded), host = (await connectOffice()).host
      const fingerprint = await documentFingerprint(f.exported)
      const operation = host.persistAcknowledgedBase!("ack-v2","original-v1",fingerprint)
      await f.saving; f.finish(); await operation
      expect(encoded ? JSON.parse(f.raw as string) : f.raw).toEqual({...linkedDocument,baseVersion:"ack-v2"})
      expect(f.events).toContain("save:Save")
      expect(f.events).not.toContain("close:Save")
    }
    const changed = installCloseHost(true), host = (await connectOffice()).host
    changed.raw = {...linkedDocument,token:"other"}
    await expect(host.persistAcknowledgedBase!("ack-v2","original-v1",await documentFingerprint(changed.exported))).rejects.toThrow("link changed")
    expect(changed.events).not.toContain("settings-set")
    const failure = installCloseHost(true), failingHost = (await connectOffice()).host
    const operation = failingHost.persistAcknowledgedBase!("ack-v2","original-v1",await documentFingerprint(failure.exported))
    await failure.saving; failure.finish("failed")
    await expect(operation).rejects.toThrow("Saved to doc.haus")
    expect(failure.events).not.toContain("save:Save")
  } finally { if (originalWindow) Object.defineProperty(globalThis,"window",originalWindow); else delete (globalThis as {window?:unknown}).window }
})

test("Done cannot stamp another document's link or close after settings persistence fails", async () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    for (const changed of [
      null,
      "invalid JSON",
      { ...linkedDocument, token: "another-synthetic-capability" },
      { ...linkedDocument, matter: "another-matter" },
      { ...linkedDocument, document: "another.docx" },
      { ...linkedDocument, origin: "http://localhost:4445" },
      { ...linkedDocument, apiOrigin: "https://localhost:4613" },
    ]) {
      const fixture = installCloseHost(true)
      const { host } = await connectOffice()
      fixture.raw = changed
      await expect(host.close!("ack-v2")).rejects.toThrow("could not update")
      expect(fixture.events).toEqual(["load:saved", "sync"])
      expect(fixture.raw).toEqual(changed)
    }
    const fixture = installCloseHost(true)
    const { host } = await connectOffice()
    const closing = host.close!("ack-v2")
    await fixture.saving
    fixture.finish("failed")
    await expect(closing).rejects.toThrow("could not update")
    expect(fixture.events).toEqual(["load:saved", "sync", "settings-set", "settings-save"])
  } finally {
    if (originalWindow) Object.defineProperty(globalThis, "window", originalWindow)
    else delete (globalThis as { window?: unknown }).window
  }
})

test("Done does not close if saving its settings fails or Word remains dirty", async () => {
  const originalWindow = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    for (const mode of ["failed", "dirty"] as const) {
      const fixture = installCloseHost(true, false, mode)
      const { host } = await connectOffice()
      const closing = host.close!("ack-v2")
      await fixture.saving
      fixture.finish()
      await expect(closing).rejects.toThrow(mode === "failed" ? "native save failure" : "new changes")
      expect(fixture.events).not.toContain("close:Save")
      expect(fixture.events).toEqual([
        "load:saved",
        "sync",
        "settings-set",
        "settings-save",
        "save:Save",
        ...(mode === "dirty" ? ["sync", "load:saved", "sync"] : []),
      ])
    }
  } finally {
    if (originalWindow) Object.defineProperty(globalThis, "window", originalWindow)
    else delete (globalThis as { window?: unknown }).window
  }
})

function installRefreshHost(options: { supported?: boolean; encoded?: boolean; failInsert?: boolean; failSettings?: boolean; exportBytes?: Uint8Array; failExport?: boolean } = {}) {
  const events: string[] = []
  let raw: unknown = options.encoded ? JSON.stringify(linkedDocument) : { ...linkedDocument }
  let tracking = "TrackAll"
  let imported: { bytes: string; location: string; options: Record<string, boolean> } | undefined
  let importedBytes: Uint8Array | undefined
  const context = {
    document: {
      saved: true,
      get changeTrackingMode() { return tracking },
      set changeTrackingMode(value: string) { tracking = value; events.push(`tracking:${value}`) },
      load: (value: string) => { events.push(`load:${value}`) },
      save: () => { events.push("native-save") },
      close: () => { events.push("native-close") },
      insertFileFromBase64: (bytes: string, location: string, optionsValue: Record<string, boolean>) => {
        expect(tracking).toBe("Off")
        events.push("document-replace")
        imported = { bytes, location, options: optionsValue }
        importedBytes = Uint8Array.from(atob(bytes), character => character.charCodeAt(0))
        if (options.failInsert) throw new Error("Synthetic unsupported native package")
        tracking = "TrackMineOnly" // Simulate the imported source tracking mode.
        raw = null // Native insertion does not preserve document settings.
      },
    },
    sync: async () => { events.push("sync") },
  }
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value: {
      location: { origin: linkedDocument.apiOrigin },
      Office: {
        onReady: async () => ({ host: "Word" }),
        FileType: { Compressed: "compressed" },
        context: {
          requirements: { isSetSupported: (name: string, version: string) => name === "WordApi" && (version !== "1.7" || options.supported !== false) },
          document: { getFileAsync: (_type: string, _options: unknown, callback: (result: unknown) => void) => {
            events.push("get-file")
            if (options.failExport) return callback({ status: "failed" })
            const bytes = importedBytes ? options.exportBytes ?? importedBytes : latestRefreshBytes()
            callback({ status: "succeeded", value: { size: bytes.length, sliceCount: 1,
              getSliceAsync: (_index: number, done: (result: unknown) => void) => { events.push("read-slice"); done({ status: "succeeded", value: { data: Array.from(bytes) } }) },
              closeAsync: (done: (result: unknown) => void) => { events.push("close-file"); done({ status: "succeeded", value: undefined }) },
            } })
          }, settings: {
            get: () => raw,
            set: (_key: string, value: unknown) => { events.push("settings-set"); raw = value },
            saveAsync: (callback: (value: { status: string; value: undefined }) => void) => {
              events.push("settings-save")
              callback({ status: options.failSettings ? "failed" : "succeeded", value: undefined })
            },
          } },
        },
      },
      Word: { run: async (action: (value: typeof context) => Promise<unknown>) => action(context) },
    },
  })
  return { events, context, get raw() { return raw }, set raw(value: unknown) { raw = value }, get imported() { return imported } }
}

const latestRefreshBytes = () => zipSync({
  "word/document.xml": strToU8('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Reviewed text</w:t></w:r><w:ins w:id="1" w:author="Earlier reviewer"><w:r><w:t> added</w:t></w:r></w:ins></w:p><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:body></w:document>'),
  "word/comments.xml": strToU8('<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1"><w:p><w:r><w:t>Retained comment</w:t></w:r></w:p></w:comment></w:comments>'),
  "word/numbering.xml": strToU8('<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="1"/></w:numbering>'),
})

test("refresh uses full-document Replace with all fidelity options, retains the link and imports source tracking without saving", async () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    for (const encoded of [false, true]) {
      const f = installRefreshHost({ encoded })
      const { host } = await connectOffice()
      const bytes = latestRefreshBytes()
      await host.replace!(bytes, "latest-v2")
      expect(f.imported).toEqual({ bytes: btoa(String.fromCharCode(...bytes)), location: "Replace", options: {
        importTheme: true, importStyles: true, importParagraphSpacing: true, importPageColor: true,
        importChangeTrackingMode: true, importCustomProperties: true, importCustomXmlParts: true, importDifferentOddEvenPages: true,
      } })
      expect(typeof f.raw).toBe(encoded ? "string" : "object")
      expect(encoded ? JSON.parse(f.raw as string) : f.raw).toEqual({ ...linkedDocument, baseVersion: "latest-v2" })
      expect(f.context.document.changeTrackingMode).toBe("TrackMineOnly")
      expect(f.events).toEqual(["load:changeTrackingMode", "sync", "tracking:Off", "sync", "document-replace", "sync", "get-file", "read-slice", "close-file", "settings-set", "settings-save"])
    }
  } finally {
    if (previous) Object.defineProperty(globalThis, "window", previous)
    else delete (globalThis as { window?: unknown }).window
  }
})

test("refresh refuses unsupported fidelity API, oversize packages and changed links before native mutation", async () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    const unsupported = installRefreshHost({ supported: false })
    expect((await connectOffice()).host.replace).toBeUndefined()
    expect(unsupported.events).toEqual([])
    for (const changed of [false, true]) {
      const f = installRefreshHost()
      const { host } = await connectOffice()
      if (changed) f.raw = { ...linkedDocument, token: "another-capability" }
      await expect(host.replace!(changed ? new Uint8Array([1]) : new Uint8Array(4 * 1024 * 1024 + 1), "v2")).rejects.toBeInstanceOf(DocumentStateError)
      expect(f.events).toEqual([])
    }
  } finally {
    if (previous) Object.defineProperty(globalThis, "window", previous)
    else delete (globalThis as { window?: unknown }).window
  }
})

test("failed native replacement or link persistence never stamps a successful base or blindly repeats import", async () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    for (const failSettings of [false, true]) {
      const f = installRefreshHost({ failInsert: !failSettings, failSettings })
      const { host } = await connectOffice()
      await expect(host.replace!(latestRefreshBytes(), "v2")).rejects.toBeInstanceOf(DocumentStateError)
      if (!failSettings) {
        expect(f.raw).toEqual(linkedDocument)
        expect(f.events).not.toContain("settings-set")
        expect(f.context.document.changeTrackingMode).toBe("TrackAll")
      }
      const count = f.events.length
      await expect(host.replace!(latestRefreshBytes(), "v2")).rejects.toThrow("previous refresh")
      expect(f.events).toHaveLength(count)
      expect(f.events).not.toContain("native-save")
      expect(f.events).not.toContain("native-close")
    }
  } finally {
    if (previous) Object.defineProperty(globalThis, "window", previous)
    else delete (globalThis as { window?: unknown }).window
  }
})

test("refresh fences a native import that changes content, revisions, comments, numbering or sections", async () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    const bytes = latestRefreshBytes()
    for (const [part, from, to] of [
      ["word/document.xml", "Reviewed text", "Changed text"],
      ["word/document.xml", "Earlier reviewer", "Different reviewer"],
      ["word/document.xml", "12240", "9999"],
      ["word/comments.xml", "Retained comment", "Changed comment"],
      ["word/numbering.xml", 'w:numId="1"', 'w:numId="2"'],
    ]) {
      const parts = unzipSync(bytes), xml = strFromU8(parts[part]!)
      expect(xml).toContain(from)
      parts[part] = strToU8(xml.replace(from, to))
      const f = installRefreshHost({ exportBytes: zipSync(parts) }), { host } = await connectOffice()
      await expect(host.replace!(bytes, "v2")).rejects.toThrow("could not verify its contents")
      expect(f.events).toContain("document-replace")
      expect(f.events).not.toContain("settings-set")
      const events = f.events.length
      await expect(host.replace!(bytes, "v2")).rejects.toThrow("previous refresh")
      expect(f.events).toHaveLength(events)
    }
    const f = installRefreshHost({ failExport: true }), { host } = await connectOffice()
    await expect(host.replace!(bytes, "v2")).rejects.toThrow("could not verify its contents")
    expect(f.events).not.toContain("settings-set")
  } finally {
    if (previous) Object.defineProperty(globalThis, "window", previous)
    else delete (globalThis as { window?: unknown }).window
  }
})

test("Load latest keeps the old base and fences saving when Word's imported document fails verification", async () => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window")
  try {
    const latest = latestRefreshBytes(), parts = unzipSync(latest)
    parts["word/comments.xml"] = strToU8(strFromU8(parts["word/comments.xml"]!).replace("Retained comment", "Changed comment"))
    const f = installRefreshHost({ exportBytes: zipSync(parts) }), { host } = await connectOffice()
    const actions: string[] = []
    let operation: { sequence: number; operationId: string | null; state: "idle" | "pending" | "confirmed" | "not-dispatched" } = { sequence: 0, operationId: null, state: "idle" }
    const connection = (base: string): WordConnectionState => ({ version: { base: { id: base, number: base === "v1" ? 1 : 2 }, current: { id: "v2", number: 2 }, state: base === "v2" ? "current" : "stale" }, lease: { state: "owned", expiresAt: 100_000, generation: "owner" }, nativeOperation: { sequence: operation.sequence, state: operation.state } })
    const transport: Transport = {
      lease: async (action, base) => {
        actions.push(`${action}:${base}`)
        return connection(base)
      },
      nativeOperation: async (base, _generation, input) => {
        actions.push(`${input.action}:${base}`)
        if (input.action === "begin") operation = { sequence: input.sequence + 1, operationId: input.operationId, state: "pending" }
        if (input.action === "settle") operation = { sequence: input.sequence, operationId: input.operationId, state: input.outcome }
        return { operation, connection: connection(base) }
      },
      save: async () => { throw new Error("unexpected upload") },
      heartbeat: async (base) => connection(base),
      done: async () => { throw new Error("unexpected done") },
    }
    let cleared = false
    const sync = new WordSync("v1", host, transport, { read: () => null, write: () => {}, clear: () => { cleared = true } }, () => {}, documentFingerprint, async () => {}, () => 50_000)
    expect(await sync.loadLatest(async () => ({ version: { id: "v2", number: 2 }, bytes: latest }), () => true)).toBe(false)
    expect(sync.snapshot().message).toContain("could not verify its contents")
    expect(sync.snapshot().message).toContain("Create saved copy")
    await sync.tick()
    expect(sync.snapshot().message).toContain("could not verify its contents")
    expect(sync.snapshot().message).toContain("Create saved copy")
    expect(sync.snapshot().canSave).toBe(false)
    expect(cleared).toBe(false)
    expect(operation.state).toBe("pending")
    expect(actions).toContain("begin:v1"); expect(actions).not.toContain("settle:v1")
    expect(actions).not.toContain("release:v1")
    expect(f.events).not.toContain("settings-set")
    expect(await sync.loadLatest(async () => ({ version: { id: "v2", number: 2 }, bytes: latest }), () => true)).toBe(false)
    expect(f.events.filter(event => event === "document-replace")).toHaveLength(1)
    operation = { sequence: operation.sequence + 1, operationId: "different", state: "pending" }
    await sync.tick()
    expect(sync.snapshot().message).not.toContain("could not verify its contents")
    expect(sync.snapshot().message).toContain("Create saved copy")
  } finally {
    if (previous) Object.defineProperty(globalThis, "window", previous)
    else delete (globalThis as { window?: unknown }).window
  }
})
