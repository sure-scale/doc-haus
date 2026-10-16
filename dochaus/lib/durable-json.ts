import { renameSync, writeFileSync } from "node:fs"
import { readRegularFile } from "./regular-file"

const durableJsonReadSnapshots: Array<Map<string, unknown>> = []
const durableJsonWritePreflights = new Map<
  string,
  (file: string, content: Buffer) => void
>()

export function readDurableJson(file: string): unknown {
  const snapshot = durableJsonReadSnapshots.at(-1)
  if (snapshot?.has(file)) return snapshot.get(file)
  const content = readRegularFile(file)
  if (!content) {
    snapshot?.set(file, undefined)
    return
  }
  try {
    const value = JSON.parse(content.toString("utf8")) as unknown
    snapshot?.set(file, value)
    return value
  } catch {
    snapshot?.set(file, undefined)
    return
  }
}

export function withDurableJsonReadSnapshot<T>(callback: () => T) {
  // Pure synchronous derivations often revisit the same large ledgers. Keep
  // one parsed view for that call stack and discard it before any later turn.
  if (durableJsonReadSnapshots.length) return callback()
  durableJsonReadSnapshots.push(new Map())
  try {
    return callback()
  } finally {
    durableJsonReadSnapshots.pop()
  }
}

export function writeDurableJson(file: string, value: unknown) {
  durableJsonReadSnapshots.forEach((snapshot) => snapshot.delete(file))
  const content = JSON.stringify(value, undefined, 2)
  if (content === undefined)
    throw new Error(
      "Durable JSON writes require a serializable value.",
    )
  const bytes = Buffer.from(content)
  durableJsonWritePreflights.forEach((preflight) =>
    preflight(file, bytes)
  )
  writeDurableBytes(file, bytes)
}

export function writeDurableBytes(
  file: string,
  bytes: Uint8Array,
) {
  durableJsonReadSnapshots.forEach((snapshot) =>
    snapshot.delete(file)
  )
  const temporary = `${file}.${process.pid}.${crypto.randomUUID()}.tmp`
  writeFileSync(temporary, bytes)
  renameSync(temporary, file)
}

export function registerDurableJsonWritePreflight(
  key: string,
  preflight: (file: string, content: Buffer) => void,
) {
  if (durableJsonWritePreflights.has(key))
    throw new Error(
      `Durable JSON write preflight ${key} is already registered.`,
    )
  durableJsonWritePreflights.set(key, preflight)
}

export function unregisterDurableJsonWritePreflight(
  key: string,
) {
  durableJsonWritePreflights.delete(key)
}
