import { AsyncLocalStorage } from "node:async_hooks"
import { mkdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from "node:fs"
import path from "node:path"

const LOCK_STALE_AFTER_MS = 60_000
const LOCK_WAIT_TIMEOUT_MS = 30_000
const LOCK_WAIT_INTERVAL_MS = 10
const HELD_DIRECTORY_LOCKS = new Set<string>()
const ASYNC_OWNED_DIRECTORY_LOCKS = new Set<string>()
const ASYNC_HELD_DIRECTORY_LOCKS = new AsyncLocalStorage<ReadonlySet<string>>()

export class DirectoryLockContentionError extends Error {
  readonly code = "directory_lock_busy"

  constructor(message: string) {
    super(message)
    this.name = "DirectoryLockContentionError"
  }
}

export function withDirectoryLock<T>(lock: string, busyMessage: string, fn: () => T) {
  const token = crypto.randomUUID()
  mkdirSync(path.dirname(lock), { recursive: true })
  acquireDirectoryLock(lock, token, busyMessage)
  try {
    const result = fn()
    if (isPromiseLike(result)) throw new Error("withDirectoryLock requires a synchronous callback.")
    return result
  } finally {
    releaseDirectoryLock(lock, token)
  }
}

// Synchronous durable-state mutations sometimes compose helpers which share a
// lock. Keep those nested calls in the same critical section while independent
// processes wait for the filesystem owner instead of failing a valid race.
export function withReentrantDirectoryLock<T>(lock: string, busyMessage: string, fn: () => T) {
  const resolved = path.resolve(lock)
  if (reentrantDirectoryLockHeld(resolved)) {
    const result = fn()
    if (isPromiseLike(result)) throw new Error("withReentrantDirectoryLock requires a synchronous callback.")
    return result
  }
  if (ASYNC_OWNED_DIRECTORY_LOCKS.has(resolved))
    throw new DirectoryLockContentionError(busyMessage)
  const token = crypto.randomUUID()
  const deadline = Date.now() + LOCK_WAIT_TIMEOUT_MS
  mkdirSync(path.dirname(resolved), { recursive: true })
  for (;;) {
    try {
      acquireDirectoryLock(resolved, token, busyMessage)
      break
    } catch (error) {
      if (!(error instanceof DirectoryLockContentionError) || Date.now() >= deadline) throw error
      waitForDirectoryLock()
    }
  }
  HELD_DIRECTORY_LOCKS.add(resolved)
  try {
    const result = fn()
    if (isPromiseLike(result)) throw new Error("withReentrantDirectoryLock requires a synchronous callback.")
    return result
  } finally {
    HELD_DIRECTORY_LOCKS.delete(resolved)
    releaseDirectoryLock(resolved, token)
  }
}

export async function withReentrantDirectoryLockAsync<T>(
  lock: string,
  busyMessage: string,
  fn: () => Promise<T>,
) {
  const resolved = path.resolve(lock)
  if (ASYNC_HELD_DIRECTORY_LOCKS.getStore()?.has(resolved)) return fn()
  if (HELD_DIRECTORY_LOCKS.has(resolved))
    throw new Error("An asynchronous directory-lock callback cannot begin inside a synchronous lock owner.")
  const token = crypto.randomUUID()
  const deadline = Date.now() + LOCK_WAIT_TIMEOUT_MS
  mkdirSync(path.dirname(resolved), { recursive: true })
  for (;;) {
    try {
      acquireDirectoryLock(resolved, token, busyMessage)
      break
    } catch (error) {
      if (!(error instanceof DirectoryLockContentionError) || Date.now() >= deadline) throw error
      await waitForDirectoryLockAsync()
    }
  }
  const held = new Set(ASYNC_HELD_DIRECTORY_LOCKS.getStore())
  held.add(resolved)
  ASYNC_OWNED_DIRECTORY_LOCKS.add(resolved)
  return ASYNC_HELD_DIRECTORY_LOCKS.run(held, async () => {
    try {
      return await fn()
    } finally {
      ASYNC_OWNED_DIRECTORY_LOCKS.delete(resolved)
      releaseDirectoryLock(resolved, token)
    }
  })
}

export function reentrantDirectoryLockHeld(lock: string) {
  const resolved = path.resolve(lock)
  return Boolean(
    HELD_DIRECTORY_LOCKS.has(resolved) ||
      ASYNC_HELD_DIRECTORY_LOCKS.getStore()?.has(resolved),
  )
}

function acquireDirectoryLock(lock: string, token: string, busyMessage: string) {
  try {
    mkdirSync(lock)
  } catch (error) {
    if (errorCode(error) !== "EEXIST") throw error
    if (!staleDirectoryLock(lock)) throw new DirectoryLockContentionError(busyMessage)
    replaceStaleDirectoryLock(lock, busyMessage)
  }
  try {
    writeFileSync(
      path.join(lock, "owner"),
      JSON.stringify({ token, pid: process.pid }),
      { flag: "wx" },
    )
  } catch (error) {
    rmSync(lock, { recursive: true, force: true })
    throw error
  }
}

function replaceStaleDirectoryLock(lock: string, busyMessage: string) {
  const stale = `${lock}.stale.${process.pid}.${crypto.randomUUID()}`
  try {
    renameSync(lock, stale)
  } catch (error) {
    if (["ENOENT", "EEXIST", "ENOTEMPTY"].includes(errorCode(error) ?? ""))
      throw new DirectoryLockContentionError(busyMessage)
    throw error
  }
  try {
    try {
      mkdirSync(lock)
    } catch (error) {
      if (errorCode(error) === "EEXIST") throw new DirectoryLockContentionError(busyMessage)
      throw error
    }
  } finally {
    rmSync(stale, { recursive: true, force: true })
  }
}

function releaseDirectoryLock(lock: string, token: string) {
  if (readDirectoryLockOwner(lock)?.token !== token) return
  rmSync(lock, { recursive: true, force: true })
}

function staleDirectoryLock(lock: string) {
  try {
    if (Date.now() - statSync(lock).mtimeMs <= LOCK_STALE_AFTER_MS) return false
  } catch {
    return false
  }
  const owner = readDirectoryLockOwner(lock)
  return !owner || !processIsAlive(owner.pid)
}

function readDirectoryLockOwner(lock: string) {
  try {
    const parsed = JSON.parse(readFileSync(path.join(lock, "owner"), "utf8")) as unknown
    if (
      !parsed ||
      typeof parsed !== "object" ||
      !("token" in parsed) ||
      !("pid" in parsed) ||
      typeof parsed.token !== "string" ||
      typeof parsed.pid !== "number" ||
      !Number.isInteger(parsed.pid) ||
      parsed.pid <= 0
    ) return
    return { token: parsed.token, pid: parsed.pid }
  } catch {
    return
  }
}

function processIsAlive(pid: number) {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return errorCode(error) !== "ESRCH"
  }
}

function waitForDirectoryLock() {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, LOCK_WAIT_INTERVAL_MS)
}

function waitForDirectoryLockAsync() {
  return new Promise<void>((resolve) => setTimeout(resolve, LOCK_WAIT_INTERVAL_MS))
}

function isPromiseLike(value: unknown): value is PromiseLike<unknown> {
  return Boolean(value && typeof value === "object" && "then" in value && typeof value.then === "function")
}

function errorCode(error: unknown) {
  if (!error || typeof error !== "object" || !("code" in error)) return
  return typeof error.code === "string" ? error.code : undefined
}
