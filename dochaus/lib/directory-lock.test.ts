import { expect, test } from "bun:test"
import { existsSync, mkdtempSync, rmSync, utimesSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import {
  DirectoryLockContentionError,
  withDirectoryLock,
  withReentrantDirectoryLock,
  withReentrantDirectoryLockAsync,
} from "./directory-lock"

test("directory lock rejects live contention and releases for an exact retry", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "directory-lock-contention-"))
  const lock = path.join(directory, "lock")
  try {
    withDirectoryLock(lock, "busy", () => {
      const old = new Date(Date.now() - 61_000)
      utimesSync(lock, old, old)
      expect(() => withDirectoryLock(lock, "busy", () => undefined)).toThrow(DirectoryLockContentionError)
      try {
        withDirectoryLock(lock, "busy", () => undefined)
        throw new Error("expected live lock contention")
      } catch (error) {
        expect(error).toMatchObject({
          name: "DirectoryLockContentionError",
          code: "directory_lock_busy",
          message: "busy",
        })
      }
      expect(existsSync(lock)).toBe(true)
    })
    expect(existsSync(lock)).toBe(false)
    expect(withDirectoryLock(lock, "busy", () => "retried")).toBe("retried")
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("a dead stale owner is recovered without stealing from a live process", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "directory-lock-takeover-"))
  const lock = path.join(directory, "lock")
  const entered = path.join(directory, "entered")
  const release = path.join(directory, "release")
  const script = path.join(directory, "holder.ts")
  writeFileSync(
    script,
    `import { existsSync, writeFileSync } from "node:fs"\n` +
      `import { withDirectoryLock } from ${JSON.stringify(path.join(import.meta.dir, "directory-lock.ts"))}\n` +
      `const [lock, entered, release] = process.argv.slice(2)\n` +
      `withDirectoryLock(lock!, "busy", () => {\n` +
      `  writeFileSync(entered!, "entered")\n` +
      `  const signal = new Int32Array(new SharedArrayBuffer(4))\n` +
      `  while (!existsSync(release!)) Atomics.wait(signal, 0, 0, 10)\n` +
      `})\n`,
  )
  const holder = Bun.spawn([process.execPath, script, lock, entered, release], { stderr: "pipe" })
  try {
    await waitForFile(entered)
    holder.kill("SIGKILL")
    await holder.exited
    const old = new Date(Date.now() - 61_000)
    utimesSync(lock, old, old)
    expect(
      withDirectoryLock(lock, "busy", () => {
        expect(existsSync(lock)).toBe(true)
        return "recovered"
      }),
    ).toBe("recovered")
    expect(existsSync(lock)).toBe(false)
  } finally {
    holder.kill()
    await holder.exited
    rmSync(directory, { recursive: true, force: true })
  }
})

test("an unrelated synchronous caller fails fast while an async context owns the lock", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "directory-lock-async-owner-"))
  const lock = path.join(directory, "lock")
  let release = () => {}
  let entered = () => {}
  const held = new Promise<void>((resolve) => {
    entered = resolve
  })
  const wait = new Promise<void>((resolve) => {
    release = resolve
  })
  try {
    const owner = withReentrantDirectoryLockAsync(lock, "busy", async () => {
      expect(
        withReentrantDirectoryLock(lock, "busy", () => "reentrant"),
      ).toBe("reentrant")
      entered()
      await wait
    })
    await held
    const started = Date.now()
    expect(() =>
      withReentrantDirectoryLock(lock, "busy", () => "unreachable"),
    ).toThrow(DirectoryLockContentionError)
    expect(Date.now() - started).toBeLessThan(1_000)
    release()
    await owner
  } finally {
    release()
    rmSync(directory, { recursive: true, force: true })
  }
})

async function waitForFile(file: string) {
  const deadline = Date.now() + 5_000
  while (!existsSync(file) && Date.now() < deadline) await Bun.sleep(10)
  expect(existsSync(file)).toBe(true)
}
