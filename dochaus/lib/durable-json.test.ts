import { describe, expect, test } from "bun:test"
import {
  lstatSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { readDurableJson } from "./durable-json"

describe("durable JSON", () => {
  test("reads regular JSON and preserves missing or malformed compatibility", () => {
    const directory = mkdtempSync(
      path.join(tmpdir(), "durable-json-"),
    )
    try {
      const valid = path.join(directory, "valid.json")
      const malformed = path.join(directory, "malformed.json")
      writeFileSync(valid, '{"valid":true}')
      writeFileSync(malformed, "{")

      expect(readDurableJson(valid)).toEqual({ valid: true })
      expect(readDurableJson(malformed)).toBeUndefined()
      expect(
        readDurableJson(path.join(directory, "missing.json")),
      ).toBeUndefined()
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })

  test("fails closed promptly on a FIFO without changing its type", async () => {
    const directory = mkdtempSync(
      path.join(tmpdir(), "durable-json-fifo-"),
    )
    let child: ReturnType<typeof spawnReadDurableJsonChild> | undefined
    try {
      const fifo = path.join(directory, "control.json")
      expect(Bun.spawnSync(["mkfifo", fifo]).exitCode).toBe(0)
      child = spawnReadDurableJsonChild(directory, fifo)

      expect(await boundedChildExit(child.process)).toBe(
        child.rejectedExitCode,
      )
      expect(lstatSync(fifo).isFIFO()).toBe(true)
    } finally {
      if (child?.process.exitCode === null) child.process.kill(9)
      if (child) await child.process.exited
      rmSync(directory, { recursive: true, force: true })
    }
  }, 20_000)
})

function spawnReadDurableJsonChild(
  directory: string,
  file: string,
) {
  const script = path.join(directory, "read-durable-json-child.ts")
  const rejectedExitCode = 89
  writeFileSync(
    script,
    `import { readDurableJson } from ${JSON.stringify(path.join(import.meta.dir, "durable-json.ts"))}\n` +
      `try {\n` +
      `  readDurableJson(process.argv[2]!)\n` +
      `} catch {\n` +
      `  process.exit(${rejectedExitCode})\n` +
      `}\n` +
      `process.exit(90)\n`,
  )
  return {
    process: Bun.spawn([
      process.execPath,
      script,
      file,
    ], { stderr: "inherit" }),
    rejectedExitCode,
  }
}

function boundedChildExit(
  process: ReturnType<typeof spawnReadDurableJsonChild>["process"],
) {
  return new Promise<number>((resolve, reject) => {
    const timeout = setTimeout(
      () => reject(new Error("Timed out waiting for bounded child.")),
      7_500,
    )
    process.exited.then(
      (code) => {
        clearTimeout(timeout)
        resolve(code)
      },
      (error) => {
        clearTimeout(timeout)
        reject(error)
      },
    )
  })
}
