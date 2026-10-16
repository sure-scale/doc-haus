import { expect, test } from "bun:test"
import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from "node:fs"
import { createRequire } from "node:module"
import { tmpdir } from "node:os"
import path from "node:path"
import { inspectParserCapabilities } from "./parser-capabilities"

test("observes metadata and PATH without executing parsers, OCR or a launcher", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "parser-inventory-"))
  const marker = path.join(directory, "executed")
  try {
    const modules = path.join(directory, "node_modules")
    const binaries = path.join(directory, "bin")
    mkdirSync(binaries)
    for (const name of ["@llamaindex/liteparse", "unpdf"]) {
      mkdirSync(path.join(modules, name), { recursive: true })
      writeFileSync(path.join(modules, name, "package.json"), JSON.stringify({ name, version: "1.2.3", main: "index.js" }))
      writeFileSync(path.join(modules, name, "index.js"), `throw new Error("Parser must not be imported")`)
    }
    for (const name of ["uvx", "pdftoppm", "tesseract"])
      writeFileSync(path.join(binaries, name), `#!/bin/sh\ntouch '${marker}'\nexit 1\n`, { mode: 0o755 })
    const require = createRequire(path.join(directory, "fixture.cjs"))
    const before = readdirSync(directory).sort()
    const snapshot = await inspectParserCapabilities({
      diskDirectory: directory, searchPath: binaries,
      resolvePackageMetadata: name => require.resolve(`${name}/package.json`),
      system: { platform: () => "darwin", arch: () => "arm64", totalmem: () => 8 * 1024 ** 3, freemem: () => 2 * 1024 ** 3 },
    })
    expect(snapshot.host).toMatchObject({ basis: "ingest_process", platform: "darwin", architecture: "arm64", physicalArchitecture: "unknown", environment: "unknown" })
    expect(snapshot.memory.total).toEqual({ status: "observed", bytes: 8 * 1024 ** 3, source: "os.totalmem" })
    expect(snapshot.memory.free).toEqual({ status: "observed", bytes: 2 * 1024 ** 3, source: "os.freemem" })
    expect(snapshot.disk.available.status).toBe("observed")
    expect(snapshot.accelerator).toEqual({ gpu: "unknown", dedicatedVramBytes: null, unifiedMemoryBytes: null, reason: "not_probed" })
    expect(snapshot.components.find(component => component.id === "liteparse")).toMatchObject({ status: "present", version: "1.2.3", basis: "package_metadata" })
    expect(snapshot.components.find(component => component.id === "unpdf")).toMatchObject({ status: "present", version: "1.2.3" })
    expect(snapshot.components.find(component => component.id === "markitdown")).toMatchObject({ status: "not_found", reason: "not_found_on_path" })
    expect(snapshot.components.find(component => component.id === "uvx")).toMatchObject({ status: "present", kind: "launcher", version: null })
    expect(snapshot.components.find(component => component.id === "pdftoppm")?.status).toBe("present")
    expect(snapshot.components.find(component => component.id === "tesseract")?.status).toBe("present")
    expect(existsSync(marker)).toBe(false)
    expect(readdirSync(directory).sort()).toEqual(before)
    expect(JSON.stringify(snapshot)).not.toContain(directory)
    expect(JSON.stringify(snapshot)).not.toContain("recommended")
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("keeps partial observations when resources or metadata cannot be established", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "parser-inventory-partial-"))
  try {
    const malformed = path.join(directory, "package.json")
    writeFileSync(malformed, "{ invalid metadata")
    const snapshot = await inspectParserCapabilities({
      diskDirectory: path.join(directory, "does-not-exist"), searchPath: directory,
      system: { platform: () => "win32", arch: () => "x64", totalmem: () => 16 * 1024 ** 3, freemem: () => { throw new Error("private probe detail") } },
      resolvePackageMetadata: name => {
        if (name === "@llamaindex/liteparse") return malformed
        throw Object.assign(new Error("private resolution detail"), { code: "MODULE_NOT_FOUND" })
      },
    })
    expect(snapshot.host.platform).toBe("win32")
    expect(snapshot.memory.total.status).toBe("observed")
    expect(snapshot.memory.free).toEqual({ status: "unknown", bytes: null, source: "os.freemem", reason: "unavailable" })
    expect(snapshot.disk.available).toEqual({ status: "unknown", bytes: null, source: "fs.statfs.bavail", reason: "unavailable" })
    expect(snapshot.components.find(component => component.id === "liteparse")).toMatchObject({ status: "unknown", version: null, reason: "metadata_unavailable" })
    expect(snapshot.components.find(component => component.id === "unpdf")).toMatchObject({ status: "not_found", reason: "not_found_in_resolution" })
    expect(snapshot.components.filter(component => component.basis === "PATH").every(component => component.status === "not_found")).toBe(true)
    expect(snapshot.accelerator.dedicatedVramBytes).toBeNull()
    expect(snapshot.accelerator.unifiedMemoryBytes).toBeNull()
    expect(JSON.stringify(snapshot)).not.toContain("private")
    expect(existsSync(path.join(directory, "does-not-exist"))).toBe(false)
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})

test("does not turn invalid byte observations or foreign metadata into capacity", async () => {
  const directory = mkdtempSync(path.join(tmpdir(), "parser-inventory-invalid-"))
  try {
    const metadata = path.join(directory, "package.json")
    writeFileSync(metadata, JSON.stringify({ name: "another-package", version: "9.0.0" }))
    const snapshot = await inspectParserCapabilities({ diskDirectory: directory, searchPath: directory,
      resolvePackageMetadata: () => metadata,
      system: { platform: () => "linux", arch: () => "x64", totalmem: () => -1, freemem: () => NaN },
    })
    expect(snapshot.memory.total).toMatchObject({ status: "unknown", bytes: null, reason: "invalid_measurement" })
    expect(snapshot.memory.free).toMatchObject({ status: "unknown", bytes: null, reason: "invalid_measurement" })
    expect(snapshot.components.filter(component => component.kind === "package").every(component => component.status === "unknown" && component.version === null && component.reason === "invalid_metadata")).toBe(true)
  } finally {
    rmSync(directory, { recursive: true, force: true })
  }
})
