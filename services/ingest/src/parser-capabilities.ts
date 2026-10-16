import { existsSync } from "node:fs"
import { statfs } from "node:fs/promises"
import { createRequire } from "node:module"
import os from "node:os"
import path from "node:path"

type ByteSource = "os.totalmem" | "os.freemem" | "fs.statfs.bavail"
export type ByteObservation =
  | { status: "observed"; bytes: number; source: ByteSource }
  | { status: "unknown"; bytes: null; source: ByteSource; reason: "unavailable" | "invalid_measurement" }

export type ParserComponentObservation = {
  id: "liteparse" | "unpdf" | "markitdown" | "pdftoppm" | "tesseract" | "uvx"
  kind: "package" | "executable" | "launcher"
  basis: "package_metadata" | "PATH"
  status: "present" | "not_found" | "unknown"
  version: string | null
  reason?: "not_found_in_resolution" | "not_found_on_path" | "metadata_unavailable" | "invalid_metadata" | "lookup_unavailable"
}

export type ParserCapabilitiesSnapshot = {
  schemaVersion: 1
  observedAt: string
  host: {
    basis: "ingest_process"
    platform: string
    architecture: string
    runtime: "Bun"
    runtimeVersion: string
    physicalArchitecture: "unknown"
    environment: "unknown"
  }
  memory: { total: ByteObservation; free: ByteObservation }
  disk: { scope: "ingest_filesystem"; available: ByteObservation }
  accelerator: { gpu: "unknown"; dedicatedVramBytes: null; unifiedMemoryBytes: null; reason: "not_probed" }
  components: ParserComponentObservation[]
}

const COMPONENTS = [
  { id: "liteparse", kind: "package", name: "@llamaindex/liteparse" },
  { id: "unpdf", kind: "package", name: "unpdf" },
  { id: "markitdown", kind: "executable", name: "markitdown" },
  { id: "pdftoppm", kind: "executable", name: "pdftoppm" },
  { id: "tesseract", kind: "executable", name: "tesseract" },
  { id: "uvx", kind: "launcher", name: "uvx" },
] as const

export async function inspectParserCapabilities(options: {
  diskDirectory?: string
  searchPath?: string
  system?: Pick<typeof os, "platform" | "arch" | "totalmem" | "freemem">
  resolvePackageMetadata?: (name: string) => string
} = {}): Promise<ParserCapabilitiesSnapshot> {
  const system = options.system ?? os
  return {
    schemaVersion: 1,
    observedAt: new Date().toISOString(),
    host: {
      basis: "ingest_process", platform: system.platform(), architecture: system.arch(),
      runtime: "Bun", runtimeVersion: Bun.version, physicalArchitecture: "unknown", environment: "unknown",
    },
    memory: {
      total: await observeBytes("os.totalmem", () => system.totalmem()),
      free: await observeBytes("os.freemem", () => system.freemem()),
    },
    disk: {
      scope: "ingest_filesystem",
      available: await observeBytes("fs.statfs.bavail", async () => {
        const filesystem = await statfs(options.diskDirectory ?? process.cwd(), { bigint: true })
        return Number(filesystem.bavail * filesystem.bsize)
      }),
    },
    // System RAM is not a measurement of dedicated VRAM or available unified memory.
    accelerator: { gpu: "unknown", dedicatedVramBytes: null, unifiedMemoryBytes: null, reason: "not_probed" },
    components: await Promise.all(COMPONENTS.map(async component => {
      if (component.kind === "package")
        return observePackage(component, options.resolvePackageMetadata ?? resolvePackageMetadata)
      const observation = { id: component.id, kind: component.kind, basis: "PATH", version: null } as const
      try {
        return Bun.which(component.name, { PATH: options.searchPath ?? process.env.PATH ?? "" })
          ? { ...observation, status: "present" as const }
          : { ...observation, status: "not_found" as const, reason: "not_found_on_path" as const }
      } catch {
        return { ...observation, status: "unknown" as const, reason: "lookup_unavailable" as const }
      }
    })),
  }
}

async function observeBytes(source: ByteSource, read: () => number | Promise<number>): Promise<ByteObservation> {
  try {
    const bytes = await read()
    return Number.isSafeInteger(bytes) && bytes >= 0
      ? { status: "observed", bytes, source }
      : { status: "unknown", bytes: null, source, reason: "invalid_measurement" }
  } catch {
    return { status: "unknown", bytes: null, source, reason: "unavailable" }
  }
}

async function observePackage(
  component: Extract<(typeof COMPONENTS)[number], { kind: "package" }>,
  resolve: (name: string) => string,
): Promise<ParserComponentObservation> {
  const observation = { id: component.id, kind: component.kind, basis: "package_metadata" } as const
  try {
    const file = Bun.file(resolve(component.name))
    if (file.size > 64 * 1024) return { ...observation, status: "unknown", version: null, reason: "invalid_metadata" }
    const metadata: unknown = await file.json()
    if (!metadata || typeof metadata !== "object" || !("name" in metadata) || metadata.name !== component.name)
      return { ...observation, status: "unknown", version: null, reason: "invalid_metadata" }
    return {
      ...observation, status: "present",
      version: "version" in metadata && typeof metadata.version === "string" && metadata.version.trim() ? metadata.version : null,
    }
  } catch (error) {
    return error && typeof error === "object" && "code" in error && error.code === "MODULE_NOT_FOUND"
      ? { ...observation, status: "not_found", version: null, reason: "not_found_in_resolution" }
      : { ...observation, status: "unknown", version: null, reason: "metadata_unavailable" }
  }
}

function resolvePackageMetadata(name: string): string {
  // Resolving entrypoints reads package maps; it never imports a parser or its native dependencies.
  const require = createRequire(import.meta.url)
  try {
    return require.resolve(`${name}/package.json`)
  } catch {
    let directory = path.dirname(require.resolve(name))
    while (directory !== path.dirname(directory)) {
      if (existsSync(path.join(directory, "package.json"))) return path.join(directory, "package.json")
      directory = path.dirname(directory)
    }
    throw new Error("Package metadata could not be located.")
  }
}
