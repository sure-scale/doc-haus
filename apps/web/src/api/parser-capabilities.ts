import { INGEST_URL } from "../config"

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

export async function getParserCapabilities(options: { signal?: AbortSignal; fetch?: typeof fetch } = {}): Promise<ParserCapabilitiesSnapshot> {
  const controller = new AbortController()
  const abort = () => controller.abort()
  const timeout = setTimeout(abort, 10_000)
  options.signal?.addEventListener("abort", abort, { once: true })
  if (options.signal?.aborted) abort()
  try {
    const response = await (options.fetch ?? fetch)(`${INGEST_URL}/host/parser-capabilities`, {
      method: "GET", cache: "no-store", signal: controller.signal,
    }).catch(() => { throw new Error("The ingestion host inventory is unavailable.") })
    if (!response.ok) throw new Error("The ingestion host inventory is unavailable.")
    const snapshot: unknown = await response.json().catch(() => null)
    if (!isParserCapabilitiesSnapshot(snapshot)) throw new Error("The ingestion host returned an unsupported inventory response.")
    return snapshot
  } finally {
    clearTimeout(timeout)
    options.signal?.removeEventListener("abort", abort)
  }
}

function isParserCapabilitiesSnapshot(value: unknown): value is ParserCapabilitiesSnapshot {
  if (!record(value) || value.schemaVersion !== 1 || typeof value.observedAt !== "string" || !Number.isFinite(Date.parse(value.observedAt))) return false
  if (!record(value.host) || value.host.basis !== "ingest_process" || !text(value.host.platform) || !text(value.host.architecture) || value.host.runtime !== "Bun" || !text(value.host.runtimeVersion) || value.host.physicalArchitecture !== "unknown" || value.host.environment !== "unknown") return false
  if (!record(value.memory) || !bytes(value.memory.total, "os.totalmem") || !bytes(value.memory.free, "os.freemem")) return false
  if (!record(value.disk) || value.disk.scope !== "ingest_filesystem" || !bytes(value.disk.available, "fs.statfs.bavail")) return false
  if (!record(value.accelerator) || value.accelerator.gpu !== "unknown" || value.accelerator.dedicatedVramBytes !== null || value.accelerator.unifiedMemoryBytes !== null || value.accelerator.reason !== "not_probed") return false
  return Array.isArray(value.components) && value.components.length === 6 && value.components.every(component) &&
    new Set(value.components.map(entry => entry.id)).size === 6
}

function bytes(value: unknown, source: ByteSource): value is ByteObservation {
  if (!record(value) || value.source !== source) return false
  if (value.status === "observed") return typeof value.bytes === "number" && Number.isSafeInteger(value.bytes) && value.bytes >= 0
  return value.status === "unknown" && value.bytes === null && (value.reason === "unavailable" || value.reason === "invalid_measurement")
}

function component(value: unknown): value is ParserComponentObservation {
  const kinds = { liteparse: "package", unpdf: "package", markitdown: "executable", pdftoppm: "executable", tesseract: "executable", uvx: "launcher" } as const
  if (!record(value) || typeof value.id !== "string" || !Object.hasOwn(kinds, value.id)) return false
  const kind = kinds[value.id as keyof typeof kinds]
  if (value.kind !== kind || value.basis !== (kind === "package" ? "package_metadata" : "PATH")) return false
  if (value.version !== null && !text(value.version)) return false
  if (value.status === "present") return kind === "package" || value.version === null
  if (value.version !== null) return false
  if (value.status === "not_found") return value.reason === (kind === "package" ? "not_found_in_resolution" : "not_found_on_path")
  return value.status === "unknown" && (kind === "package"
    ? value.reason === "metadata_unavailable" || value.reason === "invalid_metadata"
    : value.reason === "lookup_unavailable")
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value)
}

function text(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0
}
