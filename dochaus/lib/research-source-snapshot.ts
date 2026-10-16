import { createHash } from "node:crypto"
import { existsSync, mkdirSync, realpathSync, statSync, writeFileSync } from "node:fs"
import path from "node:path"
import { readRegularFile } from "./regular-file"
import { MAX_FETCH_BYTES, MAX_TEXT_CHARS, type FetchDescriptor, type FetchResult } from "./research-source-fetch"
import type { SourceSnippet } from "./research-source-snippets"

// These immutable payloads have no independent admission authority. A current
// source_fetch manifest row must bind one to the consuming research claim.
export type ResearchSourceSnapshotRef = { artifact: string; sha256: string; bytes: number }
export type ResearchSourceSnapshot = {
  version: 1
  descriptor: FetchDescriptor
  extraction: string
  completeness: "complete" | "partial" | "unknown"
  original_base64: string | null
  text: string
  locators: Array<{ id: string; start: number; end: number }>
}

const SNAPSHOT_DIRECTORY = ".dochaus/research-sources"
const MAX_SNAPSHOT_BYTES = MAX_FETCH_BYTES * 2 + MAX_TEXT_CHARS * 4

export function preserveResearchSourceSnapshot(
  directory: string,
  source: Exclude<FetchResult, { status: "blocked" }>,
  snippets: SourceSnippet[],
): ResearchSourceSnapshotRef {
  const snapshot: ResearchSourceSnapshot = {
    version: 1,
    descriptor: source.descriptor,
    extraction: source.extraction ?? "unknown",
    completeness: source.status === "partial" ? "partial" : source.sourceBytes ? "complete" : "unknown",
    original_base64: source.sourceBytes ? Buffer.from(source.sourceBytes).toString("base64") : null,
    text: source.text,
    locators: snippets.map((snippet, index) => ({ id: `snippet-${index + 1}`, start: snippet.start, end: snippet.end })),
  }
  if (!validSnapshot(snapshot)) throw new Error("Fetched source snapshot has inconsistent byte or text identity.")
  const content = Buffer.from(JSON.stringify(snapshot))
  if (content.byteLength > MAX_SNAPSHOT_BYTES) throw new Error("Fetched source snapshot exceeds its bounded storage limit.")
  const sha256 = hash(content)
  const artifact = `${SNAPSHOT_DIRECTORY}/${sha256}.json`
  mkdirSync(path.join(directory, SNAPSHOT_DIRECTORY), { recursive: true, mode: 0o700 })
  const file = containedSnapshotFile(directory, artifact)
  if (existsSync(file)) {
    if (!readRegularFile(file)?.equals(content)) throw new Error("Existing fetched source snapshot is corrupt.")
  } else {
    writeFileSync(file, content, { flag: "wx", mode: 0o600 })
  }
  return { artifact, sha256, bytes: content.byteLength }
}

export function readResearchSourceSnapshot(directory: string, ref: unknown): ResearchSourceSnapshot | undefined {
  if (!isResearchSourceSnapshotRef(ref)) return
  try {
    const file = containedSnapshotFile(directory, ref.artifact)
    if (statSync(file).size !== ref.bytes || ref.bytes > MAX_SNAPSHOT_BYTES) return
    const content = readRegularFile(file)
    if (!content || content.byteLength !== ref.bytes || hash(content) !== ref.sha256) return
    const snapshot: unknown = JSON.parse(content.toString("utf8"))
    return validSnapshot(snapshot) ? snapshot : undefined
  } catch {
    return
  }
}

export function isResearchSourceSnapshotRef(value: unknown): value is ResearchSourceSnapshotRef {
  return isRecord(value) && typeof value.sha256 === "string" && /^[a-f0-9]{64}$/.test(value.sha256) &&
    value.artifact === `${SNAPSHOT_DIRECTORY}/${value.sha256}.json` &&
    typeof value.bytes === "number" && Number.isSafeInteger(value.bytes) && value.bytes > 0 && value.bytes <= MAX_SNAPSHOT_BYTES
}

function containedSnapshotFile(directory: string, artifact: string) {
  const root = realpathSync(directory)
  const parent = realpathSync(path.join(directory, SNAPSHOT_DIRECTORY))
  if (parent !== path.join(root, SNAPSHOT_DIRECTORY)) throw new Error("Fetched source storage is outside its matter scope.")
  return path.join(parent, path.basename(artifact))
}

function validSnapshot(value: unknown): value is ResearchSourceSnapshot {
  if (!isRecord(value) || value.version !== 1 || !isRecord(value.descriptor) ||
    typeof value.extraction !== "string" ||
    !["complete", "partial", "unknown"].includes(String(value.completeness)) ||
    typeof value.text !== "string" || value.text.length > MAX_TEXT_CHARS ||
    !Array.isArray(value.locators) || !value.locators.length || value.locators.length > 10) return false
  const text = value.text
  const descriptor = value.descriptor
  if (descriptor.textHash !== (text ? hash(Buffer.from(text)) : "") || descriptor.textChars !== text.length ||
    ![descriptor.originalUrl, descriptor.finalUrl].every((url) => typeof url === "string" && /^https?:\/\//.test(url)) ||
    typeof descriptor.fetchedAt !== "string" || !Number.isFinite(Date.parse(descriptor.fetchedAt))) return false
  if (!value.locators.every((locator, index) => isRecord(locator) && locator.id === `snippet-${index + 1}` &&
    Number.isSafeInteger(locator.start) && Number.isSafeInteger(locator.end) &&
    Number(locator.start) >= 0 && Number(locator.end) >= Number(locator.start) && Number(locator.end) <= text.length)) return false
  if (value.original_base64 === null) return value.completeness !== "complete" && descriptor.bodyHash === ""
  if (typeof value.original_base64 !== "string" || value.original_base64.length > Math.ceil(MAX_FETCH_BYTES / 3) * 4) return false
  const body = Buffer.from(value.original_base64, "base64")
  return body.toString("base64") === value.original_base64 && body.byteLength <= MAX_FETCH_BYTES &&
    body.byteLength === descriptor.byteLength && descriptor.bodyHash === (body.byteLength ? hash(body) : "")
}

function hash(value: Buffer) {
  return createHash("sha256").update(value).digest("hex")
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
