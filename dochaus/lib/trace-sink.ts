import { createHash } from "node:crypto"
import { constants } from "node:fs"
import { lstat, mkdir, open, realpath, statfs } from "node:fs/promises"
import path from "node:path"

// Private, observation-only drafting trace storage. Nothing here may feed model
// context, tool results, retries, budgets, claims or settlement decisions, and
// no function here throws into the observed product path except root validation.
export const TRACE_ROOT_ENV = "DOCHAUS_PROVIDER_CAPTURE_DIR"
export const TRACE_MIN_FREE_ENV = "DOCHAUS_TRACE_MIN_FREE_BYTES"
export const TRACE_STREAMS = ["tool-calls", "generations", "recovery"] as const
export type TraceStream = (typeof TRACE_STREAMS)[number]
const REDACTED = "[REDACTED_CREDENTIAL]"
const DEFAULT_MIN_FREE = 1024 ** 3
// Compared after lowercasing and removing "-" and "_".
const SECRET_KEYS = new Set([
  "authorization", "proxyauthorization", "cookie", "setcookie", "xapikey", "xgoogapikey", "apikey",
  "accesstoken", "refreshtoken", "idtoken", "clientsecret", "privatekey", "password", "token", "secret", "secretkey", "accesskey",
])
// Each pattern is linear: bounded character classes, no nested quantifiers.
const SECRET_PATTERNS = [
  /-----BEGIN [A-Z ]{0,40}PRIVATE KEY-----[A-Za-z0-9+/=\s\\]*(?:-----END [A-Z ]{0,40}PRIVATE KEY-----)?/g,
  /\bBearer\s+(?=[A-Za-z0-9._~+/-]*[0-9._~+/-])[A-Za-z0-9._~+/-]{16,}=*/gi,
  /\bBasic\s+(?=[A-Za-z0-9+/]*[0-9+/=])[A-Za-z0-9+/]{12,}={0,2}/g,
  /\bya29\.[A-Za-z0-9._-]{20,}/g,
  /\b1\/\/0[A-Za-z0-9_-]{30,}/g,
  /\bAIza[0-9A-Za-z_-]{35}/g,
  /\bGOCSPX-[A-Za-z0-9_-]{20,}/g,
  /\b(?:AKIA|ASIA)[0-9A-Z]{16}\b/g,
  /\bsk-(?:proj-|ant-)?[A-Za-z0-9_-]{20,}/g,
  /\beyJ[A-Za-z0-9_-]{8,}\.eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}/g,
  /\bgh[pousr]_[A-Za-z0-9]{36,}/g,
  /\bgithub_pat_[A-Za-z0-9_]{20,}/g,
]

/** Full-content capture requires an explicit directory; unset and `off` disable it. */
export function resolveTraceRoot(env: Record<string, string | undefined> = process.env) {
  const configured = env[TRACE_ROOT_ENV]?.trim()
  if (configured === "off") return undefined
  if (configured) return configured
  return undefined
}

const ensured = new Map<string, Promise<string>>()

/**
 * Validate, then create (0700), a private trace root; throw a descriptive error
 * when it is unusable. Inside any git checkout the root must be git-ignored, and
 * it may never contain or sit inside a matter, the workspace or the config.
 */
export function ensureTraceRoot(destination: string, guard: { repository: string; protectedPaths: Array<string | undefined> }) {
  const key = JSON.stringify([destination, guard.repository, guard.protectedPaths])
  const existing = ensured.get(key)
  if (existing) return existing
  const pending = validateTraceRoot(destination, guard)
  ensured.set(key, pending)
  void pending.catch(() => ensured.delete(key))
  return pending
}

async function validateTraceRoot(destination: string, guard: { repository: string; protectedPaths: Array<string | undefined> }) {
  const reject = (reason: string) => new Error(`Drafting trace root ${destination} is unusable: ${reason}. Set ${TRACE_ROOT_ENV} to a private absolute directory, or to "off" to disable tracing explicitly.`)
  if (!path.isAbsolute(destination)) throw reject("the path must be absolute")
  // Check placement before creating anything, so a misplaced root leaves no directories behind.
  const planned = await canonical(destination)
  const repository = await canonical(guard.repository)
  const protectedRoots = await Promise.all(guard.protectedPaths.filter(value => value !== undefined).map(canonical))
  const overlap = protectedRoots.find(value => inside(planned, value) || inside(value, planned))
  if (overlap) throw reject(`it overlaps the protected directory ${overlap}`)
  if (inside(repository, planned)) throw reject("it contains the repository")
  await requireGitIgnored(planned, repository, reject)
  await mkdir(destination, { recursive: true, mode: 0o700 })
  const stat = await lstat(destination)
  // System ancestors may be symlinks (macOS /var); the root itself may not.
  if (stat.isSymbolicLink() || !stat.isDirectory()) throw reject("it must be a real directory, not a symlink")
  if ((stat.mode & 0o777) !== 0o700) throw reject(`its mode is ${(stat.mode & 0o777).toString(8)}, expected 700`)
  if (process.getuid && stat.uid !== process.getuid()) throw reject("it must be owned by the engine user")
  const root = await realpath(destination)
  if (root !== planned) throw reject("its location changed while it was being created")
  return root
}

// A root inside the engine checkout, or inside any other git work tree, must be
// git-ignored there so private trace content can never be committed.
async function requireGitIgnored(root: string, repository: string, reject: (reason: string) => Error) {
  const ancestor = await existingAncestor(root)
  const toplevel = await git(["-C", ancestor, "rev-parse", "--show-toplevel"])
  const trees = [...new Set([inside(root, repository) ? repository : undefined, toplevel.ok ? await canonical(toplevel.output) : undefined])]
    .filter(value => value !== undefined)
  for (const tree of trees) {
    const ignored = await git(["-C", tree, "check-ignore", "-q", root])
    if (!ignored.ok) throw reject(`it lies inside the git checkout ${tree} and is not git-ignored`)
  }
}

/** Replace credential-shaped substrings. Content is otherwise kept byte-for-byte. */
export function redactCredentials(text: string) {
  const counted = { count: 0 }
  const redacted = SECRET_PATTERNS.reduce((value, pattern) => value.replace(pattern, () => {
    counted.count++
    return REDACTED
  }), text)
  return { text: redacted, count: counted.count }
}

/** Serialize one trace record with credential redaction; bytes become base64. */
export function serializeTraceRecord(record: Record<string, unknown>) {
  return JSON.stringify(record, (key, value: unknown) => {
    if (key && SECRET_KEYS.has(key.toLowerCase().replace(/[-_]/g, "")) && value !== undefined && value !== null && value !== "") return REDACTED
    if (typeof value === "string") return redactCredentials(value).text
    if (typeof value === "bigint") return value.toString()
    if (ArrayBuffer.isView(value)) return { bytes_base64: Buffer.from(value.buffer, value.byteOffset, value.byteLength).toString("base64") }
    if (value instanceof Map) return Object.fromEntries(value)
    if (value instanceof Set) return [...value]
    return value
  })
}

/** A message for any thrown value; never throws itself. */
export function describeError(error: unknown) {
  try {
    return error instanceof Error ? error.message : String(error)
  } catch {
    return "unprintable error"
  }
}

type Queues = Map<string, Promise<void>>
const queueKey = Symbol.for("dochaus.trace-sink.queues")
// One queue per file per process, even if this module is loaded more than once.
const queues: Queues = (globalThis as typeof globalThis & { [queueKey]?: Queues })[queueKey] ??= new Map()
const warned = new Map<string, number>()

/**
 * Append one full, redacted JSON line to `<root>/streams/<stream>/<key>.jsonl`.
 * The record is serialized synchronously so later hook mutation cannot change
 * it. Serialization or sink failures and low-disk suspension are reported
 * loudly on stderr and never propagate into the observed product path.
 */
export function appendTrace(root: string, stream: TraceStream, key: string, record: Record<string, unknown>): Promise<void> {
  try {
    const line = traceLine({ version: 1, recorded_at: new Date().toISOString(), ...record })
    const file = path.join(root, "streams", stream, `${traceFileKey(key)}.jsonl`)
    const next: Promise<void> = (queues.get(file) ?? Promise.resolve()).then(() => writeLine(root, file, line)).catch((error: unknown) => {
      warnOnce(`sink:${(error as NodeJS.ErrnoException | null | undefined)?.code ?? "error"}`, `[drafting-trace] TRACE WRITE FAILED for ${file}: ${describeError(error)}`)
    }).finally(() => { if (queues.get(file) === next) queues.delete(file) })
    queues.set(file, next)
    return next
  } catch (error) {
    warnOnce("serialize", `[drafting-trace] TRACE RECORD DROPPED for ${stream}: ${describeError(error)}`)
    return Promise.resolve()
  }
}

/** Whether the trace filesystem still has the configured free-space floor. */
export async function traceSpaceAvailable(root: string) {
  const configured = Number(process.env[TRACE_MIN_FREE_ENV] ?? DEFAULT_MIN_FREE)
  const floor = Number.isFinite(configured) && configured >= 0 ? configured : DEFAULT_MIN_FREE
  if (floor !== configured) warnOnce("floor", `[drafting-trace] ${TRACE_MIN_FREE_ENV} is not a non-negative number; using ${DEFAULT_MIN_FREE} bytes.`)
  const stats = await statfs(root)
  if (stats.bavail * stats.bsize >= floor) return true
  warnOnce("disk", `[drafting-trace] TRACE CAPTURE SUSPENDED: free space on ${root} is below ${floor} bytes (${TRACE_MIN_FREE_ENV}). Existing traces are kept; new records are skipped until space is freed.`)
  return false
}

export function warnOnce(kind: string, message: string) {
  const now = Date.now()
  if (now - (warned.get(kind) ?? 0) < 60_000) return
  warned.set(kind, now)
  console.error(message)
}

export function sha256(bytes: Uint8Array | string) {
  return createHash("sha256").update(bytes).digest("hex")
}

// A cyclic record keeps its identity fields rather than disappearing.
function traceLine(record: Record<string, unknown>) {
  try {
    return serializeTraceRecord(record) + "\n"
  } catch (error) {
    const identity = Object.fromEntries(Object.entries(record).filter(([, value]) => typeof value !== "object" || value === null))
    return serializeTraceRecord({ ...identity, serialization_error: describeError(error) }) + "\n"
  }
}

function traceFileKey(key: string) {
  return /^[A-Za-z0-9_.-]{1,128}$/.test(key) && !key.startsWith(".") ? key : `key-${sha256(key).slice(0, 32)}`
}

async function writeLine(root: string, file: string, line: string) {
  if (!await traceSpaceAvailable(root)) return
  await mkdir(path.dirname(file), { recursive: true, mode: 0o700 })
  const handle = await open(file, constants.O_WRONLY | constants.O_CREAT | constants.O_APPEND | constants.O_NOFOLLOW, 0o600)
  try { await handle.writeFile(line) } finally { await handle.close() }
}

// Resolve symlinks in the longest existing prefix, then re-append the rest.
async function canonical(value: string) {
  const absolute = path.resolve(value)
  const ancestor = await existingAncestor(absolute)
  return path.join(await realpath(ancestor), path.relative(ancestor, absolute))
}

async function existingAncestor(value: string): Promise<string> {
  const found = await lstat(value).then(() => true, () => false)
  if (found || path.dirname(value) === value) return value
  return existingAncestor(path.dirname(value))
}

// A missing git binary reads as "not proven": a root inside the engine checkout
// is then rejected, while a root outside every checkout needs no git answer.
async function git(args: string[]) {
  const child = (() => { try { return Bun.spawn(["git", ...args], { stdout: "pipe", stderr: "ignore" }) } catch { return undefined } })()
  if (!child) return { ok: false, output: "" }
  const output = (await new Response(child.stdout).text()).trim()
  return { ok: await child.exited === 0, output }
}

function inside(file: string, directory: string) {
  return file === directory || file.startsWith(directory + path.sep)
}
