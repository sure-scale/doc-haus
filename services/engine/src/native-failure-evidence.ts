import { Logger, References } from "effect"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"
import { closeSync, chmodSync, constants, fchmodSync, fsyncSync, fstatSync, lstatSync, mkdirSync, openSync, writeFileSync } from "node:fs"
import path from "node:path"

const names = new Set(["Error", "TypeError", "RangeError", "ReferenceError", "SyntaxError", "URIError", "EvalError", "AggregateError", "StepFailedError"])
// Explicit shipped files only. A workspace path, stack label or arbitrary file
// below the repository is never evidence of trusted code provenance.
const files = new Set([
  "services/engine/src/server.ts",
  "packages/core/src/session/execution.ts", "packages/core/src/session/runner/index.ts",
  "packages/core/src/session/runner/step.ts", "packages/core/src/session/to-session-error.ts",
  "dochaus/plugin/legal-v2.ts", "dochaus/plugin/legal/policy.ts", "dochaus/plugin/legal/v2-context.ts",
])

function own(value: unknown, key: string): unknown {
  if (value === null || (typeof value !== "object" && typeof value !== "function")) return undefined
  try {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)
    return descriptor && "value" in descriptor ? descriptor.value : undefined
  } catch { return undefined }
}

function reference(value: string, limit: number) {
  return { ref: draftingTraceReference(value.slice(0, limit)), prefix_limit_utf16: limit, truncated: value.length > limit }
}

function errorMetadata(value: unknown, repository: string) {
  const message = typeof value === "string" ? value : own(value, "message")
  const name = own(value, "name")
  const tag = own(value, "_tag")
  const stack = own(value, "stack")
  const prototypeName = (() => {
    try {
      const prototype = Object.getPrototypeOf(value)
      return new Map<object, string>([[Error.prototype, "Error"], [TypeError.prototype, "TypeError"], [RangeError.prototype, "RangeError"],
        [ReferenceError.prototype, "ReferenceError"], [SyntaxError.prototype, "SyntaxError"],
        [URIError.prototype, "URIError"], [EvalError.prototype, "EvalError"], [AggregateError.prototype, "AggregateError"]]).get(prototype)
    } catch { return undefined }
  })()
  const frames: Array<{ file: string; line: number; column: number }> = []
  if (typeof stack === "string" && stack.length <= 8192) {
    for (const line of stack.split("\n").slice(1, 65)) {
      if (line.length > 1024) continue
      const match = line.match(/(?:\(|at\s+)([^()]+?\.(?:ts|tsx|js)):(\d{1,8}):(\d{1,8})\)?\s*$/)
      if (!match || !path.isAbsolute(match[1]!)) continue
      const file = path.relative(repository, path.resolve(match[1]!)).split(path.sep).join("/")
      const lineNumber = Number(match[2]), column = Number(match[3])
      if (!files.has(file) || lineNumber < 1 || column < 1) continue
      frames.push({ file, line: lineNumber, column })
      if (frames.length === 8) break
    }
  }
  const kind = value === null ? "null" : typeof value
  return {
    error_kind: kind,
    ...(typeof name === "string" ? names.has(name) ? { name } : { name_fingerprint: reference(name, 256) } : prototypeName ? { name: prototypeName } : {}),
    ...(typeof tag === "string" ? names.has(tag) ? { tag } : { tag_fingerprint: reference(tag, 256) } : {}),
    ...(typeof message === "string" ? { message_fingerprint: reference(message, 8192) } : {}),
    frames,
    stack_status: typeof stack !== "string" ? "unavailable" : stack.length > 8192 ? "oversized" : "inspected",
  }
}

export function sanitizeNativeDrainFailure(cause: unknown, sessionID: unknown, repository: string) {
  const reasons = own(cause, "reasons")
  const length = own(reasons, "length")
  const count = typeof length === "number" && Number.isSafeInteger(length) && length >= 0 ? length : 0
  return {
    event: "native_session_drain_failed",
    ...(typeof sessionID === "string" && /^ses_[A-Za-z0-9_]{1,160}$/.test(sessionID) ? { session_id: sessionID } : {}),
    reason_count: count,
    reasons_truncated: count > 8,
    reasons: Array.from({ length: Math.min(count, 8) }, (_, index) => {
      const reason = own(reasons, String(index))
      const kind = own(reason, "_tag")
      return { index, kind: kind === "Fail" || kind === "Die" || kind === "Interrupt" ? kind : "Unknown",
        ...errorMetadata(kind === "Fail" ? own(reason, "error") : kind === "Die" ? own(reason, "defect") : undefined, repository) }
    }),
  }
}

export function nativeFailureEvidenceLogger(input: {
  repository: string; runtime: string; loggers: ReadonlySet<Logger.Logger<unknown, unknown>>
  emit?: (line: string) => void
}) {
  const emit = input.emit ?? ((line: string) => { process.stderr.write(line) })
  return Logger.make<unknown, void>(options => {
    if (!Array.isArray(options.message) || own(options.message, "length") !== 1 || own(options.message, "0") !== "Failed to drain Session") {
      for (const logger of input.loggers) logger.log(options)
      return
    }
    // Observe the original Cause without replacing the native failure or
    // forwarding its raw payload/annotations to the default logger.
    try {
      const record = { timestamp: options.date.toISOString(), ...sanitizeNativeDrainFailure(options.cause,
        own(options.fiber.getRef(References.CurrentLogAnnotations), "sessionID"), input.repository) }
      const line = JSON.stringify(record) + "\n"
      if (Buffer.byteLength(line) > 16384) throw new Error("Evidence record bound")
      const directory = path.join(input.runtime, "diagnostics")
      mkdirSync(directory, { recursive: true, mode: 0o700 })
      const stat = lstatSync(directory)
      if (!stat.isDirectory() || stat.isSymbolicLink() || (process.getuid && stat.uid !== process.getuid())) throw new Error("Evidence directory ownership")
      chmodSync(directory, 0o700)
      const fd = openSync(path.join(directory, "native-failures.jsonl"), constants.O_WRONLY | constants.O_CREAT | constants.O_APPEND | constants.O_NOFOLLOW, 0o600)
      try {
        const file = fstatSync(fd)
        if (!file.isFile() || (process.getuid && file.uid !== process.getuid())) throw new Error("Evidence file ownership")
        fchmodSync(fd, 0o600)
        writeFileSync(fd, line)
        fsyncSync(fd)
      } finally { closeSync(fd) }
      emit(line)
    } catch {
      // Never throw into SessionExecution or recursively log a failing sink.
      try { emit("DOCHAUS_NATIVE_FAILURE_EVIDENCE_SINK_FAILED\n") } catch { /* no recursive fallback */ }
    }
  })
}
