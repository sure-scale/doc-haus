import { expect, test } from "bun:test"
import { Cause, Effect, Exit, Logger } from "effect"
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import os from "node:os"
import path from "node:path"
import { nativeFailureEvidenceLogger, sanitizeNativeDrainFailure } from "../src/native-failure-evidence"
import { draftingTraceReference } from "../../../dochaus/lib/drafting-trace-identity"

const repository = path.resolve(import.meta.dir, "../../..")

test("native failure evidence bounds Fail and Die metadata without privileged text or getters", () => {
  const secret = "PRIVATE_SENTINEL"
  const error = new TypeError(secret)
  error.stack = `${secret}\n    at ${secret} (${repository}/dochaus/plugin/legal-v2.ts:29:7)\n    at ${secret} (/foreign/${secret}.ts:3:2)\n    at ${secret} (${repository}/workspace/${secret}.ts:3:2)`
  const foreign = { name: secret, _tag: secret, message: secret, stack: error.stack }
  let getters = 0
  const accessor = Object.defineProperties({}, Object.fromEntries(["name", "message", "stack", "_tag", "cause"].map(key =>
    [key, { get() { getters++; throw new Error(secret) } }])))
  const hostile = new Proxy({}, { getOwnPropertyDescriptor() { throw new Error(secret) }, getPrototypeOf() { throw new Error(secret) } })
  const result = sanitizeNativeDrainFailure(Cause.fromReasons([Cause.fail(error), Cause.die(foreign), Cause.die(accessor), Cause.die(hostile)].flatMap(cause => cause.reasons)), "ses_neutral_failure", repository)
  expect(result.reason_count).toBe(4)
  expect(result.reasons[0]!.kind).toBe("Fail")
  expect(result.reasons[0]!.name).toBe("TypeError")
  expect(result.reasons[0]!.message_fingerprint?.ref).toBe(draftingTraceReference(secret))
  expect(result.reasons[0]!.frames).toEqual([{ file: "dochaus/plugin/legal-v2.ts", line: 29, column: 7 }])
  expect(result.reasons[1]!.kind).toBe("Die")
  expect(getters).toBe(0)
  expect(JSON.stringify(result)).not.toContain(secret)
  expect(JSON.stringify(result)).not.toContain(repository)
  expect(JSON.stringify(result)).not.toContain("/foreign/")
  const oversized = new Error(secret.repeat(2000))
  oversized.stack = secret.repeat(2000)
  const bounded = sanitizeNativeDrainFailure(Cause.fromReasons(Array.from({ length: 20 }, () => Cause.die(oversized).reasons[0]!)), secret, repository)
  expect(bounded.reason_count).toBe(20)
  expect(bounded.reasons).toHaveLength(8)
  expect(bounded.reasons_truncated).toBe(true)
  expect(bounded.reasons[0]!.message_fingerprint?.truncated).toBe(true)
  expect(bounded.reasons[0]!.stack_status).toBe("oversized")
  expect(bounded.reasons[0]!.frames).toEqual([])
  expect(JSON.stringify(bounded)).not.toContain(secret)
  expect(Buffer.byteLength(JSON.stringify(bounded))).toBeLessThan(16384)
  expect(() => sanitizeNativeDrainFailure(hostile, undefined, repository)).not.toThrow()
})

test("native failure logger delegates unrelated records and preserves original Cause while retaining one safe record", () => {
  const runtime = mkdtempSync(path.join(os.tmpdir(), "dochaus-safe-failure-"))
  const forwarded: Logger.Options<unknown>[] = []
  const emitted: string[] = []
  const delegate = Logger.make<unknown, void>(options => { forwarded.push(options) })
  const logger = nativeFailureEvidenceLogger({ repository, runtime, loggers: new Set([delegate]), emit: line => { emitted.push(line) } })
  const error = new RangeError("PRIVATE_PAYLOAD")
  try {
    Effect.runSync(Effect.logInfo("unrelated").pipe(Effect.provideService(Logger.CurrentLoggers, new Set([logger]))))
    expect(forwarded).toHaveLength(1)
    expect(forwarded[0]!.message).toEqual(["unrelated"])
    const exit = Effect.runSyncExit(Effect.fail(error).pipe(
      Effect.tapCause(cause => Effect.logError("Failed to drain Session", cause)),
      Effect.annotateLogs({ sessionID: "ses_safe_failure", privileged: "PRIVATE_PAYLOAD" }),
      Effect.provideService(Logger.CurrentLoggers, new Set([logger])),
    ))
    expect(Exit.isFailure(exit)).toBe(true)
    if (Exit.isFailure(exit)) expect(exit.cause.reasons[0]).toBeInstanceOf(Object)
    if (Exit.isFailure(exit)) expect(Cause.squash(exit.cause)).toBe(error)
    expect(forwarded).toHaveLength(1)
    expect(emitted).toHaveLength(1)
    const file = path.join(runtime, "diagnostics", "native-failures.jsonl")
    const text = readFileSync(file, "utf8")
    expect(text).toBe(emitted[0]!)
    expect(text).not.toContain("PRIVATE_PAYLOAD")
    expect(JSON.parse(text).session_id).toBe("ses_safe_failure")
    expect(statSync(file).mode & 0o777).toBe(0o600)
    expect(statSync(path.dirname(file)).mode & 0o777).toBe(0o700)
  } finally { rmSync(runtime, { recursive: true, force: true }) }
})

test("native evidence sink failure emits only its fixed marker and does not change native failure", () => {
  const runtime = mkdtempSync(path.join(os.tmpdir(), "dochaus-failure-sink-"))
  const emitted: string[] = []
  writeFileSync(path.join(runtime, "diagnostics"), "sink is intentionally not a directory")
  const error = new Error("PRIVATE_PAYLOAD")
  const logger = nativeFailureEvidenceLogger({ repository, runtime, loggers: new Set(), emit: line => { emitted.push(line) } })
  try {
    const exit = Effect.runSyncExit(Effect.fail(error).pipe(
      Effect.tapCause(cause => Effect.logError("Failed to drain Session", cause)),
      Effect.provideService(Logger.CurrentLoggers, new Set([logger])),
    ))
    expect(Exit.isFailure(exit)).toBe(true)
    if (Exit.isFailure(exit)) expect(Cause.squash(exit.cause)).toBe(error)
    expect(emitted).toEqual(["DOCHAUS_NATIVE_FAILURE_EVIDENCE_SINK_FAILED\n"])
  } finally { rmSync(runtime, { recursive: true, force: true }) }
})
