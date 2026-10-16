import { afterEach, expect, spyOn, test } from "bun:test"
import { mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import path from "node:path"
import { diagnosticErrorMetadata } from "../lib/diagnostic-metadata"
import { draftingTraceReference } from "../lib/drafting-trace-identity"
import { recordContextBudget, recordTelemetrySession, recordToolCallTelemetry, touchTelemetrySession } from "../plugin/legal/telemetry"

const directories: string[] = []
afterEach(() => directories.splice(0).forEach((directory) => rmSync(directory, { recursive: true, force: true })))

test("diagnostic errors retain fingerprints without inspecting arbitrary payloads, causes, names or getters", () => {
  const secret = "PRIVATE-EXCEPTION-TEXT"
  const error = new Error(secret, { cause: { request: secret, response: secret } })
  error.name = secret
  const metadata = diagnosticErrorMetadata(error)
  expect(metadata).toEqual({ error_kind: "error", error_ref: draftingTraceReference(secret) })
  expect(JSON.stringify(metadata)).not.toContain(secret)
  expect(diagnosticErrorMetadata(secret)).toEqual({ error_kind: "string", error_ref: draftingTraceReference(secret) })
  expect(diagnosticErrorMetadata({ get message() { throw new Error(secret) }, toJSON() { throw new Error(secret) } })).toEqual({ error_kind: "object" })
  expect(diagnosticErrorMetadata(new Proxy({}, { getOwnPropertyDescriptor() { throw new Error(secret) } }))).toEqual({ error_kind: "uninspectable" })
})

test("real telemetry storage failures never print matter paths or exception bodies", () => {
  const secret = "PRIVATE-MATTER-PATH"
  const directory = mkdtempSync(path.join(tmpdir(), `${secret}-`))
  directories.push(directory)
  writeFileSync(path.join(directory, ".dochaus"), "A file deliberately prevents creation of the telemetry directory.")
  const log = spyOn(console, "error").mockImplementation(() => {})
  try {
    recordTelemetrySession("diagnostic-session", { directory, agent: "draft", model: "configured/model" })
    recordContextBudget(directory, [{ info: { sessionID: "diagnostic-session" }, parts: [] }])
    touchTelemetrySession(directory, "diagnostic-session", "draft", "configured/model")
    recordToolCallTelemetry(directory, { tool: "read", sessionID: "diagnostic-session" }, { output: secret })
    expect(log).toHaveBeenCalledTimes(3)
    const logged = JSON.stringify(log.mock.calls)
    expect(logged).not.toContain(secret)
    expect(logged).not.toContain(directory)
    expect(logged).toContain("error_ref")
    expect(logged).toContain("recordTurn failed")
    expect(logged).toContain("touchSession failed")
    expect(logged).toContain("recordToolCall failed")
  } finally {
    log.mockRestore()
  }
})
