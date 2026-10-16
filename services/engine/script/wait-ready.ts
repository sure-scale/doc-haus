import path from "node:path"

const directory = path.resolve(process.env.WORKSPACE_ROOT ?? "workspace")
const hostname = process.env.OPENCODE_HOSTNAME ?? "127.0.0.1"
const host = hostname === "0.0.0.0" ? "127.0.0.1" : hostname.includes(":") ? `[${hostname === "::" ? "::1" : hostname}]` : hostname
const url = `http://${host}:${process.env.DOC_HAUS_ENGINE_PORT ?? "4097"}/api/dochaus/ready`
const headers: Record<string, string> = { "x-opencode-directory": directory }
if (process.env.OPENCODE_SERVER_PASSWORD)
  headers.authorization = `Basic ${btoa(`opencode:${process.env.OPENCODE_SERVER_PASSWORD}`)}`
type Outcome = "active" | "plugin_failed" | "consent_failed" | "auth_failed" | "deadline" | "json_failure"
type TransportCategory = "timeout" | "request_error" | null
type Envelope = "valid_data" | "invalid_envelope" | "invalid_json" | null
type LegalState = "active" | "error" | "failed" | "loading" | "inactive" | "absent" | "other" | null
const observation = process.env.DOCHAUS_STARTUP_OBSERVATION === "1"
const startedAt = Date.now()
let attempts = 0
let lastHTTPStatus: number | null = null
let transportCategory: TransportCategory = null
let envelope: Envelope = null
let legalState: LegalState = null
let legalPluginCount: number | null = null
let consentState: string | null = null
function receipt(outcome: Outcome) {
  if (!observation) return
  console.log(
    `DOCHAUS_ENGINE_READINESS ${JSON.stringify({
      outcome,
      attempts,
      elapsedMs: Date.now() - startedAt,
      lastHttpStatus: lastHTTPStatus,
      transportCategory,
      envelope,
      legalState,
      legalPluginCount,
      consentState,
    })}`,
  )
}
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value)
}
function state(value: unknown): LegalState {
  if (value === "active" || value === "error" || value === "failed" || value === "loading" || value === "inactive") return value
  return "other"
}
const deadline = Date.now() + 60_000
while (true) {
  attempts++
  lastHTTPStatus = null
  transportCategory = null
  envelope = null
  legalState = null
  legalPluginCount = null
  consentState = null
  const response = await fetch(url, { headers, signal: AbortSignal.timeout(5000) }).catch((error: unknown) => {
    transportCategory = error instanceof Error && error.name === "TimeoutError" ? "timeout" : "request_error"
    return undefined
  })
  if (response) lastHTTPStatus = response.status
  if (response?.status === 401) {
    receipt("auth_failed")
    throw new Error("Engine authentication failed before readiness could be checked")
  }
  if (response?.ok || response?.status === 503) {
    let inventory: unknown
    try {
      inventory = await response.json()
    } catch (error) {
      envelope = "invalid_json"
      receipt("json_failure")
      throw error
    }
    if (isRecord(inventory) && typeof inventory.ready === "boolean" && typeof inventory.legalState === "string" && typeof inventory.consentState === "string") {
      envelope = "valid_data"
      legalPluginCount = typeof inventory.legalPluginCount === "number" ? inventory.legalPluginCount : null
      legalState = inventory.legalState === "absent" ? "absent" : state(inventory.legalState)
      consentState = inventory.consentState
      if (response.ok && inventory.ready && legalState === "active" && consentState === "active") {
        receipt("active")
        break
      }
      if (legalState === "error" || legalState === "failed") {
        receipt("plugin_failed")
        throw new Error("The legal plugin failed to activate; inspect the engine log")
      }
      if (consentState === "failed" || consentState === "missing") {
        receipt("consent_failed")
        throw new Error("Research consent protection is missing or failed; inspect the engine log")
      }
    } else {
      envelope = "invalid_envelope"
    }
  }
  if (Date.now() >= deadline) {
    receipt("deadline")
    throw new Error("The legal engine did not become ready within 60 seconds")
  }
  await Bun.sleep(250)
}
console.log("doc.haus: native engine, legal plugin and research consent protection are ready")
