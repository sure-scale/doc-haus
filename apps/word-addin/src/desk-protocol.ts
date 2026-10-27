import type { WordSelectionSnapshot } from "./selection"

export type DeskView = "documents" | "compare"

export function isDeskView(value: unknown): value is DeskView {
  return value === "documents" || value === "compare"
}

export type DeskRequest =
  | { action: "selection" }
  | { action: "show"; selection: WordSelectionSnapshot }
  | { action: "sources" }
  | { action: "source"; sourceId: string; expectedHash: string }
  | { action: "versions" }
  | { action: "compare"; baseVersion: string; currentVersion: string }

const PROTOCOL = "dochaus.source-desk"
export const DESK_MESSAGE_LIMIT = 1024 * 1024

type Base = { protocol: typeof PROTOCOL; version: 1; channel: string }
export type DeskPayload =
  | { kind: "ready" | "connected" | "disconnect" }
  | { kind: "request"; id: string; request: DeskRequest }
  | { kind: "response"; id: string; ok: true; value: unknown }
  | { kind: "response"; id: string; ok: false; error: string }
export type DeskMessage = Base & DeskPayload

function record(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value))
}

function keys(value: Record<string, unknown>, allowed: string[]): boolean {
  return Object.keys(value).every((key) => allowed.includes(key))
}

function identifier(value: unknown): value is string {
  return typeof value === "string" && /^[\w-]{1,128}$/.test(value)
}

function reference(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= 512 && !/[\u0000-\u001f]/.test(value)
}

function snapshot(value: unknown): value is WordSelectionSnapshot {
  return (
    record(value) &&
    keys(value, ["text", "documentHash", "saved", "capturedAt", "navigable", "reason"]) &&
    typeof value.text === "string" &&
    value.text.length <= 10_000 &&
    typeof value.documentHash === "string" &&
    /^[a-f0-9]{64}$/.test(value.documentHash) &&
    typeof value.saved === "boolean" &&
    typeof value.navigable === "boolean" &&
    typeof value.capturedAt === "string" &&
    value.capturedAt.length <= 64 &&
    Number.isFinite(Date.parse(value.capturedAt)) &&
    (value.reason === undefined || (typeof value.reason === "string" && value.reason.length <= 500))
  )
}

export function isDeskRequest(value: unknown): value is DeskRequest {
  if (!record(value)) return false
  switch (value.action) {
    case "selection":
    case "sources":
    case "versions":
      return keys(value, ["action"])
    case "show":
      return keys(value, ["action", "selection"]) && snapshot(value.selection)
    case "source":
      return (
        keys(value, ["action", "sourceId", "expectedHash"]) &&
        reference(value.sourceId) &&
        typeof value.expectedHash === "string" &&
        /^[a-f0-9]{64}$/.test(value.expectedHash)
      )
    case "compare":
      return (
        keys(value, ["action", "baseVersion", "currentVersion"]) &&
        reference(value.baseVersion) &&
        reference(value.currentVersion) &&
        value.baseVersion !== value.currentVersion
      )
    default:
      return false
  }
}

export function deskMessage(channel: string, value: DeskPayload): string {
  const encoded = JSON.stringify({ protocol: PROTOCOL, version: 1, channel, ...value })
  if (encoded.length > DESK_MESSAGE_LIMIT) throw new Error("This preview is too large. Choose a smaller document.")
  return encoded
}

export function readDeskMessage(encoded: unknown): DeskMessage | undefined {
  if (typeof encoded !== "string" || encoded.length > DESK_MESSAGE_LIMIT) return
  let value: unknown
  try {
    value = JSON.parse(encoded)
  } catch {
    return
  }
  if (!record(value) || value.protocol !== PROTOCOL || value.version !== 1 || !identifier(value.channel)) return
  const base = ["protocol", "version", "channel", "kind"]
  if (["ready", "connected", "disconnect"].includes(String(value.kind))) {
    if (keys(value, base)) return value as DeskMessage
  } else if (value.kind === "request") {
    if (keys(value, [...base, "id", "request"]) && identifier(value.id) && isDeskRequest(value.request))
      return value as DeskMessage
  } else if (value.kind === "response" && identifier(value.id)) {
    if (value.ok === true && keys(value, [...base, "id", "ok", "value"])) return value as DeskMessage
    if (
      value.ok === false &&
      keys(value, [...base, "id", "ok", "error"]) &&
      typeof value.error === "string" &&
      value.error.length <= 500
    )
      return value as DeskMessage
  }
}
