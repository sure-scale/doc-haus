import type { WordSelectionSnapshot } from "./selection"
import type { WordChatAttachment } from "../../../dochaus/lib/word-chat-contract"

export type DeskView = "documents" | "compare"

export function isDeskView(value: unknown): value is DeskView {
  return value === "documents" || value === "compare"
}

/** Attachment mode uses the same bridge, but never mounts the document desk. */
export function isDeskDialogView(value: unknown): value is DeskView | "attachments" {
  return isDeskView(value) || value === "attachments"
}

export type DeskRequest =
  | { action: "selection" }
  | { action: "show"; selection: WordSelectionSnapshot }
  | { action: "sources" }
  | { action: "source"; sourceId: string; expectedHash: string }
  | { action: "open"; sourceId: string; expectedHash: string }
  | { action: "versions" }
  | { action: "compare"; baseVersion: string; currentVersion: string }
  | { action: "chat-materials" }
  | { action: "chat-attach"; material: WordChatAttachment }
  | { action: "chat-upload"; name: string; data: string; mode: "request" | "matter" }

const PROTOCOL = "dochaus.source-desk"
export const DESK_MESSAGE_LIMIT = 1024 * 1024
export const CHAT_UPLOAD_MAX_BYTES = 16 * 1024 * 1024
export const CHAT_UPLOAD_MAX_BASE64 = Math.ceil(CHAT_UPLOAD_MAX_BYTES / 3) * 4
// Only a validated upload request gets this allowance. Existing previews,
// responses and every other request retain the 1 MiB boundary.
export const DESK_UPLOAD_MESSAGE_LIMIT = CHAT_UPLOAD_MAX_BASE64 + 8192

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

export function isChatUploadName(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= 256 &&
    !/[\\/\u0000-\u001f\u007f]/.test(value) && /\.(docx|pdf|txt|md|csv|xlsx|rtf)$/i.test(value)
}

function attachment(value: unknown): value is WordChatAttachment {
  return record(value) && keys(value, ["kind", "id", "hash"]) &&
    ["source", "playbook", "upload"].includes(String(value.kind)) && reference(value.id) &&
    typeof value.hash === "string" && /^[a-f0-9]{64}$/.test(value.hash)
}

function uploadData(value: unknown): value is string {
  if (typeof value !== "string" || !value.length || value.length > CHAT_UPLOAD_MAX_BASE64 || value.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) return false
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0
  return value.length / 4 * 3 - padding <= CHAT_UPLOAD_MAX_BYTES
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
    case "chat-materials":
      return keys(value, ["action"])
    case "chat-attach":
      return keys(value, ["action", "material"]) && attachment(value.material)
    case "chat-upload":
      return keys(value, ["action", "name", "data", "mode"]) && isChatUploadName(value.name) &&
        uploadData(value.data) && (value.mode === "request" || value.mode === "matter")
    case "show":
      return keys(value, ["action", "selection"]) && snapshot(value.selection)
    case "source":
    case "open":
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
  const limit = value.kind === "request" && value.request.action === "chat-upload" && isDeskRequest(value.request)
    ? DESK_UPLOAD_MESSAGE_LIMIT : DESK_MESSAGE_LIMIT
  if (encoded.length > limit) throw new Error("This file is too large. Choose a smaller document.")
  return encoded
}

export function readDeskMessage(encoded: unknown): DeskMessage | undefined {
  if (typeof encoded !== "string" || encoded.length > DESK_UPLOAD_MESSAGE_LIMIT) return
  let value: unknown
  try {
    value = JSON.parse(encoded)
  } catch {
    return
  }
  if (!record(value) || value.protocol !== PROTOCOL || value.version !== 1 || !identifier(value.channel)) return
  if (encoded.length > DESK_MESSAGE_LIMIT &&
    !(value.kind === "request" && isDeskRequest(value.request) && value.request.action === "chat-upload")) return
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
