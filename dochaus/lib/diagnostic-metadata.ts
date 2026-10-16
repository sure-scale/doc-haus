import { draftingTraceReference } from "./drafting-trace-identity"

// Operational diagnostics must not serialize arbitrary exceptions: their message,
// cause, stack, custom fields, and even getters can contain privileged content.
// Fingerprints support correlation only; they are not encryption or authority.
export function diagnosticErrorMetadata(error: unknown) {
  if (typeof error === "string") return { error_kind: "string", error_ref: draftingTraceReference(error) }
  if (error === null || typeof error !== "object") return { error_kind: error === null ? "null" : typeof error }
  try {
    const message = Object.getOwnPropertyDescriptor(error, "message")?.value
    return {
      error_kind: error instanceof Error ? "error" : "object",
      ...(typeof message === "string" ? { error_ref: draftingTraceReference(message) } : {}),
    }
  } catch {
    return { error_kind: "uninspectable" }
  }
}
