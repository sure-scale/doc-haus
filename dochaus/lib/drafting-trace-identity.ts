import { createHash } from "node:crypto"

// Trace references preserve equality checks without persisting the source value.
// They are diagnostic fingerprints, not encryption or independent evidence.
export function draftingTraceReference(value: string) {
  return `sha256:${createHash("sha256").update(value).digest("hex")}`
}

export function isDraftingTraceReference(value: unknown): value is string {
  return typeof value === "string" && /^sha256:[a-f0-9]{64}$/.test(value)
}

export function normalizedDraftingTraceURL(value: unknown) {
  const raw = typeof value === "string" ? value.trim() : ""
  if (!raw) return ""
  try {
    const url = new URL(raw)
    if (!["http:", "https:"].includes(url.protocol)) return ""
    url.hash = ""
    url.searchParams.sort()
    return url.toString()
  } catch {
    return ""
  }
}
