// Relative proxy prefixes resolve before the native client validates its URL.
function backendURL(value: string) {
  const origin = typeof window === "undefined" ? "http://localhost" : window.location.origin
  return new URL(value, origin).href.replace(/\/$/, "")
}

export const OPENCODE_URL = backendURL(import.meta.env.VITE_OPENCODE_URL ?? "http://localhost:4096")
export const INGEST_URL = backendURL(import.meta.env.VITE_INGEST_URL ?? "http://localhost:4500")
