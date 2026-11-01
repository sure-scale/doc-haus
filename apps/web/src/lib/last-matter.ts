import { readBrowserValue, writeBrowserValue } from "../browser-storage"

// The matter this browser opened last, so pickers that send work into a matter
// start on the one the lawyer was just in. A per-browser convenience: without
// storage, pickers fall back to their own order.
const KEY = "dochaus.lastMatter"

export function rememberLastMatter(id: string) {
  writeBrowserValue(KEY, id)
}

/** Matters newest first, with the one opened last at the top. */
export function lastOpenedFirst<T extends { id: string; created_at: number }>(matters: T[]): T[] {
  const last = readBrowserValue(KEY)
  return [...matters].sort((a, b) => Number(b.id === last) - Number(a.id === last) || b.created_at - a.created_at)
}
