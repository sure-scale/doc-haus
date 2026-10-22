// Browser presentation, personal approval choices, unsent library-form answers,
// and migration reads belong here. Matter work uses the ingest service.
function storage(session: boolean) {
  return session ? window.sessionStorage : window.localStorage
}

export function readBrowserValue(key: string, session = false): string | null {
  try {
    return storage(session).getItem(key)
  } catch {
    return null
  }
}

export function writeBrowserValue(key: string, value: string, session = false): boolean {
  try {
    storage(session).setItem(key, value)
    return true
  } catch {
    return false
  }
}

export function removeBrowserValue(key: string): boolean {
  try {
    storage(false).removeItem(key)
    return true
  } catch {
    return false
  }
}

export function readBrowserJSON(key: string): unknown {
  try {
    return JSON.parse(readBrowserValue(key) ?? "null")
  } catch {
    return null
  }
}
