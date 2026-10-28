import { useEffect, useRef, useState, type Dispatch, type SetStateAction } from "react"
import type { LibraryAttachment } from "../../api/ingest/library"

// A pending attachment belongs to this browser's draft for one conversation.
// Keep only references here; the source bytes stay in the assistant directory.
function keyFor(directory: string, sessionID?: string) {
  return `dh.library-attachments:${JSON.stringify([directory, sessionID ?? null])}`
}
function read(key: string): LibraryAttachment[] {
  try {
    const value: unknown = JSON.parse(sessionStorage.getItem(key) ?? "[]")
    return Array.isArray(value) ? value.filter((item): item is LibraryAttachment =>
      item && typeof item.id === "string" && typeof item.name === "string" && typeof item.source === "string" && typeof item.original === "string") : []
  } catch { return [] }
}
function write(key: string, value: LibraryAttachment[]) {
  try {
    if (value.length) sessionStorage.setItem(key, JSON.stringify(value))
    else sessionStorage.removeItem(key)
  } catch { /* The composer remains usable if browser storage is unavailable. */ }
}

export function useLibraryAttachments(directory: string, sessionID?: string): [LibraryAttachment[], Dispatch<SetStateAction<LibraryAttachment[]>>] {
  const key = keyFor(directory, sessionID)
  const keyRef = useRef(key)
  const [attachments, setState] = useState<LibraryAttachment[]>(() => read(key))
  const current = useRef(attachments)
  useEffect(() => {
    if (keyRef.current !== key) {
      // The first session ID is assigned to this same composer after sending.
      write(keyRef.current, [])
      keyRef.current = key
      write(key, current.current)
    }
  }, [key])
  const setAttachments: Dispatch<SetStateAction<LibraryAttachment[]>> = (update) => {
    const next = typeof update === "function" ? update(current.current) : update
    current.current = next
    write(keyRef.current, next)
    setState(next)
  }
  return [attachments, setAttachments]
}
