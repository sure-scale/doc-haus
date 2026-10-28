import { useRef, type RefObject } from "react"

// Keep a new composer's state when its first session appears in the URL, but
// give subsequent navigation (including back to a new chat) a fresh panel.
export function useLibraryChatKey(session: string | undefined, prompt: string | undefined, created: RefObject<string | undefined>) {
  const route = session ? `session:${session}` : `new:${prompt ?? ""}`
  const previous = useRef({ route, session, key: 0 })
  if (previous.current.route !== route) {
    const firstSession = !previous.current.session && session && session === created.current
    if (!firstSession) {
      previous.current.key++
      created.current = undefined
    }
    previous.current.route = route
    previous.current.session = session
  }
  return previous.current.key
}
