import { useEffect, useState } from "react"

// Advisory only: tabs in the same browser profile can see each other. The
// server version check remains responsible for protecting every save.
export default function useReviewGridPresence(directory: string, matterId: string) {
  const [anotherWindow, setAnotherWindow] = useState(false)
  useEffect(() => {
    setAnotherWindow(false)
    if (typeof BroadcastChannel === "undefined") return
    let channel: BroadcastChannel
    try {
      channel = new BroadcastChannel(`dochaus-review-grid:${JSON.stringify([directory, matterId])}`)
    } catch {
      return
    }
    const id = crypto.randomUUID()
    const peers = new Map<string, number>()
    const post = (kind: "hello" | "present" | "bye") => channel.postMessage({ kind, id })
    const refresh = () => {
      const cutoff = Date.now() - 15000
      for (const [peer, seen] of peers) if (seen < cutoff) peers.delete(peer)
      setAnotherWindow(peers.size > 0)
    }
    channel.onmessage = ({ data }) => {
      if (!data || typeof data.id !== "string" || data.id === id) return
      if (data.kind === "bye") peers.delete(data.id)
      else if (data.kind === "hello" || data.kind === "present") {
        peers.set(data.id, Date.now())
        if (data.kind === "hello") post("present")
      } else return
      refresh()
    }
    const leave = () => post("bye")
    const returnToPage = () => post("hello")
    window.addEventListener("pagehide", leave)
    window.addEventListener("pageshow", returnToPage)
    post("hello")
    const timer = window.setInterval(() => { post("present"); refresh() }, 5000)
    return () => {
      window.clearInterval(timer)
      window.removeEventListener("pagehide", leave)
      window.removeEventListener("pageshow", returnToPage)
      leave()
      channel.close()
    }
  }, [directory, matterId])
  return anotherWindow
}
