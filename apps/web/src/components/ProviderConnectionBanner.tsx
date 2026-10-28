import { useEffect, useRef, useSyncExternalStore } from "react"
import { providerHealth } from "../api/opencode/provider-health"

export default function ProviderConnectionBanner({ enabled, onOpenSettings }: { enabled: boolean; onOpenSettings: (providerID: string) => void }) {
  const connections = useSyncExternalStore(providerHealth.subscribe, providerHealth.getSnapshot)
  useEffect(() => { if (enabled && providerHealth.getSnapshot().length === 0) void providerHealth.check() }, [enabled])
  const active = useRef(enabled)
  active.current = enabled
  const failed = connections.find(item => item.status === "auth") ?? connections.find(item => item.status === "model")
    ?? connections.find(item => item.status === "unavailable")
  useEffect(() => {
    if (!enabled || !failed) return
    const timer = window.setTimeout(() => {
      void providerHealth.check(true).then(() => {
        if (!active.current) return
        const remaining = providerHealth.getSnapshot().find(item => ["auth", "model", "unavailable"].includes(item.status))
        if (remaining) onOpenSettings(remaining.providerID)
      })
    }, 1000)
    return () => window.clearTimeout(timer)
  }, [enabled, failed?.providerID, failed?.modelID, failed?.status, onOpenSettings])
  return null
}
