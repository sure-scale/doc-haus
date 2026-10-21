import { useEffect, useSyncExternalStore } from "react"
import { connectionMessage, providerHealth, providerName } from "../api/opencode/provider-health"

export default function ProviderConnectionBanner({ enabled, onOpenSettings }: { enabled: boolean; onOpenSettings: (providerID: string) => void }) {
  const connections = useSyncExternalStore(providerHealth.subscribe, providerHealth.getSnapshot)
  useEffect(() => { if (enabled) void providerHealth.check() }, [enabled])
  if (!enabled) return null
  const failed = connections.find(item => item.status === "auth") ?? connections.find(item => item.status === "model")
    ?? connections.find(item => item.status === "unavailable")
  if (!failed) return null
  return <div className="provider-connection-banner" role="alert">
    <div><strong>{failed.providerID ? `${providerName(failed.providerID)} needs attention` : "Check model connection"}</strong><p>{connectionMessage(failed)} Your work is saved.</p></div>
    <div className="provider-connection-actions">
      <button type="button" onClick={() => onOpenSettings(failed.providerID)}>{failed.status === "auth" ? "Reconnect provider" : "Open provider settings"}</button>
      <button type="button" className="ghost" onClick={() => void providerHealth.check(true)}>Check again</button>
    </div>
  </div>
}
