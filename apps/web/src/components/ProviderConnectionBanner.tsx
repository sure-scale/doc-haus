import { useEffect, useRef, useState, useSyncExternalStore } from "react"
import { providerHealth, type ModelConnection } from "../api/opencode/provider-health"

// Only rejected credentials or an unusable model need the locked recovery Settings. Credit and
// network trouble (`unavailable`) cannot be fixed there, so it gets a calm banner instead.
export function recoveryConnection(connections: readonly ModelConnection[]) {
  return connections.find(item => item.status === "auth") ?? connections.find(item => item.status === "model")
}

const failureKey = (item?: ModelConnection) => item && `${item.providerID}/${item.modelID}/${item.status}`

export default function ProviderConnectionBanner({ enabled, onOpenSettings }: { enabled: boolean; onOpenSettings: (providerID: string) => void }) {
  const connections = useSyncExternalStore(providerHealth.subscribe, providerHealth.getSnapshot)
  useEffect(() => { if (enabled && providerHealth.getSnapshot().length === 0) void providerHealth.check() }, [enabled])
  const active = useRef(enabled)
  active.current = enabled
  // The unavailable failure a recheck has confirmed; it shows the banner and stops further automatic rechecks.
  const [settled, setSettled] = useState<string>()
  const [checking, setChecking] = useState(false)
  const failed = recoveryConnection(connections) ?? connections.find(item => item.status === "unavailable")
  const key = failureKey(failed)
  const settle = () => {
    if (!active.current) return
    const snapshot = providerHealth.getSnapshot()
    const recovery = recoveryConnection(snapshot)
    if (recovery) onOpenSettings(recovery.providerID)
    else setSettled(failureKey(snapshot.find(item => item.status === "unavailable")))
  }
  useEffect(() => {
    if (connections.length > 0 && connections.every(item => item.status === "ready")) setSettled(undefined)
  }, [connections])
  useEffect(() => {
    if (!enabled || !failed || key === settled) return
    const timer = window.setTimeout(() => { void providerHealth.check(true).then(settle) }, 1000)
    return () => window.clearTimeout(timer)
  }, [enabled, key, settled, onOpenSettings])
  if (!enabled || !(checking || (failed?.status === "unavailable" && key === settled))) return null
  const retry = () => {
    setChecking(true)
    void providerHealth.check(true).then(settle).finally(() => setChecking(false))
  }
  return <div className="provider-connection-banner" role="status" aria-live="polite">
    <div>
      <strong>The AI model can’t be reached right now.</strong>
      <p>Check your network or your AI account’s credit. Your work is saved.</p>
    </div>
    <div className="provider-connection-actions">
      <button type="button" disabled={checking} onClick={retry}>{checking ? "Checking…" : "Check again"}</button>
    </div>
  </div>
}
