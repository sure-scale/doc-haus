import { useState } from "react"
import type { Client } from "../../api/opencode/client"
import { setProviderOptions } from "../../api/opencode/config"
import { probeProvider } from "../../api/opencode/providers"
import { loadVerified, saveVerified } from "../../providers"
import type { Provider } from "./types"
import { CLOUD_SETUPS, vertexAnthropicBaseURL } from "./cloud-setups"
import NeedsSetupCard from "./NeedsSetupCard"

export default function HostProviderSetup({ client, provider, onChanged }: {
  client: Client; provider: Provider; onChanged: () => Promise<void>
}) {
  const [pending, setPending] = useState(false)
  const [error, setError] = useState("")
  const [notice, setNotice] = useState("")
  const spec = CLOUD_SETUPS.find(item => item.ids.includes(provider.id))
  if (!spec) return null

  async function configure(options: Record<string, string>, modelID?: string) {
    if (pending) return
    setPending(true); setError(""); setNotice("")
    try {
      const verified = loadVerified()
      verified.delete(provider.id); saveVerified(verified)
      await setProviderOptions(provider.id, provider.id === "google-vertex-anthropic" && options.project
        ? { ...options, baseURL: vertexAnthropicBaseURL(options.project, options.location || "global") }
        : options)
      if (modelID) {
        const result = await probeProvider(client, provider.id, modelID)
        if (!result.ok) throw new Error(result.error)
        verified.add(provider.id); saveVerified(verified)
        setNotice("Connection verified. Continue to choose your models.")
      } else setNotice("Connection settings saved. Choose a model below to test access.")
      await onChanged()
    } catch (failure) { setError(failure instanceof Error ? failure.message : "Could not configure this provider. Retry.") }
    finally { setPending(false) }
  }

  return <>
    <NeedsSetupCard spec={{ ...spec, ids: [provider.id], name: provider.name }} providers={[provider]}
      probing={pending} error={error} discoverSuggestions={false}
      onConfigure={options => void configure(options)}
      onEnable={(options, modelID) => void configure(options, modelID)} />
    <button disabled={pending} onClick={() => {
      setError("")
      void onChanged().catch(() => setError("Could not refresh the provider. Retry when the engine is ready."))
    }}>Refresh model catalog</button>
    {notice && <p role="status">{notice}</p>}
  </>
}
