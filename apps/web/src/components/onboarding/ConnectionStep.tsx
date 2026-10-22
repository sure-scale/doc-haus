import type { Client } from "../../api/opencode/client"
import NativeProviderConnect from "../settings/NativeProviderConnect"
import { MatchingSignIn, useMatchingSignIns } from "../settings/ExistingSignIn"
import type { Provider } from "./types"

export default function ConnectionStep({ client, provider, connected, onConnected }: {
  client: Client
  provider: Provider
  connected: boolean
  onConnected: (integrationID: string) => Promise<void>
}) {
  const integrationID = provider.integrationID ?? provider.id
  const matching = useMatchingSignIns(integrationID)
  const connection = <NativeProviderConnect client={client} integrationID={integrationID} showExistingSignIn={false} onboarding
    savedSignIn={matching.candidates.length > 0 ? <MatchingSignIn candidates={matching.candidates} providerName={provider.name} onReused={onConnected} /> : undefined}
    discoveryNotice={matching.error ? <p className="onboarding-discovery-note">{matching.error} <button className="linklike" onClick={() => void matching.retry()}>Retry saved sign-ins</button></p> : undefined}
    onConnected={onConnected} />

  if (!connected) return connection
  return <>
    <p className="onboarding-connection-status" role="status"><span aria-hidden="true">✓</span><span>{provider.name} is connected. You can continue with this connection or change it below.</span></p>
    <details className="onboarding-connection-options">
      <summary>Change connection (optional)</summary>
      {connection}
    </details>
  </>
}
