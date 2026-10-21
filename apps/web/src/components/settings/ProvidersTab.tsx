import { useSyncExternalStore } from "react"
import { connectionMessage, providerConnection, providerHealth } from "../../api/opencode/provider-health"
import { addLocalProvider } from "../../api/opencode/config"
import type { Client } from "../../api/opencode/client"
import { isGated } from "../../providers"
import type { CloudSpec } from "./cloud-setups"
import type { Methods, Provider } from "./types"
import type { useToast } from "../Toast"
import NeedsSetupCard from "./NeedsSetupCard"
import LocalEndpoint from "./LocalEndpoint"
import NativeProviderConnect from "./NativeProviderConnect"

export default function ProvidersTab({
  reconnectProvider,
  readyProviders,
  needsSetup,
  disabled,
  methods,
  unverify,
  toggle,
  extraReady,
  nameMemo,
  nameByID,
  setupCards,
  probing,
  probeErr,
  enable,
  client,
  toast,
  load,
}: {
  reconnectProvider?: string
  readyProviders: Provider[]
  needsSetup: Provider[]
  disabled: string[]
  methods: Methods
  unverify: (id: string) => void
  toggle: (id: string, name: string, enable: boolean) => Promise<void>
  extraReady: string[]
  nameMemo: Record<string, string>
  nameByID: Map<string, string>
  setupCards: CloudSpec[]
  probing: string
  probeErr: Record<string, string>
  enable: (targets: Provider[], options?: Record<string, string>, modelID?: string, configureOnly?: boolean) => Promise<void>
  client: Client
  toast: ReturnType<typeof useToast>
  load: () => Promise<void>
}) {
  const connections = useSyncExternalStore(providerHealth.subscribe, providerHealth.getSnapshot)
  return (
    <>
      <section className="settings-section">
        <h3>Saved connections</h3>
        <p className="settings-hint muted">
          Saved credentials do not guarantee access. Your selected models are checked when doc.haus opens and after you reconnect.
        </p>

        <div className="settings-connected">
          {readyProviders.length === 0 && needsSetup.length === 0 && disabled.length === 0 && (
            <span className="muted">None detected yet. Connect one below.</span>
          )}
          {readyProviders.map((p) => (
            <div key={p.id} className="settings-conn">
              <div className="settings-provider-copy">
                <span className="settings-conn-name">{p.name}</span>
                <span className="settings-conn-meta muted">
                  {statusLabel(p, methods)} · {Object.keys(p.models).length} models
                  {providerConnection(connections, p.id) && <span className="settings-provider-health" data-status={providerConnection(connections, p.id)?.status}>{connectionMessage(providerConnection(connections, p.id)!)}</span>}
                </span>
              </div>
              <div className="settings-provider-actions">
                {isGated(p.id) && (
                  <button className="btn-sm ghost" onClick={() => unverify(p.id)}>
                    Reconfigure
                  </button>
                )}
                <button
                  role="switch"
                  aria-checked={true}
                  aria-label={`Disable ${p.name}`}
                  className="switch on"
                  onClick={() => toggle(p.id, p.name, false)}
                />
              </div>
            </div>
          ))}
          {extraReady.map((id) => (
            <div key={id} className="settings-conn">
              <div className="settings-provider-copy">
                <span className="settings-conn-name">{nameMemo[id] ?? id}</span>
                <span className="settings-conn-meta muted">ready</span>
              </div>
              <div className="settings-provider-actions">
                <button
                  role="switch"
                  aria-checked={true}
                  aria-label={`Disable ${nameMemo[id] ?? id}`}
                  className="switch on"
                  onClick={() => toggle(id, nameMemo[id] ?? id, false)}
                />
              </div>
            </div>
          ))}
          {disabled.map((id) => (
            <div key={id} className="settings-conn off">
              <div className="settings-provider-copy">
                <span className="settings-conn-name">{nameByID.get(id) ?? nameMemo[id] ?? id}</span>
                <span className="settings-conn-meta muted">disabled</span>
              </div>
              <div className="settings-provider-actions">
                <button
                  role="switch"
                  aria-checked={false}
                  aria-label={`Enable ${nameByID.get(id) ?? nameMemo[id] ?? id}`}
                  className="switch"
                  onClick={() => toggle(id, nameByID.get(id) ?? nameMemo[id] ?? id, true)}
                />
              </div>
            </div>
          ))}
        </div>
      </section>

      {setupCards.length > 0 && (
        <section className="settings-section">
          <h3>Needs setup</h3>
          <p className="settings-hint muted">
            Vertex and Bedrock sign in from the server (gcloud ADC / AWS), not an API key. Set the project or
            region, sign in on the server, then Enable to verify — until a live check passes they stay out of
            the model picker so you can't pick a provider that will fail on the first call.
          </p>
          {setupCards.map((spec) => {
            const pending = needsSetup.filter((p) => spec.ids.includes(p.id))
            return (
              <NeedsSetupCard
                key={spec.ids[0]}
                spec={spec}
                providers={pending}
                probing={pending.some((p) => probing === p.id)}
                error={probeErr[spec.ids[0]] ?? ""}
                onEnable={(options, modelID) => enable(pending.slice(0, 1), options, modelID)}
                onConfigure={options => enable(pending.slice(0, 1), options, undefined, true)}
              />
            )
          })}
        </section>
      )}

      <NativeProviderConnect client={client} initialProviderID={[...readyProviders, ...needsSetup].find(provider => provider.id === reconnectProvider)?.integrationID ?? reconnectProvider} onConnected={async id => { await load(); if (id) void providerHealth.check(true) }} />

      <LocalEndpoint
        onAdd={async (input) => {
          await addLocalProvider(input)
          toast("success", `Added ${input.name}. Restart the engine if its models don't appear.`)
          await load()
        }}
      />
    </>
  )
}

// Plain-language note on where a ready provider's credentials came from. Env
// providers (Vertex, Bedrock) expose no auth method — their key is read from the
// host sign-in rather than typed here.
function statusLabel(provider: Provider, methods: Methods) {
  const m = methods[provider.id] ?? []
  if (m.some((x) => x.type === "api")) return "API key"
  if (m.some((x) => x.type === "oauth")) return "Signed in"
  return "Host credentials"
}
