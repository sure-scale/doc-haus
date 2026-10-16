import { useState } from "react"
import { setModelPair } from "../../api/opencode/config"
import type { ModelOption } from "../../models"
import type { Provider } from "./types"

export default function ModelPickerStep({
  connectedProviders,
  provider,
  changeProvider,
  primary,
  setPrimary,
  providerModels,
  fast,
  setFast,
  setForceConnect,
  setNotice,
  onSaved,
  hideProviderSelector = false,
}: {
  connectedProviders: Provider[]
  provider: string
  changeProvider: (id: string) => void
  primary: string
  setPrimary: (value: string) => void
  providerModels: ModelOption[]
  fast: string
  setFast: (value: string) => void
  setForceConnect: (value: boolean) => void
  setNotice: (value: string) => void
  onSaved: () => Promise<void>
  hideProviderSelector?: boolean
}) {
  const [busy, setBusy] = useState(false)
  const valid = providerModels.some(model => model.value === primary) && providerModels.some(model => model.value === (fast || primary))
  return (
    <>
      {!hideProviderSelector && <p className="settings-notice">
        Detected {connectedProviders.map((p) => p.name).join(", ")}. Pick your models to get started.
      </p>}

      {!hideProviderSelector && <section className="settings-section">
        <h3>Provider</h3>
        <p className="settings-hint muted">
          Both models run on this provider, so they can't end up on different clouds.
        </p>
        <div className="row settings-row">
          <select disabled={busy} value={provider} onChange={(e) => changeProvider(e.target.value)}>
            <option value="">Choose a provider</option>
            {connectedProviders.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </select>
        </div>
      </section>}

      <div className="onboarding-model-grid">
        <section className="settings-section">
          <h3><label htmlFor="onboarding-primary-model">Primary model</label><span className="onboarding-field-tag">Drafting &amp; analysis</span></h3>
          <p id="onboarding-primary-hint" className="settings-hint muted">Handles every matter unless an assistant uses its own model.</p>
          <div className="row settings-row">
            <select id="onboarding-primary-model" aria-describedby="onboarding-primary-hint" disabled={busy || !provider} value={primary} onChange={(e) => setPrimary(e.target.value)}>
              <option value="">Select a model</option>
              {primary && !providerModels.some(model => model.value === primary) && <option value={primary}>{primary} · unavailable</option>}
              {providerModels.map((m) => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
          </div>
        </section>

        <section className="settings-section">
          <h3><label htmlFor="onboarding-fast-model">Fast model</label><span className="onboarding-field-tag">Quick tasks</span></h3>
          <p id="onboarding-fast-hint" className="settings-hint muted">
            A smaller model for quick tasks, including conversation titles and choosing an assistant.
          </p>
          <div className="row settings-row">
            <select id="onboarding-fast-model" aria-describedby="onboarding-fast-hint" disabled={busy || !provider} value={fast} onChange={(e) => setFast(e.target.value)}>
              <option value="">Same as primary</option>
              {fast && !providerModels.some(model => model.value === fast) && <option value={fast}>{fast} · unavailable</option>}
              {providerModels.map((m) => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
          </div>
        </section>
      </div>

      <div className="onboarding-actions">
        <button disabled={busy} onClick={() => setForceConnect(true)}>{hideProviderSelector ? "Back to connection" : "Use a different provider"}</button>
        <button
          className="primary"
          disabled={busy || !valid}
          onClick={async () => {
            if (busy || !valid) return
            setBusy(true)
            try { await setModelPair(primary, fast || primary); await onSaved(); setNotice("") }
            catch (error) { setNotice(error instanceof Error ? error.message : "Could not save models. Retry.") }
            finally { setBusy(false) }
          }}
        >
          {busy ? "Saving models…" : "Save and continue"}{!busy && <span aria-hidden="true"> →</span>}
        </button>
      </div>
    </>
  )
}
