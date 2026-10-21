import { useState } from "react"
import { providerHealth } from "../../api/opencode/provider-health"
import { setModelPair } from "../../api/opencode/config"
import type { ModelOption } from "../../models"
import type { useToast } from "../Toast"
import type { Provider, Tab } from "./types"
import { FieldRow } from "./controls"

export default function ModelsTab({
  models,
  switchTab,
  provider,
  changeProvider,
  readyProviders,
  model,
  setModel,
  providerModels,
  smallModel,
  setSmallModelValue,
  toast,
  firstRun,
  onClose,
}: {
  models: ModelOption[]
  switchTab: (next: Tab) => void
  provider: string
  changeProvider: (id: string) => void
  readyProviders: Provider[]
  model: string
  setModel: (value: string) => void
  providerModels: ModelOption[]
  smallModel: string
  setSmallModelValue: (value: string) => void
  toast: ReturnType<typeof useToast>
  firstRun: boolean
  onClose: () => void
}) {
  const [saving, setSaving] = useState(false)
  const valid = providerModels.some(item => item.value === model) && providerModels.some(item => item.value === (smallModel || model))
  return (
    <section className="settings-section">
      {models.length === 0 && (
        <p className="settings-hint muted">
          No provider is ready yet.{" "}
          <button className="linklike" onClick={() => switchTab("providers")}>
            Connect one under Providers.
          </button>
        </p>
      )}
      <FieldRow id="settings-model-provider" label="Provider" hint="The service that runs your models. Connect and manage services under Providers.">
        <select id="settings-model-provider" aria-describedby="settings-model-provider-hint" value={provider} onChange={(e) => changeProvider(e.target.value)} disabled={models.length === 0}>
          <option value="">{models.length ? "Select a provider" : "Add a provider first"}</option>
          {readyProviders.map((p) => (
            <option key={p.id} value={p.id}>
              {p.name}
            </option>
          ))}
        </select>
      </FieldRow>
      <FieldRow id="settings-default-model" label="Primary model" hint="The model used for assistant replies.">
        <select id="settings-default-model" aria-describedby="settings-default-model-hint" value={model} onChange={(e) => setModel(e.target.value)} disabled={!provider}>
          <option value="">Select a model</option>
          {model && !providerModels.some(item => item.value === model) && <option value={model}>{model} · unavailable</option>}
          {providerModels.map((m) => (
            <option key={m.value} value={m.value}>
              {m.label}
            </option>
          ))}
        </select>
      </FieldRow>
      <FieldRow id="settings-fast-model" label="Fast model" hint="Used for background tasks, including progress updates and Auto routing.">
        <select id="settings-fast-model" aria-describedby="settings-fast-model-hint" value={smallModel} onChange={(e) => setSmallModelValue(e.target.value)} disabled={!provider}>
          <option value="">Same as primary</option>
          {smallModel && !providerModels.some(item => item.value === smallModel) && <option value={smallModel}>{smallModel} · unavailable</option>}
          {providerModels.map((m) => (
            <option key={m.value} value={m.value}>
              {m.label}
            </option>
          ))}
        </select>
      </FieldRow>
      <div className="settings-actions">
        <button
          className="primary"
          disabled={saving || !valid}
          onClick={async () => {
            if (saving || !valid) return
            setSaving(true)
            try {
              await setModelPair(model, smallModel || model)
              void providerHealth.check(true)
              if (firstRun) return onClose()
              toast("success", "Models saved.")
            } catch (error) { toast("error", error instanceof Error ? error.message : "Could not save models. Retry.") }
            finally { setSaving(false) }
          }}
        >
          Save
        </button>
      </div>
    </section>
  )
}
