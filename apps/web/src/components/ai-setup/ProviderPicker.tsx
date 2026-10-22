import { useId } from "react"
import { FieldRow } from "../settings/controls"
import type { Provider } from "../settings/types"
import "./ai-setup.css"

export type ProviderPickerProps = {
  providers: Provider[]
  provider: string
  onChange: (id: string) => void
  reusableProviderIDs?: readonly string[]
  disabled?: boolean
}

export default function ProviderPicker({ providers, provider, onChange, reusableProviderIDs = [], disabled }: ProviderPickerProps) {
  const id = useId()
  const reusable = (item: Provider) => reusableProviderIDs.includes(item.id) || reusableProviderIDs.includes(item.integrationID ?? item.id)
  const detected = (item: Provider) => item.nativeAvailable || item.credentialSources.length > 0 || Boolean(item.host?.sources.length) || reusable(item)
  const selected = providers.find(item => item.id === provider)
  const detectedProviders = providers.filter(detected)
  const otherProviders = providers.filter(item => !detected(item))
  const sources = selected ? [...new Set([...selected.credentialSources, ...(selected.host?.sources ?? [])])] : []
  const note = provider === "__local_endpoint__"
    ? "Connect a local or custom OpenAI-compatible endpoint."
    : selected
      ? `${sources.length ? `Detected: ${sources.join(", ")}. ` : selected.nativeAvailable ? "Available in this installation. " : ""}${reusable(selected) ? "A saved sign-in is available to reuse. " : ""}${detected(selected) ? "Detection alone does not confirm model access." : "Set up a connection to use this provider."}`
      : provider ? "This saved provider is unavailable. Choose a provider to continue." : ""

  return <div className="ai-setup-provider">
    <FieldRow id={id} label="Provider" hint="Choose the service that will run both models.">
      <select
        id={id}
        aria-describedby={`${id}-hint${note ? ` ${id}-source` : ""}`}
        value={provider}
        disabled={disabled}
        onChange={event => onChange(event.target.value)}
      >
        <option value="">Select a provider</option>
        {provider && provider !== "__local_endpoint__" && !selected && <option value={provider}>{provider} · unavailable</option>}
        {detectedProviders.length > 0 && <optgroup label="Detected">
          {detectedProviders.map(item => <option key={item.id} value={item.id}>{item.name}</option>)}
        </optgroup>}
        <optgroup label="Other providers">
          {otherProviders.map(item => <option key={item.id} value={item.id}>{item.name}</option>)}
          <option value="__local_endpoint__">Local or custom endpoint</option>
        </optgroup>
      </select>
      {note && <p id={`${id}-source`} className="settings-hint muted ai-setup-source">{note}</p>}
    </FieldRow>
  </div>
}
