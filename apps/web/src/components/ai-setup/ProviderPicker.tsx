import { useId } from "react"
import { FieldRow } from "../settings/controls"
import type { Provider } from "../settings/types"
import "./ai-setup.css"

export function providerDescription(count: number) {
  const rounded = Math.floor((count - 1) / 10) * 10
  return rounded > 0 ? `Choose from over ${rounded} cloud providers or run models locally.` : "Choose a cloud provider or run models locally."
}

export type ProviderPickerProps = {
  providers: Provider[]
  provider: string
  onChange: (id: string) => void
  onDetectedSelect?: (id: string) => void
  reusableProviderIDs?: readonly string[]
  disabled?: boolean
}

export default function ProviderPicker({ providers, provider, onChange, onDetectedSelect, reusableProviderIDs = [], disabled }: ProviderPickerProps) {
  const id = useId()
  const reusable = (item: Provider) => reusableProviderIDs.includes(item.id) || reusableProviderIDs.includes(item.integrationID ?? item.id)
  const detected = (item: Provider) => item.nativeAvailable || item.credentialSources.length > 0 || Boolean(item.host?.sources.length) || reusable(item)
  const selected = providers.find(item => item.id === provider)
  const sortedProviders = [...providers].sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base", numeric: true }))
  const sources = selected ? [...new Set([...selected.credentialSources, ...(selected.host?.sources ?? [])])] : []
  const note = provider === "__local_endpoint__"
    ? "Connect a local or custom OpenAI-compatible endpoint."
    : selected
      ? `${sources.length ? `Detected: ${sources.join(", ")}. ` : selected.nativeAvailable ? "Available in this installation. " : ""}${reusable(selected) ? "A saved sign-in is available to reuse. " : ""}${detected(selected) ? "Detection alone does not confirm model access." : "Set up a connection to use this provider."}`
      : provider ? "This saved provider is unavailable. Choose a provider to continue." : ""

  const detectedProviders = sortedProviders.filter(detected)

  return <div className="ai-setup-provider">
    {detectedProviders.length > 0 && <div className="ai-setup-detected" role="group" aria-label="We detected these on your system already">
      <p>We detected these on your system already</p>
      <div className="ai-setup-detected-options">{detectedProviders.map(item => <button type="button" key={item.id} disabled={disabled} aria-pressed={provider === item.id} onClick={() => (onDetectedSelect ?? onChange)(item.id)}>{item.name}</button>)}</div>
    </div>}
    <FieldRow id={id} label="Provider">
      <select
        id={id}
        aria-describedby={note ? `${id}-source` : undefined}
        value={provider}
        disabled={disabled}
        onChange={event => onChange(event.target.value)}
      >
        <option value="">Select a provider</option>
        {provider && provider !== "__local_endpoint__" && !selected && <option value={provider}>{provider} · unavailable</option>}
        <option value="__local_endpoint__">Local or custom endpoint</option>
        {sortedProviders.map(item => <option key={item.id} value={item.id}>{item.name}</option>)}
      </select>
      {note && <p id={`${id}-source`} className="settings-hint muted ai-setup-source">{note}</p>}
    </FieldRow>
  </div>
}
