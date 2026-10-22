import { useId } from "react"
import type { ModelOption } from "../../models"
import { FieldRow } from "../settings/controls"
import "./ai-setup.css"

export type ModelFieldsProps = {
  models: ModelOption[]
  primaryModels?: ModelOption[]
  primary: string
  fast: string
  onPrimaryChange: (value: string) => void
  onFastChange: (value: string) => void
  disabled?: boolean
}

export default function ModelFields({ models, primaryModels = models, primary, fast, onPrimaryChange, onFastChange, disabled }: ModelFieldsProps) {
  const id = useId()
  const primaryId = `${id}-primary`
  const fastId = `${id}-fast`
  return <div className="ai-setup-models">
    <FieldRow id={primaryId} label="Primary model" hint="Used for drafting, analysis and assistant replies. Supports text and tools.">
      <select id={primaryId} aria-describedby={`${primaryId}-hint`} value={primary} disabled={disabled} onChange={event => onPrimaryChange(event.target.value)}>
        <option value="">{primaryModels.length ? "Select a model" : "No models available"}</option>
        {primary && !primaryModels.some(item => item.value === primary) && <option value={primary}>{primary} · unavailable</option>}
        {primaryModels.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
      </select>
    </FieldRow>
    <FieldRow id={fastId} label="Fast model" hint="Used for quick background tasks. Choose a smaller model, or use your primary model.">
      <select id={fastId} aria-describedby={`${fastId}-hint`} value={fast} disabled={disabled} onChange={event => onFastChange(event.target.value)}>
        <option value="">Same as primary</option>
        {fast && !models.some(item => item.value === fast) && <option value={fast}>{fast} · unavailable</option>}
        {models.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
      </select>
    </FieldRow>
  </div>
}
