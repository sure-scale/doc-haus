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
  showEffort?: boolean
}

export default function ModelFields({ models, primaryModels = models, primary, fast, onPrimaryChange, onFastChange, disabled, showEffort = false }: ModelFieldsProps) {
  const id = useId()
  const primaryId = `${id}-primary`
  const fastId = `${id}-fast`
  const base = (value: string) => value.split("#", 1)[0] ?? ""
  const baseModels = (options: ModelOption[]) => options.filter(item => !item.value.includes("#"))
  const effort = (value: string, options: ModelOption[], role: "Primary" | "Fast", onChange: (value: string) => void) => {
    const model = base(value)
    const variants = options.filter(item => item.value.startsWith(`${model}#`))
    const current = value.includes("#") ? value.slice(model.length + 1) : ""
    if (!showEffort || !model || (!variants.length && !current)) return null
    const effortId = `${id}-${role.toLowerCase()}-effort`
    return <FieldRow id={effortId} label={`${role} reasoning effort`} hint="Higher effort may take longer and cost more.">
      <select id={effortId} aria-describedby={`${effortId}-hint`} value={current} disabled={disabled} onChange={event => onChange(event.target.value ? `${model}#${event.target.value}` : model)}>
        <option value="">Default</option>
        {current && !variants.some(item => item.value === value) && <option value={current}>{current} · unavailable</option>}
        {variants.map(item => <option key={item.value} value={item.value.slice(model.length + 1)}>{item.label.split(" · ").at(-1)}</option>)}
      </select>
    </FieldRow>
  }
  const primaryBase = base(primary), fastBase = base(fast)
  const primaryOptions = baseModels(primaryModels), fastOptions = baseModels(models)
  return <div className="ai-setup-models">
    <FieldRow id={primaryId} label="Primary model" hint="Used for drafting, analysis and assistant replies. Supports text and tools.">
      <select id={primaryId} aria-describedby={`${primaryId}-hint`} value={primaryBase} disabled={disabled} onChange={event => onPrimaryChange(event.target.value)}>
        <option value="">{primaryOptions.length ? "Select a model" : "No models available"}</option>
        {primaryBase && !primaryOptions.some(item => item.value === primaryBase) && <option value={primaryBase}>{primaryBase} · unavailable</option>}
        {primaryOptions.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
      </select>
    </FieldRow>
    {effort(primary, primaryModels, "Primary", onPrimaryChange)}
    <FieldRow id={fastId} label="Fast model" hint="Used for quick background tasks. Choose a smaller model, or use your primary model.">
      <select id={fastId} aria-describedby={`${fastId}-hint`} value={fastBase} disabled={disabled} onChange={event => onFastChange(event.target.value)}>
        <option value="">Same as primary</option>
        {fastBase && !fastOptions.some(item => item.value === fastBase) && <option value={fastBase}>{fastBase} · unavailable</option>}
        {fastOptions.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
      </select>
    </FieldRow>
    {fast && effort(fast, models, "Fast", onFastChange)}
  </div>
}
