import { useId } from "react"
import type { FormAnswer, IntegrationInfo, IntegrationMethod } from "@opencode/client"
import type { Provider } from "../settings/types"
import type { MatchingSignInCandidate } from "./connection-fields"
import { visibleFields } from "./connection-fields"
import { CLOUD_SETUPS } from "../settings/cloud-setups"
import { FieldRow, PasswordInput } from "../settings/controls"
import type { CandidateAuthResult } from "../../api/opencode/provider-candidate"
import "./connection.css"

export type ConnectionDraft = {
  choice: string
  key: string
  answer: FormAnswer
  options: Record<string, string>
  name: string
}
export type ConnectionChoice = { value: string; label: string; method?: IntegrationMethod }

export function connectionChoices(provider: Provider | undefined, integration: IntegrationInfo | undefined, reusable: MatchingSignInCandidate[], endpoint = false): ConnectionChoice[] {
  if (!provider) return []
  const saved: ConnectionChoice[] = (integration?.connections ?? []).map(connection => connection.type === "credential"
    ? { value: `credential:${connection.id}`, label: `${connection.label} · saved ${connection.method === "oauth" ? "sign-in" : "API key"}` }
    : { value: `env:${connection.name}`, label: `Server environment · ${connection.name}` })
  if (provider.host && !integration?.connections.length) saved.push({ value: "host", label: `Server credentials${provider.host.sources.length ? " · detected" : ""}` })
  saved.push(...reusable.map(connection => ({ value: `reuse:${connection.sourceId}:${connection.connectionId}`, label: `${connection.label} · from ${connection.sourceLabel}` })))
  if (endpoint && !integration?.connections.length) saved.push({ value: "host", label: "Endpoint without API key" })
  const methods = [...(integration?.methods ?? []), ...(endpoint && !integration?.methods.some(method => method.type === "key") ? [{ type: "key" as const, label: "New API key" }] : [])]
  return [...saved, ...methods.flatMap((method, index) => method.type === "env" ? [] : [{ value: `method:${index}`, label: method.label || (method.type === "key" ? "New API key" : "Sign in"), method }])]
}

export function methodDefaults(method?: IntegrationMethod): FormAnswer {
  return method && "form" in method ? Object.fromEntries((method.form ?? []).flatMap(field => "default" in field && field.default !== undefined ? [[field.key, field.default]] : [])) : {}
}

export default function ConnectionEditor({ provider, integration, reusable, draft, onChange, disabled = false, local = false, endpoint = false, sourceOptions = {}, attempt, authCode, onAuthCode, onAuthAction }: {
  provider?: Provider
  integration?: IntegrationInfo
  reusable: MatchingSignInCandidate[]
  draft: ConnectionDraft
  onChange: (draft: ConnectionDraft) => void
  disabled?: boolean
  local?: boolean
  endpoint?: boolean
  sourceOptions?: Record<string, string>
  attempt?: CandidateAuthResult & { type: "oauth" | "command" }
  authCode: string
  onAuthCode: (value: string) => void
  onAuthAction: (action: "complete" | "status" | "cancel") => void
}) {
  const id = useId()
  const choices = connectionChoices(provider, integration, reusable, endpoint)
  const method = choices.find(choice => choice.value === draft.choice)?.method
  const spec = CLOUD_SETUPS.find(item => item.ids.includes(provider?.id ?? ""))
  const fields = method && "form" in method ? method.form ?? [] : []
  const setOption = (key: string, value: string) => onChange({ ...draft, options: { ...draft.options, [key]: value } })
  return <div className="ai-setup-connection">
    {local ? <>
      <FieldRow id={`${id}-name`} label="Connection name" hint="A name for this local or custom provider.">
        <input id={`${id}-name`} aria-describedby={`${id}-name-hint`} value={draft.name} disabled={disabled} onChange={event => onChange({ ...draft, name: event.target.value })} placeholder="My local models" />
      </FieldRow>
      <FieldRow id={`${id}-url`} label="Endpoint" hint="An OpenAI-compatible address reachable from the computer running doc.haus.">
        <input id={`${id}-url`} aria-describedby={`${id}-url-hint`} value={draft.options.baseURL ?? ""} disabled={disabled} onChange={event => setOption("baseURL", event.target.value)} placeholder="http://localhost:1234/v1" />
      </FieldRow>
      <FieldRow id={`${id}-key`} label="API key" hint="Leave blank if this endpoint does not require a key.">
        <PasswordInput id={`${id}-key`} aria-describedby={`${id}-key-hint`} value={draft.key} disabled={disabled} onChange={key => onChange({ ...draft, key })} />
      </FieldRow>
    </> : <>
      <FieldRow id={`${id}-method`} label="Connection method" hint="Choose the connection to use. Access is tested after you choose your models.">
        <select id={`${id}-method`} aria-describedby={`${id}-method-hint`} value={draft.choice} disabled={disabled} onChange={event => {
          const method = choices.find(item => item.value === event.target.value)?.method
          onChange({ ...draft, choice: event.target.value, key: "", answer: methodDefaults(method) })
        }}>
          <option value="">Choose a connection method</option>
          {choices.map(choice => <option key={choice.value} value={choice.value}>{choice.label}</option>)}
        </select>
      </FieldRow>
      {endpoint && <FieldRow id={`${id}-url`} label="Endpoint" hint="An OpenAI-compatible address reachable from the computer running doc.haus."><input id={`${id}-url`} aria-describedby={`${id}-url-hint`} disabled={disabled} value={draft.options.baseURL ?? ""} onChange={event => setOption("baseURL", event.target.value)} /></FieldRow>}
      {draft.choice.startsWith("env:") && <p className="settings-hint muted">Uses the detected environment value. When you save, a copy is stored as a connection in this installation.</p>}
      {draft.choice.startsWith("reuse:") && <p className="settings-hint muted">Only this selected sign-in will be reused. It becomes active in this installation when you save.</p>}
      {method?.type === "key" && <FieldRow id={`${id}-key`} label="API key" hint="The new key is used for testing and becomes active when you save.">
        <PasswordInput id={`${id}-key`} aria-describedby={`${id}-key-hint`} value={draft.key} disabled={disabled} onChange={key => onChange({ ...draft, key })} />
      </FieldRow>}
      {visibleFields(fields, draft.answer).map(field => field.type === "external"
        ? <p key={field.key}><a href={field.url} target="_blank" rel="noreferrer">{field.title || field.key}</a> {field.description}</p>
        : <FieldRow key={field.key} id={`${id}-${field.key}`} label={field.title || field.key} hint={field.description}>
          {field.type === "boolean" ? <input id={`${id}-${field.key}`} type="checkbox" disabled={disabled} checked={draft.answer[field.key] === true} onChange={event => onChange({ ...draft, answer: { ...draft.answer, [field.key]: event.target.checked } })} />
            : field.type === "multiselect" ? <select id={`${id}-${field.key}`} multiple required={field.required} disabled={disabled} value={Array.isArray(draft.answer[field.key]) ? draft.answer[field.key] as string[] : []} onChange={event => onChange({ ...draft, answer: { ...draft.answer, [field.key]: Array.from(event.target.selectedOptions, option => option.value) } })}>{field.options.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}</select>
              : field.type === "string" && field.options && !field.custom ? <select id={`${id}-${field.key}`} required={field.required} disabled={disabled} value={String(draft.answer[field.key] ?? "")} onChange={event => onChange({ ...draft, answer: { ...draft.answer, [field.key]: event.target.value } })}><option value="">Choose</option>{field.options.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}</select>
                : <input id={`${id}-${field.key}`} required={field.required} disabled={disabled} type={field.type === "number" || field.type === "integer" ? "number" : "text"} value={String(draft.answer[field.key] ?? "")} min={"minimum" in field ? field.minimum : undefined} max={"maximum" in field ? field.maximum : undefined} step={field.type === "integer" ? 1 : "any"} minLength={field.type === "string" ? field.minLength : undefined} maxLength={field.type === "string" ? field.maxLength : undefined} pattern={field.type === "string" ? field.pattern : undefined} onChange={event => {
                  const answer = { ...draft.answer }
                  if (event.target.value === "") delete answer[field.key]
                  else answer[field.key] = field.type === "number" || field.type === "integer" ? Number(event.target.value) : event.target.value
                  onChange({ ...draft, answer })
                }} />}
        </FieldRow>)}
      {spec && <>
        <p className="settings-hint muted">{provider?.host?.sources.length ? `Detected on the server: ${provider.host.sources.join(", ")}.` : "Uses credentials on the computer running doc.haus."}</p>
        {spec.fields.map(field => <FieldRow key={field.key} id={`${id}-${field.key}`} label={field.label} hint={`${sourceOptions[field.key] ?? (field.fallback ? "Default" : "Required")}${field.required ? ". Required for this provider." : "."}`}>
          {field.options ? <select id={`${id}-${field.key}`} aria-describedby={`${id}-${field.key}-hint`} value={draft.options[field.key] ?? ""} disabled={disabled} onChange={event => setOption(field.key, event.target.value)}><option value="">Choose</option>{field.options.map(value => <option key={value} value={value}>{value}</option>)}</select>
            : <input id={`${id}-${field.key}`} aria-describedby={`${id}-${field.key}-hint`} required={field.required} value={draft.options[field.key] ?? ""} disabled={disabled} placeholder={field.placeholder} onChange={event => setOption(field.key, event.target.value)} />}
        </FieldRow>)}
      </>}
      {!choices.length && <p className="settings-hint muted">No interactive connection method is available. Configure this provider’s credentials on the computer running doc.haus, then refresh discovery.</p>}
      {method?.type === "command" && <p className="settings-hint muted">This starts a sign-in command on the computer running doc.haus: <code>{method.command.join(" ")}</code>.</p>}
    </>}
    {attempt && <div className="ai-setup-auth" aria-live="polite">
      <p>{attempt.status === "complete" ? "Sign-in prepared. Choose your models to test access." : `Sign-in ${attempt.status || "pending"}.`}</p>
      {(attempt.instructions || attempt.message) && <p className="settings-auth-instructions">{attempt.instructions || attempt.message}</p>}
      {attempt.url && /^https?:\/\//.test(attempt.url) && <a href={attempt.url} target="_blank" rel="noreferrer">Open provider sign-in</a>}
      {attempt.type === "oauth" && attempt.mode === "code" && attempt.status !== "complete" && <label>Authorization code<input type="password" autoComplete="off" disabled={disabled} value={authCode} onChange={event => onAuthCode(event.target.value)} /></label>}
      {attempt.status !== "complete" && <div className="ai-setup-actions">
        {attempt.type === "oauth" && <button type="button" disabled={disabled || attempt.mode === "code" && !authCode.trim()} onClick={() => onAuthAction("complete")}>Complete sign-in</button>}
        <button type="button" disabled={disabled} onClick={() => onAuthAction("status")}>Check sign-in status</button>
        <button type="button" className="ghost" disabled={disabled} onClick={() => onAuthAction("cancel")}>Cancel sign-in</button>
      </div>}
    </div>}
  </div>
}
