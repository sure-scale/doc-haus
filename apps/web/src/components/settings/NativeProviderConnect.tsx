import { useEffect, useState, type ReactNode } from "react"
import type { FormAnswer, FormField, IntegrationInfo, IntegrationMethod } from "@opencode/client"
import { location, type Client } from "../../api/opencode/client"
import { listIntegrations } from "../../api/opencode/providers"
import { OPENCODE_URL } from "../../config"
import { ExistingSignInDisclosure } from "./ExistingSignIn"

type Attempt = { integrationID: string; attemptID: string; type: "oauth" | "command"; mode?: "auto" | "code" }
const attemptKey = `dochaus.auth-attempt:${OPENCODE_URL}`
function savedAttempt(): Attempt | null {
  try {
    const value = JSON.parse(sessionStorage.getItem(attemptKey) ?? "null")
    return value && typeof value.integrationID === "string" && typeof value.attemptID === "string" && ["oauth", "command"].includes(value.type) ? value : null
  } catch { return null }
}
export function visibleFields(fields: readonly FormField[], answer: FormAnswer, includeHidden = false) {
  return fields.filter(field => (includeHidden || !("hidden" in field && field.hidden)) && (!("when" in field) || !field.when || field.when.every(condition => {
    const value = answer[condition.key]
    if (value === undefined) return false
    const equal = Array.isArray(value) ? value.includes(String(condition.value)) : value === condition.value
    return condition.op === "eq" ? equal : !equal
  })))
}
function initialAnswer(method?: IntegrationMethod): FormAnswer {
  return method && "form" in method ? Object.fromEntries((method.form ?? []).flatMap(field => "default" in field && field.default !== undefined ? [[field.key, field.default]] : [])) : {}
}

export default function NativeProviderConnect({ client, onConnected, showExistingSignIn = true, integrationID, onboarding = false, savedSignIn, discoveryNotice }: { client: Client; onConnected: (integrationID: string) => Promise<void>; showExistingSignIn?: boolean; integrationID?: string; onboarding?: boolean; savedSignIn?: ReactNode; discoveryNotice?: ReactNode }) {
  const [integrations, setIntegrations] = useState<IntegrationInfo[]>([])
  const [pick, setPick] = useState("")
  const [methodIndex, setMethodIndex] = useState("")
  const [key, setKey] = useState("")
  const [code, setCode] = useState("")
  const [answer, setAnswer] = useState<FormAnswer>({})
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [loading, setLoading] = useState(true)
  const [attempt, setAttempt] = useState<Attempt | null>(savedAttempt)
  const [instructions, setInstructions] = useState("")
  const [url, setURL] = useState("")
  const [status, setStatus] = useState("")
  const integration = integrations.find(item => item.id === (integrationID ?? pick))
  const methods = integration?.methods.flatMap((method, index) => method.type === "env" ? [] : [{ value: String(index), label: method.label || "API key" }]) ?? []
  const selection = onboarding ? methodIndex || methods[0]?.value || (!loading && savedSignIn ? "__saved__" : "") : methodIndex
  const method = integration?.methods[Number(selection)]
  const fields = method && "form" in method ? method.form ?? [] : []
  const pending = attempt && (integrationID === undefined || attempt.integrationID === integrationID)
  async function refresh() {
    setLoading(true)
    try { setIntegrations(await listIntegrations(client)); setError("") }
    catch { setError("Could not load sign-in methods. Retry.") }
    finally { setLoading(false) }
  }
  useEffect(() => { void refresh() }, [client])
  useEffect(() => { setMethodIndex(""); setAnswer({}); setKey(""); setCode("") }, [integrationID])
  useEffect(() => { if (onboarding) setAnswer(initialAnswer(method)) }, [onboarding, integrationID, selection])
  function remember(value: Attempt | null) {
    setAttempt(value)
    try { if (value) sessionStorage.setItem(attemptKey, JSON.stringify(value)); else sessionStorage.removeItem(attemptKey) } catch { /* session storage is optional */ }
  }
  async function run(action: () => Promise<void>) {
    if (busy) return
    setBusy(true); setError("")
    try { await action() }
    catch { setError("The sign-in request failed or timed out. Check its status before starting another attempt.") }
    finally { setBusy(false) }
  }
  async function check() {
    if (!attempt) return
    const input = { integrationID: attempt.integrationID, attemptID: attempt.attemptID, location: location(client) }
    const result = attempt.type === "oauth"
      ? await client.integration.oauth.status(input, { signal: AbortSignal.timeout(10000) })
      : await client.integration.command.status(input, { signal: AbortSignal.timeout(10000) })
    if (attempt.type === "command" && "message" in result.data && typeof result.data.message === "string") setInstructions(result.data.message)
    setStatus(result.data.status)
    if (result.data.status === "complete") { remember(null); await refresh(); await onConnected(attempt.integrationID) }
    if (result.data.status === "failed" || result.data.status === "expired") {
      setError(result.data.status === "expired" ? "This attempt expired. Start a new sign-in when ready." : "Sign-in failed. Check the provider account or host setup, then start a new attempt.")
      remember(null)
    }
  }
  return <><section className={`settings-section${onboarding ? " onboarding-native-connect" : ""}`}>
    {!onboarding && <><h3>Connect a provider</h3>
    <p className="settings-hint muted">Choose a sign-in method for this provider. Credentials are stored in this installation. Connecting does not run a model test.</p></>}
    {onboarding && loading && !integration && <p className="settings-hint muted" role="status">Loading connection options…</p>}
    {onboarding && !pending && (methods.length + (savedSignIn ? 1 : 0) > 1) && <label className="onboarding-field">Connection method<select disabled={busy} value={selection} onChange={event => { setMethodIndex(event.target.value); setAnswer(initialAnswer(integration?.methods[Number(event.target.value)])); setKey("") }}>
      {savedSignIn && <option value="__saved__">Use saved sign-in</option>}
      {methods.map(item => <option key={item.value} value={item.value}>{item.label}</option>)}
    </select></label>}
    {onboarding && discoveryNotice}
    {error && <p role="alert" className="settings-warning">{error}</p>}
    {instructions && <p className="settings-auth-instructions" style={{ whiteSpace: "pre-wrap" }}>{instructions}</p>}
    {onboarding && !integration && !loading && <button onClick={() => void refresh()}>Retry connection options</button>}
    {pending ? <>
      <p role="status">Sign-in {status || "pending"}. Returning here resumes this attempt; it never reruns a host command.</p>
      {url && /^https?:\/\//.test(url) && <a href={url} target="_blank" rel="noreferrer">Open provider sign-in</a>}
      {attempt.type === "oauth" && attempt.mode === "code" && <label>Authorization code<input type="password" value={code} onChange={event => setCode(event.target.value)} /></label>}
      <div className="settings-actions">
        {attempt.type === "oauth" && <button disabled={busy || (attempt.mode === "code" && !code.trim())} onClick={() => void run(async () => {
          await client.integration.oauth.complete({ ...attempt, code: code || undefined, location: location(client) }, { signal: AbortSignal.timeout(15000) }); setCode(""); await check()
        })}>Complete sign-in</button>}
        <button disabled={busy} onClick={() => void run(check)}>Check status</button>
        <button disabled={busy} onClick={() => void run(async () => {
          const input = { integrationID: attempt.integrationID, attemptID: attempt.attemptID, location: location(client) }
          if (attempt.type === "oauth") await client.integration.oauth.cancel(input, { signal: AbortSignal.timeout(10000) })
          else await client.integration.command.cancel(input, { signal: AbortSignal.timeout(10000) })
          remember(null); setStatus("cancelled"); setCode(""); setInstructions(""); setURL("")
        })}>Cancel sign-in</button>
      </div>
    </> : onboarding && selection === "__saved__" ? savedSignIn : <form onSubmit={event => { event.preventDefault(); void run(async () => {
      if (!integration || !method || method.type === "env") return
      const active = new Set(visibleFields(fields, answer, true).map(field => field.key))
      const submitted = Object.fromEntries(Object.entries(answer).filter(([name]) => active.has(name)))
      if (method.type === "key") {
        await client.integration.connect.key({ integrationID: integration.id, key: key.trim(), answer: submitted, location: location(client) }, { signal: AbortSignal.timeout(15000) })
        setKey(""); setAnswer({}); setInstructions(""); setURL(""); setStatus("complete"); await refresh(); await onConnected(integration.id); return
      }
      if (method.type === "oauth") {
        const result = await client.integration.oauth.connect({ integrationID: integration.id, methodID: method.id, answer: submitted, location: location(client) }, { signal: AbortSignal.timeout(15000) })
        remember({ integrationID: integration.id, attemptID: result.data.attemptID, type: "oauth", mode: result.data.mode }); setInstructions(result.data.instructions); setURL(result.data.url); setStatus("pending"); return
      }
      const result = await client.integration.command.connect({ integrationID: integration.id, methodID: method.id, location: location(client) }, { signal: AbortSignal.timeout(15000) })
      remember({ integrationID: integration.id, attemptID: result.data.attemptID, type: "command" }); setInstructions(""); setURL(""); setStatus("pending")
    }) }}>
      {integrationID === undefined && <label>Provider<select value={pick} disabled={busy} onChange={event => { setPick(event.target.value); setMethodIndex(""); setAnswer({}); setKey("") }}><option value="">Choose a provider</option>{integrations.map(item => <option key={item.id} value={item.id}>{item.name}{item.connections.length ? " · connected" : ""}</option>)}</select></label>}
      {!onboarding && integration && <label>Sign-in method<select disabled={busy} value={methodIndex} onChange={event => { setMethodIndex(event.target.value); setAnswer(initialAnswer(integration.methods[Number(event.target.value)])); setKey("") }}><option value="">Choose a method</option>{integration.methods.map((item, index) => <option key={index} value={index}>{item.type === "env" ? "Host environment" : item.label || "API key"}</option>)}</select></label>}
      {method && selection !== "" && <>
        {method.type === "env" ? <p>Set {method.names.join(", ")} on the engine host, then refresh the catalog. No browser key is needed.</p> : <>
          {method.type === "key" && <label>API key<input type="password" autoComplete="off" required value={key} onChange={event => setKey(event.target.value)} /></label>}
          {visibleFields(fields, answer).map(field => field.type === "external" ? <p key={field.key}><a href={field.url} target="_blank" rel="noreferrer">{field.title || field.key}</a>{field.description}</p> : <label key={field.key}>{field.title || field.key}
            {field.type === "boolean" ? <input type="checkbox" checked={answer[field.key] === true} onChange={event => setAnswer(value => ({ ...value, [field.key]: event.target.checked }))} />
              : field.type === "multiselect" ? <select multiple required={field.required} value={Array.isArray(answer[field.key]) ? answer[field.key] as string[] : []} onChange={event => setAnswer(value => ({ ...value, [field.key]: Array.from(event.target.selectedOptions, option => option.value) }))}>{field.options.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}</select>
              : field.type === "string" && field.options && !field.custom ? <select required={field.required} value={String(answer[field.key] ?? "")} onChange={event => setAnswer(value => ({ ...value, [field.key]: event.target.value }))}><option value="">Select</option>{field.options.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}</select>
              : <input required={field.required} type={field.type === "number" || field.type === "integer" ? "number" : "text"} value={String(answer[field.key] ?? "")} min={"minimum" in field ? field.minimum : undefined} max={"maximum" in field ? field.maximum : undefined} step={field.type === "integer" ? 1 : "any"} minLength={field.type === "string" ? field.minLength : undefined} maxLength={field.type === "string" ? field.maxLength : undefined} pattern={field.type === "string" ? field.pattern : undefined} onChange={event => setAnswer(value => { const next = { ...value }; if (event.target.value === "") delete next[field.key]; else next[field.key] = field.type === "number" || field.type === "integer" ? Number(event.target.value) : event.target.value; return next })} />}
            {field.description && <span className="settings-hint muted">{field.description}</span>}
          </label>)}
          {method.type === "command" && <p>This starts a command on the engine host: <code>{method.command.join(" ")}</code>. Start it only when you are ready.</p>}
          <button className="primary" disabled={busy || (method.type === "key" && !key.trim())}>{busy ? "Connecting…" : method.type === "command" ? "Run sign-in command" : "Connect"}</button>
        </>}
      </>}
    </form>}
    {onboarding && integration?.methods.some(item => item.type === "env") && <details className="onboarding-environment"><summary>Environment setup</summary>{integration.methods.flatMap((item, index) => item.type === "env" ? [<p className="settings-hint muted" key={index}>Set {item.names.map(name => <code key={name}>{name} </code>)}on the computer running doc.haus, then restart doc.haus to load them. No API key is needed in this form.</p>] : [])}</details>}
    {!onboarding && <button className="ghost" disabled={busy} onClick={() => void run(async () => { await refresh(); await onConnected("") })}>Refresh provider catalog</button>}
  </section>{showExistingSignIn && <ExistingSignInDisclosure onReused={async integrationID => { await refresh(); await onConnected(integrationID) }} />}</>
}
