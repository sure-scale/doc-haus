import { useEffect, useMemo, useRef, useState } from "react"
import type { IntegrationInfo } from "@opencode/client"
import { getConfig, setDisabledProviders, type Config } from "../../api/opencode/config"
import { listIntegrations, listProviders } from "../../api/opencode/providers"
import { settingsClient } from "../../api/opencode/client"
import { isRemovedModel, isRemovedProvider } from "../../api/opencode/provider-readiness"
import { providerHealth } from "../../api/opencode/provider-health"
import { listReusableSignIns } from "../../api/ingest/credential-reuse"
import { candidateAuth, candidateModels, createCandidate, discardCandidate, saveCandidate, SaveUnconfirmedError, testCandidate, type CandidateAuthResult, type CandidateConnection, type CandidateInput } from "../../api/opencode/provider-candidate"
import { OPENCODE_URL } from "../../config"
import { providerOf, type ModelOption } from "../../models"
import type { Provider } from "../settings/types"
import type { MatchingSignInCandidate } from "./connection-fields"
import { CLOUD_SETUPS } from "../settings/cloud-setups"
import { visibleFields } from "./connection-fields"
import { FieldRow } from "../settings/controls"
import ConnectionEditor, { connectionChoices, methodDefaults, type ConnectionDraft } from "./ConnectionEditor"
import ProviderPicker from "./ProviderPicker"
import ModelFields from "./ModelFields"
import TestResults, { type ModelTestResult } from "./TestResults"

export type SetupPhase = "provider" | "connection" | "models" | "test"
export type AISetupProgress = { provider: boolean; connection: boolean; models: boolean; test: boolean; loading: boolean; busy: boolean; saving?: boolean }
const emptyDraft = (): ConnectionDraft => ({ choice: "", key: "", answer: {}, options: {}, name: "" })
const message = (error: unknown, fallback: string) => error instanceof Error ? error.message : fallback
const localID = (name: string) => name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "")
const modelID = (reference: string) => reference.slice(reference.indexOf("/") + 1)

type PendingSignIn = { sessionID: string; providerID: string; methodID: string; options: Record<string, string>; attempt: CandidateAuthResult & { type: "oauth" | "command" } }
const pendingKey = `dochaus.setup-sign-in:${OPENCODE_URL}`
function pendingSignIn(): PendingSignIn | undefined {
  try {
    const saved = JSON.parse(sessionStorage.getItem(pendingKey) ?? "null")
    if (saved && typeof saved.sessionID === "string" && typeof saved.providerID === "string" && typeof saved.methodID === "string" &&
      saved.attempt && typeof saved.attempt.attemptID === "string" && ["oauth", "command"].includes(saved.attempt.type) &&
      saved.options && typeof saved.options === "object" && Object.entries(saved.options).every(([key, value]) => ["project", "location", "region", "profile", "baseURL"].includes(key) && typeof value === "string")) return saved
  } catch { /* A pending sign-in can be restarted if browser storage is unavailable. */ }
}
function forgetSignIn(id: string) {
  try { if (pendingSignIn()?.sessionID === id) sessionStorage.removeItem(pendingKey) } catch { /* Optional browser storage. */ }
}

const endpointProvider = (id: string, config: Config) => id === "ollama" || /(?:openai-compatible|ollama)$/.test(config.provider?.[id]?.npm ?? "")

function prefill(provider: Provider | undefined, integration: IntegrationInfo | undefined, config: Config): ConnectionDraft {
  if (!provider) return emptyDraft()
  const options: Record<string, string> = {}
  const spec = CLOUD_SETUPS.find(item => item.ids.includes(provider.id))
  for (const field of spec?.fields ?? []) {
    const saved = config.provider?.[provider.id]?.options?.[field.key]
    options[field.key] = typeof saved === "string" ? saved : provider.host?.options[field.key] ?? field.fallback ?? ""
  }
  const savedEndpoint = config.provider?.[provider.id]?.options?.baseURL
  if (typeof savedEndpoint === "string") options.baseURL = savedEndpoint
  const choices = connectionChoices(provider, integration, [], endpointProvider(provider.id, config))
  // Reusable external sign-ins are never selected or copied by discovery.
  const choice = choices.find(item => item.value.startsWith("credential:") || item.value.startsWith("env:"))
    ?? (provider.host?.sources.length || endpointProvider(provider.id, config) ? choices.find(item => item.value === "host") : undefined)
    ?? (choices.length === 1 ? choices[0] : undefined)
  return { ...emptyDraft(), options, choice: choice?.value ?? "", answer: methodDefaults(choice?.method) }
}

export default function AISetup({ mode = "settings", reconnectProvider, phase = "provider", onPhaseChange, onProgress, onSaved }: {
  mode?: "settings" | "onboarding"
  reconnectProvider?: string
  phase?: SetupPhase
  onPhaseChange?: (phase: SetupPhase) => void
  onProgress?: (progress: AISetupProgress) => void
  onSaved: () => void | Promise<void>
}) {
  const client = useMemo(settingsClient, [])
  const [providers, setProviders] = useState<Provider[]>([])
  const [integrations, setIntegrations] = useState<IntegrationInfo[]>([])
  const [reusable, setReusable] = useState<MatchingSignInCandidate[]>([])
  const [config, setConfig] = useState<Config>({})
  const [provider, setProvider] = useState("")
  const [draft, setDraft] = useState<ConnectionDraft>(emptyDraft)
  const [models, setModels] = useState<ModelOption[]>([])
  const [primaryChoices, setPrimaryChoices] = useState<ModelOption[]>([])
  const [primary, setPrimary] = useState("")
  const [fast, setFast] = useState("")
  const [customModels, setCustomModels] = useState(false)
  const [loaded, setLoaded] = useState(false)
  const [loadVersion, setLoadVersion] = useState(0)
  const [modelReady, setModelReady] = useState(false)
  const [saveUnconfirmed, setSaveUnconfirmed] = useState(false)
  const [continuePending, setContinuePending] = useState(false)
  const [loading, setLoading] = useState(true)
  const [busy, setBusy] = useState(false)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState("")
  const [discoveryNotice, setDiscoveryNotice] = useState("")
  const [sessionID, setSessionID] = useState("")
  const [results, setResults] = useState<ModelTestResult[]>([])
  const [verified, setVerified] = useState("")
  const [saved, setSaved] = useState(false)
  const [attempt, setAttempt] = useState<(CandidateAuthResult & { type: "oauth" | "command" })>()
  const [authCode, setAuthCode] = useState("")
  const revision = useRef(0)
  const session = useRef("")
  const autoPrepare = useRef(false)
  const callbacks = useRef({ onProgress, onPhaseChange })
  callbacks.current = { onProgress, onPhaseChange }
  const local = provider === "__local_endpoint__"
  const selected = providers.find(item => item.id === provider)
  const actualProvider = local ? localID(draft.name) : provider
  const endpoint = endpointProvider(provider, config)
  const integration = integrations.find(item => item.id === (selected?.integrationID ?? provider))
  const matching = reusable.filter(item => item.integrationID === (selected?.integrationID ?? provider))
  const method = connectionChoices(selected, integration, matching, endpoint).find(item => item.value === draft.choice)?.method
  const chosenFast = fast || primary
  const choiceValid = Boolean(primary && chosenFast && providerOf(primary) === actualProvider && providerOf(chosenFast) === actualProvider && (customModels || primaryChoices.some(item => item.value === primary) && models.some(item => item.value === chosenFast)))
  const fingerprint = JSON.stringify([sessionID, primary, chosenFast])
  const passed = Boolean(sessionID && choiceValid && verified === fingerprint)
  const configured = Boolean(sessionID && modelReady)
  const progress = { provider: Boolean(provider), connection: configured, models: choiceValid, test: passed, loading, busy, saving }
  useEffect(() => { callbacks.current.onProgress?.(progress) }, [provider, configured, choiceValid, passed, loading, busy, saving])

  function clearTest() { revision.current++; setResults([]); setVerified(""); setSaved(false); setError(""); setContinuePending(false) }
  function resetCandidate() {
    clearTest()
    const old = session.current
    session.current = ""; setSessionID(""); setModelReady(false); setModels([]); setPrimaryChoices([]); setAttempt(undefined); setAuthCode(""); setCustomModels(false)
    if (old) { forgetSignIn(old); void discardCandidate(old) }
  }
  async function load(expected?: { providerID: string; models: string[] }) {
    setLoading(true); setError("")
    try {
      const [initialCatalog, initialAuth, cfg, reuse] = await Promise.all([
        listProviders(client), listIntegrations(client), getConfig(),
        listReusableSignIns().then(sources => ({ sources, error: "" }), () => ({ sources: [], error: "Saved sign-in discovery is unavailable. You can still use a connection from this installation or sign in directly." })),
      ])
      let catalog = initialCatalog, auth = initialAuth
      // The settings rename is synchronous; native provider discovery follows a
      // debounced watcher. Read metadata until the saved model pair appears.
      if (expected) for (let retry = 0; ; retry++) {
        const found = catalog.all.find(item => item.id === expected.providerID && item.activation !== "disabled")
        if (found && expected.models.every(id => Boolean(found.models[id.split("#")[0]!]))) break
        if (retry === 15) throw new Error("AI setup was saved. Provider discovery is still updating. Refresh discovery to load it.")
        await new Promise(resolve => setTimeout(resolve, 200))
        const [nextCatalog, nextAuth] = await Promise.all([listProviders(client), listIntegrations(client)])
        catalog = nextCatalog
        auth = nextAuth
      }
      const all = catalog.all.filter(item => !isRemovedProvider(item.id))
      const pending = pendingSignIn()
      const target = expected?.providerID ?? (reconnectProvider ? all.find(item => item.id === reconnectProvider)?.id ?? all.find(item => item.integrationID === reconnectProvider)?.id ?? "" : pending?.providerID ?? providerOf(cfg.model ?? ""))
      const picked = all.find(item => item.id === target && item.activation !== "disabled")
      setProviders(all); setIntegrations(auth); setConfig(cfg)
      setReusable(reuse.sources.filter(source => source.status === "ready").flatMap(source => source.connections.map(connection => ({ sourceId: source.id, connectionId: connection.id, integrationID: connection.integrationID, label: connection.label, authType: connection.authType, sourceLabel: source.label }))))
      setDiscoveryNotice(reuse.error || (reuse.sources.some(source => source.status === "unavailable") ? "Some saved sign-in sources could not be read. The connections listed are available to choose." : ""))
      setProvider(picked?.id ?? "")
      const next = prefill(picked, auth.find(item => item.id === (picked?.integrationID ?? picked?.id)), cfg)
      const pendingMethod = picked && pending?.providerID === picked.id ? auth.find(item => item.id === (picked.integrationID ?? picked.id))?.methods.find(item => "id" in item && item.id === pending.methodID && item.type === pending.attempt.type) : undefined
      if (pending && pendingMethod) {
        next.choice = connectionChoices(picked, auth.find(item => item.id === (picked?.integrationID ?? picked?.id)), []).find(item => item.method === pendingMethod)?.value ?? ""
        next.options = pending.options
        session.current = pending.sessionID; setSessionID(pending.sessionID); setAttempt({ ...pending.attempt, status: "pending" }); setModelReady(false)
      }
      setDraft(next)
      setPrimary(!isRemovedModel(cfg.model) && picked?.id === providerOf(cfg.model ?? "") ? cfg.model ?? "" : "")
      setFast(!isRemovedModel(cfg.small_model) && picked?.id === providerOf(cfg.small_model ?? "") && cfg.small_model !== cfg.model ? cfg.small_model ?? "" : "")
      autoPrepare.current = Boolean(!pendingMethod && picked && (next.choice.startsWith("credential:") || next.choice.startsWith("env:") || next.choice === "host"))
      setLoaded(true); setLoadVersion(value => value + 1)
    } catch (failure) { setError(message(failure, "Could not load AI setup. Retry when the engine is ready.")) }
    finally { setLoading(false) }
  }
  useEffect(() => {
    void load()
    return () => { revision.current++; if (session.current && pendingSignIn()?.sessionID !== session.current) void discardCandidate(session.current) }
  }, [client])
  useEffect(() => {
    if (!loaded || loading || busy || continuePending || !autoPrepare.current) return
    autoPrepare.current = false
    void prepare(false)
  }, [loadVersion, loading, busy, continuePending])

  function changeProvider(id: string) {
    resetCandidate(); setProvider(id); setPrimary(""); setFast("")
    const item = providers.find(item => item.id === id)
    setDraft(id === "__local_endpoint__" ? { ...emptyDraft(), options: { baseURL: "http://localhost:1234/v1" } } : prefill(item, integrations.find(integration => integration.id === (item?.integrationID ?? id)), config))
  }
  function changeDraft(next: ConnectionDraft) { resetCandidate(); setDraft(next) }
  function changePrimary(value: string) { clearTest(); setPrimary(value) }
  function changeFast(value: string) { clearTest(); setFast(value) }

  function input(): CandidateInput {
    if (!actualProvider || isRemovedProvider(actualProvider)) throw new Error("Choose a provider first.")
    const options: Record<string, string | null> = Object.fromEntries(Object.entries(draft.options).filter(([key, value]) => value.trim() || typeof config.provider?.[provider]?.options?.[key] === "string").map(([key, value]) => [key, value.trim() || null]))
    if (local) {
      if (!draft.name.trim() || typeof options.baseURL !== "string" || !options.baseURL) throw new Error("Enter a connection name and endpoint.")
      return { providerID: actualProvider, options: { baseURL: options.baseURL }, connection: draft.key.trim() ? { type: "key", key: draft.key.trim() } : { type: "host" }, local: { name: draft.name.trim(), baseURL: options.baseURL, modelIDs: [] } }
    }
    if (endpoint && !options.baseURL) throw new Error("Enter the model endpoint.")
    const spec = CLOUD_SETUPS.find(item => item.ids.includes(actualProvider))
    if (spec?.fields.some(field => field.required && !options[field.key])) throw new Error("Complete the required connection fields.")
    let connection: CandidateConnection
    if (draft.choice.startsWith("credential:")) connection = { type: "credential", id: draft.choice.slice(11) }
    else if (draft.choice.startsWith("env:")) connection = { type: "env", name: draft.choice.slice(4) }
    else if (draft.choice === "host") connection = { type: "host" }
    else if (draft.choice.startsWith("reuse:")) {
      const candidate = matching.find(item => `reuse:${item.sourceId}:${item.connectionId}` === draft.choice)
      if (!candidate) throw new Error("Choose a saved sign-in from the current list.")
      connection = { type: "reuse", sourceId: candidate.sourceId, connectionId: candidate.connectionId }
    } else if (method?.type === "key") {
      if (!draft.key.trim()) throw new Error("Enter your API key.")
      const fields = method.form ?? []
      const active = new Set(visibleFields(fields, draft.answer, true).map(field => field.key))
      connection = { type: "key", key: draft.key.trim(), answer: Object.fromEntries(Object.entries(draft.answer).filter(([key]) => active.has(key))) }
    } else if (method?.type === "oauth" || method?.type === "command") connection = { type: "new" }
    else throw new Error("Choose a connection method.")
    return { providerID: actualProvider, options, connection }
  }

  async function refreshModels(id: string, current: number) {
    const catalog = await candidateModels(id)
    if (revision.current !== current) return
    const text = catalog.data.filter(item => !item.capabilities || item.capabilities.input.includes("text") && item.capabilities.output.includes("text"))
    const option = (item: typeof text[number]): ModelOption[] => [{ value: `${item.providerID}/${item.id}`, label: item.name, providerId: item.providerID, modelId: item.id }, ...(item.variants ?? []).map(variant => ({ value: `${item.providerID}/${item.id}#${variant.id}`, label: `${item.name} · ${variant.name ?? variant.id}`, providerId: item.providerID, modelId: `${item.id}#${variant.id}` }))]
    const next = text.flatMap(option).sort((a, b) => a.label.localeCompare(b.label))
    const primaries = text.filter(item => !item.capabilities || item.capabilities.tools).flatMap(option).sort((a, b) => a.label.localeCompare(b.label))
    setModels(next); setPrimaryChoices(primaries); setCustomModels(catalog.customModels); setModelReady(true)
    if (primaries.length === 1) setPrimary(value => value || primaries[0]!.value)
  }
  async function prepare(advance = true) {
    if (busy) return
    setBusy(true); setError("")
    const current = revision.current
    try {
      const created = await createCandidate(input())
      if (revision.current !== current) { void discardCandidate(created.sessionID); return }
      if (session.current) void discardCandidate(session.current)
      session.current = created.sessionID; setSessionID(created.sessionID); setResults([]); setVerified("")
      if (method?.type === "oauth" || method?.type === "command") {
        const result = await candidateAuth(created.sessionID, `${method.type}.connect`, { methodID: method.id, answer: Object.fromEntries(Object.entries(draft.answer).filter(([key]) => visibleFields("form" in method ? method.form ?? [] : [], draft.answer, true).some(field => field.key === key))) })
        if (revision.current !== current) return
        const pending = { ...result.data, type: method.type, status: result.data.status ?? "pending" }
        setAttempt(pending)
        if (pending.attemptID) try {
          // Resume native attempts without storing keys, form answers or codes.
          sessionStorage.setItem(pendingKey, JSON.stringify({ sessionID: created.sessionID, providerID: actualProvider, methodID: method.id, options: draft.options, attempt: { type: method.type, attemptID: pending.attemptID, mode: pending.mode } }))
        } catch { /* Optional browser storage. */ }
      } else {
        await refreshModels(created.sessionID, current)
        if (advance && revision.current === current) callbacks.current.onPhaseChange?.("models")
      }
    } catch (failure) { if (revision.current === current) setError(message(failure, "Could not prepare this connection. Check the details and retry.")) }
    finally { if (revision.current === current) setBusy(false) }
  }
  async function authAction(action: "complete" | "status" | "cancel") {
    if (!attempt || !sessionID || busy) return
    setBusy(true); setError("")
    const current = revision.current
    try {
      if (action === "complete" && attempt.type !== "oauth") return
      const actionName = action === "complete" ? "oauth.complete" : `${attempt.type}.${action}` as const
      const result = await candidateAuth(sessionID, actionName, { attemptID: attempt.attemptID, ...(action === "complete" ? { code: authCode.trim() || undefined } : {}) })
      if (revision.current !== current) return
      if (action === "cancel") { resetCandidate(); return }
      const next = result.data
      setAttempt({ ...attempt, ...next })
      setAuthCode("")
      if (next.status === "complete") { forgetSignIn(sessionID); await refreshModels(sessionID, current); callbacks.current.onPhaseChange?.("models") }
      else if (next.status === "failed" || next.status === "expired") setError("Sign-in did not complete. Cancel this attempt, then start a new sign-in.")
    } catch (failure) { if (action === "cancel") resetCandidate(); else setError(message(failure, "Could not check sign-in. Check its status before starting another attempt.")) }
    finally { setBusy(false) }
  }
  async function runTests() {
    if (!configured || !choiceValid || busy) return
    setBusy(true); setError(""); setVerified(""); setSaved(false)
    const current = revision.current
    const pair = [{ role: "primary" as const, model: primary }, { role: "fast" as const, model: chosenFast }]
    setResults(pair.map(item => ({ ...item, status: "checking", ...(primary === chosenFast && item.role === "fast" ? { shared: true } : {}) })))
    try {
      if (customModels) await candidateModels(sessionID, [...new Set(pair.map(item => modelID(item.model)))])
      let allPassed = true
      for (const selected of [...new Set(pair.map(item => item.model))]) {
        const outcome = await testCandidate(sessionID, modelID(selected))
        if (revision.current !== current) return
        allPassed &&= outcome.ok
        setResults(results => results.map(item => item.model !== selected ? item : { ...item, status: outcome.ok ? "ready" : outcome.kind, ...(!outcome.ok ? { message: outcome.error } : {}) }))
      }
      if (allPassed && revision.current === current) setVerified(fingerprint)
    } catch (failure) {
      if (revision.current === current) {
        const detail = message(failure, "The connection test could not complete.")
        setError(detail); setResults(results => results.map(item => item.status === "checking" ? { ...item, status: "unavailable", message: detail } : item))
      }
    } finally { if (revision.current === current) setBusy(false) }
  }
  async function continueAfterSave() {
    setContinuePending(false)
    try { await onSaved() }
    catch { setContinuePending(true); setError("AI setup was saved. Setup progress could not be updated. Continue again to retry that step.") }
  }
  async function save() {
    if (!passed || busy || saveUnconfirmed) return
    setSaving(true); setBusy(true); setError("")
    try {
      await saveCandidate(sessionID, modelID(primary), modelID(chosenFast))
    } catch (failure) {
      if (failure instanceof SaveUnconfirmedError) setSaveUnconfirmed(true)
      setError(message(failure, "Could not save AI setup. Retry.")); setBusy(false); setSaving(false); return
    }
    // Save is committed. Refresh/progress failures must never invite another
    // credential activation against the now-closed candidate.
    setSaved(true); session.current = ""; setSessionID(""); setModelReady(false)
    providerHealth.confirmed([{ providerID: actualProvider, modelID: modelID(primary) }, { providerID: actualProvider, modelID: modelID(chosenFast) }])
    await load({ providerID: actualProvider, models: [modelID(primary), modelID(chosenFast)] })
    await continueAfterSave()
    setBusy(false); setSaving(false)
  }

  const show = (section: SetupPhase) => mode === "settings" || phase === section
  const sourceOptions = Object.fromEntries((CLOUD_SETUPS.find(item => item.ids.includes(provider))?.fields ?? []).map(field => [field.key,
    typeof config.provider?.[provider]?.options?.[field.key] === "string" ? "Prefilled from this installation" : selected?.host?.options[field.key] ? "Prefilled from server detection" : "Default",
  ]))
  const nativeAuth = method?.type === "oauth" || method?.type === "command"
  const prepareLabel = nativeAuth ? method.type === "command" ? "Run sign-in command" : "Sign in" : mode === "onboarding" ? "Continue to models" : configured ? "Reload models" : "Load models"
  if (loading && !loaded) return <p className="settings-hint muted" role="status">Loading providers and saved connections…</p>
  if (!loaded) return <div role="alert"><p>{error}</p><button type="button" onClick={() => void load()}>Retry AI setup</button></div>
  return <div className="ai-setup" aria-busy={loading || busy}>
    {mode === "settings" && <div className="ai-setup-summary">
      <h3>AI setup</h3>
      <p className="settings-hint muted">Choose a provider, configure its connection, then select and test your models.</p>
      {config.model && <p className="settings-hint muted">Current primary model: <strong>{config.model}</strong></p>}
    </div>}
    {error && <p role="alert" className="settings-warning">{error}</p>}
    {saveUnconfirmed && <button type="button" onClick={() => { setSaveUnconfirmed(false); resetCandidate(); void load() }}>Reload saved AI setup</button>}
    {saved && error && mode === "settings" && <button type="button" disabled={busy || loading} onClick={() => void load()}>Retry provider discovery</button>}
    {continuePending && <button type="button" onClick={() => void continueAfterSave()}>Continue</button>}
    {discoveryNotice && <p className="settings-hint muted">{discoveryNotice}</p>}
    {(isRemovedModel(config.model) || isRemovedModel(config.small_model)) && <p role="alert" className="settings-warning">The previously selected provider has been removed. Choose another provider and test both model choices.</p>}
    {show("provider") && <section className="settings-section">
      {mode === "settings" && <h3 className="ai-setup-section-title"><span>1</span> Provider</h3>}
      <ProviderPicker providers={providers.filter(item => item.activation !== "disabled" && !config.disabled_providers?.includes(item.id))} provider={provider} onChange={changeProvider} reusableProviderIDs={reusable.map(item => item.integrationID)} disabled={busy || loading} />
      {mode === "onboarding" && <div className="ai-setup-actions"><button type="button" className="primary" disabled={!provider || busy} onClick={() => onPhaseChange?.("connection")}>Continue to connection</button></div>}
    </section>}
    {show("connection") && <section className="settings-section">
      {mode === "settings" && <h3 className="ai-setup-section-title"><span>2</span> Connection</h3>}
      {provider ? <>
        <ConnectionEditor provider={selected} integration={integration} reusable={matching} draft={draft} onChange={changeDraft} disabled={busy || loading} local={local} endpoint={endpoint} sourceOptions={sourceOptions} attempt={attempt} authCode={authCode} onAuthCode={setAuthCode} onAuthAction={action => void authAction(action)} />
        {!attempt && <div className="ai-setup-actions">
          {mode === "onboarding" && <button type="button" disabled={busy} onClick={() => onPhaseChange?.("provider")}>Back to providers</button>}
          <button type="button" className="primary" disabled={busy || (!local && !draft.choice)} onClick={() => configured && mode === "onboarding" ? onPhaseChange?.("models") : void prepare()}>{busy ? "Preparing connection…" : configured && mode === "onboarding" ? "Continue to models" : prepareLabel}</button>
        </div>}
        {attempt?.status === "complete" && <div className="ai-setup-actions"><button type="button" disabled={busy} onClick={() => {
          if (configured && mode === "onboarding") { onPhaseChange?.("models"); return }
          setBusy(true); setError("")
          void refreshModels(sessionID, revision.current).then(() => { if (mode === "onboarding") onPhaseChange?.("models") }).catch(failure => setError(message(failure, "Could not load models. Retry."))).finally(() => setBusy(false))
        }}>{configured && mode === "onboarding" ? "Continue to models" : "Retry loading models"}</button></div>}
      </> : <p className="settings-hint muted">Choose a provider to see its connection options.</p>}
    </section>}
    {show("models") && <section className="settings-section">
      {mode === "settings" && <h3 className="ai-setup-section-title"><span>3</span> Models</h3>}
      {configured ? <>
        {customModels ? <>
          <FieldRow id="ai-setup-primary-id" label="Primary model ID" hint="Enter the model identifier served by this endpoint."><input id="ai-setup-primary-id" aria-describedby="ai-setup-primary-id-hint" disabled={busy} value={primary ? modelID(primary) : ""} onChange={event => changePrimary(event.target.value.trim() ? `${actualProvider}/${event.target.value.trim()}` : "")} /></FieldRow>
          <FieldRow id="ai-setup-fast-id" label="Fast model ID" hint="Leave blank to use the primary model for background tasks."><input id="ai-setup-fast-id" aria-describedby="ai-setup-fast-id-hint" disabled={busy} placeholder="Same as primary" value={fast ? modelID(fast) : ""} onChange={event => changeFast(event.target.value.trim() ? `${actualProvider}/${event.target.value.trim()}` : "")} /></FieldRow>
        </> : <ModelFields models={models} primaryModels={primaryChoices} primary={primary} fast={fast} onPrimaryChange={changePrimary} onFastChange={changeFast} disabled={busy} />}
        {!models.length && !customModels && <p className="settings-hint muted">No models are listed for this connection. Check the provider connection or refresh discovery.</p>}
      </> : <p className="settings-hint muted">Configure the connection to load its model choices.</p>}
      {mode === "onboarding" && <div className="ai-setup-actions"><button type="button" disabled={busy} onClick={() => onPhaseChange?.("connection")}>Back to connection</button><button type="button" className="primary" disabled={!choiceValid || !configured || busy} onClick={() => onPhaseChange?.("test")}>Continue to test</button></div>}
    </section>}
    {show("test") && <section className="settings-section">
      {mode === "settings" && <h3 className="ai-setup-section-title"><span>4</span> Test</h3>}
      {choiceValid && <p className="settings-hint ai-setup-selected-models">Primary: <strong>{primary}</strong><br />Fast: <strong>{chosenFast}</strong>{primary === chosenFast ? " (same as primary)" : ""}</p>}
      <p className="settings-hint muted">Tests the selected connection, primary model and fast model with one short request per distinct model. Usage charges may apply.</p>
      <TestResults results={results} connectionStatus={passed ? "Connection verified for both selected model roles." : results.some(item => item.status === "auth") ? "The selected connection needs attention." : undefined} />
      {saved && <p role="status">AI setup saved.</p>}
      <div className="ai-setup-actions">
        {mode === "onboarding" && <button type="button" disabled={busy} onClick={() => onPhaseChange?.("models")}>Back to models</button>}
        <button type="button" disabled={busy || !choiceValid || !configured} onClick={() => void runTests()}>{busy && results.some(item => item.status === "checking") ? "Testing connection…" : "Test connection"}</button>
        <button type="button" className="primary" disabled={busy || !passed || saved || saveUnconfirmed} onClick={() => void save()}>{busy && passed ? "Saving…" : mode === "onboarding" ? "Save and continue" : "Save"}</button>
      </div>
    </section>}
    {mode === "settings" && <details className="ai-setup-manage settings-section"><summary>Manage connections</summary>
      <p className="settings-hint muted">Saved connections remain available when you change provider. Enable or disable other providers here.</p>
      {providers.map(item => {
        const disabled = config.disabled_providers?.includes(item.id) ?? false
        return <div className="settings-conn" key={item.id}><span className="settings-conn-name">{item.name}</span><button type="button" className="btn-sm ghost" disabled={busy || item.id === provider} onClick={async () => {
          setBusy(true); setError("")
          try { await setDisabledProviders(disabled ? (config.disabled_providers ?? []).filter(id => id !== item.id) : [...(config.disabled_providers ?? []), item.id]); setConfig(await getConfig()) }
          catch (failure) { setError(message(failure, "Could not update this provider.")) }
          finally { setBusy(false) }
        }}>{disabled ? "Enable" : "Disable"}</button></div>
      })}
      <button type="button" className="ghost" disabled={busy} onClick={() => { resetCandidate(); void load() }}>Refresh discovery</button>
    </details>}
  </div>
}
