import { useEffect, useMemo, useRef, useState } from "react"
import { getConfig } from "../../api/opencode/config"
import { listProviders } from "../../api/opencode/providers"
import { settingsClient, type Client } from "../../api/opencode/client"
import { getPreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { listJurisdictions, type Jurisdiction } from "../../api/ingest/matters"
import { getSetup, saveSetup, type SetupProgress } from "../../api/ingest/setup"
import { pickForProvider, providerOf } from "../../models"
import { loadPrefs } from "../../prefs"
import { isGated, loadVerified } from "../../providers"
import type { Provider } from "./types"
import DraftingDetailsStep from "./DraftingDetailsStep"
import ModelPickerStep from "./ModelPickerStep"
import ConnectionStep from "./ConnectionStep"
import ResearchTab from "../settings/ResearchTab"
import LocalEndpoint from "../settings/LocalEndpoint"
import { addLocalProvider } from "../../api/opencode/config"
import HostProviderSetup from "../settings/HostProviderSetup"
import "./onboarding.css"

const stages = [
  { id: "provider", label: "Provider", title: "Choose your model provider", description: "Connect the service you want to use for drafting and working with your matters." },
  { id: "connection", label: "Connection", title: "Connect your provider", description: "Use an existing connection or sign in securely. Credentials are stored in this installation." },
  { id: "models", label: "Models", title: "Choose your models", description: "Pick a primary model for your matters and a fast model for smaller tasks." },
  { id: "research", label: "Research", title: "Set up legal research", description: "Choose research sources and add service keys." },
  { id: "details", label: "Details", title: "Make doc.haus yours", description: "Add the details you use in your drafts." },
]

export default function Onboarding({ onClose }: { onClose: () => void }) {
  const client = useMemo<Client>(() => settingsClient(), [])
  const [all, setAll] = useState<Provider[]>([])
  const [connected, setConnected] = useState<Set<string>>(new Set())
  const [provider, setProvider] = useState("")
  const [primary, setPrimary] = useState("")
  const [fast, setFast] = useState("")
  const [notice, setNotice] = useState("")
  // A failed step save clears on the next save attempt; other notices stay.
  const [saveNotice, setSaveNotice] = useState("")
  const [loading, setLoading] = useState(true)
  const [loaded, setLoaded] = useState(false)
  const [failed, setFailed] = useState(false)
  const [phase, setPhase] = useState<"provider" | "connection" | "models">("provider")
  const [step, setStep] = useState<"setup" | "details" | "research">("setup")
  const [drafting, setDrafting] = useState<DraftingPreferences | null>(null)
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [selectedJx, setSelectedJx] = useState<string[]>(() => loadPrefs().defaultJurisdictions)
  // The firm's own date format pick; until it picks, the Details step follows
  // its default jurisdictions.
  const [dateFormatChoice, setDateFormatChoice] = useState<DraftingPreferences["dateFormat"] | null>(null)
  const [verified, setVerified] = useState<Set<string>>(loadVerified)
  const [advancing, setAdvancing] = useState(false)
  const [progress, setProgress] = useState<SetupProgress | null>(null)
  const [mounted, setMounted] = useState({ connection: false, research: false })
  const panel = useRef<HTMLDivElement>(null)
  const body = useRef<HTMLDivElement>(null)
  const heading = useRef<HTMLHeadingElement>(null)
  const currentStage = step === "setup" ? phase : step
  const current = stages.findIndex(stage => stage.id === currentStage)
  const stage = stages[current]!

  useEffect(() => {
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const overflow = document.body.style.overflow
    document.body.style.overflow = "hidden"
    function onKey(event: KeyboardEvent) {
      if (event.key !== "Tab" || !panel.current) return
      // The jurisdiction picker portals outside the scrolling dialog.
      const containers = [panel.current, ...document.querySelectorAll<HTMLElement>(".jx-popover")]
      const controls = containers.flatMap(container => Array.from(container.querySelectorAll<HTMLElement>('button:not([disabled]), a[href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), summary, [tabindex="0"]'))).filter(control => control.getClientRects().length > 0)
      const first = controls[0]
      const last = controls.at(-1)
      if (!first || !last) { event.preventDefault(); heading.current?.focus(); return }
      if (event.shiftKey && (document.activeElement === first || document.activeElement === heading.current)) { event.preventDefault(); last.focus() }
      if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
    }
    document.addEventListener("keydown", onKey)
    return () => { document.body.style.overflow = overflow; document.removeEventListener("keydown", onKey); previous?.focus({ preventScroll: true }) }
  }, [])
  useEffect(() => {
    if (currentStage === "connection" || currentStage === "research") setMounted(previous => previous[currentStage] ? previous : { ...previous, [currentStage]: true })
    body.current?.scrollTo({ top: 0 })
    heading.current?.focus({ preventScroll: true })
  }, [current])

  async function load(justConnected = "", resume = false) {
    setLoading(true); setFailed(false); setNotice("")
    try {
      const [providers, cfg, progress] = await Promise.all([listProviders(client), getConfig(), getSetup()])
      setAll(providers.all); setConnected(new Set(providers.connected)); setVerified(loadVerified()); setLoaded(true)
      setProgress(progress)
      setPrimary(cfg.model ?? ""); setFast(cfg.small_model ?? "")
      const ready = providers.all.filter(item => providers.connected.includes(item.id) && (!isGated(item.id) || loadVerified().has(item.id)))
      const choices = ready.flatMap(item => Object.values(item.models).map(model => ({ value: `${item.id}/${model.id}`, label: model.name, providerId: item.id, modelId: model.id })))
      const newly = justConnected ? ready.filter(item => item.integrationID === justConnected || item.id === justConnected) : []
      const selected = newly.length === 1 ? newly[0]!.id : provider === "__local_endpoint__" || providers.all.some(item => item.id === provider && item.activation !== "disabled") ? provider : providers.all.some(item => item.id === providerOf(cfg.model ?? "")) ? providerOf(cfg.model ?? "") : ""
      setProvider(selected)
      if (selected && selected !== providerOf(cfg.model ?? "")) { const defaults = pickForProvider(choices, selected); setPrimary(defaults.primary); setFast(defaults.fast) }
      if (resume && cfg.model && selected) setPhase("models")
      if (resume) setStep(progress.next === "details" || progress.next === "research" ? progress.next : "setup")
    } catch (error) { setFailed(true); setNotice(error instanceof Error ? error.message : "Could not load setup. Retry.") }
    finally { setLoading(false) }
  }
  useEffect(() => {
    void load("", true)
    void Promise.all([getPreferences(), listJurisdictions()]).then(([preferences, list]) => { setDrafting(preferences); setJurisdictions(list) }).catch(() => setNotice("Could not load drafting details. Retry setup."))
  }, [client])
  const ready = all.filter(item => connected.has(item.id) && (!isGated(item.id) || verified.has(item.id)))
  const models = ready.flatMap(item => Object.values(item.models).map(model => ({ value: `${item.id}/${model.id}`, label: `${item.name} — ${model.name}`, providerId: item.id, modelId: model.id }))).sort((a, b) => a.label.localeCompare(b.label))
  const selected = all.find(item => item.id === provider)
  const active = all.filter(item => item.activation !== "disabled").sort((a, b) => a.name.localeCompare(b.name))
  const detected = (item: Provider) => item.nativeAvailable || item.credentialSources.length > 0 || !!item.host?.sources.length
  const selectable = models.some(model => model.providerId === provider)
  const modelsSaved = progress?.ready === true && progress.completed.includes("models")
  function unavailable(id: string) {
    if (loading || !loaded) return "Wait for your saved setup to load."
    if (failed) return "Retry setup before changing steps."
    if (advancing) return "Wait for setup progress to finish saving."
    if (id === "connection" && !provider) return "Choose a provider first."
    if (id === "models" && !selectable) return "Connect a provider with available models first."
    if ((id === "research" || id === "details") && !modelsSaved) return "Save your model choices first."
    return ""
  }
  function navigate(id: string) {
    if (unavailable(id)) return
    if (id === "provider" || id === "connection" || id === "models") { setPhase(id); setStep("setup"); return }
    if (id === "research" || id === "details") setStep(id)
  }
  function completed(id: string) {
    if (id === "provider") return !!provider
    if (id === "connection") return selectable
    if (id === "models") return modelsSaved
    return (id === "research" || id === "details") && !!progress?.completed.includes(id)
  }
  function changeProvider(id: string) {
    setProvider(id)
    const defaults = pickForProvider(models, id)
    setPrimary(value => models.some(model => model.value === value && model.providerId === id) ? value : defaults.primary)
    setFast(value => models.some(model => model.value === value && model.providerId === id) ? value : defaults.fast)
  }
  async function advance(which: "details" | "research", status: "completed" | "skipped") {
    if (advancing) return
    setAdvancing(true); setSaveNotice("")
    try {
      const next = await saveSetup(which, status)
      setProgress(next)
      if (next.next === "done") return onClose()
      setStep(next.next === "models" ? "setup" : next.next)
    } finally { setAdvancing(false) }
  }
  return <div className="viewer-overlay onboarding-overlay">
    <div ref={panel} className="onboarding-panel" role="dialog" aria-modal="true" aria-labelledby="onboarding-title" aria-describedby="onboarding-description" onKeyDown={event => {
      if (event.key === "Escape" && !document.querySelector(".jx-popover")) { event.stopPropagation(); onClose() }
    }}>
      <header className="onboarding-header">
        <div className="onboarding-topline"><div><span className="onboarding-brand">doc.haus</span><h1 id="onboarding-title">Let's get you set up</h1></div><button className="ghost" onClick={onClose}>Set up later <span aria-hidden="true">×</span></button></div>
        <nav aria-label="Setup progress" className="onboarding-progress">
          <ol>{stages.map((item, index) => {
            const reason = unavailable(item.id)
            const done = completed(item.id)
            const skipped = (item.id === "research" || item.id === "details") && progress?.skipped.includes(item.id)
            return <li key={item.id} className={index === current ? "is-current" : done ? "is-complete" : ""} aria-current={index === current ? "step" : undefined} title={reason || undefined}>
              <button type="button" className="onboarding-step-button" disabled={!!reason} aria-describedby={reason ? `onboarding-${item.id}-unavailable` : undefined} onClick={() => navigate(item.id)}>
                <span className="onboarding-step-number" aria-hidden="true">{done && index !== current ? <svg viewBox="0 0 16 16" fill="none"><path d="m4 8 2.5 2.5L12 5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" /></svg> : index + 1}</span>
                <span className="onboarding-step-label"><span className="onboarding-sr-only">{done ? "Completed: " : skipped ? "Skipped: " : `Step ${index + 1}: `}</span>{item.label}</span>
              </button>
              {reason && <span id={`onboarding-${item.id}-unavailable`} className="onboarding-sr-only">{reason}</span>}
            </li>
          })}</ol>
        </nav>
        <p className="onboarding-mobile-progress">{stage.label}</p>
      </header>
      <div ref={body} className="onboarding-body" aria-busy={loading}>
        <div className="onboarding-stage-heading"><h2 ref={heading} tabIndex={-1}>{step === "setup" && phase === "connection" ? `Connect ${selected?.name ?? "a local endpoint"}` : stage.title}</h2><p id="onboarding-description">{stage.description}</p></div>
        {notice && <p role="alert" className="settings-notice">{notice}</p>}
        {saveNotice && <p role="alert" className="settings-notice">{saveNotice}</p>}
        {loading && loaded && <p className="settings-notice" role="status">Refreshing providers and your saved choices…</p>}
        {failed && loaded && <div className="onboarding-status"><p>Could not refresh setup. Retry to update the provider catalog.</p><button onClick={() => void load()}>Retry setup</button></div>}
        {loading && !loaded ? <div className="onboarding-status" role="status"><span className="onboarding-loading-dot" aria-hidden="true" />Loading providers and your saved choices…</div>
          : failed && !loaded ? <div className="onboarding-status"><p>Setup could not be loaded. Your saved choices are still available when you retry.</p><button onClick={() => void load("", true)}>Retry setup</button></div>
          : <>
          {currentStage === "provider" && <>
            <section className="settings-section onboarding-provider-card">
            <label className="onboarding-field">Model provider<select disabled={loading} value={provider} onChange={event => changeProvider(event.target.value)}>
              <option value="">Choose a provider</option>
              <optgroup label="Detected">{active.filter(detected).map(item => <option key={item.id} value={item.id}>{item.name}{item.publicModels ? " · Public models" : item.host?.type === "google-adc" && !item.host.options.project ? " · Project required" : item.host && !verified.has(item.id) ? " · Setup / verification required" : " · Detected"}</option>)}</optgroup>
              <optgroup label="Other providers">{active.filter(item => !detected(item)).map(item => <option key={item.id} value={item.id}>{item.name}</option>)}</optgroup>
              <option value="__local_endpoint__">Local / OpenAI-compatible endpoint</option>
            </select></label>
            <div className="onboarding-note"><span aria-hidden="true">↳</span><p>Use a cloud provider or connect a local, OpenAI-compatible endpoint.</p></div>
            </section>
            <div className="onboarding-actions"><p className="onboarding-action-hint">You can change your provider in Settings.</p><button className="primary" disabled={loading || failed || !provider} onClick={() => setPhase("connection")}>Continue to connection <span aria-hidden="true">→</span></button></div>
          </>}
          {(mounted.connection || currentStage === "connection") && <div className="onboarding-stage-content" hidden={currentStage !== "connection"} inert={currentStage !== "connection"}>
            {provider === "__local_endpoint__" ? <LocalEndpoint onAdd={async input => { await addLocalProvider(input); await load(input.id) }} />
              : selected && isGated(selected.id) ? verified.has(selected.id) ? <>
                <p className="onboarding-connection-status" role="status"><span aria-hidden="true">✓</span><span>Connection verified. Continue to choose your models.</span></p>
                <details className="onboarding-connection-options"><summary>Change cloud connection (optional)</summary><section className="settings-section onboarding-cloud-connection"><HostProviderSetup client={client} provider={selected} onChanged={async () => { await load() }} /></section></details>
              </> : <section className="settings-section onboarding-cloud-connection"><HostProviderSetup client={client} provider={selected} onChanged={async () => { await load() }} /></section>
              : selected && <ConnectionStep key={selected.id} client={client} provider={selected} connected={connected.has(selected.id)} onConnected={async id => { await load(id) }} />}
            <div className="onboarding-actions"><button onClick={() => setPhase("provider")}>Back to providers</button><button className="primary" disabled={loading || failed || !selectable} onClick={() => setPhase("models")}>Continue to models <span aria-hidden="true">→</span></button></div>
          </div>}
          {currentStage === "models" && <>
            <div className="onboarding-provider-summary"><span className="onboarding-provider-icon" aria-hidden="true">✓</span><div><span className="onboarding-eyebrow">Model provider</span><strong>{selected?.name ?? provider}</strong></div><span className="onboarding-provider-ready">Connected</span></div>
            <ModelPickerStep hideProviderSelector connectedProviders={ready.filter(item => item.id === provider)} provider={provider} changeProvider={changeProvider} primary={primary} setPrimary={setPrimary} providerModels={models.filter(model => model.providerId === provider)} fast={fast} setFast={setFast} setForceConnect={() => setPhase("connection")} setNotice={setNotice} onSaved={async () => { const next = await saveSetup("models", "completed"); setProgress(next); if (next.next === "done") onClose(); else setStep(next.next === "models" ? "setup" : next.next) }} />
          </>}
          {currentStage === "details" && <DraftingDetailsStep drafting={drafting} setDrafting={setDrafting} jurisdictions={jurisdictions} selectedJx={selectedJx} setSelectedJx={setSelectedJx} dateFormatChoice={dateFormatChoice} setDateFormatChoice={setDateFormatChoice} setStep={setStep} onClose={() => advance("details", "completed")} />}
          {(mounted.research || currentStage === "research") && <div className="onboarding-stage-content onboarding-research" hidden={currentStage !== "research"} inert={currentStage !== "research"}>
            <ResearchTab heading="Sources and connections" actionLabel="Save and continue" secondaryActions={busy => <button disabled={busy || advancing} onClick={() => { setPhase("models"); setStep("setup") }}>Back to models</button>} onComplete={() => { void advance("research", "completed").catch(error => setSaveNotice(error instanceof Error ? error.message : "Could not save setup progress. Retry.")) }} />
          </div>}
        </>}
      </div>
      <footer className="onboarding-footer"><p>Your saved setup resumes when you return.</p></footer>
    </div>
  </div>
}
