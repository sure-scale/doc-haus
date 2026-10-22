import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent } from "react"
import { createPortal } from "react-dom"
import { getConfig, setDisabledProviders, setProviderOptions } from "../../api/opencode/config"
import { listAuthMethods, listProviders, probeProvider } from "../../api/opencode/providers"
import { settingsClient, type Client } from "../../api/opencode/client"
import { defaultProvider, pickForProvider, probeModel, providerOf } from "../../models"
import { isRemovedModel, isRemovedProvider } from "../../api/opencode/provider-readiness"
import { isGated, loadVerified, saveVerified } from "../../providers"
import { useToast } from "../Toast"
import type { Methods, Provider, Tab } from "./types"
import { CLOUD_SETUPS, vertexAnthropicBaseURL } from "./cloud-setups"
import ModelsTab from "./ModelsTab"
import ProvidersTab from "./ProvidersTab"
import DraftingTab from "./DraftingTab"
import ResearchTab from "./ResearchTab"
import StatisticsTab from "./StatisticsTab"
import MatterDefaultsTab from "./MatterDefaultsTab"
import ApprovalsTab from "./ApprovalsTab"
import AppearanceTab from "./AppearanceTab"

// The settings modal: a left rail of tabs over one shared shell. Models and
// Providers act on the engine's global config + auth store (so they share the
// header-less settings client and provider state); Drafting and Research persist
// standing instructions through the ingest service; Matter defaults, Approvals,
// and Appearance are per-browser preferences (see prefs.ts).
const TABS: { id: Tab; label: string }[] = [
  { id: "models", label: "Models" },
  { id: "providers", label: "Providers" },
  { id: "drafting", label: "Drafting" },
  { id: "research", label: "Research" },
  { id: "data", label: "Statistics" },
  { id: "matters", label: "Matter defaults" },
  { id: "approvals", label: "Approvals" },
  { id: "appearance", label: "Appearance" },
]

export default function Settings({ onClose, firstRun = false, initialTab, reconnectProvider }: { onClose: () => void; firstRun?: boolean; initialTab?: Tab; reconnectProvider?: string }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const titleId = useId()
  const tabsId = useId()
  const [narrow, setNarrow] = useState(() => window.matchMedia("(max-width: 700px)").matches)
  useLayoutEffect(() => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = dialog.current
    modal?.showModal()
    modal?.querySelector<HTMLButtonElement>(".modal-close")?.focus({ preventScroll: true })
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [])

  // First run is about getting a provider connected, so land there; otherwise
  // open on the models people come back to tweak.
  const [tab, setTab] = useState<Tab>(initialTab ?? (firstRun ? "providers" : "models"))
  useEffect(() => {
    const query = window.matchMedia("(max-width: 700px)")
    const update = () => setNarrow(query.matches)
    query.addEventListener("change", update)
    return () => query.removeEventListener("change", update)
  }, [])
  const toast = useToast()
  const client = useMemo<Client>(() => settingsClient(), [])
  const [all, setAll] = useState<Provider[]>([])
  const [connected, setConnected] = useState<Set<string>>(new Set())
  const [methods, setMethods] = useState<Methods>({})
  // Both models are locked to one provider so they can never straddle clouds
  // (the Auto router runs the fast model — see routeAgent).
  const [provider, setProvider] = useState("")
  const [model, setModel] = useState("")
  // The engine's small/fast model, used for title generation and for routing the
  // chat's "Auto" assistant. Picked from the same provider models as the default.
  const [smallModel, setSmallModelValue] = useState("")
  const [disabled, setDisabled] = useState<string[]>([])
  // Host-credential providers (Vertex/Bedrock) report "connected" on project/region
  // presence alone — the engine never checks the sign-in actually resolves. So we
  // keep them out of the model picker until a live probe (probeProvider) confirms
  // real credentials, and remember the ones that passed across reloads.
  const [verified, setVerified] = useState<Set<string>>(loadVerified)
  const [probing, setProbing] = useState("")
  const [probeErr, setProbeErr] = useState<Record<string, string>>({})

  async function load() {
    const [providers, auth, cfg] = await Promise.all([
      listProviders(client),
      listAuthMethods(client),
      getConfig(),
    ])
    setAll(providers.all)
    setConnected(new Set(providers.connected))
    setMethods(auth)
    setModel(cfg.model ?? "")
    setSmallModelValue(cfg.small_model ?? "")
    setDisabled((cfg.disabled_providers ?? []).filter(id => !isRemovedProvider(id)))
  }

  // Enable/disable a provider by editing config.disabled_providers. A disabled
  // provider drops out of the catalog entirely, so we flip local state straight
  // away (the engine reports stale `connected` right after the write) and keep
  // the name around to label its row once it leaves `all`.
  const nameByID = new Map(all.map((p) => [p.id, p.name] as const))
  // Remember names of providers we disable, since they vanish from `all`.
  const [nameMemo, setNameMemo] = useState<Record<string, string>>({})
  async function toggle(id: string, name: string, enable: boolean) {
    const next = enable ? disabled.filter((x) => x !== id) : [...disabled, id]
    setDisabled(next)
    setConnected((prev) => {
      const n = new Set(prev)
      if (enable) n.add(id)
      else n.delete(id)
      return n
    })
    setNameMemo((m) => ({ ...m, [id]: name }))
    await setDisabledProviders(next)
  }

  useEffect(() => {
    void load().catch(() => toast("error", "Could not load settings. Close settings and retry."))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function switchTab(next: Tab) {
    setTab(next)
  }

  function navigateTabs(event: KeyboardEvent<HTMLDivElement>) {
    const index = TABS.findIndex(item => item.id === tab)
    const previous = narrow ? "ArrowLeft" : "ArrowUp"
    const next = narrow ? "ArrowRight" : "ArrowDown"
    let target = index
    if (event.key === previous) target = (index + TABS.length - 1) % TABS.length
    else if (event.key === next) target = (index + 1) % TABS.length
    else if (event.key === "Home") target = 0
    else if (event.key === "End") target = TABS.length - 1
    else return
    event.preventDefault()
    const item = TABS[target]
    setTab(item.id)
    const control = document.getElementById(`${tabsId}-${item.id}`)
    control?.focus({ preventScroll: true })
    control?.scrollIntoView({ block: "nearest", inline: "nearest" })
  }

  const connectedProviders = all.filter((p) => connected.has(p.id))

  // Vertex/Bedrock need a passing credential probe before they count as ready;
  // every other (API-key/OAuth) provider is ready the moment it's connected.
  const readyProviders = connectedProviders.filter((p) => !isGated(p.id) || verified.has(p.id))

  // Host-credential providers in the catalog that haven't passed a probe yet —
  // shown under "Needs setup" with their project/region fields and an Enable button.
  const needsSetup = all.filter((p) => isGated(p.id) && p.activation !== "disabled" && !disabled.includes(p.id) && !verified.has(p.id))

  // A cloud card shows only while its primary provider is unverified. Secondary
  // ids on the same card (Claude on Vertex) verify silently when the shared
  // check passes and silently stay out of the model picker when it doesn't —
  // showing their failure under a card that just verified reads as the whole
  // setup being broken when it's only an optional model family without quota.
  const setupCards = CLOUD_SETUPS.flatMap(spec => spec.ids.filter(id => needsSetup.some(p => p.id === id)).map(id => ({ ...spec, ids: [id], name: all.find(p => p.id === id)?.name ?? spec.name })))

  // Save the card's project/region to every provider it covers, then probe each
  // one's credentials with one real call; each that passes is marked verified
  // (persisted) and joins the picker, each that fails stays gated. Only the
  // card's primary id surfaces its error — a secondary like Claude-on-Vertex
  // (separate provider id, same GCP project) fails silently and its models stay
  // out of the picker until a later Reconfigure passes. The probe model must be
  // the curated one — the first id in the merged Vertex catalog is a
  // Claude-on-Vertex model that hits a different endpoint and 400s even when
  // the Gemini setup is correct.
  async function enable(targets: Provider[], options?: Record<string, string>, testModelID?: string, configureOnly = false) {
    if (!targets.length || probing) return
    setProbing(targets[0].id)
    try {
    setProbeErr((e) => ({ ...e, ...Object.fromEntries(targets.map((p) => [p.id, ""])) }))
    // Sequential: setProviderOptions does a read-merge-write of the whole
    // provider map, so concurrent saves would drop each other's entry.
    // Claude-on-Vertex needs the endpoint spelled out: its SDK builds the host
    // as "{location}-aiplatform.googleapis.com", which is wrong for "global"
    // (the global endpoint has no location prefix), so the probe hangs on a
    // bogus host instead of failing with a real credential/quota error.
    const optionsFor = (id: string) =>
      id === "google-vertex-anthropic" && options?.project
        ? { ...options, baseURL: vertexAnthropicBaseURL(options.project, options.location || "global") }
        : options
    if (options && Object.keys(options).length) for (const p of targets) {
      unverify(p.id)
      await setProviderOptions(p.id, optionsFor(p.id)!)
    }
    if (configureOnly) {
      await load()
      toast("success", "Connection settings saved. Refresh the catalog if needed, then choose a model to test.")
      return
    }
    const results = await Promise.all(
      targets.map(async (p) => {
        const selected = [model, smallModel].find(value => providerOf(value) === p.id && Object.keys(p.models).some(id => value === `${p.id}/${id}`))
        const modelID = testModelID || (selected ? selected.slice(p.id.length + 1) : probeModel(p.id, Object.keys(p.models)))
        if (!modelID) return { p, ok: false as const, error: "Provider has no models to test." }
        // All providers use the engine's native model resolution, credentials
        // and bounded, stateless provider check against the saved settings.
        const result = await probeProvider(client, p.id, modelID)
        if (!result.ok) return { p, ok: false as const, error: result.error || "Verification failed." }
        return { p, ok: true as const, error: "" }
      }),
    )
    setProbing("")
    setProbeErr((e) => ({ ...e, ...Object.fromEntries(results.map((r) => [r.p.id, r.error])) }))
    const passed = results.filter((r) => r.ok)
    if (!passed.length) return
    setVerified((prev) => {
      const next = new Set(prev)
      passed.forEach((r) => next.add(r.p.id))
      saveVerified(next)
      return next
    })
    toast("success", `${passed.map((r) => r.p.name).join(", ")} verified.`)
    await load()
    } catch { toast("error", "Verification failed or timed out. Check the provider setup and retry.") }
    finally { setProbing("") }
  }

  // Drop a host provider back to "Needs setup" so its project/region can be changed
  // and re-verified.
  function unverify(id: string) {
    setVerified((prev) => {
      const next = new Set(prev)
      next.delete(id)
      saveVerified(next)
      return next
    })
  }

  // Providers re-enabled this session won't be back in `all` until the engine
  // refreshes its catalog, so render them from the connected set directly.
  const extraReady = [...connected].filter((id) => !isRemovedProvider(id) && !all.some((p) => p.id === id))

  // Every model across connected providers, as the "providerID/modelID" strings
  // the engine expects, for the model pickers.
  const models = readyProviders
    .flatMap((p) =>
      Object.values(p.models).map((m) => ({
        value: `${p.id}/${m.id}`,
        label: `${p.name} — ${m.name}`,
        providerId: p.id,
        modelId: m.id,
      })),
    )
    .sort((a, b) => a.label.localeCompare(b.label))

  // A removed saved choice requires an explicit replacement, even when only
  // one remaining provider is connected. Keep choices made in this session.
  useEffect(() => {
    if (!models.length) return
    setProvider((v) => v || (isRemovedModel(model) || isRemovedModel(smallModel) ? "" : providerOf(model) || defaultProvider(models)))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [models.length])

  // Switching provider re-clamps both slots into it: keep a model already on the
  // new provider, otherwise drop to that provider's curated default.
  function changeProvider(id: string) {
    setProvider(id)
    const defaults = pickForProvider(models, id)
    setModel((m) => (providerOf(m) === id ? m : defaults.primary))
    setSmallModelValue((s) => (providerOf(s) === id ? s : defaults.fast))
  }

  const providerModels = models.filter((m) => m.providerId === provider)

  return createPortal(
    <dialog
      ref={dialog}
      className="viewer-overlay settings-overlay"
      aria-labelledby={titleId}
      onCancel={(event) => { event.preventDefault(); onClose() }}
      onClick={(event) => { if (event.target === event.currentTarget) onClose() }}
    >
      <div className="picker-panel settings-panel" onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <span className="viewer-title" id={titleId}>Settings</span>
          <button className="modal-close" aria-label="Close" onClick={onClose}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12" /></svg>
          </button>
        </div>
        <div className="settings-body">
          <div className="settings-nav" role="tablist" aria-label="Settings categories" aria-orientation={narrow ? "horizontal" : "vertical"} onKeyDown={navigateTabs}>
            <p className="settings-nav-caption" aria-hidden="true">Your workspace</p>
            {TABS.map((t) => (
              <button type="button" key={t.id} id={`${tabsId}-${t.id}`} role="tab" aria-selected={tab === t.id} aria-controls={`${tabsId}-panel`} tabIndex={tab === t.id ? 0 : -1} className={tab === t.id ? "active" : ""} onClick={() => switchTab(t.id)}>
                {t.label}
              </button>
            ))}
          </div>
          <div className="settings-content" id={`${tabsId}-panel`} role="tabpanel" aria-labelledby={`${tabsId}-${tab}`} tabIndex={0}>
            {firstRun && tab === "providers" && (
              <p className="settings-notice">
                Welcome to doc.haus. Connect a model provider — from a host sign-in (gcloud/AWS), an API key, or a
                local endpoint — then pick your primary model under Models.
              </p>
            )}

            {tab === "models" && (
              <ModelsTab
                models={models}
                switchTab={switchTab}
                provider={provider}
                changeProvider={changeProvider}
                readyProviders={readyProviders}
                model={model}
                setModel={setModel}
                providerModels={providerModels}
                smallModel={smallModel}
                setSmallModelValue={setSmallModelValue}
                toast={toast}
                firstRun={firstRun}
                onClose={onClose}
              />
            )}

            {tab === "providers" && (
              <ProvidersTab
                reconnectProvider={reconnectProvider}
                readyProviders={readyProviders}
                needsSetup={needsSetup}
                disabled={disabled}
                methods={methods}
                unverify={unverify}
                toggle={toggle}
                extraReady={extraReady}
                nameMemo={nameMemo}
                nameByID={nameByID}
                setupCards={setupCards}
                probing={probing}
                probeErr={probeErr}
                enable={enable}
                client={client}
                toast={toast}
                load={load}
              />
            )}

            {tab === "drafting" && <DraftingTab onSaved={(m) => toast("success", m)} />}
            {tab === "research" && <ResearchTab onSaved={(m) => toast("success", m)} />}
            {tab === "data" && <StatisticsTab onSaved={(m) => toast("success", m)} onError={(m) => toast("error", m)} />}
            {tab === "matters" && <MatterDefaultsTab />}
            {tab === "approvals" && <ApprovalsTab />}
            {tab === "appearance" && <AppearanceTab />}
          </div>
        </div>
      </div>
    </dialog>,
    document.body,
  )
}
