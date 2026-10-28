import { useEffect, useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from "react"
import { createPortal } from "react-dom"
import { useToast } from "../Toast"
import type { Tab } from "./types"
import AISetup from "../ai-setup/AISetup"
import DraftingTab from "./DraftingTab"
import ResearchTab from "./ResearchTab"
import StatisticsTab from "./StatisticsTab"
import MatterDefaultsTab from "./MatterDefaultsTab"
import ApprovalsTab from "./ApprovalsTab"
import AppearanceTab from "./AppearanceTab"
import WordAddin from "./WordAddin"
import { connectionMessage, providerHealth } from "../../api/opencode/provider-health"
import backdropClose from "../backdropClose"
import { SettingsFooterContext } from "./SettingsFooter"

// AI setup shares its provider, connection, model and test flow with onboarding.
const TABS: { id: Tab; label: string }[] = [
  { id: "ai", label: "AI setup" },
  { id: "drafting", label: "Drafting" },
  { id: "research", label: "Research" },
  { id: "data", label: "Statistics" },
  { id: "matters", label: "Matter defaults" },
  { id: "addins", label: "Add-ins" },
  { id: "approvals", label: "Approvals" },
  { id: "appearance", label: "Appearance" },
]

export default function Settings({ onClose, closing = false, firstRun = false, recoveryRequired = false, initialTab, reconnectProvider, onSampleAdded }: { onClose: () => void; closing?: boolean; firstRun?: boolean; recoveryRequired?: boolean; initialTab?: Tab; reconnectProvider?: string; onSampleAdded?: () => void }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const titleId = useId()
  const tabsId = useId()
  const [footer, setFooter] = useState<HTMLDivElement | null>(null)
  const [narrow, setNarrow] = useState(() => window.matchMedia("(max-width: 700px)").matches)
  // Closes, and returns focus, as the exit starts; only the picture lingers.
  useLayoutEffect(() => {
    if (closing) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = dialog.current
    modal?.showModal();
    (modal?.querySelector<HTMLElement>(".modal-close") ?? modal?.querySelector<HTMLElement>(".viewer-title"))?.focus({ preventScroll: true })
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [closing])

  const [tab, setTab] = useState<Tab>(() => initialTab ?? "ai")
  useEffect(() => { if (recoveryRequired) setTab("ai") }, [recoveryRequired])
  useEffect(() => {
    const query = window.matchMedia("(max-width: 700px)")
    const update = () => setNarrow(query.matches)
    query.addEventListener("change", update)
    return () => query.removeEventListener("change", update)
  }, [])
  const toast = useToast()
  const [aiSaving, setAISaving] = useState(false)
  const [wordBusy, setWordBusy] = useState(false)
  const [matterBusy, setMatterBusy] = useState(false)
  const busy = aiSaving || wordBusy || matterBusy
  function switchTab(next: Tab) {
    if (busy || (recoveryRequired && next !== "ai")) return
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
    switchTab(TABS[target].id)
    const control = document.getElementById(`${tabsId}-${TABS[target].id}`)
    control?.focus({ preventScroll: true })
    control?.scrollIntoView({ block: "nearest", inline: "nearest" })
  }

  return createPortal(
    <dialog
      ref={dialog}
      className={`viewer-overlay settings-overlay${closing ? " is-closing" : ""}`}
      inert={closing}
      aria-labelledby={titleId}
      onCancel={(event) => { event.preventDefault(); if (!busy && !recoveryRequired) onClose() }}
      {...backdropClose(() => { if (!busy && !recoveryRequired) onClose() })}
    >
      <div className="picker-panel settings-panel" onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <span className="viewer-title" id={titleId} tabIndex={-1}>Settings</span>
          {!recoveryRequired && <button className="modal-close" aria-label="Close" disabled={busy} onClick={onClose}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12" /></svg>
          </button>}
        </div>
        <div className="settings-body">
          <div className="settings-nav" role="tablist" aria-label="Settings categories" aria-orientation={narrow ? "horizontal" : "vertical"} onKeyDown={navigateTabs}>
            {TABS.map((t) => (
              <button type="button" key={t.id} id={`${tabsId}-${t.id}`} role="tab" aria-selected={tab === t.id} aria-controls={`${tabsId}-panel`} tabIndex={tab === t.id ? 0 : -1} disabled={busy || (recoveryRequired && t.id !== "ai")} className={tab === t.id ? "active" : ""} onClick={() => switchTab(t.id)}>
                {t.label}
              </button>
            ))}
          </div>
          <div className="settings-main"><SettingsFooterContext.Provider value={{ target: footer, tab }}><div className="settings-content" id={`${tabsId}-panel`} role="tabpanel" aria-labelledby={`${tabsId}-${tab}`} tabIndex={0}>
            <div hidden={tab !== "ai"} inert={tab !== "ai"}>{recoveryRequired && <div role="alert" className="settings-notice"><p>{connectionMessage(providerHealth.getSnapshot().find(item => ["auth", "model", "unavailable"].includes(item.status)) ?? { providerID: reconnectProvider ?? "", modelID: "", status: "unavailable" })}</p><button type="button" disabled={busy} onClick={() => void providerHealth.check(true).then(() => { if (providerHealth.getSnapshot().length > 0 && providerHealth.getSnapshot().every(item => item.status === "ready")) onClose() })}>Check again</button></div>}<AISetup key={`${recoveryRequired}:${reconnectProvider ?? ""}`} reconnectProvider={reconnectProvider} onProgress={progress => setAISaving(progress.saving === true)} onSaved={() => { toast("success", "AI setup saved."); if (firstRun || recoveryRequired) onClose() }} /></div>

            {tab === "drafting" && <DraftingTab onSaved={(m) => toast("success", m)} />}
            {tab === "research" && <ResearchTab onSaved={(m) => toast("success", m)} />}
            {tab === "data" && <StatisticsTab onSaved={(m) => toast("success", m)} onError={(m) => toast("error", m)} />}
            {tab === "matters" && <MatterDefaultsTab onSampleAdded={() => { toast("success", "Sample matter added to your matters."); onSampleAdded?.() }} onError={(m) => toast("error", m)} onSaved={() => toast("success", "Matter settings saved.")} onBusy={setMatterBusy} />}
            {tab === "addins" && <WordAddin onBusy={setWordBusy} />}
            {tab === "approvals" && <ApprovalsTab />}
            {tab === "appearance" && <AppearanceTab />}
          </div><div className="settings-footer" ref={setFooter} /></SettingsFooterContext.Provider></div>
        </div>
      </div>
    </dialog>,
    document.body,
  )
}
