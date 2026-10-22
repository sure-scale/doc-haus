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
import backdropClose from "../backdropClose"

// AI setup shares its provider, connection, model and test flow with onboarding.
const TABS: { id: Tab; label: string }[] = [
  { id: "ai", label: "AI setup" },
  { id: "drafting", label: "Drafting" },
  { id: "research", label: "Research" },
  { id: "data", label: "Statistics" },
  { id: "matters", label: "Matter defaults" },
  { id: "approvals", label: "Approvals" },
  { id: "appearance", label: "Appearance" },
]

export default function Settings({ onClose, closing = false, firstRun = false, initialTab, reconnectProvider }: { onClose: () => void; closing?: boolean; firstRun?: boolean; initialTab?: Tab; reconnectProvider?: string }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const titleId = useId()
  const tabsId = useId()
  const [narrow, setNarrow] = useState(() => window.matchMedia("(max-width: 700px)").matches)
  // Closes, and returns focus, as the exit starts; only the picture lingers.
  useLayoutEffect(() => {
    if (closing) return
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const modal = dialog.current
    modal?.showModal()
    modal?.querySelector<HTMLButtonElement>(".modal-close")?.focus({ preventScroll: true })
    return () => {
      modal?.close()
      if (opener?.isConnected) opener.focus({ preventScroll: true })
    }
  }, [closing])

  const [tab, setTab] = useState<Tab>(() => initialTab ?? "ai")
  useEffect(() => {
    const query = window.matchMedia("(max-width: 700px)")
    const update = () => setNarrow(query.matches)
    query.addEventListener("change", update)
    return () => query.removeEventListener("change", update)
  }, [])
  const toast = useToast()
  const [busy, setBusy] = useState(false)
  function switchTab(next: Tab) {
    if (busy) return
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
      onCancel={(event) => { event.preventDefault(); if (!busy) onClose() }}
      {...backdropClose(() => { if (!busy) onClose() })}
    >
      <div className="picker-panel settings-panel" onClick={(e) => e.stopPropagation()}>
        <div className="viewer-bar">
          <span className="viewer-title" id={titleId}>Settings</span>
          <button className="modal-close" aria-label="Close" disabled={busy} onClick={onClose}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12" /></svg>
          </button>
        </div>
        <div className="settings-body">
          <div className="settings-nav" role="tablist" aria-label="Settings categories" aria-orientation={narrow ? "horizontal" : "vertical"} onKeyDown={navigateTabs}>
            <p className="settings-nav-caption" aria-hidden="true">Your workspace</p>
            {TABS.map((t) => (
              <button type="button" key={t.id} id={`${tabsId}-${t.id}`} role="tab" aria-selected={tab === t.id} aria-controls={`${tabsId}-panel`} tabIndex={tab === t.id ? 0 : -1} disabled={busy} className={tab === t.id ? "active" : ""} onClick={() => switchTab(t.id)}>
                {t.label}
              </button>
            ))}
          </div>
          <div className="settings-content" id={`${tabsId}-panel`} role="tabpanel" aria-labelledby={`${tabsId}-${tab}`} tabIndex={0}>
            <div hidden={tab !== "ai"} inert={tab !== "ai"}><AISetup reconnectProvider={reconnectProvider} onProgress={progress => setBusy(progress.saving === true)} onSaved={() => { toast("success", "AI setup saved."); if (firstRun) onClose() }} /></div>

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
