import { useContext, useEffect, useRef, useState } from "react"
import { getPreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { listJurisdictions, type Jurisdiction } from "../../api/ingest/matters"
import { getSetup, saveSetup, type SetupProgress } from "../../api/ingest/setup"
import { loadPrefs } from "../../prefs"
import { Closing } from "../useExit"
import DraftingDetailsStep from "./DraftingDetailsStep"
import ResearchTab from "../settings/ResearchTab"
import AISetup, { type AISetupProgress, type SetupPhase } from "../ai-setup/AISetup"
import "./onboarding.css"

const stages = [
  { id: "provider", label: "Provider", title: "Choose your model provider", description: "Connect the service you want to use for drafting and working with your matters." },
  { id: "connection", label: "Connection", title: "Connect your provider", description: "Use an existing connection or sign in securely. Credentials are stored in this installation." },
  { id: "models", label: "Models", title: "Choose your models", description: "Pick a primary model for your matters and a fast model for smaller tasks." },
  { id: "test", label: "Test", title: "Test your connection", description: "Check that your connection works with both selected model roles, then save your setup." },
  { id: "research", label: "Research", title: "Set up legal research", description: "Choose research sources and add service keys." },
  { id: "details", label: "Details", title: "Make doc.haus yours", description: "Add the details you use in your drafts." },
]

export default function Onboarding({ onClose }: { onClose: () => void }) {
  // The trap lets go, and focus returns, as the exit starts.
  const closing = useContext(Closing)
  const [notice, setNotice] = useState("")
  const [saveNotice, setSaveNotice] = useState("")
  const [loading, setLoading] = useState(true)
  const [failed, setFailed] = useState(false)
  const [phase, setPhase] = useState<SetupPhase>("provider")
  const [step, setStep] = useState<"setup" | "details" | "research">("setup")
  const [ai, setAI] = useState<AISetupProgress>({ provider: false, connection: false, models: false, test: false, loading: true, busy: false })
  const [drafting, setDrafting] = useState<DraftingPreferences | null>(null)
  const [jurisdictions, setJurisdictions] = useState<Jurisdiction[]>([])
  const [selectedJx, setSelectedJx] = useState<string[]>(() => loadPrefs().defaultJurisdictions)
  // The firm's own date format pick; until it picks, the Details step follows
  // its default jurisdictions.
  const [dateFormatChoice, setDateFormatChoice] = useState<DraftingPreferences["dateFormat"] | null>(null)
  const [advancing, setAdvancing] = useState(false)
  const [progress, setProgress] = useState<SetupProgress | null>(null)
  const [researchMounted, setResearchMounted] = useState(false)
  const [aiMounted, setAIMounted] = useState(false)
  const panel = useRef<HTMLDivElement>(null)
  const body = useRef<HTMLDivElement>(null)
  const heading = useRef<HTMLHeadingElement>(null)
  const currentStage = step === "setup" ? phase : step
  const current = stages.findIndex(stage => stage.id === currentStage)
  const stage = stages[current]!

  useEffect(() => {
    if (closing) return
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
  }, [closing])
  useEffect(() => {
    if (currentStage === "research") setResearchMounted(true)
    body.current?.scrollTo({ top: 0 })
    heading.current?.focus({ preventScroll: true })
  }, [current])

  async function load() {
    setLoading(true); setFailed(false); setNotice("")
    try {
      const [next, preferences, list] = await Promise.all([getSetup(), getPreferences(), listJurisdictions()])
      setProgress(next); setDrafting(preferences); setJurisdictions(list)
      if (next.ready && (next.next === "details" || next.next === "research")) setStep(next.next)
    } catch (error) { setFailed(true); setNotice(error instanceof Error ? error.message : "Could not load setup. Retry.") }
    finally { setLoading(false) }
  }
  useEffect(() => { void load() }, [])
  useEffect(() => { if (!loading && step === "setup") setAIMounted(true) }, [loading, step])
  const modelsSaved = progress?.ready === true && progress.completed.includes("models")
  function unavailable(id: string) {
    if (loading) return "Wait for your saved setup to load."
    if (failed) return "Retry setup before changing steps."
    if (advancing || ai.busy) return "Wait for the current action to finish."
    if (id === "connection" && !ai.provider) return "Choose a provider first."
    if (id === "models" && !ai.connection) return "Prepare a provider connection first."
    if (id === "test" && (!ai.connection || !ai.models)) return "Choose your primary and fast models first."
    if ((id === "research" || id === "details") && !modelsSaved) return "Test and save your model choices first."
    return ""
  }
  function navigate(id: string) {
    if (unavailable(id)) return
    if (id === "provider" || id === "connection" || id === "models" || id === "test") { setPhase(id); setStep("setup"); return }
    if (id === "research" || id === "details") setStep(id)
  }
  function completed(id: string) {
    if (id === "provider" || id === "connection" || id === "models") return ai[id]
    if (id === "test") return modelsSaved
    return (id === "research" || id === "details") && !!progress?.completed.includes(id)
  }
  async function modelsComplete() {
    const next = await saveSetup("models", "completed")
    setProgress(next)
    if (next.next === "done") onClose()
    else setStep(next.next === "models" ? "setup" : next.next)
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
  return <div className={`viewer-overlay onboarding-overlay${closing ? " is-closing" : ""}`} inert={closing}>
    <div ref={panel} className="onboarding-panel" role="dialog" aria-modal="true" aria-labelledby="onboarding-title" aria-describedby="onboarding-description" onKeyDown={event => {
      if (event.key === "Escape" && !ai.busy && !advancing && !document.querySelector(".jx-popover")) { event.stopPropagation(); onClose() }
    }}>
      <header className="onboarding-header">
        <div className="onboarding-topline"><div><span className="onboarding-brand">doc.haus</span><h1 id="onboarding-title">Let's get you set up</h1></div><button className="ghost" disabled={ai.busy || advancing} onClick={onClose}>Set up later <span aria-hidden="true">×</span></button></div>
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
        <div className="onboarding-stage-heading"><h2 ref={heading} tabIndex={-1}>{stage.title}</h2><p id="onboarding-description">{stage.description}</p></div>
        {notice && <p role="alert" className="settings-notice">{notice}</p>}
        {saveNotice && <p role="alert" className="settings-notice">{saveNotice}</p>}
        {loading ? <div className="onboarding-status" role="status"><span className="onboarding-loading-dot" aria-hidden="true" />Loading your saved setup…</div>
          : failed ? <div className="onboarding-status"><p>Setup could not be loaded. Retry to continue.</p><button onClick={() => void load()}>Retry setup</button></div>
          : <>
          {(aiMounted || step === "setup") && <div className="onboarding-stage-content" hidden={step !== "setup"} inert={step !== "setup"}>
            <AISetup mode="onboarding" phase={phase} onPhaseChange={setPhase} onProgress={setAI} onSaved={modelsComplete} />
          </div>}
          {currentStage === "details" && <DraftingDetailsStep drafting={drafting} setDrafting={setDrafting} jurisdictions={jurisdictions} selectedJx={selectedJx} setSelectedJx={setSelectedJx} dateFormatChoice={dateFormatChoice} setDateFormatChoice={setDateFormatChoice} setStep={setStep} onClose={() => advance("details", "completed")} />}
          {(researchMounted || currentStage === "research") && <div className="onboarding-stage-content onboarding-research" hidden={currentStage !== "research"} inert={currentStage !== "research"}>
            <ResearchTab heading="Sources and connections" actionLabel="Save and continue" secondaryActions={busy => <button disabled={busy || advancing} onClick={() => { setPhase("provider"); setStep("setup") }}>Back to AI setup</button>} onComplete={() => { void advance("research", "completed").catch(error => setSaveNotice(error instanceof Error ? error.message : "Could not save setup progress. Retry.")) }} />
          </div>}
        </>}
      </div>
      <footer className="onboarding-footer"><p>Your saved setup resumes when you return.</p></footer>
    </div>
  </div>
}
