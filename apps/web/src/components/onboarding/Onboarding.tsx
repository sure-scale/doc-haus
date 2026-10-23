import { useContext, useEffect, useRef, useState } from "react"
import { getPreferences, type DraftingPreferences } from "../../api/ingest/preferences"
import { addSampleMatter, listJurisdictions, listMatters, type Jurisdiction } from "../../api/ingest/matters"
import { humanizeError } from "../../api/ingest/errors"
import { getSetup, saveSetup, type SetupProgress } from "../../api/ingest/setup"
import { loadPrefs } from "../../prefs"
import { Closing } from "../useExit"
import { useToast } from "../Toast"
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

export default function Onboarding({ onClose, onSampleAdded }: { onClose: () => void; onSampleAdded?: () => void }) {
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
  // The sample matter choice on the last step: pre-selected only for an empty
  // workspace, and offered only while no sample exists. Kept across steps.
  const [addSample, setAddSample] = useState<boolean | null>(null)
  const [sampleExists, setSampleExists] = useState(false)
  const [sampleState, setSampleState] = useState<"idle" | "adding" | "failed">("idle")
  const [sampleError, setSampleError] = useState("")
  // Where setup goes once the sample is in, so a failed seed can still move on.
  const [sampleNext, setSampleNext] = useState<SetupProgress["next"]>("done")
  const toast = useToast()
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
      const [next, preferences, list, matters] = await Promise.all([getSetup(), getPreferences(), listJurisdictions(), listMatters().catch(() => null)])
      setProgress(next); setDrafting(preferences); setJurisdictions(list)
      if (matters) { setSampleExists(matters.some(matter => matter.sample)); setAddSample(choice => choice ?? matters.length === 0) }
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
  function leave(next: SetupProgress["next"]) {
    if (next === "done") onClose()
    else setStep(next === "models" ? "setup" : next)
  }
  // Setup progress is saved before seeding, so a failed seed never undoes it.
  async function seedSample(next: SetupProgress["next"]) {
    setSampleState("adding"); setSampleError(""); setSampleNext(next)
    try {
      await addSampleMatter()
      setSampleExists(true); setSampleState("idle")
      onSampleAdded?.()
      toast("success", "Sample matter added to your matters.")
      leave(next)
    } catch (error) {
      setSampleState("failed"); setSampleError(humanizeError(error, "Could not add the sample matter."))
    }
  }
  async function finishDetails() {
    if (advancing) return
    setAdvancing(true); setSaveNotice("")
    try {
      const next = await saveSetup("details", "completed")
      setProgress(next)
      if (addSample && !sampleExists) await seedSample(next.next)
      else leave(next.next)
    } finally { setAdvancing(false) }
  }
  async function retrySample() {
    if (advancing) return
    setAdvancing(true)
    try { await seedSample(sampleNext) } finally { setAdvancing(false) }
  }
  const sampleChoice = !sampleExists && addSample !== null && <section className="settings-section onboarding-sample">
    <h3>Sample matter</h3>
    <label className="page-layout-check" htmlFor="onboarding-sample">
      <input id="onboarding-sample" type="checkbox" checked={addSample} disabled={advancing} aria-describedby="onboarding-sample-hint" onChange={event => setAddSample(event.target.checked)} />
      <span>Add a sample matter to explore</span>
    </label>
    <p id="onboarding-sample-hint" className="settings-hint muted">A fictional engagement letter, ready for cited questions and a legal review. Delete it like any matter when you are done.</p>
    {sampleState === "adding" && <p role="status" className="settings-hint">Adding the sample matter…</p>}
    {sampleState === "failed" && <div role="alert" className="settings-warning onboarding-sample-error">
      <p>{sampleError} Your setup is saved.</p>
      <div className="onboarding-sample-actions"><button disabled={advancing} onClick={() => void retrySample()}>Retry</button><button disabled={advancing} onClick={() => { setSampleState("idle"); leave(sampleNext) }}>Continue without it</button></div>
    </div>}
  </section>
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
          {currentStage === "details" && <DraftingDetailsStep drafting={drafting} setDrafting={setDrafting} jurisdictions={jurisdictions} selectedJx={selectedJx} setSelectedJx={setSelectedJx} dateFormatChoice={dateFormatChoice} setDateFormatChoice={setDateFormatChoice} setStep={setStep} onClose={finishDetails} extra={sampleChoice} busyLabel={sampleState === "adding" ? "Adding sample matter…" : undefined} />}
          {(researchMounted || currentStage === "research") && <div className="onboarding-stage-content onboarding-research" hidden={currentStage !== "research"} inert={currentStage !== "research"}>
            <ResearchTab heading="Sources and connections" actionLabel="Save and continue" secondaryActions={busy => <button disabled={busy || advancing} onClick={() => { setPhase("provider"); setStep("setup") }}>Back to AI setup</button>} onComplete={() => { void advance("research", "completed").catch(error => setSaveNotice(error instanceof Error ? error.message : "Could not save setup progress. Retry.")) }} />
          </div>}
        </>}
      </div>
      <footer className="onboarding-footer"><p>Your saved setup resumes when you return.</p></footer>
    </div>
  </div>
}
