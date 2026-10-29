import { useEffect, useState, type FormEvent } from "react"
import { Link, useNavigate } from "react-router-dom"
import { addSampleMatter, type Matter } from "../../api/ingest/matters"
import { savePreferences } from "../../api/ingest/preferences"
import type { SetupProgress } from "../../api/ingest/setup"
import { loadPrefs, savePrefs } from "../../prefs"
import JurisdictionSelect from "../JurisdictionSelect"
import { SUGGESTED_PRACTICE_AREAS } from "../settings/MatterDefaultsTab"
import AISetup from "../ai-setup/AISetup"
import SettingsDialog from "../settings/SettingsDialog"
import ResearchTab from "../settings/ResearchTab"
import WordAddin from "../settings/WordAddin"
import DraftingTab from "../settings/DraftingTab"
import { FieldRow } from "../settings/controls"
import { SettingsFooter, SettingsFooterContext } from "../settings/SettingsFooter"
import type { DashboardData } from "./data"
import { icons } from "./icons"
import { SetupChecklist, type SetupChecklistConfig } from "./visual/SetupCards"
import type { WidgetContext } from "./widgets"

function completion(data: DashboardData, setup: SetupProgress | null, jurisdictions: string[]) {
  const prefs = data.preferences
  const hasChoice = (choices: string[]) => choices.some((choice) => typeof choice === "string" && !!choice.trim())
  return [
    setup?.ready === true,
    !!prefs?.attorney.trim() && !!prefs.firm.trim() && hasChoice(jurisdictions),
    !!prefs?.searchApiKey.trim(),
    data.word?.status === "ready" || (data.word?.status === "external" && data.word.available),
    prefs?.draftingConfigured === true,
  ]
}

// An unavailable service stays incomplete; opening Settings never marks it done.
export function setupTabProgress(data: DashboardData, setup: SetupProgress | null, jurisdictions = loadPrefs().defaultJurisdictions): { done: number; total: 5 } {
  return { done: completion(data, setup, jurisdictions).filter(Boolean).length, total: 5 }
}


export function SetupTab({ context, onSampleAdded, onHideSetup, onSetupRefresh }: { context: WidgetContext; onSampleAdded?: () => void; onHideSetup?: () => void; onSetupRefresh?: () => void }) {
  const { data, setup, onPreferences, onRetry } = context
  const navigate = useNavigate()
  const prefs = data.preferences
  const [attorney, setAttorney] = useState(prefs?.attorney ?? "")
  const [firm, setFirm] = useState(prefs?.firm ?? "")
  const [jurisdictions, setJurisdictions] = useState(() => loadPrefs().defaultJurisdictions)
  const [areas, setAreas] = useState(prefs?.practiceAreas ?? [])
  const [identityBusy, setIdentityBusy] = useState(false)
  const [identityError, setIdentityError] = useState("")
  const [identitySaved, setIdentitySaved] = useState("")
  const [exaDraft, setExaDraft] = useState<string>()
  const [stageBusy, setStageBusy] = useState(false)
  const [stage, setStage] = useState("")
  const [footer, setFooter] = useState<HTMLDivElement | null>(null)
  const saving = identityBusy || stageBusy
  const [adding, setAdding] = useState(false)
  const [demoError, setDemoError] = useState("")
  const [added, setAdded] = useState<Matter>()
  function closeStage() {
    setStage("")
    setStageBusy(false)
    onRetry()
    onSetupRefresh?.()
  }
  function openStage(id: string) {
    setStageBusy(false)
    setStage(id)
  }
  // Saving an Exa key or refreshing unrelated settings must not discard an
  // unfinished identity form. Synchronize only the saved fields that changed.
  useEffect(() => { setAttorney(prefs?.attorney ?? "") }, [prefs?.attorney])
  useEffect(() => { setFirm(prefs?.firm ?? "") }, [prefs?.firm])
  const savedAreas = JSON.stringify(prefs?.practiceAreas ?? [])
  useEffect(() => { setAreas(prefs?.practiceAreas ?? []) }, [savedAreas])
  const savedJurisdictions = JSON.stringify(loadPrefs().defaultJurisdictions)
  useEffect(() => { setJurisdictions(loadPrefs().defaultJurisdictions) }, [savedJurisdictions])
  const done = completion(data, setup, loadPrefs().defaultJurisdictions)
  const sample = data.matters.find((row) => row.matter.sample)?.matter ?? added
  const changed = attorney.trim() !== prefs?.attorney.trim() || firm.trim() !== prefs?.firm.trim()
    || JSON.stringify(areas) !== JSON.stringify(prefs?.practiceAreas) || JSON.stringify(jurisdictions) !== JSON.stringify(loadPrefs().defaultJurisdictions)
  const packs = data.jurisdictions ?? []
  const areaOptions = [...areas, ...SUGGESTED_PRACTICE_AREAS].filter((area, index, all) => all.findIndex((other) => other.toLocaleLowerCase() === area.toLocaleLowerCase()) === index).map((name) => ({ code: name, name }))
  async function saveIdentity(event: FormEvent) {
    event.preventDefault()
    if (!prefs || saving) return
    if (!attorney.trim() || !firm.trim()) {
      setIdentityError("Enter your name and firm name before saving.")
      setIdentitySaved("")
      return
    }
    setIdentityBusy(true); setIdentityError(""); setIdentitySaved("")
    try {
      const saved = await savePreferences({ attorney: attorney.trim(), firm: firm.trim(), practiceAreas: areas })
      try {
        savePrefs({ defaultJurisdictions: jurisdictions }, { requireStorage: true })
        // Publish after both saves, so the parent's progress reads the new
        // jurisdiction defaults on the very same render as the saved identity.
        onPreferences?.(saved)
        setIdentitySaved("Practice details saved.")
        closeStage()
      } catch {
        // The server save succeeded even when this browser could not save law defaults.
        onPreferences?.(saved)
        setIdentityError("Your name, firm and practice areas were saved. Jurisdiction defaults could not be saved in this browser. Try again.")
      }
    } catch {
      setIdentityError("Could not confirm that your practice details were saved. Your changes are still here. Try again.")
    } finally { setIdentityBusy(false) }
  }
  async function populate() {
    if (adding || sample) return
    setAdding(true); setDemoError("")
    try { setAdded(await addSampleMatter()); onSampleAdded?.(); onRetry() }
    catch { setDemoError("Could not add the sample matter. Try again.") }
    finally { setAdding(false) }
  }
  const config: SetupChecklistConfig = {
    title: "Finish setting up",
    description: "Connect your tools and defaults so every draft and search is ready.",
    items: [
      { id: "models", icon: "model", title: "AI provider & models", detail: done[0] ? "Your AI tools are connected and ready to use." : setup ? "Connect your AI tools so you can start drafting and searching." : "Open AI settings to check your connection.", done: !!done[0], action: { id: "ai", label: done[0] ? "Change" : "Connect" } },
      { id: "identity", icon: "person", title: "Practice identity & default law", detail: "Add your name, firm and the law you usually work with.", done: !!done[1], editable: true },
      { id: "research", icon: "search", title: "Web research with Exa", detail: done[2] ? "Your Exa API key is saved. Test or change it in Research settings." : "Add your Exa API key to search the web for legal and company information.", done: !!done[2], editable: !done[2], ...(done[2] ? { action: { id: "research", label: "Change" } } : {}) },
      { id: "word", icon: "word", title: "Microsoft Word add-in", detail: done[3] ? "Word is connected and ready for document review." : data.word?.status === "word_missing" ? "Install Word on this computer, then connect the add-in" : data.word?.status === "unsupported" ? "See add-in settings for supported ways to connect Word" : "Review documents and apply tracked changes directly in Microsoft Word.", done: !!done[3], action: { id: "addins", label: done[3] ? "Manage Word" : "Set up Word", variant: "outline" } },
      { id: "style", icon: "template", title: "Drafting preferences", detail: "Set how doc.haus writes, explains and formats your drafts.", done: !!done[4], action: { id: "drafting", label: "Configure", variant: "outline" } },
    ],
  }
  const selected = config.items.find(item => item.id === stage)
  const tab = stage === "models" ? "ai" : stage === "style" ? "drafting" : stage === "word" ? "addins" : stage
  const identity = prefs ? <form id="dash-setup-identity-form" onSubmit={saveIdentity} aria-label="Practice identity & default law">
    <fieldset disabled={saving} className="settings-drafting-fields">
      <FieldRow id="dash-setup-attorney" label="Your full name"><input id="dash-setup-attorney" value={attorney} onChange={event => { setAttorney(event.target.value); setIdentitySaved("") }} placeholder="Your full name" autoComplete="name" maxLength={200} required /></FieldRow>
      <FieldRow id="dash-setup-firm" label="Firm name"><input id="dash-setup-firm" value={firm} onChange={event => { setFirm(event.target.value); setIdentitySaved("") }} placeholder="Firm name" autoComplete="organization" maxLength={200} required /></FieldRow>
      <FieldRow id="dash-setup-jurisdictions" label="Jurisdictions" hint="Preselected when you create a new matter."><JurisdictionSelect id="dash-setup-jurisdictions" aria-describedby="dash-setup-jurisdictions-hint" jurisdictions={packs} selected={jurisdictions} onChange={values => { setJurisdictions(values); setIdentitySaved("") }} disabled={saving || !packs.length} /></FieldRow>
      <FieldRow id="dash-setup-practice-areas" label="Practice areas" hint="Optional. Select the areas you work in, or add your own."><JurisdictionSelect id="dash-setup-practice-areas" aria-describedby="dash-setup-practice-areas-hint" jurisdictions={areaOptions} selected={areas} onChange={values => { setAreas(values); setIdentitySaved("") }} practiceAreas disabled={saving} /></FieldRow>
    </fieldset>
    {identityError && <p className="settings-error" role="alert">{identityError}</p>}
    <SettingsFooter tab="identity"><div className="settings-actions"><button type="submit" form="dash-setup-identity-form" className="primary" disabled={saving || !changed} aria-busy={identityBusy}>{identityBusy ? "Saving…" : "Save"}</button></div></SettingsFooter>
  </form> : <p className="settings-error">Could not load your practice details. <button type="button" onClick={onRetry}>Retry</button></p>
  return <div className="dash-setup-flow">
    <div className="dash-setup-demo-row" role="group" aria-label="Demo data">
      <span className="dash-setup-demo-row-icon" aria-hidden="true">{icons.sample}</span>
      <div className="dash-setup-demo-copy"><p>Want to see doc.haus in action first?</p><span>{sample ? "A sample matter, documents and data are ready to explore." : "Populate your workspace with a sample matter, documents and data."}</span></div>
      <button type="button" className="dash-setup-demo-primary" disabled={adding} aria-busy={adding} onClick={() => { if (sample) navigate(`/matter/${sample.id}?view=chat&agent=ask`); else void populate() }}>{sample ? icons.arrowRight : icons.plus}{adding ? "Adding demo data…" : sample ? "Open sample matter" : "Add sample data"}</button>
      {onHideSetup && <button type="button" className="dash-text-button dash-setup-hide" onClick={onHideSetup}>Hide setup</button>}
      {demoError && <p className="dash-setup-error" role="alert">{demoError}</p>}
    </div>
    <p className="sr-only" role="status">{identitySaved}</p>
    <SetupChecklist onOpen={openStage} config={config} />
    {selected && <SettingsDialog title={selected.title} onClose={closeStage} busy={saving} focused>
      <div className="settings-body"><div className="settings-main"><SettingsFooterContext.Provider value={{ target: footer, tab }}>
        <div className="settings-content">
          {stage === "identity" && identity}
          {stage === "models" && <AISetup onProgress={progress => setStageBusy(progress.busy || progress.saving === true)} onSaved={closeStage} />}
          {stage === "research" && <ResearchTab connectionOnly showIntro={false} showCourtListener={false} actionLabel="Save key" onBusy={setStageBusy} exaDraft={exaDraft} onExaDraftChange={setExaDraft} onComplete={closeStage} />}
          {stage === "word" && <WordAddin onBusy={setStageBusy} />}
          {stage === "style" && <><DraftingTab onBusy={setStageBusy} onSaved={closeStage} /><div className="dash-setup-style-links"><Link to="/templates" onClick={closeStage}>Manage firm templates{icons.arrowRight}</Link></div></>}
        </div>
        <div className="settings-footer" ref={setFooter} />
      </SettingsFooterContext.Provider></div></div>
    </SettingsDialog>}
  </div>
}
