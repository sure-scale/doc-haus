import { useEffect, useState, type FormEvent, type ReactNode } from "react"
import { Link, useNavigate } from "react-router-dom"
import { addSampleMatter, type Matter } from "../../api/ingest/matters"
import { savePreferences } from "../../api/ingest/preferences"
import type { SetupProgress } from "../../api/ingest/setup"
import { loadPrefs, savePrefs } from "../../prefs"
import JurisdictionFlag from "../JurisdictionFlag"
import JurisdictionSelect from "../JurisdictionSelect"
import { SUGGESTED_PRACTICE_AREAS } from "../settings/MatterDefaultsTab"
import { useOpenSettings } from "../settings/navigation"
import type { DashboardData } from "./data"
import { icons } from "./icons"
import { SetupChecklist, type SetupChecklistConfig } from "./visual/SetupCards"
import type { WidgetContext } from "./widgets"

function completion(data: DashboardData, setup: SetupProgress | null, jurisdictions: string[]) {
  const prefs = data.preferences
  const hasChoice = (choices: string[]) => choices.some((choice) => typeof choice === "string" && !!choice.trim())
  return [
    setup?.ready === true,
    !!prefs?.attorney.trim() && !!prefs.firm.trim() && hasChoice(jurisdictions) && hasChoice(prefs.practiceAreas),
    !!prefs?.searchApiKey.trim(),
    data.word?.status === "ready" || (data.word?.status === "external" && data.word.available),
    !!prefs?.houseStyle.trim() && (data.templates ?? 0) > 0,
  ]
}

// An unavailable service stays incomplete; opening Settings never marks it done.
export function setupTabProgress(data: DashboardData, setup: SetupProgress | null, jurisdictions = loadPrefs().defaultJurisdictions): { done: number; total: 5 } {
  return { done: completion(data, setup, jurisdictions).filter(Boolean).length, total: 5 }
}


export function SetupTab({ context, onSampleAdded, onHideSetup }: { context: WidgetContext; onSampleAdded?: () => void; onHideSetup?: () => void }) {
  const { data, setup, onPreferences, onRetry } = context
  const openSettings = useOpenSettings()
  const navigate = useNavigate()
  const prefs = data.preferences
  const [attorney, setAttorney] = useState(prefs?.attorney ?? "")
  const [firm, setFirm] = useState(prefs?.firm ?? "")
  const [jurisdictions, setJurisdictions] = useState(() => loadPrefs().defaultJurisdictions)
  const [areas, setAreas] = useState(prefs?.practiceAreas ?? [])
  const [identityBusy, setIdentityBusy] = useState(false)
  const [identityError, setIdentityError] = useState("")
  const [identitySaved, setIdentitySaved] = useState("")
  const [key, setKey] = useState("")
  const [keyBusy, setKeyBusy] = useState(false)
  const [keyError, setKeyError] = useState("")
  const [keySaved, setKeySaved] = useState("")
  // Both endpoints return full preferences; serialize writes to avoid an older
  // response replacing a newer save in the live checklist.
  const saving = identityBusy || keyBusy
  const [adding, setAdding] = useState(false)
  const [demoError, setDemoError] = useState("")
  const [added, setAdded] = useState<Matter>()
  const [expanded, setExpanded] = useState(() => ["models", "identity", "research", "word", "style"][completion(data, setup, loadPrefs().defaultJurisdictions).findIndex(value => !value)] ?? "")
  function finishStep(id: string) {
    setExpanded(current => current === id ? "" : current)
    requestAnimationFrame(() => { if (!document.querySelector('.dash-setup-toggle[aria-expanded="true"]')) document.getElementById(`dash-setup-toggle-${id}`)?.focus() })
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
  // Show saved choices and a small useful selection; the searchable picker has all packs.
  const shownPacks = [...new Set([...jurisdictions, ...packs.slice(0, 3).map((pack) => pack.code)])].slice(0, Math.max(jurisdictions.length, 3))
  const shownAreas = [...new Set([...areas, "Commercial", "Corporate", "Employment"])].slice(0, Math.max(areas.length, 3))
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
        finishStep("identity")
        onRetry()
      } catch {
        // The server save succeeded even when this browser could not save law defaults.
        onPreferences?.(saved)
        setIdentityError("Your name, firm and practice areas were saved. Jurisdiction defaults could not be saved in this browser. Try again.")
      }
    } catch {
      setIdentityError("Could not confirm that your practice details were saved. Your changes are still here. Try again.")
    } finally { setIdentityBusy(false) }
  }
  async function saveKey(event: FormEvent) {
    event.preventDefault()
    if (!key.trim() || saving) return
    setKeyBusy(true); setKeyError(""); setKeySaved("")
    try {
      const saved = await savePreferences({ searchApiKey: key.trim() })
      if (!saved.searchApiKey.trim()) throw new Error("Key was not saved")
      onPreferences?.(saved)
      setKey("")
      setKeySaved("Exa key saved. You can test the connection in Research settings.")
      finishStep("research")
    } catch {
      // Never echo a service response which could include the submitted key.
      setKeyError("Could not confirm that the Exa key was saved. Try again.")
    } finally { setKeyBusy(false) }
  }
  async function populate() {
    if (adding || sample) return
    setAdding(true); setDemoError("")
    try { setAdded(await addSampleMatter()); onSampleAdded?.(); onRetry() }
    catch { setDemoError("Could not add the sample matter. Try again.") }
    finally { setAdding(false) }
  }
  const toggle = (values: string[], value: string) => values.includes(value) ? values.filter((other) => other !== value) : [...values, value]
  const config: SetupChecklistConfig = {
    title: "Finish setting up",
    description: "Connect your tools and defaults so every draft and search is ready.",
    items: [
      { id: "models", icon: "model", title: "AI provider & models", detail: done[0] ? "Your AI tools are connected and ready to use." : setup ? "Connect your AI tools so you can start drafting and searching." : "Open AI settings to check your connection.", done: !!done[0], action: { id: "ai", label: done[0] ? "Change" : "Connect" } },
      { id: "identity", icon: "person", title: "Practice identity & default law", detail: "Add your name, firm and the law you usually work with.", done: !!done[1], editable: true },
      { id: "research", icon: "search", title: "Web research with Exa", detail: done[2] ? "Web research is connected. Test or change it in Research settings." : "Add your Exa API key to search the web for legal and company information.", done: !!done[2], editable: !done[2], ...(done[2] ? { action: { id: "research", label: "Change" } } : {}) },
      { id: "word", icon: "word", title: "Microsoft Word add-in", detail: done[3] ? "Word is connected and ready for document review." : data.word?.status === "word_missing" ? "Install Word on this computer, then connect the add-in" : data.word?.status === "unsupported" ? "See add-in settings for supported ways to connect Word" : "Review documents and apply tracked changes directly in Microsoft Word.", done: !!done[3], action: { id: "addins", label: done[3] ? "Manage Word" : "Set up Word", variant: "outline" } },
      { id: "style", icon: "template", title: "House style & firm templates", detail: "Choose how your drafts look and upload templates you use regularly.", done: !!done[4], action: { id: "drafting", label: "Configure", variant: "outline" } },
    ],
  }
  function content(id: string): ReactNode {
    if (id === "identity") return prefs ? <form className="dash-setup-inline" onSubmit={saveIdentity} aria-labelledby="dash-setup-identity">
      <div className="dash-setup-input-row"><label><span className="sr-only">Your full name</span><input value={attorney} onChange={(event) => { setAttorney(event.target.value); setIdentitySaved("") }} placeholder="Your full name" autoComplete="name" disabled={saving} maxLength={200} required /></label><label><span className="sr-only">Firm name</span><input value={firm} onChange={(event) => { setFirm(event.target.value); setIdentitySaved("") }} placeholder="Firm name" autoComplete="organization" disabled={saving} maxLength={200} required /></label><button type="submit" disabled={saving || !changed} aria-busy={identityBusy}>{identityBusy ? "Saving…" : "Save"}</button></div>
      <div className="dash-setup-chips" role="group" aria-label="Default jurisdictions">{shownPacks.map((code) => <button type="button" key={code} aria-pressed={jurisdictions.includes(code)} disabled={saving} onClick={() => { setJurisdictions(toggle(jurisdictions, code)); setIdentitySaved("") }}><JurisdictionFlag code={code} region />{packs.find((pack) => pack.code === code)?.name ?? code}{jurisdictions.includes(code) && icons.check}</button>)}<JurisdictionSelect jurisdictions={packs} selected={jurisdictions} onChange={setJurisdictions} disabled={saving || !packs.length} trigger={<span>All jurisdictions…</span>} triggerTitle="Choose default jurisdictions" /></div>
      <div className="dash-setup-chips" role="group" aria-label="Practice areas">{shownAreas.map((area) => <button type="button" key={area} aria-pressed={areas.includes(area)} disabled={saving} onClick={() => { setAreas(toggle(areas, area)); setIdentitySaved("") }}>{area}{areas.includes(area) && icons.check}</button>)}<JurisdictionSelect jurisdictions={areaOptions} selected={areas} onChange={setAreas} practiceAreas disabled={saving} trigger={<span>More…</span>} triggerTitle="Choose practice areas" /></div>
      {identityError && <p className="dash-setup-error" role="alert">{identityError}</p>}{identitySaved && <p className="dash-setup-saved" role="status">{identitySaved}</p>}
    </form> : <p className="dash-setup-unavailable">Could not load your practice details. <button type="button" className="dash-text-button" onClick={onRetry}>Retry</button></p>
    if (id === "research") return <>
      {!done[2] && <form className="dash-setup-exa" onSubmit={saveKey} aria-labelledby="dash-setup-research"><label><span className="sr-only">Exa API key</span><input type="password" value={key} onChange={(event) => { setKey(event.target.value); setKeyError("") }} placeholder="Paste Exa API key" autoComplete="off" spellCheck={false} disabled={saving || !prefs} maxLength={4096} required /></label><button type="submit" disabled={saving || !key.trim() || !prefs} aria-busy={keyBusy}>{keyBusy ? "Saving…" : "Save key"}</button></form>}
      {keyError && <p className="dash-setup-error" role="alert">{keyError}</p>}{keySaved && <p className="dash-setup-saved" role="status">{keySaved}</p>}
    </>
    if (id === "models") return <p className="dash-setup-panel-note">{done[0] ? "Your primary and fast models are connected. Open AI settings to review the connection or change models." : "Choose your provider, connect your account and test your primary and fast models in AI settings."}</p>
    if (id === "word") return <p className="dash-setup-panel-note">Connect Microsoft Word to review documents and apply tracked changes. Add-in settings will guide you through installation.</p>
    if (id === "style") return <div className="dash-setup-style-links"><span>{prefs?.houseStyle.trim() ? "House style saved" : "House style to add"}</span><Link to="/templates">{data.templates === undefined ? "View templates" : `${data.templates} ${data.templates === 1 ? "template" : "templates"}`}{icons.arrowRight}</Link></div>
    return null
  }
  return <div className="dash-setup-flow">
    <div className="dash-setup-demo-row" role="group" aria-label="Demo data">
      <span className="dash-setup-demo-row-icon" aria-hidden="true">{icons.sample}</span>
      <div className="dash-setup-demo-copy"><p>Want to see doc.haus in action first?</p><span>{sample ? "Explore your sample matter, documents and data." : "Populate your workspace with a sample matter, documents and data."}</span></div>
      <button type="button" className="dash-setup-demo-primary" disabled={adding} aria-busy={adding} onClick={() => { if (sample) navigate(`/matter/${sample.id}?view=chat&agent=ask`); else void populate() }}>{sample ? icons.arrowRight : icons.plus}{adding ? "Adding demo data…" : sample ? "Open sample matter" : "Add sample data"}</button>
      {onHideSetup && <button type="button" className="dash-text-button dash-setup-hide" onClick={onHideSetup}>Hide setup</button>}
      {demoError && <p className="dash-setup-error" role="alert">{demoError}</p>}
    </div>
    <p className="sr-only" role="status">{[identitySaved, keySaved].filter(Boolean).join(" ")}</p>
    <SetupChecklist accordion={{ expanded, onToggle: setExpanded }} config={config} onAction={(action) => { if (action === "ai" || action === "research" || action === "addins" || action === "drafting") openSettings(action) }} renderContent={content} />
  </div>
}
