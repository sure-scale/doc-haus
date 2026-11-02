import { useEffect, useRef, useState } from "react"
import { getPreferences, savePreferences, type BackgroundAgents } from "../../api/ingest/preferences"
import { fetchWordChatPreferences, saveWordChatPreferences } from "../../api/ingest/word"
import { ToggleRow } from "./controls"

// Optional jobs that start without a click. Each switch saves at once. Firm preferences merge
// top-level keys only, so the whole backgroundAgents object is sent; Word suggestions keep their
// own saved setting, which the Word pane reads.
type Switches = BackgroundAgents & { wordSuggestions: boolean }
const AGENTS: { key: keyof Switches; title: string; when: string; description: string }[] = [
  { key: "chronologySuggestion", title: "Suggest a chronology", when: "When you open a matter, add documents or close a chat.", description: "Checks whether a timeline of events would help and suggests one. If off, open it yourself from the matter menu." },
  { key: "reviewCompare", title: "Compare review answers", when: "When you fill a review table.", description: "Shows which documents give the same answer to each question and which differ. If off, use Compare with AI on a column." },
  { key: "narration", title: "Narrate drafting progress", when: "While a draft is being prepared.", description: "Adds a short plain-English update as each step finishes, and every few minutes during long steps. If off, you still see each step." },
  { key: "wordSuggestions", title: "Suggest ideas in Word", when: "While the doc.haus pane is open in Word.", description: "Offers up to three ideas about your document or the passage you've selected. If off, ask in the Word chat." },
]

const STATES = { on: "On", off: "Off" }

// Missing keys mean on, matching the server default.
function normalize(saved?: Partial<BackgroundAgents>): BackgroundAgents {
  return {
    chronologySuggestion: saved?.chronologySuggestion !== false,
    reviewCompare: saved?.reviewCompare !== false,
    narration: saved?.narration !== false,
  }
}

export default function BackgroundAgentsTab({ onSaved }: { onSaved?: (text: string) => void }) {
  const [agents, setAgents] = useState<Switches | null>(null)
  const [loadError, setLoadError] = useState("")
  const [error, setError] = useState("")
  const [saving, setSaving] = useState(false)
  const [reload, setReload] = useState(0)
  const mounted = useRef(true)

  useEffect(() => {
    mounted.current = true
    let active = true
    setLoadError("")
    Promise.all([getPreferences(), fetchWordChatPreferences()]).then(([prefs, word]) => {
      if (active) setAgents({ ...normalize(prefs.backgroundAgents), wordSuggestions: word.suggestions })
    }).catch(() => {
      if (active) setLoadError("Could not load background agent settings.")
    })
    return () => {
      active = false
      mounted.current = false
    }
  }, [reload])

  if (!agents) return (
    <section className="settings-section">
      {loadError ? (
        <>
          <p className="settings-error" role="alert">{loadError}</p>
          <button type="button" onClick={() => setReload((value) => value + 1)}>Try again</button>
        </>
      ) : <p className="muted" role="status">Loading background agents…</p>}
    </section>
  )

  const toggle = async (agent: (typeof AGENTS)[number], on: boolean) => {
    const previous = agents
    const next = { ...agents, [agent.key]: on }
    setAgents(next)
    setError("")
    setSaving(true)
    try {
      if (agent.key === "wordSuggestions") {
        const saved = await saveWordChatPreferences({ suggestions: on })
        if (!mounted.current) return
        setAgents({ ...next, wordSuggestions: saved.suggestions })
      } else {
        const { wordSuggestions, ...backgroundAgents } = next
        const saved = await savePreferences({ backgroundAgents })
        if (!mounted.current) return
        setAgents({ ...normalize(saved.backgroundAgents ?? backgroundAgents), wordSuggestions })
      }
      onSaved?.(`${agent.title} turned ${on ? "on" : "off"}.`)
    } catch {
      if (!mounted.current) return
      setAgents(previous)
      setError("Could not save this setting. Try again.")
    } finally {
      if (mounted.current) setSaving(false)
    }
  }

  return (
    <section className="settings-section">
      <p className="settings-hint muted">
        These run by themselves while you work, without you asking. Each one sends some of the matter's text to your AI provider. Turn one off to stop it.
      </p>
      {AGENTS.map((agent) => (
        <ToggleRow
          key={agent.key}
          title={agent.title}
          description={<><strong>{agent.when}</strong> {agent.description}</>}
          on={agents[agent.key]}
          states={STATES}
          disabled={saving}
          onChange={(on) => void toggle(agent, on)}
        />
      ))}
      {error && <p className="settings-error" role="alert">{error}</p>}
    </section>
  )
}
