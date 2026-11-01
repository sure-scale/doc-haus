import { useEffect, useRef, useState } from "react"
import { getPreferences, savePreferences, type BackgroundAgents } from "../../api/ingest/preferences"
import { ToggleRow } from "./controls"

// Optional jobs that start without a click. Each switch saves at once; the
// server merges top-level keys only, so the whole object is always sent.
const AGENTS: { key: keyof BackgroundAgents; title: string; description: string }[] = [
  { key: "chronologySuggestion", title: "Suggest a chronology", description: "Suggests a chronology when a matter's documents would benefit from one." },
  { key: "reviewCompare", title: "Compare review answers", description: "Groups matching answers when a review column finishes. You can still compare a column yourself." },
  { key: "narration", title: "Narrate drafting progress", description: "Writes short notes on how a draft is going. Progress steps still show." },
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
  const [agents, setAgents] = useState<BackgroundAgents | null>(null)
  const [loadError, setLoadError] = useState("")
  const [error, setError] = useState("")
  const [saving, setSaving] = useState(false)
  const [reload, setReload] = useState(0)
  const mounted = useRef(true)

  useEffect(() => {
    mounted.current = true
    let active = true
    setLoadError("")
    getPreferences().then((prefs) => {
      if (active) setAgents(normalize(prefs.backgroundAgents))
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
      const saved = await savePreferences({ backgroundAgents: next })
      if (!mounted.current) return
      setAgents(normalize(saved.backgroundAgents ?? next))
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
        These agents start on their own, without a click. Each sends matter text to your AI provider.
      </p>
      {AGENTS.map((agent) => (
        <ToggleRow
          key={agent.key}
          title={agent.title}
          description={agent.description}
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
