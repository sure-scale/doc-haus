import { useEffect, useState, type ReactNode } from "react"
import { OFFICIAL_SOURCE_HOSTS, OFFICIAL_SOURCE_SUFFIXES } from "../../../../../dochaus/lib/research-sources"
import { getPreferences, savePreferences, testResearchConnection, type DraftingPreferences } from "../../api/ingest/preferences"
import { FieldRow, PasswordInput, SegRow } from "./controls"
import "./ResearchTab.css"

// Where the assistants may read on the live web, and the key that powers it.
// Saved through the same ingest preferences as Drafting (it's one standing-
// instructions file), but split into its own tab because web search is a
// distinct, service-backed capability — not a drafting style choice.
type Connection = "exa" | "courtlistener" | "courtlistener-mcp"

export default function ResearchTab({ onSaved, onComplete, heading = "Research", actionLabel = "Save", secondaryActions }: {
  onSaved?: (text: string) => void
  onComplete?: () => void
  heading?: string
  actionLabel?: string
  secondaryActions?: (busy: boolean) => ReactNode
}) {
  const [prefs, setPrefs] = useState<DraftingPreferences | null>(null)
  const [saved, setSaved] = useState<DraftingPreferences | null>(null)
  const [saving, setSaving] = useState(false)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState("")
  const [testing, setTesting] = useState<Connection | null>(null)
  const [tests, setTests] = useState<Partial<Record<Connection, { status: string; message: string }>>>({})
  const load = async () => {
    setLoading(true)
    setError("")
    try {
      const loaded = await getPreferences()
      setPrefs(loaded)
      setSaved(loaded)
    } catch {
      setError("Could not load research settings. Retry.")
    } finally {
      setLoading(false)
    }
  }
  useEffect(() => {
    void load()
  }, [])
  if (loading) return <><p className="muted" role="status">Loading research preferences...</p>{secondaryActions && <div className="onboarding-actions">{secondaryActions(false)}</div>}</>
  if (!prefs) return <section><p role="alert">{error}</p><button onClick={() => void load()}>Retry</button>{secondaryActions && <div className="onboarding-actions">{secondaryActions(false)}</div>}</section>
  const set = (patch: Partial<DraftingPreferences>) => {
    setPrefs({ ...prefs, ...patch })
    setTests({})
    setError("")
  }
  const connectionTest = (service: Connection, label: string, saveLabel: string) => {
    const patch = connectionPreferences(prefs, service)
    const changed = !saved || JSON.stringify(patch) !== JSON.stringify(connectionPreferences(saved, service))
    return <div>
      <button type="button" disabled={saving || testing !== null} onClick={async () => {
        setTesting(service)
        setError("")
        try {
          if (changed) {
            await savePreferences(patch)
            const updated = await getPreferences()
            if (JSON.stringify(connectionPreferences(updated, service)) !== JSON.stringify(patch)) throw new Error("Service settings were not confirmed")
            setSaved(updated)
          }
          const result = await testResearchConnection(service)
          setTests(previous => ({ ...previous, [service]: result }))
        } catch {
          setTests(previous => ({ ...previous, [service]: { status: "failed", message: "Could not save or test this connection. Try again." } }))
        } finally {
          setTesting(null)
        }
      }}>{testing === service ? "Testing..." : changed ? saveLabel : label}</button>
      {tests[service] && <p role="status" className={`settings-hint${tests[service]?.status === "verified" ? " muted" : " settings-warning"}`}>{tests[service]?.message}</p>}
    </div>
  }
  return (
    <section className="settings-section">
      <h3>{heading}</h3>
      <p className="settings-hint muted">
        Connect services for live legal research. Keys stay in this installation.
      </p>
      <SegRow
        label="Web research"
        hint="Read across the web or limit research to approved sources."
        value={prefs.webResearch}
        disabled={saving || testing !== null}
        onChange={(webResearch) => set({ webResearch: webResearch as DraftingPreferences["webResearch"] })}
        options={[
          { value: "open", label: "Entire web" },
          { value: "official", label: "Approved sources only" },
        ]}
      />
      {prefs.webResearch === "official" && <>
        <div className="settings-form-row">
          <div className="settings-field-copy"><span className="settings-label">Built-in approved sources</span></div>
          <div className="settings-field-control">
            <details className="research-source-list">
              <summary>View {OFFICIAL_SOURCE_HOSTS.length} hosts and {OFFICIAL_SOURCE_SUFFIXES.length} domain endings</summary>
              <p className="settings-hint muted">Listed hosts include subdomains. Domain endings match hostnames ending in the full suffix shown, including the dot.</p>
              <div className="research-source-columns">
                <div>
                  <h4>Hosts and subdomains</h4>
                  <ul>{OFFICIAL_SOURCE_HOSTS.map(host => <li key={host}><code>{host}</code></li>)}</ul>
                </div>
                <div>
                  <h4>Domain endings</h4>
                  <ul>{OFFICIAL_SOURCE_SUFFIXES.map(suffix => <li key={suffix}><code>{suffix}</code></li>)}</ul>
                </div>
              </div>
            </details>
          </div>
        </div>
        <FieldRow
          id="research-approved-sources"
          label="Additional approved sources"
          hint="One hostname per line; no URLs, ports, paths or wildcards. Subdomains are included."
        >
          <textarea
            id="research-approved-sources"
            aria-describedby="research-approved-sources-hint"
            className="settings-textarea"
            rows={3}
            placeholder="trusted-source.org"
            value={prefs.approvedSources.join("\n")}
            disabled={saving || testing !== null}
            onChange={(e) => set({ approvedSources: e.target.value.split("\n") })}
          />
        </FieldRow>
      </>}
      <FieldRow
        id="research-exa-key"
        label="Exa API key"
        labelAction={<a className="settings-hint" href="https://dashboard.exa.ai/" target="_blank" rel="noopener noreferrer">Sign up for Exa ↗</a>}
        hint="For citation search and source text. Testing uses a small amount of your Exa search allowance."
      >
        <PasswordInput
          id="research-exa-key"
          aria-describedby="research-exa-key-hint"
          placeholder="Exa API key (exa.ai)"
          value={prefs.searchApiKey}
          disabled={saving || testing !== null}
          onChange={(searchApiKey) => set({ searchApiKey })}
        />
        {!!prefs.searchApiKey.trim() && connectionTest("exa", "Test Exa key", "Save key and test")}
        {!prefs.searchApiKey.trim() && (
          <p className="settings-hint settings-warning">
            Without a key, search uses limited keyless access. Exa source text requires a key.
          </p>
        )}
      </FieldRow>
      <SegRow
        label="CourtListener connector"
        hint="Connect U.S. case-law research through CourtListener."
        value={prefs.courtlistenerMcpEnabled ? "on" : "off"}
        disabled={saving || testing !== null}
        onChange={(value) => set({ courtlistenerMcpEnabled: value === "on" })}
        options={[
          { value: "off", label: "Off" },
          { value: "on", label: "On" },
        ]}
      />
      {prefs.courtlistenerMcpEnabled && <FieldRow
        id="research-courtlistener-token"
        label="CourtListener API token"
        labelAction={<a className="settings-hint" href="https://www.courtlistener.com/register/" target="_blank" rel="noopener noreferrer">Sign up for CourtListener ↗</a>}
        hint="Required for the connector. Testing checks API access."
      >
        <PasswordInput
          id="research-courtlistener-token"
          aria-describedby="research-courtlistener-token-hint"
          placeholder="CourtListener API token (courtlistener.com/profile/api/)"
          value={prefs.caseLawApiKey}
          disabled={saving || testing !== null}
          onChange={(caseLawApiKey) => set({ caseLawApiKey })}
        />
        {prefs.caseLawApiKey.trim() ? <>
          {connectionTest("courtlistener", "Test token", "Save token and test")}
          {connectionTest("courtlistener-mcp", "Test connector", "Save connector and test")}
        </> : <p className="settings-hint muted">Add a token to connect.</p>}
      </FieldRow>}
      {error && <p className="settings-hint settings-warning" role="alert">{error}</p>}
      <div className="settings-actions">
        {secondaryActions?.(saving || testing !== null)}
        <button
          className={`primary${saving ? " btn-loading" : ""}`}
          disabled={saving || testing !== null}
          onClick={async () => {
            setSaving(true)
            setError("")
            try {
              const updated = await savePreferences(researchPreferences(prefs))
              setPrefs(updated)
              setSaved(updated)
              setTests({})
              onSaved?.("Research preferences saved.")
              onComplete?.()
            } catch {
              setError("Could not save research settings. Check additional source hostnames, then retry.")
            } finally {
              setSaving(false)
            }
          }}
        >
          {actionLabel}
        </button>
      </div>
    </section>
  )
}

function connectionPreferences(prefs: DraftingPreferences, service: Connection) {
  if (service === "exa") return { searchApiKey: prefs.searchApiKey.trim() }
  if (service === "courtlistener") return { caseLawApiKey: prefs.caseLawApiKey.trim() }
  return { caseLawApiKey: prefs.caseLawApiKey.trim(), courtlistenerMcpEnabled: prefs.courtlistenerMcpEnabled }
}

function researchPreferences(prefs: DraftingPreferences) {
  return {
    webResearch: prefs.webResearch,
    approvedSources: prefs.approvedSources.map(host => host.trim()).filter(Boolean),
    searchApiKey: prefs.searchApiKey,
    caseLawApiKey: prefs.caseLawApiKey,
    courtlistenerMcpEnabled: prefs.courtlistenerMcpEnabled,
  }
}
