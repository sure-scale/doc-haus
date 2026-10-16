import { useEffect, useState } from "react"
import { getProviderOptions } from "../../api/opencode/config"
import { listAwsProfiles, listGcpProjects } from "../../api/ingest/host"
import type { CloudSpec } from "./cloud-setups"
import type { Provider } from "./types"
import { FieldRow } from "./controls"

// One card per cloud setup: the sign-in it expects on the server, labeled
// project/region fields shared by every provider the card covers, and a single
// Enable that saves the fields and runs the live probe in one go. The old split
// (Save here, Enable elsewhere) let you probe values you had typed but never
// saved. Only the primary provider's result shows — the card verifies or fails
// on it alone, and a secondary like Claude-on-Vertex rides along silently.
export default function NeedsSetupCard({
  spec,
  providers,
  probing,
  error,
  onEnable,
  onConfigure,
  discoverSuggestions = true,
}: {
  spec: CloudSpec
  providers: Provider[]
  probing: boolean
  error: string
  onEnable: (options: Record<string, string>, modelID: string) => void
  onConfigure: (options: Record<string, string>) => void
  discoverSuggestions?: boolean
}) {
  const [testModel, setTestModel] = useState("")
  const [values, setValues] = useState<Record<string, string>>({})
  const [loadError, setLoadError] = useState("")
  // What the host's sign-in can see for suggest fields (GCP projects, AWS
  // profiles), keyed by field. Empty means nothing to suggest — type instead.
  const [suggestions, setSuggestions] = useState<Record<string, string[]>>({})
  const providerID = spec.ids[0]
  const host = providers[0]?.host
  const models = Object.values(providers[0]?.models ?? {})
  // Prefill from whatever the engine already has saved (falling back to the
  // field default) so the fields show the live config rather than resetting to
  // blank each time Settings opens. The card's providers share their settings,
  // so the first id is the source.
  useEffect(() => {
    getProviderOptions(providerID).then((opts) => {
      setValues(Object.fromEntries(spec.fields.map((f) => [f.key, opts[f.key] ?? host?.options[f.key] ?? f.fallback ?? ""])))
      setLoadError("")
    }).catch(() => {
      setValues(Object.fromEntries(spec.fields.map(f => [f.key, host?.options[f.key] ?? f.fallback ?? ""])))
      setLoadError("Could not load saved connection settings. Detected defaults are shown; check them before saving.")
    })
    spec.fields
      .filter((f) => discoverSuggestions && f.suggest)
      .forEach((f) =>
        (f.suggest === "gcp-projects" ? listGcpProjects() : listAwsProfiles()).then((items) =>
          setSuggestions((s) => ({ ...s, [f.key]: items })),
        ),
      )
  }, [providerID])
  const ready = spec.fields.every((f) => !f.required || values[f.key]?.trim() || f.fallback)
  const options = () => Object.fromEntries(spec.fields.map(f => [f.key, (values[f.key]?.trim() || f.fallback) ?? ""]))
  return (
    <div className="settings-needs">
      <div className="settings-provider-copy">
        <span className="settings-conn-name">{spec.name}</span>
        <span className="settings-conn-meta muted">Server sign-in: {spec.signin}</span>
      </div>
      <p className="settings-hint muted">{host?.sources.length ? `Detected on the server: ${host.sources.join(", ")}. Access has not been tested.` : "No host sign-in detected yet. Existing server credentials can be used after the connection is configured."}{host?.type === "google-adc" && !host.options.project ? " Choose a GCP project; an ADC sign-in does not select one." : ""}</p>
      {loadError && <p role="alert" className="settings-warning">{loadError}</p>}
      {spec.fields.map((f) => {
        const set = (value: string) => setValues((v) => ({ ...v, [f.key]: value }))
        const fieldID = `cloud-${spec.ids[0]}-${f.key}`
        const listID = `${fieldID}-suggestions`
        return (
          <FieldRow
            key={f.key}
            id={fieldID}
            label={f.label}
            hint={f.fallback ? `Defaults to ${f.fallback}.` : f.required ? "Required to connect this provider." : "Optional. Uses the server's default credentials when blank."}
          >
            {f.options ? (
              <select id={fieldID} aria-describedby={`${fieldID}-hint`} value={values[f.key] ?? f.fallback ?? ""} onChange={(e) => set(e.target.value)}>
                {f.options.map((o) => (
                  <option key={o} value={o}>
                    {o}
                  </option>
                ))}
              </select>
            ) : (
              <input
                id={fieldID}
                aria-describedby={`${fieldID}-hint`}
                placeholder={f.placeholder}
                value={values[f.key] ?? ""}
                list={suggestions[f.key]?.length ? listID : undefined}
                onChange={(e) => set(e.target.value)}
              />
            )}
            {(suggestions[f.key]?.length ?? 0) > 0 && (
              <datalist id={listID}>
                {suggestions[f.key]?.map((o) => (
                  <option key={o} value={o} />
                ))}
              </datalist>
            )}
          </FieldRow>
        )
      })}
      {models.length > 0 ? <>
        <label>Model to test<select value={testModel} onChange={event => setTestModel(event.target.value)} disabled={probing}><option value="">Choose a model</option>{models.map(model => <option key={model.id} value={model.id}>{model.name}</option>)}</select></label>
        <p className="settings-hint muted">This makes one short model request and may incur usage charges. Opening settings makes no model requests.</p>
      </> : <p className="settings-hint muted">Save the connection settings to load model choices. Saving makes no model requests.</p>}
      <div className="settings-actions">
        <button
          className={`primary${probing ? " btn-loading" : ""}`}
          disabled={probing || !ready || (models.length > 0 && !testModel)}
          onClick={() => models.length ? onEnable(options(), testModel) : onConfigure(options())}
        >
          {models.length ? "Test and enable" : "Save connection settings"}
        </button>
      </div>
      {error && <span className="settings-error">{error}</span>}
    </div>
  )
}
