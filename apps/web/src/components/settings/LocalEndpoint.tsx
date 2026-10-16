import { useState } from "react"
import { FieldRow } from "./controls"

export default function LocalEndpoint({
  onAdd,
}: {
  onAdd: (input: { id: string; name: string; baseURL: string; modelID: string }) => Promise<void>
}) {
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState("")
  const [name, setName] = useState("")
  const [baseURL, setBaseURL] = useState("http://localhost:1234/v1")
  const [modelID, setModelID] = useState("")
  const id = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-")
  const ready = name.trim() && baseURL.trim() && modelID.trim()
  return (
    <section className="settings-section">
      <h3>Local endpoint</h3>
      {error && <p role="alert" className="settings-warning">{error}</p>}
      <p className="settings-hint muted">
        OpenAI-compatible servers — LM Studio, Ollama, vLLM. The URL must be reachable from the machine running the
        engine, not the browser.
      </p>
      <FieldRow id="local-provider-name" label="Name" hint="The name shown for this provider in the model picker.">
        <input id="local-provider-name" aria-describedby="local-provider-name-hint" placeholder="LM Studio" value={name} onChange={(e) => setName(e.target.value)} />
      </FieldRow>
      <FieldRow id="local-provider-url" label="Base URL" hint="The OpenAI-compatible API address, reachable from the engine.">
        <input id="local-provider-url" aria-describedby="local-provider-url-hint" placeholder="http://localhost:1234/v1" value={baseURL} onChange={(e) => setBaseURL(e.target.value)} />
      </FieldRow>
      <FieldRow id="local-provider-model" label="Model ID" hint="Use the exact model identifier served by the endpoint.">
        <input id="local-provider-model" aria-describedby="local-provider-model-hint" placeholder="llama-3.1-8b" value={modelID} onChange={(e) => setModelID(e.target.value)} />
      </FieldRow>
      <div className="settings-actions">
        <button
          className="primary"
          disabled={busy || !ready}
          onClick={async () => {
            if (busy) return
            setBusy(true); setError("")
            try { await onAdd({ id, name: name.trim(), baseURL: baseURL.trim(), modelID: modelID.trim() }); setName(""); setModelID("") }
            catch { setError("Could not add this endpoint. Check the URL and model ID, then retry.") }
            finally { setBusy(false) }
          }}
        >
          Add provider
        </button>
      </div>
    </section>
  )
}
