import { addLocalProvider } from "../../api/opencode/config"
import { setProviderKey } from "../../api/opencode/providers"
import type { Client } from "../../api/opencode/client"
import type { ModelOption } from "../../models"
import type { Provider } from "./types"

export default function ConnectProviderStep({
  pick,
  setPick,
  keyProviders,
  apiKey: key,
  setKey,
  all,
  client,
  setForceConnect,
  setNotice,
  load,
  localName,
  setLocalName,
  localURL,
  setLocalURL,
  localModel,
  setLocalModel,
  forceConnect,
  models,
}: {
  pick: string
  setPick: (value: string) => void
  keyProviders: Provider[]
  apiKey: string
  setKey: (value: string) => void
  all: Provider[]
  client: Client
  setForceConnect: (value: boolean) => void
  setNotice: (value: string) => void
  load: () => Promise<void>
  localName: string
  setLocalName: (value: string) => void
  localURL: string
  setLocalURL: (value: string) => void
  localModel: string
  setLocalModel: (value: string) => void
  forceConnect: boolean
  models: ModelOption[]
}) {
  return (
    <>
      <p className="settings-notice">
        No model provider detected. Connect a cloud provider with an API key, or point doc.haus at a local
        endpoint like Ollama.
      </p>

      <section className="settings-section">
        <h3>Cloud provider</h3>
        <p className="settings-hint muted">
          Google AI Studio and other API-key providers. The key is stored on the engine, never in the browser.
        </p>
        <div className="row settings-row">
          <select value={pick} onChange={(e) => setPick(e.target.value)}>
            <option value="">Choose a provider</option>
            {keyProviders.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </select>
          <input
            type="password"
            placeholder="API key"
            value={key}
            onChange={(e) => setKey(e.target.value)}
            disabled={!pick}
          />
          <button
            className="primary"
            disabled={!pick || !key.trim()}
            onClick={async () => {
              const name = all.find((p) => p.id === pick)?.name ?? pick
              await setProviderKey(client, pick, key.trim())
              setPick("")
              setKey("")
              setForceConnect(false)
              setNotice(`Connected ${name}. Pick your models below.`)
              await load()
            }}
          >
            Connect
          </button>
        </div>
      </section>

      <section className="settings-section">
        <h3>Local endpoint (Ollama)</h3>
        <p className="settings-hint muted">
          Any OpenAI-compatible server — Ollama, LM Studio, vLLM. The URL must be reachable from the machine
          running the engine, not the browser.
        </p>
        <div className="settings-grid">
          <label className="settings-label">Name</label>
          <input value={localName} onChange={(e) => setLocalName(e.target.value)} />
          <label className="settings-label">Base URL</label>
          <input value={localURL} onChange={(e) => setLocalURL(e.target.value)} />
          <label className="settings-label">Model id</label>
          <input placeholder="llama3.1" value={localModel} onChange={(e) => setLocalModel(e.target.value)} />
          <button
            className="primary settings-add"
            disabled={!localName.trim() || !localURL.trim() || !localModel.trim()}
            onClick={async () => {
              const id = localName.trim().toLowerCase().replace(/[^a-z0-9]+/g, "-")
              await addLocalProvider({
                id,
                name: localName.trim(),
                baseURL: localURL.trim(),
                modelID: localModel.trim(),
              })
              setLocalModel("")
              setForceConnect(false)
              setNotice(`Added ${localName.trim()}. Pick your models below — restart the engine if it's missing.`)
              await load()
            }}
          >
            Add endpoint
          </button>
        </div>
      </section>

      {forceConnect && models.length > 0 && (
        <div className="row" style={{ justifyContent: "flex-end", marginTop: 12 }}>
          <button onClick={() => setForceConnect(false)}>Back to model picker</button>
        </div>
      )}
    </>
  )
}
