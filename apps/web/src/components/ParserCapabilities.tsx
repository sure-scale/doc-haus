import { useEffect, useState } from "react"
import { getParserCapabilities, type ByteObservation, type ParserCapabilitiesSnapshot, type ParserComponentObservation } from "../api/parser-capabilities"
import { formatDateTime } from "../lib/format"

export type ParserCapabilitiesState =
  | { status: "loading" }
  | { status: "unavailable" }
  | { status: "observed"; snapshot: ParserCapabilitiesSnapshot }

export default function ParserCapabilities() {
  const [state, setState] = useState<ParserCapabilitiesState>({ status: "loading" })
  useEffect(() => {
    const controller = new AbortController()
    getParserCapabilities({ signal: controller.signal })
      .then(snapshot => { if (!controller.signal.aborted) setState({ status: "observed", snapshot }) })
      .catch(() => { if (!controller.signal.aborted) setState({ status: "unavailable" }) })
    return () => controller.abort()
  }, [])
  return <ParserCapabilitiesView state={state} />
}

export function ParserCapabilitiesView({ state }: { state: ParserCapabilitiesState }) {
  const names = { liteparse: "LiteParse", unpdf: "unpdf", markitdown: "MarkItDown", pdftoppm: "Poppler (pdftoppm)", tesseract: "Tesseract", uvx: "uvx" }
  return (
    <section className="settings-section" aria-labelledby="parser-capabilities-heading">
      <h3 id="parser-capabilities-heading">PDF parser components</h3>
      <p className="settings-hint muted">Observed on the ingestion host. Availability does not establish parsing quality or readiness.</p>
      {state.status === "loading" && <p role="status">Reading ingestion host capabilities…</p>}
      {state.status === "unavailable" && <p role="status">The ingestion host inventory is unavailable. Capabilities are unknown.</p>}
      {state.status === "observed" && <>
        <dl className="settings-summary">
          <div className="settings-summary-row"><dt>Process OS</dt><dd>{state.snapshot.host.platform}</dd></div>
          <div className="settings-summary-row"><dt>Process architecture</dt><dd>{state.snapshot.host.architecture}</dd></div>
          <div className="settings-summary-row"><dt>Physical architecture</dt><dd>Unknown</dd></div>
          <div className="settings-summary-row"><dt>Execution environment</dt><dd>Unknown</dd></div>
          <div className="settings-summary-row"><dt>Runtime</dt><dd>{state.snapshot.host.runtime} {state.snapshot.host.runtimeVersion}</dd></div>
          <div className="settings-summary-row"><dt>Total system RAM</dt><dd>{formatBytes(state.snapshot.memory.total)}</dd></div>
          <div className="settings-summary-row"><dt>OS-reported free RAM</dt><dd>{formatBytes(state.snapshot.memory.free)}</dd></div>
          <div className="settings-summary-row"><dt>Disk available on ingest filesystem</dt><dd>{formatBytes(state.snapshot.disk.available)}</dd></div>
          <div className="settings-summary-row"><dt>GPU</dt><dd>Unknown</dd></div>
          <div className="settings-summary-row"><dt>Dedicated GPU VRAM</dt><dd>Unknown</dd></div>
          <div className="settings-summary-row"><dt>Unified memory availability</dt><dd>Unknown</dd></div>
        </dl>
        <p className="settings-hint muted">Free RAM is an OS observation, not a parser memory budget. System RAM does not measure available GPU or unified memory.</p>
        <ul className="settings-component-list">
          {state.snapshot.components.map(component => <li key={component.id} className="settings-component">
            <div className="settings-component-copy">
              <strong>{names[component.id]}</strong>
              <span className="muted">{componentStatus(component)}{component.version && ` · ${component.version}`}</span>
            </div>
          </li>)}
        </ul>
        <p className="settings-hint muted">PATH results cover the ingestion process PATH. uvx is a launcher; its presence does not establish that MarkItDown is installed.</p>
        <p className="settings-hint muted">Observed <time dateTime={state.snapshot.observedAt}>{formatDateTime(state.snapshot.observedAt)}</time>.</p>
      </>}
    </section>
  )
}

function componentStatus(component: ParserComponentObservation) {
  if (component.status === "unknown") return "Unknown"
  if (component.status === "not_found") return component.basis === "PATH" ? "Not found on ingest PATH" : "Not found in package resolution"
  return component.basis === "PATH" ? "Found on ingest PATH" : "Package metadata found"
}

function formatBytes(observation: ByteObservation) {
  if (observation.status === "unknown") return "Unknown"
  const units = ["B", "KiB", "MiB", "GiB", "TiB"]
  const index = observation.bytes === 0 ? 0 : Math.min(Math.floor(Math.log(observation.bytes) / Math.log(1024)), units.length - 1)
  return `${new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 }).format(observation.bytes / 1024 ** index)} ${units[index]}`
}
