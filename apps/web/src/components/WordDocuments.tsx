import { createContext, useContext, useEffect, useRef, useState, type ReactNode } from "react"
import { fetchWordConfig, fetchWordDocuments, linkedWordDocument, revokeWordDocument, type WordConfig, type WordOpenDocument } from "../api/ingest/word"
import { humanizeError } from "../api/ingest/errors"
import { useToast } from "./Toast"

type WordState = {
  config?: WordConfig
  setupError?: string
  documents: WordOpenDocument[]
  statusError?: string
  refreshKey: number
  refresh: () => void
  retrySetup: () => void
}
const WordDocuments = createContext<WordState>({ documents: [], refreshKey: 0, refresh: () => {}, retrySetup: () => {} })

// Heartbeats only say that Word is in contact. Version identity, rather than a
// save-time timestamp, tells the app that ingest actually received new bytes.
export function wordVersionsChanged(before: WordOpenDocument[], after: WordOpenDocument[]) {
  return after.some((next) => {
    const previous = before.find((entry) => entry.name === next.name)
    return previous ? previous.currentVersion !== next.currentVersion : next.currentVersion !== null
  })
}

export function WordDocumentsProvider({ matterId, onChanged, children }: { matterId: string; onChanged: () => void; children: ReactNode }) {
  const [config, setConfig] = useState<WordConfig>()
  const [setupError, setSetupError] = useState<string>()
  const [documents, setDocuments] = useState<WordOpenDocument[]>([])
  const [statusError, setStatusError] = useState<string>()
  const [refreshKey, setRefreshKey] = useState(0)
  const [setupAttempt, setSetupAttempt] = useState(0)
  const changed = useRef(onChanged)
  changed.current = onChanged
  const refresh = useRef<() => void>(() => {})

  useEffect(() => {
    let cancelled = false
    fetchWordConfig().then((value) => {
      if (cancelled) return
      setConfig(value)
      setSetupError(undefined)
    }, (error) => !cancelled && setSetupError(humanizeError(error, "Could not check Word setup.")))
    return () => { cancelled = true }
  }, [matterId, setupAttempt])

  useEffect(() => {
    let cancelled = false
    let running = false
    let focused = false
    let previous: WordOpenDocument[] | undefined
    setDocuments([])
    setStatusError(undefined)
    async function read(fromFocus = false) {
      focused ||= fromFocus
      if (running) return
      running = true
      let received = false
      try {
        const next = await fetchWordDocuments(matterId)
        if (cancelled) return
        received = previous !== undefined && wordVersionsChanged(previous, next)
        previous = next
        setDocuments(next)
        setStatusError(undefined)
      } catch {
        if (!cancelled) setStatusError("Word status could not be refreshed. The last known status is shown.")
      } finally {
        running = false
        if (!cancelled && (received || focused)) {
          focused = false
          setRefreshKey((value) => value + 1)
          changed.current()
        }
      }
    }
    const focus = () => { void read(true) }
    refresh.current = () => { void read() }
    void read()
    const timer = window.setInterval(() => { void read() }, 5000)
    window.addEventListener("focus", focus)
    return () => {
      cancelled = true
      window.clearInterval(timer)
      window.removeEventListener("focus", focus)
    }
  }, [matterId])

  return <WordDocuments.Provider value={{ config, setupError, documents, statusError, refreshKey, refresh: () => refresh.current(), retrySetup: () => setSetupAttempt((value) => value + 1) }}>{children}</WordDocuments.Provider>
}

// A refresh never remounts the viewer or clears a typed edit request. Defer
// until an in-flight write completes so a stale read cannot replace its result.
export function useWordRefresh(onRefresh: () => void, defer = false) {
  const { refreshKey } = useContext(WordDocuments)
  const seen = useRef(refreshKey)
  const pending = useRef(false)
  useEffect(() => {
    if (seen.current !== refreshKey) pending.current = true
    seen.current = refreshKey
    if (!defer && pending.current) {
      pending.current = false
      onRefresh()
    }
  }, [refreshKey, defer, onRefresh])
}

export function wordWarning(documents: WordOpenDocument[], name?: string) {
  const open = documents.filter((entry) => entry.open)
  if (!open.length) return
  const local = name && open.some((entry) => entry.name === name)
  const names = local ? name : open.map((entry) => entry.name).join(", ")
  return `${names} ${open.length === 1 || local ? "is" : "are"} open in Word. Use Save to doc.haus in Word and close the document before changing it here. Changes made in both places are kept as separate versions.`
}

export function WordWarning({ name }: { name?: string }) {
  const { documents, statusError } = useContext(WordDocuments)
  const warning = wordWarning(documents, name)
  if (!warning && !statusError) return null
  return <p className="word-warning" role="status">{warning}{statusError && ` ${statusError}`}</p>
}

export function WordBadge({ name }: { name: string }) {
  const status = useContext(WordDocuments).documents.find((entry) => entry.name === name)
  if (!status?.open && !status?.overlap) return null
  return <span className={`word-badge${status.overlap ? " has-overlap" : ""}`} title={status.overlap ? "Word and doc.haus changed this document. Compare its versions before continuing." : "The Word add-in is in contact with doc.haus; this does not confirm every local save was received."}>
    {status.overlap ? "Overlapping changes" : "Open in Word"}
  </span>
}

export function wordSetupReason(config?: WordConfig, error?: string) {
  if (error) return error
  if (!config) return "Checking Word setup…"
  if (!config.available) return config.reason || "Word needs a trusted HTTPS connection to this doc.haus instance. Complete Word setup on the server, then try again."
}

function WordInstallInstructions() {
  return <>
    <p>If doc.haus runs on this Mac or Windows computer, open Terminal on Mac or PowerShell on Windows. In your doc.haus folder, run <code>bun scripts/setup-word-addin.ts</code>, then restart doc.haus and Word. Setup installs the add-in and asks your computer to trust the local Word connection.</p>
    <p>If your firm runs doc.haus on another computer, ask your administrator to set up the trusted Word connection and install its add-in on this computer.</p>
  </>
}

// Kept nearby instead of repeated on every row. It remains discoverable when
// Word cannot be used, with the actual reason and a way to check after setup.
export function WordSetup() {
  const { config, setupError, retrySetup } = useContext(WordDocuments)
  const reason = wordSetupReason(config, setupError)
  if (reason) return <div className="word-setup" role="status">
    <span>{reason}</span>
    {config && !config.available && <details><summary>Set up Word</summary><WordInstallInstructions /></details>}
    {(config || setupError) && <button type="button" className="linklike" onClick={retrySetup}>Check again</button>}
  </div>
  return <details className="word-setup">
    <summary>Using Word for the first time?</summary>
    <p>Install the doc.haus Word add-in on this computer once, then open the downloaded document in Word. Use <strong>Save to doc.haus</strong> in its pane and wait for confirmation before closing. Background updates may combine saves or stop when the document closes.</p>
    <WordInstallInstructions />
    {config?.manifestUrl && <a href={config.manifestUrl} download="doc-haus-word.xml">Download the Word add-in manifest</a>}
  </details>
}

export function WordOpenControl({ matterId, name, disabled = false, compact = false }: { matterId: string; name: string; disabled?: boolean; compact?: boolean }) {
  const { config, setupError, documents, refresh } = useContext(WordDocuments)
  const [busy, setBusy] = useState<"open" | "revoke">()
  const toast = useToast()
  if (!/\.docx$/i.test(name)) return null
  const linked = documents.some((entry) => entry.name === name)
  const available = config?.available && !setupError
  async function open() {
    setBusy("open")
    try {
      const bytes = await linkedWordDocument(matterId, name, window.location.origin)
      const url = URL.createObjectURL(bytes)
      const a = document.createElement("a")
      a.href = url
      a.download = name
      a.click()
      window.setTimeout(() => URL.revokeObjectURL(url), 1000)
      refresh()
      toast("success", `Open the downloaded ${name} in Word. Use Save to doc.haus in its pane and wait for confirmation before closing.`)
    } catch (error) {
      toast("error", humanizeError(error, `Could not open ${name} in Word.`))
    } finally { setBusy(undefined) }
  }
  async function revoke() {
    setBusy("revoke")
    try {
      await revokeWordDocument(matterId, name)
      refresh()
      toast("success", `Disconnected the Word copies of ${name}. Their local files are kept; they can no longer save to doc.haus.`)
    } catch (error) {
      toast("error", humanizeError(error, `Could not disconnect ${name} from Word.`))
    } finally { setBusy(undefined) }
  }
  return <span className={`word-controls${compact ? " is-compact" : ""}`}>
    {available ? <button type="button" className={compact ? "linklike word-open" : "word-open"} onClick={open} disabled={disabled || busy !== undefined} aria-label={`Open ${name} in Word`} title="Download a linked copy to open in Word">
      {busy === "open" ? "Preparing Word copy…" : "Open in Word"}
    </button> : <span className="word-unavailable" title={wordSetupReason(config, setupError)}>{config || setupError ? "Word setup needed" : "Checking Word…"}</span>}
    {linked && <button type="button" className="linklike word-revoke" onClick={revoke} disabled={busy !== undefined} aria-label={`Disconnect ${name} from Word`} title="Stop all existing Word copies of this document from saving to doc.haus">
      {busy === "revoke" ? "Disconnecting…" : "Disconnect Word"}
    </button>}
  </span>
}
