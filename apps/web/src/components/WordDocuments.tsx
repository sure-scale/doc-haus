import { createContext, useContext, useEffect, useRef, useState, type ReactNode } from "react"
import { fetchWordConfig, fetchWordDocuments, linkedWordDocument, recoverWordEditing, type WordConfig, type WordOpenDocument } from "../api/ingest/word"
import { humanizeError } from "../api/ingest/errors"
import { useToast } from "./Toast"
import { useOpenSettings } from "./settings/navigation"

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
    const changed = () => setSetupAttempt(value => value + 1)
    window.addEventListener("dochaus:word-setup", changed)
    return () => window.removeEventListener("dochaus:word-setup", changed)
  }, [])

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
    let pending = false
    let epoch = 0
    let focused = false
    let previous: WordOpenDocument[] | undefined
    setDocuments([])
    setStatusError(undefined)
    async function read(fromFocus = false) {
      focused ||= fromFocus
      if (running) { pending = true; return }
      running = true
      const requested = epoch
      let received = false
      try {
        const next = await fetchWordDocuments(matterId)
        if (cancelled || requested !== epoch) return
        received = previous !== undefined && wordVersionsChanged(previous, next)
        previous = next
        setDocuments(next)
        setStatusError(undefined)
      } catch {
        if (!cancelled && requested === epoch) setStatusError("Word status could not be refreshed. The last known status is shown.")
      } finally {
        running = false
        if (!cancelled && (received || focused)) {
          focused = false
          setRefreshKey((value) => value + 1)
          changed.current()
        }
        if (!cancelled && pending) {
          pending = false
          void read()
        }
      }
    }
    const focus = () => { void read(true) }
    refresh.current = () => { epoch++; void read() }
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
  const locked = documents.filter((entry) => entry.lease?.state === "other" && (!name || entry.name === name))
  if (locked.length) return `${locked.map((entry) => entry.name).join(", ")} ${locked.length === 1 ? "is" : "are"} being edited in Word. Changes here are locked until the Word editor finishes or you take over editing. Local Word files are kept when ownership changes.`
  const open = documents.filter((entry) => entry.open)
  if (!open.length) return
  const local = name && open.some((entry) => entry.name === name)
  const names = local ? name : open.map((entry) => entry.name).join(", ")
  return `${names} ${open.length === 1 || local ? "is" : "are"} open in Word. Use Save to doc.haus in Word and close the document before changing it here. A Word copy that falls behind must be compared with the current version before saving.`
}

export function useWordEditingLocked(name: string) {
  return useContext(WordDocuments).documents.some((entry) => entry.name === name && entry.lease?.state === "other")
}

export function WordWarning({ name }: { name?: string }) {
  const { documents, statusError } = useContext(WordDocuments)
  const warning = wordWarning(documents, name)
  if (!warning && !statusError) return null
  return <p className="word-warning" role="status">{warning}{statusError && ` ${statusError}`}</p>
}

export function WordBadge({ name }: { name: string }) {
  const status = useContext(WordDocuments).documents.find((entry) => entry.name === name)
  const locked = status?.lease?.state === "other"
  if (!status?.open && !status?.overlap && !locked) return null
  return <span className={`word-badge${status.overlap ? " has-overlap" : ""}`} title={status.overlap ? "Word and doc.haus changed this document. Compare its versions before continuing." : "The Word add-in is in contact with doc.haus; this does not confirm every local save was received."}>
    {status.overlap ? "Overlapping changes" : locked ? "Editing in Word" : "Open in Word"}{status.currentNumber !== undefined && ` · Version ${status.currentNumber}`}
  </span>
}

export function wordSetupReason(config?: WordConfig, error?: string) {
  if (error) return error
  if (!config) return "Checking Word setup…"
  if (!config.available) return "Word is not ready. Set it up in Settings → Add-ins."
}

// Kept nearby instead of repeated on every row. It remains discoverable when
// Word cannot be used, with the actual reason and a way to check after setup.
export function WordSetup() {
  const { config, setupError, retrySetup } = useContext(WordDocuments)
  const openSettings = useOpenSettings()
  const reason = wordSetupReason(config, setupError)
  if (reason) return <div className="word-setup" role="status">
    <span>{reason}</span>
    {config && !config.available && <button type="button" className="linklike" onClick={() => openSettings("addins")}>Set up Word</button>}
    {(config || setupError) && <button type="button" className="linklike" onClick={retrySetup}>Check again</button>}
  </div>
  return <details className="word-setup">
    <summary>Using Word for the first time?</summary>
    <p>Choose Open in Word and open the downloaded copy. Keep doc.haus running, then choose <strong>Save</strong> in Word and wait for <strong>Saved</strong> before closing.</p>
    <button type="button" className="linklike" onClick={() => openSettings("addins")}>Word add-in settings</button>
  </details>
}

export function WordOpenControl({ matterId, name, disabled = false, compact = false }: { matterId: string; name: string; disabled?: boolean; compact?: boolean }) {
  const { config, setupError, documents, refresh } = useContext(WordDocuments)
  const [busy, setBusy] = useState<"open" | "release" | "takeover">()
  const [recovery, setRecovery] = useState<{ action: "release" | "takeover"; document: WordOpenDocument }>()
  const toast = useToast()
  if (!/\.docx$/i.test(name)) return null
  const status = documents.find((entry) => entry.name === name)
  const locked = status?.lease?.state === "other"
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
  async function recover() {
    if (!recovery) return
    const { action, document } = recovery
    setBusy(action)
    try {
      await recoverWordEditing(matterId, document, action)
      setRecovery(undefined)
      refresh()
      toast("success", action === "takeover"
        ? "Editing is available here. Earlier Word copies are disconnected; their local files are kept. Open a fresh Word copy to continue there."
        : "Word editing lock released. A Word copy must check its version and claim editing again before saving.")
    } catch (error) {
      refresh()
      toast("error", humanizeError(error, "Could not change Word editing ownership."))
    } finally { setBusy(undefined) }
  }
  return <span className={`word-controls${compact ? " is-compact" : ""}`}>
    {available ? <button type="button" className={compact ? "linklike word-open" : "word-open"} onClick={open} disabled={disabled || busy !== undefined} aria-label={`Open ${name} in Word`} title="Download a linked copy to open in Word">
      {busy === "open" ? "Preparing Word copy…" : "Open in Word"}
    </button> : <span className="word-unavailable" title={wordSetupReason(config, setupError)}>{config || setupError ? "Word setup needed" : "Checking Word…"}</span>}
    {locked && <>
      <button type="button" className="linklike" disabled={busy !== undefined || disabled} onClick={() => setRecovery(status ? { action: "takeover", document: status } : undefined)}>Take over editing</button>
      <button type="button" className="linklike" disabled={busy !== undefined || disabled} onClick={() => setRecovery(status ? { action: "release", document: status } : undefined)}>Release Word lock</button>
    </>}
    {recovery && <span className="word-recovery" role="group" aria-label="Confirm editing ownership change">
      <span>{recovery.action === "takeover" ? "Disconnect earlier Word copies and continue editing here? Their unsent edits remain in their local files." : "Release this editing lock? The Word editor will need to check its version and claim editing again."}</span>
      <button type="button" onClick={recover} disabled={busy !== undefined}>{busy === recovery.action ? "Updating…" : "Confirm"}</button>
      <button type="button" className="linklike" onClick={() => setRecovery(undefined)} disabled={busy !== undefined}>Cancel</button>
    </span>}
  </span>
}
