import { useContext, useEffect, useLayoutEffect, useRef, useState } from "react"
import { Closing } from "./useExit"
import { INGEST_URL } from "../config"
import { readBrowserValue, writeBrowserValue } from "../browser-storage"

type Release = { version: number; name: string; notes: string }
type Status = { current: Release; latest?: Release; target?: string; available: boolean; error?: string; phase?: string }
const PENDING = "dh.update-pending"

export function useUpdates(enabled: boolean) {
  const [status, setStatus] = useState<Status>()
  const [open, setOpen] = useState(false)
  const [updating, setUpdating] = useState(() => Boolean(readBrowserValue(PENDING, true)))
  const [requesting, setRequesting] = useState(false)
  const [error, setError] = useState("")
  const [checkError, setCheckError] = useState("")
  const [checking, setChecking] = useState(false)
  const retry = useRef(() => {})
  const prompted = useRef(false)

  useEffect(() => {
    if (!enabled) return
    let live = true
    const started = Date.now()
    let inFlight = false
    async function check(refresh = false) {
      if (inFlight) return
      inFlight = true
      setChecking(true)
      if (refresh) setError("")
      try {
        const response = await fetch(`${INGEST_URL}/updates${refresh ? "?refresh=1" : ""}`, { signal: AbortSignal.timeout(40_000) })
        if (!response.ok) throw new Error("Could not check for updates.")
        const value: Status = await response.json()
        if (!live) return
        setStatus(value)
        if (!updating) {
          setCheckError(value.error || (value.available && !applicable(value) ? "The update information is incomplete. Check again before updating." : ""))
        }
        if (updating) {
          if (value.phase === "failed") {
            writeBrowserValue(PENDING, "", true)
            setUpdating(false)
            setError(value.error ?? "The update failed. Please try again.")
            setOpen(true)
            return
          }
          if (value.current.version >= Number(readBrowserValue(PENDING, true))) {
            writeBrowserValue(PENDING, "", true)
            window.location.reload()
            return
          }
        }
        if (!prompted.current && !value.error && applicable(value) && value.latest) {
          prompted.current = true
          if (readBrowserValue("dh.update-seen", true) !== String(value.latest.version)) {
            writeBrowserValue("dh.update-seen", String(value.latest.version), true)
            setOpen(true)
          }
        }
      } catch {
        // A restart can disconnect an accepted update. Keep reconnecting without
        // offering another install while the launcher may still be working.
        if (live && !updating) setCheckError("Could not reach the update service. Check your connection and try again.")
      } finally {
        inFlight = false
        if (live) setChecking(false)
      }
      if (live && updating && Date.now() - started > 180_000) {
        writeBrowserValue(PENDING, "", true)
        setUpdating(false)
        setError("The app has not restarted yet. Start Doc.Haus again if needed, then reload this page.")
        setOpen(true)
      }
    }
    retry.current = () => { void check(true) }
    void check()
    const timer = window.setInterval(() => { void check() }, updating ? 2000 : 60 * 60_000)
    const focus = () => { void check() }
    window.addEventListener("focus", focus)
    return () => { live = false; clearInterval(timer); window.removeEventListener("focus", focus) }
  }, [enabled, updating])

  const canApply = applicable(status) && !checkError

  async function apply() {
    if (!canApply || checking || updating || requesting) return
    setError("")
    setRequesting(true)
    try {
      const response = await fetch(`${INGEST_URL}/updates/apply`, {
        method: "POST", headers: { "Content-Type": "application/json", "x-dochaus-update": "1" },
        body: JSON.stringify({ target: status?.target }), signal: AbortSignal.timeout(40_000),
      })
      const value = await response.json()
      if (!response.ok) {
        setError(value.error ?? "Could not start the update.")
        return
      }
      // The session marker survives Vite's reload while the code is pulled.
      writeBrowserValue(PENDING, String(status?.latest?.version), true)
      setUpdating(true)
    } catch {
      // The response may have been lost after acceptance. Reconnect instead
      // of offering a second install while the launcher may be updating.
      writeBrowserValue(PENDING, String(status?.latest?.version), true)
      setUpdating(true)
    } finally {
      setRequesting(false)
    }
  }

  return { status, open, updating: updating || requesting, checking, canApply, checkFailed: Boolean(checkError), error: error || checkError, apply, retry: () => retry.current(), show: () => setOpen(true), close: () => setOpen(false) }
}

function applicable(status?: Status) {
  return Boolean(status?.available && !status.error && status.latest && Number.isSafeInteger(status.latest.version) && status.latest.version > status.current?.version && /^[a-f0-9]{40,64}$/.test(status.target ?? ""))
}

export function UpdateButton({ updates }: { updates: ReturnType<typeof useUpdates> }) {
  if (!updates.canApply && !updates.updating && !updates.error) return null
  const label = updates.updating ? "Updating…" : updates.canApply && !updates.error ? "Update available" : updates.checkFailed ? "Update check failed" : "Update status"
  return <button type="button" className="nav-item sidebar-update" onClick={updates.show} title={label} aria-label={label}>
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 16V4m-4 4 4-4 4 4M5 15v5h14v-5" /></svg>
    <span>{label}</span>
  </button>
}

export function UpdateModal({ updates }: { updates: ReturnType<typeof useUpdates> }) {
  const dialog = useRef<HTMLDialogElement>(null)
  const heading = useRef<HTMLHeadingElement>(null)
  // Closes as the exit starts; only the picture lingers.
  const closing = useContext(Closing)
  useLayoutEffect(() => {
    if (closing) return
    const modal = dialog.current
    modal?.showModal()
    heading.current?.focus()
    return () => modal?.close()
  }, [closing])
  const available = updates.canApply && !updates.error
  return <dialog ref={dialog} className={`launch-notice update-modal${closing ? " is-closing" : ""}`} inert={closing} aria-labelledby="update-title" aria-describedby="update-intro" onCancel={event => { event.preventDefault(); if (!updates.updating) updates.close() }}>
    <div className="launch-notice-body">
      <span className="launch-notice-eyebrow">Doc.Haus updates</span>
      <h1 ref={heading} tabIndex={-1} id="update-title">{updates.updating ? "Updating Doc.Haus" : available ? "Update available" : "Update status"}</h1>
      <p id="update-intro">{updates.updating ? "Installing the update and restarting the app. This page will reload when it is ready." : available ? "A new version of Doc.Haus is ready to install. Update now, or keep working and update later." : updates.error ? "You can keep working and try checking for updates again." : "Doc.Haus is up to date."}</p>
      {available && updates.status?.latest && <div className="update-release"><strong>{updates.status.latest.name}</strong><p>{updates.status.latest.notes}</p></div>}
      {!updates.updating && available && <p>Finish any uploads, drafts and answers before updating. Your matters and settings will stay in place.</p>}
      {updates.error && <p className="update-error" role="alert">{updates.error}</p>}
      {updates.updating && <p role="status">Please keep this tab open…</p>}
    </div>
    <div className="launch-notice-actions">
      <button type="button" onClick={updates.close} disabled={updates.updating}>{available ? "Later" : "Close"}</button>
      {!updates.updating && !available ? <button type="button" className="primary" disabled={updates.checking} onClick={updates.retry}>{updates.checking ? "Checking…" : "Check again"}</button> :
        <button type="button" className="primary" disabled={updates.updating || updates.checking || !updates.canApply} onClick={() => { void updates.apply() }}>{updates.updating ? "Updating…" : "Update now"}</button>}
    </div>
  </dialog>
}
