import { useEffect, useState } from "react"
import { Route, Routes, useLocation } from "react-router-dom"
import Matters from "./pages/Matters"
import Templates from "./pages/Templates"
import Playbooks from "./pages/Playbooks"
import Workflows from "./pages/Workflows"
import Skills from "./pages/Skills"
import Agents from "./pages/Agents"
import MatterDetail from "./pages/MatterDetail"
import NotFound from "./pages/NotFound"
import Settings from "./components/settings/Settings"
import Onboarding from "./components/onboarding/Onboarding"
import Sidebar from "./components/Sidebar"
import BootScreen from "./components/BootScreen"
import ConnectionBanner from "./components/ConnectionBanner"
import ProviderConnectionBanner from "./components/ProviderConnectionBanner"
import LaunchNotice, { hasAcceptedLaunchNotice } from "./components/LaunchNotice"
import { ToastProvider } from "./components/Toast"
import useExit, { Exit } from "./components/useExit"
import { getSetup, type SetupProgress } from "./api/ingest/setup"
import { OPENCODE_URL } from "./config"
import { applyAppearance } from "./prefs"
import { useUpdates, UpdateButton, UpdateModal } from "./components/Updates"

export default function App() {
  const [acknowledged, setAcknowledged] = useState(hasAcceptedLaunchNotice)
  const [settings, setSettings] = useState(false)
  const [settingsShown, settingsClosing] = useExit(settings)
  const [reconnectProvider, setReconnectProvider] = useState<string | undefined>()
  // Text size lands as a data attribute on <html> the stylesheet keys off.
  useEffect(() => {
    applyAppearance()
  }, [])
  const [onboarding, setOnboarding] = useState(false)
  const [setup, setSetup] = useState<SetupProgress | null>(null)
  const [engineReady, setEngineReady] = useState(false)
  const [ingestReady, setIngestReady] = useState(false)
  const [bootError, setBootError] = useState("")
  const [bootRetry, setBootRetry] = useState(0)
  useEffect(() => {
    if (!acknowledged) return
    const controller = new AbortController()
    async function boot() {
      setBootError("")
      for (let attempt = 0; attempt < 3 && !controller.signal.aborted; attempt++) {
        try {
          const response = await fetch(`${OPENCODE_URL}/api/dochaus/ready`, { signal: AbortSignal.any([controller.signal, AbortSignal.timeout(10000)]) })
          if (!response.ok || !(await response.json()).ready) throw new Error("The engine is not ready yet.")
          if (controller.signal.aborted) return
          setEngineReady(true)
          const progress = await getSetup()
          if (controller.signal.aborted) return
          setSetup(progress); setIngestReady(true); setOnboarding(progress.next !== "done")
          return
        } catch {
          if (controller.signal.aborted) return
        }
      }
      if (!controller.signal.aborted) setBootError("The services did not become ready. Check the engine and document service, then retry.")
    }
    void boot()
    return () => controller.abort()
  }, [acknowledged, bootRetry])
  const updates = useUpdates(acknowledged && setup?.next === "done")
  const booted = engineReady && ingestReady && setup !== null
  async function refreshSetup() {
    try {
      const progress = await getSetup(); setSetup(progress)
    } catch { setBootError("Could not check saved setup. Retry before generating work."); setSetup(null) }
  }
  // Bumped when a matter mints a new conversation, so the sidebar re-lists its
  // chats the instant one starts rather than waiting for the turn to settle.
  const [sessionsVersion, setSessionsVersion] = useState(0)
  // Bumped when the sample matter is added, so the matters list re-loads.
  const [mattersVersion, setMattersVersion] = useState(0)
  const location = useLocation()
  if (!acknowledged) return <LaunchNotice onAccept={() => setAcknowledged(true)} />
  if (!booted)
    return (
      <ToastProvider>
        <BootScreen engineReady={engineReady} ingestReady={ingestReady} error={bootError} onRetry={() => setBootRetry(value => value + 1)} />
      </ToastProvider>
    )
  return (
    <ToastProvider>
      <div className={`shell${location.pathname.startsWith("/matter/") ? " shell-matter" : ""}`}>
        <a className="skip-link" href="#main-content">Skip to workspace</a>
        <Sidebar onOpenSettings={() => { setReconnectProvider(undefined); setSettings(true) }} sessionsVersion={sessionsVersion} updateButton={<UpdateButton updates={updates} />} />
        {/* The open matter's documents drawer portals here, between the rail and the canvas. */}
        <div id="matter-drawer-slot" className="matter-drawer-slot" />
        <main id="main-content" className="content" tabIndex={-1}>
          <ConnectionBanner />
          <ProviderConnectionBanner enabled={setup?.next === "done" && !onboarding} onOpenSettings={providerID => { setReconnectProvider(providerID); setSettings(true) }} />
          {setup?.next !== "done" && !onboarding && <p className="settings-notice">Setup is incomplete. {setup?.ready ? "Your models are saved. Finish or skip the remaining optional steps before starting assistant work." : "Save your model choices before starting assistant work."} <button onClick={() => setOnboarding(true)}>Resume setup</button></p>}
          {/* Keyed on the path so each page settles in with a short fade. */}
          <div className={`container route-fade${location.pathname.startsWith("/matter/") ? " matter-container" : ""}`} key={location.pathname}>
            <Routes>
              <Route path="/" element={<Matters refresh={mattersVersion} />} />
              <Route path="/templates" element={<Templates />} />
              <Route path="/playbooks" element={<Playbooks />} />
              <Route path="/workflows" element={<Workflows />} />
              <Route path="/skills" element={<Skills />} />
              <Route path="/agents" element={<Agents />} />
              <Route
                path="/matter/:id"
                element={setup?.ready && setup.next === "done" ? <MatterDetail onSessionsChanged={() => setSessionsVersion((v) => v + 1)} /> : <p className="settings-notice">Finish setup before opening assistant work. Optional steps can be skipped. <button onClick={() => setOnboarding(true)}>Set up models</button><button onClick={() => setSettings(true)}>Open settings</button></p>}
              />
              <Route path="*" element={<NotFound />} />
            </Routes>
          </div>
        </main>
        <Exit when={updates.open && setup?.next === "done" && !onboarding && !settingsShown}>{() => <UpdateModal updates={updates} />}</Exit>
        <Exit when={onboarding}>{() => (
          <Onboarding
            onClose={() => { setOnboarding(false); void refreshSetup() }}
            onSampleAdded={() => setMattersVersion((v) => v + 1)}
          />
        )}</Exit>
        {settingsShown && (
          <Settings
            closing={settingsClosing}
            initialTab={reconnectProvider !== undefined ? "ai" : undefined}
            reconnectProvider={reconnectProvider}
            onSampleAdded={() => setMattersVersion((v) => v + 1)}
            onClose={() => {
              // A second close during the exit animation changes nothing.
              if (!settings) return
              setSettings(false)
              void refreshSetup()
            }}
          />
        )}
      </div>
    </ToastProvider>
  )
}
