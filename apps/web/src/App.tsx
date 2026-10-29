import { useCallback, useEffect, useState } from "react"
import { Route, Routes, useLocation, useNavigate } from "react-router-dom"
import Matters from "./pages/Matters"
import Dashboard from "./pages/Dashboard"
import Clients from "./pages/Clients"
import Templates from "./pages/Templates"
import Playbooks from "./pages/Playbooks"
import Workflows from "./pages/Workflows"
import Skills from "./pages/Skills"
import Agents from "./pages/Agents"
import MatterDetail from "./pages/MatterDetail"
import NotFound from "./pages/NotFound"
import Settings from "./components/settings/Settings"
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
import { SettingsNavigation } from "./components/settings/navigation"
import type { Tab } from "./components/settings/types"

export default function App() {
  const [acknowledged, setAcknowledged] = useState(hasAcceptedLaunchNotice)
  const [settings, setSettings] = useState(false)
  const [settingsTab, setSettingsTab] = useState<Tab>("ai")
  const [settingsShown, settingsClosing] = useExit(settings)
  const [recoveryRequired, setRecoveryRequired] = useState(false)
  const [reconnectProvider, setReconnectProvider] = useState<string | undefined>()
  // Text size lands as a data attribute on <html> the stylesheet keys off.
  useEffect(() => {
    applyAppearance()
  }, [])
  const openRecovery = useCallback((providerID: string) => {
    setReconnectProvider(providerID); setSettingsTab("ai"); setRecoveryRequired(true); setSettings(true)
  }, [])
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
          setSetup(progress); setIngestReady(true)
          // First-run setup lives in the dashboard. Keep a deliberately opened
          // page, and do not redirect again when Settings refreshes progress.
          if (!progress.ready && window.location.pathname === "/") navigate("/dashboard?tab=setup", { replace: true })
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
  const updates = useUpdates(acknowledged && setup?.ready === true)
  const booted = engineReady && ingestReady && setup !== null
  async function refreshSetup() {
    try {
      const progress = await getSetup(); setSetup(progress); return progress
    } catch { setBootError("Could not check saved setup. Retry before generating work."); setSetup(null) }
  }
  // Bumped when a matter mints a new conversation, so the sidebar re-lists its
  // chats the instant one starts rather than waiting for the turn to settle.
  const [sessionsVersion, setSessionsVersion] = useState(0)
  // Bumped when the sample matter is added, so the matters list re-loads.
  const [mattersVersion, setMattersVersion] = useState(0)
  const location = useLocation()
  const navigate = useNavigate()
  if (!acknowledged) return <LaunchNotice onAccept={() => setAcknowledged(true)} />
  if (!booted)
    return (
      <ToastProvider>
        <BootScreen engineReady={engineReady} ingestReady={ingestReady} error={bootError} onRetry={() => setBootRetry(value => value + 1)} />
      </ToastProvider>
    )
  return (
    <ToastProvider>
      <SettingsNavigation.Provider value={tab => { setSettingsTab(tab); setReconnectProvider(undefined); setSettings(true) }}>
      <div className={`shell${location.pathname.startsWith("/matter/") ? " shell-matter" : ""}`}>
        <a className="skip-link" href="#main-content">Skip to workspace</a>
        <Sidebar onOpenSettings={() => { setSettingsTab("ai"); setReconnectProvider(undefined); setSettings(true) }} sessionsVersion={sessionsVersion} updateButton={<UpdateButton updates={updates} />} />
        {/* The open matter's documents drawer portals here, between the rail and the canvas. */}
        <div id="matter-drawer-slot" className="matter-drawer-slot" />
        <main id="main-content" className="content" tabIndex={-1}>
          <ConnectionBanner />
          <ProviderConnectionBanner enabled={setup?.ready === true && !recoveryRequired} onOpenSettings={openRecovery} />
          {/* Keyed on the path so each page settles in with a short fade. */}
          <div className={`container route-fade${location.pathname.startsWith("/matter/") ? " matter-container" : ""}`} key={location.pathname}>
            <Routes>
              <Route path="/" element={<Matters refresh={mattersVersion} />} />
              <Route path="/dashboard" element={<Dashboard setup={setup} refresh={mattersVersion} onSampleAdded={() => setMattersVersion((v) => v + 1)} />} />
              <Route path="/clients" element={<Clients />} />
              <Route path="/templates" element={<Templates />} />
              <Route path="/playbooks" element={<Playbooks />} />
              <Route path="/workflows" element={<Workflows />} />
              <Route path="/skills" element={<Skills />} />
              <Route path="/agents" element={<Agents />} />
              <Route
                path="/matter/:id"
                element={setup?.ready ? <MatterDetail onSessionsChanged={() => setSessionsVersion((v) => v + 1)} /> : <p className="settings-notice">Connect an AI provider and save your primary and fast models to start working in this matter. <button onClick={() => { setSettingsTab("ai"); setReconnectProvider(undefined); setSettings(true) }}>Set up AI</button></p>}
              />
              <Route path="*" element={<NotFound />} />
            </Routes>
          </div>
        </main>
        <Exit when={updates.open && setup?.ready && !settingsShown}>{() => <UpdateModal updates={updates} />}</Exit>
        {settingsShown && (
          <Settings
            closing={settingsClosing}
            firstRun={!setup.ready}
            recoveryRequired={recoveryRequired}
            initialTab={reconnectProvider !== undefined ? "ai" : settingsTab}
            reconnectProvider={reconnectProvider}
            onSampleAdded={() => setMattersVersion((v) => v + 1)}
            onClose={() => {
              // A second close during the exit animation changes nothing.
              if (!settings) return
              setRecoveryRequired(false)
              setSettings(false)
              void refreshSetup()
            }}
          />
        )}
      </div>
      </SettingsNavigation.Provider>
    </ToastProvider>
  )
}
