import { useEffect, useRef, useState } from "react"
import { DashboardTabConflict, getDashboardTabs, putDashboardTabs, type DashboardTabSnapshot, type DashboardTabState } from "../../api/ingest/dashboard-tabs"
import { readBrowserValue } from "../../browser-storage"
import { loadDashboards, readDashboards } from "./layout"

const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
function normalized(state: DashboardTabState): DashboardTabState {
  state = { ...state, workspace: readDashboards(state.workspace) }
  const custom = state.workspace.dashboards.filter(tab => tab.id !== "overview")
  const active = state.workspace.active === "overview" && custom[0] ? custom[0].id : state.workspace.active
  return { ...state, workspace: { ...state.workspace, active }, view: state.view ?? (!state.setupHidden && !custom.some(tab => tab.id === active) ? "setup" : "dashboard") }
}
function legacy(): DashboardTabState {
  return normalized({ version: 1, workspace: loadDashboards(), setupHidden: readBrowserValue("dh.dashboard.setup-hidden.v1") === "true" })
}

// The database owns presentation state. Legacy browser values only seed an empty
// database; revision checks and a single writer keep queued edits in order.
export function useDashboardTabs() {
  const [state, setState] = useState<DashboardTabState>(legacy)
  const current = useRef(state)
  const [ready, setReady] = useState(false)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState("")
  const [conflict, setConflict] = useState(false)
  const revision = useRef(0)
  const confirmed = useRef<DashboardTabState | null>(null)
  const attempted = useRef<DashboardTabState | null>(null)
  const pending = useRef<DashboardTabState | undefined>(undefined)
  const saving = useRef(false)
  const blocked = useRef(false)
  const mounted = useRef(false)
  const generation = useRef(0)

  function adopt(saved: DashboardTabSnapshot) {
    if (!saved.state) throw new Error("Could not load saved dashboards.")
    revision.current = saved.revision
    confirmed.current = saved.state
    current.current = normalized(saved.state)
    setState(current.current)
    setError(""); setConflict(false); blocked.current = false
    pending.current = undefined
    setReady(true)
  }
  async function load() {
    const request = ++generation.current
    setLoading(true); setError("")
    try {
      let saved = await getDashboardTabs()
      if (!mounted.current || request !== generation.current) return
      if (!saved.state) {
        try { saved = await putDashboardTabs(current.current, 0) }
        catch (failure) {
          if (!(failure instanceof DashboardTabConflict)) throw failure
          // Another browser seeded first. It wins over this browser's legacy data.
          saved = await getDashboardTabs()
        }
      }
      if (mounted.current && request === generation.current) adopt(saved)
    } catch { if (mounted.current && request === generation.current) setError("Could not load your dashboards.") }
    finally { if (mounted.current && request === generation.current) setLoading(false) }
  }
  useEffect(() => {
    mounted.current = true
    void load()
    return () => { mounted.current = false; generation.current++ }
  }, [])

  async function flush() {
    if (saving.current || blocked.current || !mounted.current) return
    saving.current = true
    try {
      while (pending.current && !blocked.current) {
        const next = pending.current
        pending.current = undefined
        attempted.current = next
        try {
          const saved = await putDashboardTabs(next, revision.current)
          revision.current = saved.revision
          confirmed.current = saved.state
          attempted.current = null
          if (mounted.current) { setError(""); setConflict(false) }
        } catch (failure) {
          pending.current ??= next
          blocked.current = true
          if (mounted.current) {
            const changed = failure instanceof DashboardTabConflict
            setConflict(changed)
            setError(changed ? "These dashboards changed in another window. Your unsaved changes are still here." : "Could not save your dashboard changes. Your changes are still here.")
          }
        }
      }
    } finally { saving.current = false }
  }
  function change(input: DashboardTabState | ((saved: DashboardTabState) => DashboardTabState)) {
    if (!ready) return
    const next = normalized(typeof input === "function" ? input(current.current) : input)
    current.current = next
    setState(next)
    pending.current = next
    void flush()
  }
  async function retrySave() {
    if (saving.current) return
    setLoading(true)
    try {
      const saved = await getDashboardTabs()
      // A lost PUT response may already have committed. Reconcile that exact
      // attempt without replaying it over edits made by a different browser.
      if (!equal(saved.state, confirmed.current) && !equal(saved.state, attempted.current)) {
        setConflict(true)
        setError("These dashboards changed in another window. Your unsaved changes are still here.")
        return
      }
      revision.current = saved.revision
      confirmed.current = saved.state
      attempted.current = null
      blocked.current = false
      setError(""); setConflict(false)
      if (equal(saved.state, current.current)) pending.current = undefined
      else pending.current = current.current
      void flush()
    } catch { setError("Could not save your dashboard changes. Your changes are still here.") }
    finally { setLoading(false) }
  }
  return { state, ready, loading, error, conflict, change, retry: ready ? retrySave : load, loadSaved: load }
}
