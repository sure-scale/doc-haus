import { useState } from "react"
import { loadPrefs, savePrefs, type Prefs } from "../../prefs"
import { ToggleRow } from "./controls"

// Personal browser choices for individual permission requests. Native saved
// allowances and the separate Draft workflow retain their own approval behavior.
export default function ApprovalsTab() {
  const [prefs, setPrefs] = useState<Prefs>(loadPrefs)
  const [error, setError] = useState("")
  const set = (patch: Partial<Prefs>) => {
    try {
      setPrefs(savePrefs(patch, { requireStorage: true }))
      setError("")
    } catch (error) {
      setError(error instanceof Error ? error.message : "Could not save this preference. Please try again.")
    }
  }
  return (
    <section className="settings-section">
      <h3>Approvals</h3>
      <p className="settings-hint muted">
        Choose which requests this browser automatically allows, across all matters and conversations. Redlines remain
        proposed changes for you to review in the document.
      </p>
      <ToggleRow
        title="Redlines & tracked changes"
        description="Proposing a tracked change or redline in a document."
        on={prefs.autoApproveRedlines}
        onChange={(on) => set({ autoApproveRedlines: on })}
      />
      <ToggleRow
        title="New documents & templates"
        description="Drafting a new document or creating a template."
        on={prefs.autoApproveDrafting}
        onChange={(on) => set({ autoApproveDrafting: on })}
      />
      {error && <p className="settings-hint" role="alert">{error}</p>}
      <p className="settings-hint muted">
        Draft runs follow their own review workflow and do not use these switches. Required questions still need your answers.
      </p>
      <p className="settings-hint muted">
        Direct edits and redactions are outside these switches. Separately saved “Always allow” decisions still apply.
      </p>
    </section>
  )
}
