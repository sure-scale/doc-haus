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
      <p className="settings-hint muted">
        Automatically allow these requests in this browser across all matters. Redlines still need your review.
      </p>
      <ToggleRow
        title="Redlines & tracked changes"
        description="Propose changes for your review."
        on={prefs.autoApproveRedlines}
        onChange={(on) => set({ autoApproveRedlines: on })}
      />
      <ToggleRow
        title="New documents & templates"
        description="Create documents and templates."
        on={prefs.autoApproveDrafting}
        onChange={(on) => set({ autoApproveDrafting: on })}
      />
      {error && <p className="settings-hint" role="alert">{error}</p>}
      <p className="settings-hint muted">
        Draft runs have their own review process, and required questions still need answers. These switches don't cover
        direct edits or redactions. Saved “Always allow” decisions also apply.
      </p>
    </section>
  )
}
