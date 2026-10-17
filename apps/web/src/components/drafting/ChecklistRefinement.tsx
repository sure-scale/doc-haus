import { useId, useState } from "react"
import type { LeanDraftingStatus } from "../../api/ingest/lean"
import type { RunActions } from "./useRunActions"

export default function ChecklistRefinement({ status, actions }: { status: LeanDraftingStatus; actions: RunActions }) {
  const [selected, setSelected] = useState("")
  const id = useId()
  const refinement = status.checklist_refinement
  if (!refinement?.available || status.status !== "blocked" || status.in_flight || status.pending_requests > 0 ||
    status.stages["checklist-check"]?.status !== "blocked") return null
  const parents = refinement.parents.filter(parent => parent.scopes.includes("target_body") &&
    (parent.finding.status === "unsatisfied" || parent.finding.status === "indeterminate"))
  if (!parents.length) return null
  const parent = parents.find(parent => parent.parent_id === selected)
  const titles = new Map((status.plan?.documents ?? []).map(document => [document.id, document.title]))

  return (
    <form className="run-checklist-refinement" aria-label="Check a requirement against the full draft" onSubmit={async event => {
      event.preventDefault()
      if (!parent || actions.busy || actions.resuming) return
      if (await actions.refine(parent.parent_id)) setSelected("")
    }}>
      <fieldset disabled={actions.busy || actions.resuming}>
        <legend>Check a requirement against the full draft</legend>
        <div className="run-checklist-choices">
          <p className="run-note">Select the requirement to compare with the complete draft and the work recorded for it. The earlier finding stays in the record.</p>
          {parents.map((parent, index) => (
            <label className="run-checklist-choice" key={parent.parent_id}>
              <input type="radio" name={id} value={index} checked={selected === parent.parent_id} onChange={() => setSelected(parent.parent_id)} />
              <span>
                <strong>{parent.check}</strong>
                {titles.has(parent.target.document) && <span className="run-note">{titles.get(parent.target.document)}</span>}
                <span className="run-note">{parent.finding.status === "unsatisfied" ? "Not yet met" : "Could not confirm"}: {parent.finding.reason}</span>
              </span>
            </label>
          ))}
        </div>
        <button className="primary" type="submit" disabled={!parent || actions.busy || actions.resuming}>
          {actions.busy ? "Requesting check…" : "Check the full draft"}
        </button>
      </fieldset>
    </form>
  )
}
