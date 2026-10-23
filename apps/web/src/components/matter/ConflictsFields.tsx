import { useId, type ReactNode } from "react"
import { conflictsProblems, type Conflicts, type ConflictsInput, type ConflictsProblems } from "../../api/ingest/conflicts"

// "none" leaves the check unrecorded: the new-matter default, so creating a
// matter stays one click.
export type ConflictsChoice = Conflicts["status"] | "none"
export type ConflictsDraft = { status: ConflictsChoice; by: string; owner: string; due: string; reason: string }

export function conflictsDraft(current: Conflicts | undefined, initial: ConflictsChoice, attorney = ""): ConflictsDraft {
  return {
    status: current?.status ?? initial,
    by: current?.by ?? attorney,
    owner: current?.owner ?? attorney,
    due: current?.due ?? "",
    reason: current?.reason ?? "",
  }
}

// The draft as the endpoint's input; undefined while it is left unrecorded.
export function draftInput(draft: ConflictsDraft): ConflictsInput | undefined {
  return draft.status === "none" ? undefined : { ...draft, status: draft.status }
}

// The names the draft still needs; none while it is left unrecorded.
export function draftProblems(draft: ConflictsDraft): ConflictsProblems {
  const input = draftInput(draft)
  return input ? conflictsProblems(input) : {}
}

const STATES: { value: ConflictsChoice; label: string }[] = [
  { value: "none", label: "Not recorded" },
  { value: "cleared", label: "Cleared" },
  { value: "pending", label: "Pending" },
  { value: "bypassed", label: "Bypassed" },
]

// The conflicts-check fields shared by the header's dialog and the new and
// edit matter forms: the state, then only the names that state needs. The host
// passes its own field class so the fields match the form around them, and
// shows problems once a save has been tried.
export default function ConflictsFields({
  draft,
  onChange,
  problems = {},
  legend,
  className = "",
  fieldClassName,
  allowNone = false,
  attorney,
  autoFocus = false,
}: {
  draft: ConflictsDraft
  onChange: (draft: ConflictsDraft) => void
  problems?: ConflictsProblems
  legend: ReactNode
  className?: string
  fieldClassName: string
  allowNone?: boolean
  // The firm's attorney name, filled into a blank name when a state is chosen.
  attorney?: string
  autoFocus?: boolean
}) {
  const id = useId()
  const set = (fields: Partial<ConflictsDraft>) => onChange({ ...draft, ...fields })
  const choose = (status: ConflictsChoice) => set({ status, by: draft.by || attorney || "", owner: draft.owner || attorney || "" })

  function nameField(key: "by" | "owner" | "reason", label: string, placeholder: string, focus: boolean) {
    const problem = problems[key]
    return (
      <div className={fieldClassName}>
        <label htmlFor={`${id}-${key}`}>{label}</label>
        <input
          id={`${id}-${key}`}
          autoFocus={autoFocus && focus}
          value={draft[key]}
          placeholder={placeholder}
          aria-required="true"
          aria-invalid={problem ? true : undefined}
          aria-describedby={problem ? `${id}-${key}-error` : undefined}
          onChange={(e) => set({ [key]: e.target.value })}
        />
        {problem && (
          <small id={`${id}-${key}-error`} className="field-error">
            {problem}
          </small>
        )}
      </div>
    )
  }

  return (
    <fieldset className={`conflicts-fields ${className}`.trim()}>
      <legend id={`${id}-legend`}>{legend}</legend>
      <div className="conflicts-states" role="radiogroup" aria-labelledby={`${id}-legend`}>
        {STATES.filter((row) => allowNone || row.value !== "none").map((row) => (
          <label key={row.value}>
            <input type="radio" name={`${id}-status`} value={row.value} checked={draft.status === row.value} onChange={() => choose(row.value)} />
            {row.label}
          </label>
        ))}
      </div>
      {draft.status === "cleared" && nameField("by", "Checked by", "Name", true)}
      {draft.status === "pending" && (
        <>
          {nameField("owner", "Owner", "Who will complete the check", true)}
          <div className={fieldClassName}>
            <label htmlFor={`${id}-due`}>Due (optional)</label>
            <input id={`${id}-due`} type="date" value={draft.due} onChange={(e) => set({ due: e.target.value })} />
          </div>
          <p className="conflicts-warning" role="note">
            Pending blocks redlines and drafting tools on this matter until it is cleared or bypassed.
          </p>
        </>
      )}
      {draft.status === "bypassed" && (
        <>
          {nameField("by", "Bypassed by", "Name", true)}
          {nameField("reason", "Reason", "Why the check is bypassed", false)}
          <p className="conflicts-note">Reviews will state who bypassed the check and why.</p>
        </>
      )}
      {draft.status !== "none" && <p className="conflicts-note">This records that a check happened. doc.haus does not check for conflicts.</p>}
    </fieldset>
  )
}
