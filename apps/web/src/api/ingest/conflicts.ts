import { INGEST_URL } from "../../config"
import { formatDate } from "../../lib/format"
import { ingestError } from "./errors"
import type { Matter } from "./matters"

// The matter's conflicts-check record (matter.json `conflicts`). doc.haus only
// records that a check happened and how it ended; it never judges conflicts.
export type Conflicts = {
  status: "cleared" | "pending" | "bypassed"
  by?: string
  owner?: string
  due?: string
  reason?: string
  recorded_at: number
}

export type ConflictsInput = { status: Conflicts["status"]; by?: string; owner?: string; due?: string; reason?: string }

// Each state sends only the fields it needs; blank ones drop out of the body.
export function conflictsBody(input: ConflictsInput): ConflictsInput {
  const clean = (value?: string) => value?.trim() || undefined
  if (input.status === "cleared") return { status: "cleared", by: clean(input.by) }
  if (input.status === "pending") return { status: "pending", owner: clean(input.owner), due: clean(input.due) }
  return { status: "bypassed", by: clean(input.by), reason: clean(input.reason) }
}

// The names a state needs and lacks, worded for the field that is blank. The
// endpoint applies the same rules; checking first gives an inline message.
export type ConflictsProblems = Partial<Record<"by" | "owner" | "reason", string>>
export function conflictsProblems(input: ConflictsInput): ConflictsProblems {
  const body = conflictsBody(input)
  if (body.status === "cleared") return body.by ? {} : { by: "Name who ran the check." }
  if (body.status === "pending") return body.owner ? {} : { owner: "Name who will complete the check." }
  return {
    ...(body.by ? {} : { by: "Name who bypassed the check." }),
    ...(body.reason ? {} : { reason: "Give the reason for the bypass." }),
  }
}

// Whether saving the input would change the recorded check. Recording stamps a
// new date, so an edit form records only a real change.
export function conflictsChanged(input: ConflictsInput, current: Conflicts | undefined): boolean {
  if (!current) return true
  const next = conflictsBody(input)
  const now = conflictsBody(current)
  return (["status", "by", "owner", "due", "reason"] as const).some((key) => next[key] !== now[key])
}

// The endpoint validates too; its message is shown as returned.
export async function recordConflicts(id: string, input: ConflictsInput): Promise<Matter> {
  const res = await fetch(`${INGEST_URL}/matters/${encodeURIComponent(id)}/conflicts`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(conflictsBody(input)),
  })
  if (!res.ok) throw new Error(await ingestError(res, "Could not record the conflicts check"))
  return res.json()
}

// The header chip: its words and whether it asks for attention.
export function conflictsChip(conflicts: Conflicts | undefined): { text: string; tone: "none" | "ok" | "pending" | "bypassed" } {
  if (!conflicts) return { text: "Conflicts: not recorded", tone: "none" }
  if (conflicts.status === "cleared") {
    const on = formatDate(conflicts.recorded_at)
    return { text: `Conflicts: cleared by ${conflicts.by}${on ? ` on ${on}` : ""}`, tone: "ok" }
  }
  if (conflicts.status === "pending") {
    const due = conflicts.due ? formatDate(conflicts.due) : ""
    return { text: `Conflicts: pending — owner ${conflicts.owner}${due ? `, due ${due}` : ""}`, tone: "pending" }
  }
  return { text: `Conflicts: bypassed by ${conflicts.by}`, tone: "bypassed" }
}
