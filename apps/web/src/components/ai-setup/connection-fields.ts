import type { FormAnswer, FormField } from "@opencode/client"

export function visibleFields(fields: readonly FormField[], answer: FormAnswer, includeHidden = false) {
  return fields.filter(field => (includeHidden || !("hidden" in field && field.hidden)) && (!("when" in field) || !field.when || field.when.every(condition => {
    const value = answer[condition.key]
    if (value === undefined) return false
    const equal = Array.isArray(value) ? value.includes(String(condition.value)) : value === condition.value
    return condition.op === "eq" ? equal : !equal
  })))
}

export type MatchingSignInCandidate = {
  sourceId: string
  connectionId: string
  integrationID: string
  label: string
  authType: "key" | "oauth"
  sourceLabel: string
}
