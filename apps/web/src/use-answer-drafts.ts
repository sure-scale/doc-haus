import { useEffect, useMemo, useSyncExternalStore } from "react"
import { INGEST_URL } from "./config"
import { AnswerDraftStore, type DraftAnswer, type DraftAnswers } from "./answer-draft-store"
import { readBrowserJSON, removeBrowserValue, writeBrowserValue } from "./browser-storage"

const stores = new Map<string, AnswerDraftStore>()

function validAnswer(value: unknown): value is DraftAnswer {
  if (typeof value === "string") return true
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false
  return Object.values(value).every(
    (field) =>
      typeof field === "string" ||
      typeof field === "number" ||
      typeof field === "boolean" ||
      (Array.isArray(field) && field.every((item) => typeof item === "string")),
  )
}

export function useAnswerDrafts(matterId: string | undefined, scope: string, legacyKey: string) {
  const store = useMemo(() => {
    const url = `${INGEST_URL}/matters/${encodeURIComponent(matterId ?? "")}/answer-drafts/${encodeURIComponent(scope)}`
    const key = matterId ? url : `browser:${legacyKey}`
    const existing = stores.get(key)
    if (existing) return existing
    const created = new AnswerDraftStore(
      async (input) => {
        // Library conversations have native session forms without a matter.
        // Keep their unsent answers in this browser, never at a fake matter URL.
        if (!matterId) {
          if (input) {
            const next = { ...input, revision: input.revision + 1 }
            if (!writeBrowserValue(legacyKey, JSON.stringify(next))) throw new Error("Could not save answers in this browser. Keep this page open and retry.")
            return next
          }
          const saved = readBrowserJSON(legacyKey) as { answers?: unknown; revision?: unknown } | null
          const answers = saved && typeof saved.answers === "object" && saved.answers !== null && !Array.isArray(saved.answers)
            ? Object.fromEntries(Object.entries(saved.answers).filter((entry): entry is [string, DraftAnswer] => validAnswer(entry[1])))
            : {}
          return { answers, revision: typeof saved?.revision === "number" ? saved.revision : 0 }
        }
        const response = await fetch(url, {
          method: input ? "PUT" : "GET",
          ...(input ? { headers: { "Content-Type": "application/json" }, body: JSON.stringify(input) } : {}),
          signal: AbortSignal.timeout(10000),
        })
        if (!response.ok) {
          const body = (await response.json().catch(() => null)) as { error?: string } | null
          throw new Error(body?.error ?? "Could not save answers. Keep this page open and retry.")
        }
        return response.json() as Promise<{ answers: DraftAnswers; revision: number }>
      },
      () => {
        if (!matterId) return {}
        const saved = readBrowserJSON(legacyKey)
        if (typeof saved !== "object" || saved === null || Array.isArray(saved)) return {}
        return Object.fromEntries(
          Object.entries(saved).filter(
            (entry): entry is [string, DraftAnswer] =>
              validAnswer(entry[1]) &&
              (scope.startsWith("lean.") ? typeof entry[1] === "string" : typeof entry[1] !== "string"),
          ),
        )
      },
      () => {
        if (matterId) removeBrowserValue(legacyKey)
      },
    )
    stores.set(key, created)
    return created
  }, [matterId, scope, legacyKey])
  const snapshot = useSyncExternalStore(store.subscribe, store.getSnapshot, store.getSnapshot)
  useEffect(() => {
    if (!scope) return
    void store.load()
    window.addEventListener("focus", store.load)
    return () => window.removeEventListener("focus", store.load)
  }, [store, matterId, scope])
  return { ...snapshot, set: store.set, retry: store.retry }
}
