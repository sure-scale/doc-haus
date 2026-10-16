import { useEffect, useMemo, useSyncExternalStore } from "react"
import { INGEST_URL } from "./config"
import { AnswerDraftStore, type DraftAnswer, type DraftAnswers } from "./answer-draft-store"
import { readBrowserJSON, removeBrowserValue } from "./browser-storage"

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

export function useAnswerDrafts(matterId: string, scope: string, legacyKey: string) {
  const store = useMemo(() => {
    const url = `${INGEST_URL}/matters/${encodeURIComponent(matterId)}/answer-drafts/${encodeURIComponent(scope)}`
    const existing = stores.get(url)
    if (existing) return existing
    const created = new AnswerDraftStore(
      async (input) => {
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
        removeBrowserValue(legacyKey)
      },
    )
    stores.set(url, created)
    return created
  }, [matterId, scope, legacyKey])
  const snapshot = useSyncExternalStore(store.subscribe, store.getSnapshot, store.getSnapshot)
  useEffect(() => {
    if (!matterId || !scope) return
    void store.load()
    window.addEventListener("focus", store.load)
    return () => window.removeEventListener("focus", store.load)
  }, [store, matterId, scope])
  return { ...snapshot, set: store.set, retry: store.retry }
}
