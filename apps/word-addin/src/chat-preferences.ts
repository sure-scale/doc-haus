import type { WordChatScope } from "../../../dochaus/lib/word-chat-contract"

export const chatPreferencesKey = "dochaus.word.chat-preferences.v1"
export type ChatPreferences = { scope: WordChatScope; suggestions: boolean; target: "document" | "selection" }
type PreferencesStorage = Pick<Storage, "getItem" | "setItem">

export function readChatPreferences(storage?: PreferencesStorage): ChatPreferences {
  const defaults: ChatPreferences = { scope: "document", suggestions: true, target: "document" }
  try {
    const value: unknown = JSON.parse(storage?.getItem(chatPreferencesKey) ?? "null")
    if (!value || typeof value !== "object" || Array.isArray(value)) return defaults
    const item = value as Record<string, unknown>
    return { scope: item.scope === "matter" ? "matter" : "document", suggestions: typeof item.suggestions === "boolean" ? item.suggestions : true, target: item.target === "selection" ? "selection" : "document" }
  } catch { return defaults }
}

/** Installation preferences only; never persist prompts, selections or files. */
export function writeChatPreferences(storage: PreferencesStorage | undefined, value: ChatPreferences): void {
  try { storage?.setItem(chatPreferencesKey, JSON.stringify({ scope: value.scope, suggestions: value.suggestions, target: value.target })) }
  catch { /* Chat still works when Office blocks browser storage. */ }
}
