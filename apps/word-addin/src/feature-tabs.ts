/** Installation preference only. Never persist a document, conversation or source. */
export const FEATURE_TAB_KEY = "dochaus.word.feature-tab"
export type FeatureTab = "chat" | "documents" | "compare"
export function featureTab(value: unknown): FeatureTab | undefined {
  return value === "chat" || value === "documents" || value === "compare" ? value : undefined
}
export function readFeatureTab(storage?: Pick<Storage, "getItem">): { tab: FeatureTab; obsolete: boolean } {
  try {
    const saved = storage?.getItem(FEATURE_TAB_KEY)
    return { tab: featureTab(saved) ?? "chat", obsolete: Boolean(saved && !featureTab(saved)) }
  } catch { return { tab: "chat", obsolete: false } }
}
export function rememberFeatureTab(tab: FeatureTab, storage?: Pick<Storage, "setItem">): void {
  try { storage?.setItem(FEATURE_TAB_KEY, tab) } catch { /* Host storage can be disabled. */ }
}

/** Ask once at pane startup. Word owns the actual width and subsequent resizing. */
export function preferRoomyPane(office: unknown): void {
  const runtime = office as {
    context?: { requirements?: { isSetSupported(name: string, version: string): boolean } }
    extensionLifeCycle?: { taskpane?: { setWidth(width: number): void } }
  } | undefined
  try {
    if (runtime?.context?.requirements?.isSetSupported("TaskPaneApi", "1.1"))
      runtime.extensionLifeCycle?.taskpane?.setWidth(500)
  } catch { /* Older hosts retain their normal manually resizable pane. */ }
}
