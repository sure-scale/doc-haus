type SelectionHandler = (event: unknown) => void
type SelectionResult = { status: string }

/** The common Office event surface, kept separate from global Office types. */
export type WordChatSelectionOffice = {
  EventType?: { DocumentSelectionChanged?: string }
  context?: {
    document?: {
      addHandlerAsync?(event: string, handler: SelectionHandler, callback: (result: SelectionResult) => void): void
      removeHandlerAsync?(event: string, options: { handler: SelectionHandler }, callback: (result: SelectionResult) => void): void
    }
  }
}

/** Observe native selection changes without reading, editing or polling Word.
 * Content-only changes are checked by the normal capture/apply paths. */
export function watchWordChatSelection(
  onChange: () => void,
  office: WordChatSelectionOffice | undefined = typeof window === "undefined"
    ? undefined
    : (window as unknown as { Office?: WordChatSelectionOffice }).Office,
): () => void {
  const document = office?.context?.document
  if (!document?.addHandlerAsync || !document.removeHandlerAsync) return () => {}
  const event = office?.EventType?.DocumentSelectionChanged ?? "documentSelectionChanged"
  let disposed = false
  let registered = false
  let removed = false
  const handler: SelectionHandler = () => { if (!disposed) onChange() }
  const remove = () => {
    if (!registered || removed) return
    removed = true
    try {
      // Omitting handler would remove other features' selection observers.
      document.removeHandlerAsync!(event, { handler }, () => {})
    } catch { /* A closed host may no longer accept event removal. */ }
  }
  try {
    document.addHandlerAsync(event, handler, result => {
      if (result.status !== "succeeded") return
      registered = true
      if (disposed) remove()
    })
  } catch { /* Startup capture and normal document context remain available. */ }
  return () => {
    if (disposed) return
    disposed = true
    remove()
  }
}
