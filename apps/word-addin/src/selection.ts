/** A transient view of Word's current body, never a saved matter version. */
export type WordSelectionSnapshot = {
  text: string
  documentHash: string
  saved: boolean
  capturedAt: string
  navigable: boolean
  reason?: string
}

type Range = { text: string; load(properties: string): void; select(): void }
type SelectionContext = {
  document: {
    saved: boolean
    load(properties: string): void
    getSelection(): Range
    body: {
      text: string
      load(properties: string): void
      search(
        text: string,
        options: {
          matchCase: boolean
          matchWildcards: boolean
          matchWholeWord: boolean
          ignorePunct: boolean
          ignoreSpace: boolean
        },
      ): { items: Range[]; load(properties: string): void }
    }
  }
  sync(): Promise<void>
}

export type SelectionWord = { run<T>(callback: (context: SelectionContext) => Promise<T>): Promise<T> }

export class SelectionError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "SelectionError"
  }
}

const LIMIT = 10_000
const BODY_LIMIT = 16 * 1024 * 1024

function wordRuntime(): SelectionWord {
  const word = (window as unknown as { Word?: SelectionWord }).Word
  if (!word) throw new SelectionError("Open the Source Desk from this document in Word.")
  return word
}

async function hash(text: string): Promise<string> {
  if (text.length > BODY_LIMIT)
    throw new SelectionError("This document is too large to navigate safely. Select the passage directly in Word.")
  // Office can serialize different OOXML for the same unchanged body. Passage
  // navigation identifies exact text; saved-source provenance still hashes bytes.
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text))
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("")
}

function unsupported(text: string): string | undefined {
  if (!text.trim()) return "Select a passage in Word, then refresh the selection."
  if (text.length > 255 || /[\u0000-\u001f\u007f^]/.test(text))
    return "For exact navigation, select a single-line passage of up to 255 characters without Word find codes."
}

function occurrences(body: string, text: string): number {
  let count = 0
  let at = body.indexOf(text)
  while (at !== -1 && count < 2) {
    count++
    at = body.indexOf(text, at + 1)
  }
  return count
}

export async function captureSelection(word = wordRuntime()): Promise<WordSelectionSnapshot> {
  try {
    return await word.run(async (context) => {
      const document = context.document
      const selected = document.getSelection()
      selected.load("text")
      document.load("saved")
      document.body.load("text")
      await context.sync()
      const text = selected.text.slice(0, LIMIT)
      let reason =
        selected.text.length > LIMIT
          ? "This selection is too long. Select a shorter passage in Word."
          : unsupported(text)
      if (!reason) {
        const count = occurrences(document.body.text, text)
        if (count !== 1)
          reason =
            count > 1
              ? "This passage appears more than once. Select a longer unique passage in Word."
              : "This selection is outside the document body. Select the passage directly in Word."
      }
      return {
        text,
        documentHash: await hash(document.body.text),
        saved: document.saved,
        capturedAt: new Date().toISOString(),
        navigable: !reason,
        ...(reason ? { reason } : {}),
      }
    })
  } catch (error) {
    if (error instanceof SelectionError) throw error
    throw new SelectionError("Word could not read the selection. Select the passage and refresh it in the desk.")
  }
}

export async function showSelection(snapshot: WordSelectionSnapshot, word = wordRuntime()): Promise<void> {
  if (!snapshot.navigable || unsupported(snapshot.text) || !/^[a-f0-9]{64}$/.test(snapshot.documentHash))
    throw new SelectionError(snapshot.reason || "Refresh a unique passage in Word before navigating to it.")
  try {
    await word.run(async (context) => {
      const body = context.document.body
      body.load("text")
      const matches = body.search(snapshot.text, {
        matchCase: true,
        matchWildcards: false,
        matchWholeWord: false,
        ignorePunct: false,
        ignoreSpace: false,
      })
      matches.load("items/text")
      await context.sync()
      if ((await hash(body.text)) !== snapshot.documentHash)
        throw new SelectionError("The Word text has changed. Refresh the selection before navigating.")
      if (
        occurrences(body.text, snapshot.text) !== 1 ||
        matches.items.length !== 1 ||
        matches.items[0]?.text !== snapshot.text
      )
        throw new SelectionError(
          "Word could not identify one exact passage. Select it in Word and refresh the selection.",
        )
      // Only this explicit action selects text. Capturing/refreshing never changes
      // selection, formatting, settings, saved bytes or doc.haus editing ownership.
      matches.items[0].select()
      await context.sync()
    })
  } catch (error) {
    if (error instanceof SelectionError) throw error
    throw new SelectionError("Word could not show this passage. Select it directly in Word and refresh the selection.")
  }
}
