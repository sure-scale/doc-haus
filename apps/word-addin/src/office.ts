import { WORD_LINK_SETTING } from "../../../dochaus/lib/word-addin-contract"
import { readLink } from "./client"

type Result<T> = { status: string; value: T }
type Callback<T> = (result: Result<T>) => void
export type OfficeFile = {
  size: number
  sliceCount: number
  getSliceAsync(index: number, callback: Callback<{ data: number[] }>): void
  closeAsync(callback: Callback<void>): void
}
type WordContext = {
  document: { saved: boolean; load(properties: string): void; save(): void; close(behavior: "Save"): void }
  sync(): Promise<void>
}
type OfficeRuntime = {
  onReady(): Promise<{ host: string }>
  context: {
    document: {
      settings: {
        get(key: string): unknown
        set(key: string, value: unknown): void
        saveAsync(callback: Callback<void>): void
      }
      getFileAsync(type: string, options: { sliceSize: number }, callback: Callback<OfficeFile>): void
    }
    requirements: { isSetSupported(name: string, version: string): boolean }
  }
  FileType: { Compressed: string }
}

declare global {
  interface Window {
    Office?: OfficeRuntime
    Word?: { run<T>(callback: (context: WordContext) => Promise<T>): Promise<T> }
  }
}

export interface WordHost {
  saved(): Promise<boolean>
  save(): Promise<void>
  export(): Promise<Uint8Array>
  close?(base: string): Promise<void>
}

export class DocumentStateError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "DocumentStateError"
  }
}

function asyncResult<T>(start: (callback: Callback<T>) => void, message: string): Promise<T> {
  // Keep the host error out of UI/logs: it may include document paths.
  return new Promise((resolve, reject) =>
    start((result) => (result.status === "succeeded" ? resolve(result.value) : reject(new Error(message)))),
  )
}

export async function exportFile(file: OfficeFile): Promise<Uint8Array> {
  try {
    if (!Number.isSafeInteger(file.size) || file.size < 1 || file.size > 128 * 1024 * 1024) {
      throw new Error("This document is too large to send. Keep the Word file and return to doc.haus.")
    }
    const bytes = new Uint8Array(file.size)
    let offset = 0
    for (let index = 0; index < file.sliceCount; index++) {
      const slice = await asyncResult<{ data: number[] }>(
        (callback) => file.getSliceAsync(index, callback),
        "Word could not read the document. Try saving again.",
      )
      if (offset + slice.data.length > bytes.length)
        throw new Error("Word could not read the complete document. Try saving again.")
      bytes.set(slice.data, offset)
      offset += slice.data.length
    }
    if (offset !== bytes.length) throw new Error("Word could not read the complete document. Try saving again.")
    return bytes
  } finally {
    await asyncResult<void>(
      (callback) => file.closeAsync(callback),
      "Word could not finish reading the document. Try saving again.",
    )
  }
}

export async function connectOffice(): Promise<{ host: WordHost; link: unknown }> {
  const office = window.Office
  if (!office) throw new Error("Open this pane in Microsoft Word.")
  const ready = await office.onReady()
  if (ready.host !== "Word" || !window.Word || !office.context.requirements.isSetSupported("WordApi", "1.1")) {
    throw new Error("Use a current desktop version of Microsoft Word to save to doc.haus.")
  }
  const word = window.Word
  const link = office.context.document.settings.get(WORD_LINK_SETTING)
  const initial = readLink(link, window.location.origin)
  const identity = initial && [initial.apiOrigin, initial.origin, initial.matter, initial.document, initial.token]
  return {
    link,
    host: {
      saved: () =>
        word.run(async (context) => {
          context.document.load("saved")
          await context.sync()
          return context.document.saved
        }),
      save: () =>
        word.run(async (context) => {
          context.document.save()
          await context.sync()
          context.document.load("saved")
          await context.sync()
          if (!context.document.saved)
            throw new Error("Word has not saved the document yet. Try Save to doc.haus again.")
        }),
      export: async () => {
        const file = await asyncResult<OfficeFile>(
          (callback) =>
            office.context.document.getFileAsync(office.FileType.Compressed, { sliceSize: 1024 * 1024 }, callback),
          "Word could not prepare the document. Try saving again.",
        )
        return exportFile(file)
      },
      ...(office.context.requirements.isSetSupported("WordApi", "1.5")
        ? {
            close: (base: string) =>
              word.run(async (context) => {
                context.document.load("saved")
                await context.sync()
                if (!context.document.saved)
                  throw new DocumentStateError("Word has new changes. Choose Save to doc.haus before closing.")
                const settings = office.context.document.settings
                const raw = settings.get(WORD_LINK_SETTING)
                const current = readLink(raw, window.location.origin)
                if (
                  !base ||
                  !identity ||
                  !current ||
                  ![current.apiOrigin, current.origin, current.matter, current.document, current.token].every(
                    (value, index) => value === identity[index],
                  )
                )
                  throw new Error("Word could not update this document's link. Close it in Word.")
                const updated = { ...current, baseVersion: base }
                settings.set(WORD_LINK_SETTING, typeof raw === "string" ? JSON.stringify(updated) : updated)
                await asyncResult<void>(
                  (callback) => settings.saveAsync(callback),
                  "Word could not update this document's link. Close it in Word.",
                )
                context.document.close("Save")
                await context.sync()
              }),
          }
        : {}),
    },
  }
}
