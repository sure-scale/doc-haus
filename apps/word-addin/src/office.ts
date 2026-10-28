import { WORD_LINK_SETTING } from "../../../dochaus/lib/word-addin-contract"
import { readLink } from "./client"
import { documentFingerprint } from "./fingerprint"

type Result<T> = { status: string; value: T }
type Callback<T> = (result: Result<T>) => void
export type OfficeFile = {
  size: number
  sliceCount: number
  getSliceAsync(index: number, callback: Callback<{ data: number[] }>): void
  closeAsync(callback: Callback<void>): void
}
type WordContext = {
  document: {
    saved: boolean
    changeTrackingMode: "Off" | "TrackAll" | "TrackMineOnly"
    load(properties: string): void
    save(behavior?: "Save"): void
    close(behavior: "Save"): void
    insertFileFromBase64(base64: string, location: "Replace", options: {
      importTheme: boolean
      importStyles: boolean
      importParagraphSpacing: boolean
      importPageColor: boolean
      importChangeTrackingMode: boolean
      importCustomProperties: boolean
      importCustomXmlParts: boolean
      importDifferentOddEvenPages: boolean
    }): unknown
  }
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
  replace?(bytes: Uint8Array, base: string): Promise<void>
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
  const sameIdentity = (value: ReturnType<typeof readLink>) => Boolean(identity && value &&
    [value.apiOrigin, value.origin, value.matter, value.document, value.token].every((part, index) => part === identity[index]))
  let refreshUncertain = false
  const exportCurrent = async () => {
    const file = await asyncResult<OfficeFile>(
      callback => office.context.document.getFileAsync(office.FileType.Compressed, { sliceSize: 1024 * 1024 }, callback),
      "Word could not prepare the document. Try saving again.",
    )
    return exportFile(file)
  }
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
      export: exportCurrent,
      // Full-document import copies sections and their headers/footers. The
      // complete options set needs 1.7 (odd/even-page headers); body/range import
      // is not an equivalent fallback. Word limits this API to a 4 MB DOCX.
      ...(office.context.requirements.isSetSupported("WordApi", "1.7")
        ? {
            replace: async (bytes: Uint8Array, base: string) => {
              if (refreshUncertain)
                throw new DocumentStateError("Word could not confirm the previous refresh. Check this document and open the latest copy from doc.haus before continuing.")
              if (!base || !bytes.length || bytes.length > 4 * 1024 * 1024)
                throw new DocumentStateError("This copy cannot be refreshed here. Open the latest copy from doc.haus instead; this Word document is unchanged.")
              let expected: string
              try { expected = await documentFingerprint(bytes) }
              catch { throw new DocumentStateError("The latest copy could not be verified. Open it from doc.haus instead; this Word document is unchanged.") }
              const settings = office.context.document.settings
              const raw = settings.get(WORD_LINK_SETTING)
              const current = readLink(raw, window.location.origin)
              if (!sameIdentity(current))
                throw new DocumentStateError("This document's link changed. Reopen the latest copy from doc.haus before continuing.")
              let binary = ""
              for (let offset = 0; offset < bytes.length; offset += 0x8000)
                binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
              const encoded = btoa(binary)
              await word.run(async (context) => {
                context.document.load("changeTrackingMode")
                await context.sync()
                if (!sameIdentity(readLink(settings.get(WORD_LINK_SETTING), window.location.origin)))
                  throw new DocumentStateError("This document's link changed. Reopen the latest copy from doc.haus before continuing.")
                const tracking = context.document.changeTrackingMode
                let inserted = false
                let verified = false
                try {
                  // Loading an existing version must not redline the whole old
                  // document. Import the source's tracking mode and revisions.
                  context.document.changeTrackingMode = "Off"
                  await context.sync()
                  refreshUncertain = true
                  context.document.insertFileFromBase64(encoded, "Replace", {
                    importTheme: true,
                    importStyles: true,
                    importParagraphSpacing: true,
                    importPageColor: true,
                    importChangeTrackingMode: true,
                    importCustomProperties: true,
                    importCustomXmlParts: true,
                    importDifferentOddEvenPages: true,
                  })
                  await context.sync()
                  inserted = true
                  // A successful native batch does not establish fidelity.
                  // Compare the downloaded package with Word's current export
                  // before changing the saved link or acknowledging the base.
                  if (await documentFingerprint(await exportCurrent()) !== expected)
                    throw new Error("Imported document differs from the reviewed copy")
                  verified = true
                  // Import does not preserve document settings. Keep the same
                  // scoped capability; only its successfully loaded base moves.
                  const after = settings.get(WORD_LINK_SETTING)
                  if (after != null && !sameIdentity(readLink(after, window.location.origin)))
                    throw new Error("Document link changed during refresh")
                  const updated = { ...current!, baseVersion: base }
                  settings.set(WORD_LINK_SETTING, typeof raw === "string" ? JSON.stringify(updated) : updated)
                  await asyncResult<void>(
                    callback => settings.saveAsync(callback),
                    "The latest contents loaded, but Word could not preserve this document's link. Keep it open and reopen the latest copy from doc.haus before saving.",
                  )
                  refreshUncertain = false
                } catch {
                  if (!inserted) {
                    try {
                      context.document.changeTrackingMode = tracking
                      await context.sync()
                    } catch { /* A failed native batch may leave Word unavailable. */ }
                  }
                  throw new DocumentStateError(inserted
                    ? verified
                      ? "The latest contents loaded, but Word could not preserve this document's link. Keep it open and reopen the latest copy from doc.haus before saving."
                      : "Word loaded the latest copy but could not verify its contents. Check this document and reopen the latest copy from doc.haus before saving."
                    : "Word could not confirm the refresh. Check this document before trying again, or open the latest copy from doc.haus. Linked content, protection or unsupported controls can prevent refreshing here.")
                }
              })
            },
          }
        : {}),
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
                // Settings persistence dirties the document. Save that change
                // explicitly before closing, rather than relying on close to save.
                context.document.save("Save")
                await context.sync()
                context.document.load("saved")
                await context.sync()
                if (!context.document.saved)
                  throw new DocumentStateError("Word has new changes. Choose Save to doc.haus before closing.")
                context.document.close("Save")
                await context.sync()
              }),
          }
        : {}),
    },
  }
}
