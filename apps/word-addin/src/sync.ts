import type { WordDocumentLink, WordSaveResponse } from "../../../dochaus/lib/word-addin-contract"
import { ConnectionError } from "./client"
import { documentFingerprint } from "./fingerprint"
import { DocumentStateError, type WordHost } from "./office"

export type Acknowledgement = {
  base: string
  fingerprint: string
  number: number
  at: number
  overlap: boolean
}
export interface Checkpoints {
  read(): Acknowledgement | null
  write(value: Acknowledgement): void
}
type Transport = {
  save(bytes: Uint8Array, base: string): Promise<WordSaveResponse>
  heartbeat(): Promise<void>
  done(): Promise<void>
}
export type PaneState = {
  status: "not-saved" | "saving" | "saved" | "done"
  busy: boolean
  message: string
  acknowledgement?: Acknowledgement
}
type Pending = { bytes: Uint8Array; fingerprint: string; base: string }

/** Store no document bytes or bearer token. Failure to store is non-fatal. */
export async function browserCheckpoints(link: WordDocumentLink, storage: Storage | undefined): Promise<Checkpoints> {
  const identity = JSON.stringify([link.apiOrigin, link.matter, link.document, link.baseVersion, link.token])
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(identity))
  const key = `dochaus.word.v1.${Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("")}`
  return {
    read() {
      try {
        const value = JSON.parse(storage?.getItem(key) ?? "null") as Acknowledgement | null
        if (
          value &&
          typeof value.base === "string" &&
          /^[a-f0-9]{64}$/.test(value.fingerprint) &&
          Number.isSafeInteger(value.number) &&
          Number.isFinite(value.at) &&
          typeof value.overlap === "boolean"
        )
          return value
      } catch {
        /* A cleared or unavailable store is an ordinary fresh open. */
      }
      return null
    },
    write(value) {
      try {
        storage?.setItem(key, JSON.stringify(value))
      } catch {
        /* Reopening can recover the saved file. */
      }
    },
  }
}

export class WordSync {
  private queue: Promise<unknown> = Promise.resolve()
  private queued = 0
  private completed = false
  private completedCheckedAt = -Infinity
  private captured = false
  private base: string
  private pending?: Pending
  private acknowledgement?: Acknowledgement
  private nextRetry = 0
  private heartbeatAt = -Infinity
  private state: PaneState = { status: "not-saved", busy: false, message: "Checking the saved Word document…" }

  constructor(
    base: string,
    private readonly host: WordHost,
    private readonly transport: Transport,
    private readonly checkpoints: Checkpoints,
    private readonly render: (state: PaneState) => void,
    private readonly fingerprint = documentFingerprint,
    private readonly wait = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)),
    private readonly now = Date.now,
  ) {
    this.base = base
  }

  snapshot(): PaneState {
    return { ...this.state }
  }

  private show(status: PaneState["status"], message: string): void {
    this.state = { status, message, busy: this.queued > 0, acknowledgement: this.acknowledgement }
    this.render(this.snapshot())
  }

  private enqueue(task: () => Promise<void>): Promise<boolean> {
    this.queued++
    this.render({ ...this.state, busy: true })
    const result = this.queue.then(async () => {
      try {
        await task()
        return true
      } catch (error) {
        this.nextRetry = this.now() + 30_000
        const message =
          error instanceof ConnectionError
            ? error.message
            : error instanceof Error && error.name === "DocumentStateError"
              ? error.message
              : "Word could not finish the save. Keep this document open and try Save to doc.haus again."
        this.show("not-saved", message)
        return false
      } finally {
        this.queued--
        this.state = { ...this.state, busy: this.queued > 0 }
        this.render(this.snapshot())
      }
    })
    this.queue = result
    return result
  }

  private async heartbeat(force = false): Promise<void> {
    if (!force && this.now() - this.heartbeatAt < 20_000) return
    this.heartbeatAt = this.now()
    await this.transport.heartbeat()
  }

  start(): Promise<boolean> {
    return this.enqueue(async () => {
      await this.heartbeat()
      await this.capture()
    })
  }

  /** Poll saved state and content: lastSaveTime is minute-rounded on some hosts. */
  tick(): Promise<boolean> {
    if (this.queued > 0) return Promise.resolve(false)
    return this.enqueue(async () => {
      if (this.completed) {
        // Keep a cheap saved-state watch after Done. Check content less often
        // too: an edit followed by native Save can occur between saved polls.
        const saved = await this.host.saved()
        if (saved && this.now() - this.completedCheckedAt < 20_000) return
        if (saved && (await this.matchesAcknowledgement())) {
          this.completedCheckedAt = this.now()
          return
        }
        this.completed = false
        this.show("not-saved", "You have new changes in Word. Choose Save to doc.haus before closing.")
        await this.heartbeat(true)
      }
      await this.heartbeat()
      if (this.now() < this.nextRetry) return
      await this.flush()
      await this.capture()
    })
  }

  save(): Promise<boolean> {
    return this.enqueue(() => this.explicitSave())
  }

  done(): Promise<boolean> {
    return this.enqueue(async () => {
      await this.explicitSave()
      this.show("saving", "Finishing in Word…")
      try {
        // Clearing the remote marker can take time. Changes during that request
        // must restore the marker before allowing any completion message.
        await this.transport.done()
        if (!(await this.matchesAcknowledgement())) {
          throw documentStateError("Word has new changes. Choose Done in Word again before closing.")
        }
      } catch (error) {
        // A lost response may still have cleared the remote marker.
        this.completed = false
        try {
          await this.heartbeat(true)
        } catch {
          /* Retry contact on the next poll. */
        }
        throw error
      }
      this.completed = true
      this.completedCheckedAt = this.now()
      const manualClose = "Your document is saved to doc.haus. Close it in Word."
      if (!this.host.close) {
        this.show("done", manualClose)
        return
      }
      this.show("saving", "Your document is saved to doc.haus. Closing this document…")
      try {
        await this.host.close()
        this.show("done", "Your document is saved to doc.haus.")
      } catch (error) {
        let unchanged: boolean | undefined
        if (error instanceof DocumentStateError) unchanged = false
        else {
          try {
            unchanged = await this.matchesAcknowledgement()
          } catch {
            /* A close error can also make the remaining Word context unavailable. */
          }
        }
        if (unchanged) {
          this.show("done", manualClose)
          return
        }
        this.completed = false
        const message =
          unchanged === false
            ? "Your last save is in doc.haus. Word has new changes. Choose Save to doc.haus before closing."
            : "Your last save is in doc.haus. Check this document for new changes in Word before closing it."
        this.show("not-saved", message)
        try {
          await this.heartbeat(true)
        } catch {
          /* Keep the new-change watch live even when contact needs retrying. */
        }
        throw documentStateError(message)
      }
    })
  }

  private async explicitSave(): Promise<void> {
    this.completed = false
    this.show("saving", "Saving in Word…")
    await this.host.save()
    await this.flush()
    if (!(await this.capture(true)))
      throw documentStateError("Word has new changes. Try Save to doc.haus again before closing.")
  }

  private async capture(requireAcknowledgement = false): Promise<boolean> {
    if (!(await this.host.saved())) {
      this.show("not-saved", "You have changes in Word. Choose Save to doc.haus before closing.")
      return false
    }
    const bytes = await this.host.export()
    if (!(await this.host.saved())) {
      this.show("not-saved", "You have new changes in Word. Choose Save to doc.haus before closing.")
      return false
    }
    const fingerprint = await this.fingerprint(bytes)
    if (!(await this.host.saved())) {
      this.show("not-saved", "You have new changes in Word. Choose Save to doc.haus before closing.")
      return false
    }
    if (!this.captured) {
      this.captured = true
      const previous = this.checkpoints.read()
      // A forwarded or separately edited physical copy may share the token.
      // Reuse an acknowledgement only when this saved file actually matches it.
      if (previous?.fingerprint === fingerprint) {
        this.acknowledgement = previous
        this.base = previous.base
      }
    }
    if (!requireAcknowledgement && fingerprint === this.acknowledgement?.fingerprint) {
      this.showSaved()
      return true
    }
    this.pending = { bytes, fingerprint, base: this.base }
    await this.flush()
    if (!(await this.matchesAcknowledgement())) {
      this.show("not-saved", "You have new changes in Word. Choose Save to doc.haus before closing.")
      return false
    }
    this.showSaved()
    return true
  }

  private async matchesAcknowledgement(): Promise<boolean> {
    if (!(await this.host.saved())) return false
    const bytes = await this.host.export()
    const matches = (await this.fingerprint(bytes)) === this.acknowledgement?.fingerprint
    return matches && (await this.host.saved())
  }

  private async flush(): Promise<void> {
    const pending = this.pending
    if (!pending) return
    this.show("saving", "Saving to doc.haus…")
    let response: WordSaveResponse | undefined
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        response = await this.transport.save(pending.bytes, pending.base)
        break
      } catch (error) {
        if (!(error instanceof ConnectionError) || !error.retryable || attempt === 1) throw error
        await this.wait(650)
      }
    }
    if (!response) throw new ConnectionError("doc.haus did not confirm the save. Try Save to doc.haus again.")
    this.base = response.version.id
    this.acknowledgement = {
      base: this.base,
      fingerprint: pending.fingerprint,
      number: response.version.number,
      at: this.now(),
      overlap: response.overlap,
    }
    this.pending = undefined
    this.nextRetry = 0
    this.checkpoints.write(this.acknowledgement)
  }

  private showSaved(): void {
    this.show(
      "saved",
      this.acknowledgement?.overlap
        ? "Saved as a separate version. The document also changed in doc.haus; compare the versions there."
        : "The saved Word document is in doc.haus.",
    )
  }
}

function documentStateError(message: string): Error {
  return new DocumentStateError(message)
}
