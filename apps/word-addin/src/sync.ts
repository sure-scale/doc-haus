import type {
  WordConnectionState,
  WordLeaseAction,
  WordDocumentLink,
  WordSaveResponse,
} from "../../../dochaus/lib/word-addin-contract"
import { ConnectionError, type LatestWordContent } from "./client"
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
  clear?(): void
  rebase?(base: string): Promise<void>
}
export type Transport = {
  lease(
    action: WordLeaseAction,
    base: string,
    generation?: string,
    expected?: { current: string; generation: string },
  ): Promise<WordConnectionState>
  save(bytes: Uint8Array, base: string, generation: string): Promise<WordSaveResponse>
  heartbeat(base: string, generation: string): Promise<WordConnectionState>
  done(base: string, generation: string): Promise<WordConnectionState>
}
export type PaneState = {
  status: "not-saved" | "saving" | "saved" | "done"
  busy: boolean
  message: string
  acknowledgement?: Acknowledgement
  connection?: WordConnectionState
  canSave: boolean
  /** Presentation only: version equality does not verify the current Word text. */
  content: "unverified" | "changed" | "confirmed"
  problem?: boolean
  connectionFailed?: boolean
}
type Pending = { bytes: Uint8Array; fingerprint: string; base: string; generation: string }

/** Store no document bytes or bearer token. Failure to store is non-fatal. */
export async function browserCheckpoints(link: WordDocumentLink, storage: Storage | undefined): Promise<Checkpoints> {
  const keyFor = async (base: string) => {
    const identity = JSON.stringify([link.apiOrigin, link.matter, link.document, base, link.token])
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(identity))
    return `dochaus.word.v1.${Array.from(new Uint8Array(digest), value => value.toString(16).padStart(2, "0")).join("")}`
  }
  let key = await keyFor(link.baseVersion)
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
    clear() {
      try { storage?.removeItem(key) } catch { /* Storage may be unavailable. */ }
    },
    async rebase(base) {
      const next = await keyFor(base)
      try {
        storage?.removeItem(key)
        storage?.removeItem(next)
      } catch { /* Storage may be unavailable. */ }
      key = next
    },
  }
}

export class WordSync {
  private queue: Promise<unknown> = Promise.resolve()
  private queued = 0
  private foreground = 0
  private completed = false
  private completedCheckedAt = -Infinity
  private checkpointChecked = false
  private base: string
  private connection?: WordConnectionState
  private pending?: Pending
  private acknowledgement?: Acknowledgement
  // Add-in mutations are local until a deliberate pane Save/Done. Native
  // AutoSave or Undo can report saved=true without granting publication.
  private explicitSaveRequired = false
  private latestUncertain = false
  private nextRetry = 0
  private heartbeatAt = -Infinity
  private content: PaneState["content"] = "unverified"
  private state: PaneState = {
    status: "not-saved",
    busy: false,
    canSave: false,
    message: "Checking the saved Word document…",
    content: "unverified",
  }

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

  private show(status: PaneState["status"], message: string, problem = false, connectionFailed = false): void {
    this.state = {
      status,
      message,
      busy: this.foreground > 0 || status === "saving",
      acknowledgement: this.acknowledgement,
      connection: this.connection,
      canSave: this.canSave(),
      content: this.content,
      problem,
      ...(connectionFailed ? { connectionFailed: true } : {}),
    }
    this.render(this.snapshot())
  }

  private enqueue(task: () => Promise<void>, quiet = false): Promise<boolean> {
    this.queued++
    if (!quiet) {
      this.foreground++
      this.state = { ...this.state, busy: true }
      this.render(this.snapshot())
    }
    const result = this.queue.then(async () => {
      try {
        await task()
        return true
      } catch (error) {
        if (error instanceof ConnectionError) this.connection = error.connection
        this.nextRetry = this.now() + 30_000
        const message =
          error instanceof ConnectionError
            ? error.message
            : error instanceof Error && error.name === "DocumentStateError"
              ? error.message
              : "Word could not finish the save. Keep this document open and try Save again."
        this.show(
          "not-saved",
          error instanceof ConnectionError && this.connection && !this.canWrite() ? this.blockedMessage() : message,
          true,
          error instanceof ConnectionError && error.retryable,
        )
        return false
      } finally {
        this.queued--
        if (!quiet) this.foreground--
        this.state = { ...this.state, busy: this.foreground > 0 }
        this.render(this.snapshot())
      }
    })
    this.queue = result
    return result
  }

  private async heartbeat(force = false): Promise<void> {
    if (!this.canWrite() || (!force && this.now() - this.heartbeatAt < 20_000)) return
    this.heartbeatAt = this.now()
    this.connection = await this.transport.heartbeat(this.base, this.connection!.lease.generation!)
  }

  private canSave(): boolean {
    return (
      !this.latestUncertain &&
      this.connection?.version.state === "current" && (this.connection.lease.state === "available" || this.canWrite())
    )
  }

  private canWrite(): boolean {
    return (
      !this.latestUncertain &&
      this.connection?.version.state === "current" &&
      this.connection.lease.state === "owned" &&
      Boolean(this.connection.lease.generation)
    )
  }

  private blockedMessage(): string {
    if (this.latestUncertain)
      return "Word could not confirm the refresh. Check this document and open the latest copy from doc.haus before saving or applying changes."
    const state = this.connection
    if (!state) return "Check the connection before saving. Your Word changes stay in this file."
    if (state.version.state === "removed")
      return "This copy predates a redaction or removal. Keep your local file and reopen the current document from doc.haus."
    if (state.version.state === "unrecorded-change")
      return "The file in doc.haus changed outside its version history. Keep your Word file and resolve that change in doc.haus before saving."
    if (state.version.state === "stale")
      return "A newer version is in doc.haus. Keep this Word file using Save As. In doc.haus, use Versions to compare or Open in Word for a fresh copy. This pane will not overwrite it."
    if (state.lease.state === "other")
      return "Another editing session owns saving. Keep your Word edits here. Wait for that session to finish, or take over editing here."
    return "Ready to save. Choose Save and wait for Saved before closing Word."
  }

  private async inspect(claim: boolean): Promise<void> {
    this.connection = await this.transport.lease("inspect", this.base)
    if (claim && !this.latestUncertain && this.connection.version.state === "current" && this.connection.lease.state === "available")
      this.connection = await this.transport.lease("claim", this.base)
  }

  refresh(): Promise<boolean> {
    const claim = this.connection?.version.state === "current" && this.connection.lease.state === "available"
    return this.enqueue(async () => {
      this.completed = false
      await this.restoreCheckpoint()
      if (this.pending) await this.flush()
      await this.inspect(claim)
      if (!this.canWrite()) {
        this.show("not-saved", this.blockedMessage())
        return
      }
      await this.capture()
    })
  }

  /** Inspect only; this command never claims, exports or retries a save. */
  checkStatus(): Promise<boolean> {
    return this.enqueue(async () => {
      await this.inspect(false)
      const confirmed = this.canSave() && (this.state.status === "saved" || this.state.status === "done")
      this.show(
        confirmed ? this.state.status : "not-saved",
        confirmed ? this.state.message : this.canSave() ? "Editing status checked." : this.blockedMessage(),
      )
    })
  }

  /** Replace this open Word document; this command never claims or publishes. */
  async loadLatest(
    load: (base: string) => Promise<LatestWordContent>,
    confirm: (version: LatestWordContent["version"]) => boolean | Promise<boolean>,
  ): Promise<boolean> {
    let loaded = false
    await this.enqueue(async () => {
      if (this.latestUncertain) throw documentStateError(this.blockedMessage())
      if (this.pending) throw documentStateError("Finish the pending Save before loading the latest version.")
      if (!this.host.replace)
        throw documentStateError("This version of Word cannot refresh the whole document here. Open the latest copy from doc.haus instead; this Word document is unchanged.")
      await this.inspect(false)
      if (this.connection?.version.state !== "stale" || !this.connection.version.base || !this.connection.version.current)
        throw documentStateError(this.connection?.version.state === "current" ? "This copy is already on the latest version." : this.blockedMessage())
      const oldBase = this.base
      const before = await this.fingerprint(await this.host.export())
      const latest = await load(oldBase)
      if (!(await confirm(latest.version))) {
        this.show("not-saved", "Loading cancelled. Your Word contents are unchanged.")
        return
      }
      // Download/confirmation can take time. Check both the recorded version
      // and the complete live Word package again before destructive import.
      const current = await this.transport.lease("inspect", latest.version.id)
      if (current.version.state !== "current" || current.version.base?.id !== latest.version.id || current.version.current?.id !== latest.version.id)
        throw documentStateError("The latest version changed while loading. Review it and choose Load latest version again.")
      if ((await this.fingerprint(await this.host.export())) !== before)
        throw documentStateError("This Word document changed while loading. Your changes are still here. Choose Load latest version again if you want to replace them.")
      const previousOwner = current.lease.state === "owned" ? current.lease.generation : undefined
      this.show("not-saved", "Loading the latest contents into this Word document…")
      // Native import/settings failures may happen after Word has changed. Do
      // not let a later native Save automatically publish that uncertain buffer.
      this.explicitSaveRequired = true
      this.latestUncertain = true
      await this.host.replace(latest.bytes, latest.version.id)
      this.latestUncertain = false
      this.base = latest.version.id
      this.acknowledgement = undefined
      this.pending = undefined
      try { this.checkpoints.clear?.() } catch { /* Storage failure is non-fatal. */ }
      try { await this.checkpoints.rebase?.(this.base) } catch { /* Reopening can recover without a checkpoint. */ }
      this.checkpointChecked = true
      this.completed = false
      this.nextRetry = 0
      this.content = "changed"
      loaded = true
      if (previousOwner) {
        try {
          // The release is fenced to this session's old generation. It cannot
          // release another editor that acquired ownership during the import.
          await this.transport.lease("release", oldBase, previousOwner)
        } catch { /* Inspect below; the imported document already exists. */ }
      }
      const after = await this.transport.lease("inspect", this.base)
      this.connection = after
      this.show("not-saved", after.version.state === "current"
        ? "Latest version loaded in Word. Choose Save when you’re ready."
        : "The version loaded in Word, and doc.haus has changed again. Check the latest version before saving.")
    })
    // Replacement invalidates the previous chat context even if the final
    // connection check failed; enqueue already reports that failure in the pane.
    return loaded
  }

  /** Deliberate ownership claim only; current text still needs confirmation. */
  startEditing(): Promise<boolean> {
    return this.enqueue(async () => {
      this.completed = false
      await this.inspect(true)
      this.discardFencedPending()
      this.show("not-saved", this.canWrite() ? "Editing started. Choose Save to confirm this copy in doc.haus." : this.blockedMessage())
    })
  }

  takeover(): Promise<boolean> {
    const observed = this.connection
    return this.enqueue(async () => {
      if (this.latestUncertain) throw documentStateError(this.blockedMessage())
      this.completed = false
      if (
        !observed ||
        observed.version.state !== "current" ||
        observed.lease.state !== "other" ||
        !observed.version.current ||
        !observed.lease.generation
      )
        throw documentStateError("Check editing status before taking over.")
      this.connection = await this.transport.lease("takeover", this.base, undefined, {
        current: observed.version.current.id,
        generation: observed.lease.generation,
      })
      this.discardFencedPending()
      this.show(
        "not-saved",
        this.canWrite()
          ? "You now own saving. Your local edits have not been sent; choose Save."
          : this.blockedMessage(),
      )
    })
  }

  release(): Promise<boolean> {
    return this.enqueue(async () => {
      if (this.connection?.lease.state !== "owned" || !this.connection.lease.generation)
        throw documentStateError("This pane does not own the editing session.")
      this.connection = await this.transport.lease("release", this.base, this.connection.lease.generation)
      this.show("not-saved", "Saving released. Your local Word edits remain in this file.")
    })
  }

  start(): Promise<boolean> {
    return this.enqueue(async () => {
      await this.restoreCheckpoint()
      // Opening/reconstructing a pane is read-only. It must not create a
      // version or compete for ownership before an explicit saving action.
      await this.inspect(false)
      this.show(
        "not-saved",
        this.canWrite() ? "Ready to save. Choose Save to send this document." : this.blockedMessage(),
      )
    })
  }

  /** Poll saved state and content without flashing the controls for read-only
   * checks. The same queue still serializes exports and explicit commands;
   * flush() visibly enters Saving whenever there are actual bytes to upload.
   * lastSaveTime is minute-rounded on some hosts, so content checks remain. */
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
        this.content = "changed"
        this.show("not-saved", "You have new changes in Word. Choose Save before closing.")
        await this.heartbeat(true)
      }
      if (this.pending && this.now() >= this.nextRetry) await this.flush()
      await this.inspect(false)
      if (!this.canWrite()) {
        this.show("not-saved", this.blockedMessage())
        return
      }
      await this.heartbeat()
      if (this.now() < this.nextRetry) return
      await this.capture()
    }, true)
  }

  save(): Promise<boolean> {
    return this.enqueue(() => this.explicitSave())
  }

  /** Reviewed changes share Save's queue but never export, save or publish.
   * The callback receives the live identity inspected inside that queue. */
  async localEdit<T>(edit: (identity: { base: string; generation: string }) => Promise<T>): Promise<T> {
    let value!: T
    let rejected: unknown
    const complete = await this.enqueue(async () => {
      try {
        if (this.pending) throw documentStateError("Finish the pending Save before applying a change.")
        await this.inspect(true)
        if (!this.canWrite()) throw documentStateError(this.blockedMessage())
        this.completed = false
        // Establish the fence before calling Word: a lost response can mean
        // the edit happened, and Undo may restore Word's saved flag later.
        this.explicitSaveRequired = true
        this.content = "changed"
        this.show("not-saved", "Applying the reviewed change in Word…")
        value = await edit({ base: this.base, generation: this.connection!.lease.generation! })
        this.content = "changed"
        this.show("not-saved", "Applied in Word. Choose Save when you’re ready.")
      } catch (error) {
        rejected = error
        if (error instanceof ConnectionError || error instanceof DocumentStateError) throw error
        throw documentStateError(error instanceof Error ? error.message : "Word could not apply this change. Review it again.")
      }
    })
    if (!complete) throw rejected instanceof Error ? rejected : documentStateError(this.state.message)
    return value
  }

  done(): Promise<boolean> {
    return this.enqueue(async () => {
      await this.explicitSave()
      this.show("saving", "Finishing in Word…")
      try {
        // Releasing editing can take time. Changes during that request must
        // keep Word open; release never grants a later automatic claim.
        this.connection = await this.transport.done(this.base, this.connection!.lease.generation!)
        if (!(await this.matchesAcknowledgement())) {
          this.content = "changed"
          throw documentStateError("Word has new changes. Choose Save and wait for Saved before closing.")
        }
      } catch (error) {
        // A lost response may still have released editing. A heartbeat can
        // renew only a lease that this session still owns.
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
        await this.host.close(this.acknowledgement!.base)
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
        if (unchanged === false) this.content = "changed"
        const message =
          unchanged === false
            ? "Your last save is in doc.haus. Word has new changes. Choose Save before closing."
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
    if (this.latestUncertain) throw documentStateError(this.blockedMessage())
    this.completed = false
    this.show("saving", "Saving in Word…")
    await this.restoreCheckpoint(true)
    // Retry already-sent bytes with their original fence before inspecting a
    // possibly advanced server base; only the server may confirm a lost ACK.
    try {
      await this.flush()
    } catch (error) {
      const state = error instanceof ConnectionError && !error.retryable ? error.connection : undefined
      if (
        !this.pending ||
        state?.version.state !== "current" ||
        state.lease.state !== "available" ||
        state.version.current?.id !== this.pending.base ||
        this.base !== this.pending.base
      )
        throw error
      // The refused generation cannot be reused. Only explicit saving may
      // reclaim an unchanged base; retain pending bytes until that claim wins.
      this.connection = state
      await this.inspect(true)
      if (!this.discardFencedPending()) throw documentStateError(this.blockedMessage())
    }
    await this.inspect(true)
    if (!this.canWrite()) throw documentStateError(this.blockedMessage())
    await this.host.save()
    if (!(await this.capture(true)))
      throw documentStateError("Word has new changes. Try Save again before closing.")
  }

  private discardFencedPending(): boolean {
    const pending = this.pending
    if (
      !pending ||
      !this.canWrite() ||
      this.base !== pending.base ||
      this.connection!.version.current?.id !== pending.base ||
      this.connection!.lease.generation === pending.generation
    )
      return false
    // No version advanced, and a deliberate claim/takeover has a new fence.
    // The local Word document is untouched; export it anew rather than replay
    // an unacknowledged snapshot under a different generation.
    this.pending = undefined
    return true
  }

  private async capture(requireAcknowledgement = false): Promise<boolean> {
    if (!this.canWrite()) {
      this.show("not-saved", this.blockedMessage())
      return false
    }
    if (this.explicitSaveRequired && !requireAcknowledgement) {
      this.content = "changed"
      this.show("not-saved", "Your Word changes are local. Choose Save when you’re ready.")
      return false
    }
    await this.restoreCheckpoint(requireAcknowledgement)
    if (!(await this.host.saved())) {
      this.content = "changed"
      this.show("not-saved", "You have changes in Word. Choose Save before closing.")
      return false
    }
    const bytes = await this.host.export()
    if (!(await this.host.saved())) {
      this.content = "changed"
      this.show("not-saved", "You have new changes in Word. Choose Save before closing.")
      return false
    }
    const fingerprint = await this.fingerprint(bytes)
    if (!(await this.host.saved())) {
      this.content = "changed"
      this.show("not-saved", "You have new changes in Word. Choose Save before closing.")
      return false
    }
    if (!requireAcknowledgement && fingerprint === this.acknowledgement?.fingerprint) {
      this.showSaved()
      return true
    }
    if (!this.canWrite()) {
      this.show("not-saved", this.blockedMessage())
      return false
    }
    this.pending = { bytes, fingerprint, base: this.base, generation: this.connection!.lease.generation! }
    await this.flush()
    if (!this.canWrite()) {
      this.show("not-saved", this.blockedMessage())
      return false
    }
    if (!(await this.matchesAcknowledgement())) {
      this.content = "changed"
      this.show("not-saved", "You have new changes in Word. Choose Save before closing.")
      return false
    }
    this.showSaved()
    return true
  }

  private async restoreCheckpoint(allowExport = false): Promise<void> {
    if (this.checkpointChecked || this.explicitSaveRequired && !allowExport) return
    const previous = this.checkpoints.read()
    if (previous) {
      // Probe before editing/saving can replace an initially matching snapshot.
      // A dirty native flag is not an upload permission or a new acknowledgement.
      const fingerprint = await this.fingerprint(await this.host.export())
      if (previous.fingerprint === fingerprint) {
        this.acknowledgement = previous
        this.base = previous.base
      }
    }
    this.checkpointChecked = true
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
        response = await this.transport.save(pending.bytes, pending.base, pending.generation)
        break
      } catch (error) {
        if (!(error instanceof ConnectionError) || !error.retryable || attempt === 1) throw error
        await this.wait(650)
      }
    }
    if (!response) throw new ConnectionError("doc.haus did not confirm the save. Try Save again.")
    this.base = response.version.id
    this.connection = response.connection
    this.acknowledgement = {
      base: this.base,
      fingerprint: pending.fingerprint,
      number: response.version.number,
      at: this.now(),
      overlap: response.overlap,
    }
    this.pending = undefined
    this.explicitSaveRequired = false
    this.nextRetry = 0
    this.checkpoints.write(this.acknowledgement)
  }

  private showSaved(): void {
    this.content = "confirmed"
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
