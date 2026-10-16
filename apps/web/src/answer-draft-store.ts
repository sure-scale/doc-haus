export type DraftAnswer = string | Record<string, string | number | boolean | readonly string[]>
export type DraftAnswers = Record<string, DraftAnswer>
export type DraftSnapshot = { answers: DraftAnswers; loaded: boolean; saving: boolean; error: string }
type DraftResponse = { answers: DraftAnswers; revision: number }

// One queue per scope survives dialog unmounts. An in-flight save always finishes
// before the next edit (including clearing a submitted answer) is written.
export class AnswerDraftStore {
  private snapshot: DraftSnapshot = { answers: {}, loaded: false, saving: false, error: "" }
  private listeners = new Set<() => void>()
  private revision = 0
  private pending = false
  private running = false
  private loading = false
  constructor(
    private request: (input?: DraftResponse) => Promise<DraftResponse>,
    private legacy: () => DraftAnswers = () => ({}),
    private migrated: () => void = () => {},
  ) {}

  getSnapshot = () => this.snapshot
  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  private update(patch: Partial<DraftSnapshot>) {
    this.snapshot = { ...this.snapshot, ...patch }
    this.listeners.forEach((listener) => listener())
  }

  load = async () => {
    if (this.snapshot.loaded || this.loading || this.running || this.pending) return
    this.loading = true
    try {
      const response = await this.request()
      this.revision = response.revision
      const legacy = this.legacy()
      // Server entries take precedence; existing browser work is removed only
      // after the merged answer set has been acknowledged by SQLite.
      const answers = { ...legacy, ...response.answers }
      this.update({ answers, loaded: true, error: "" })
      if (Object.keys(legacy).length) {
        this.pending = true
        await this.save()
      } else this.migrated()
    } catch {
      this.update({ error: "Could not load saved answers. Retry before editing." })
    } finally {
      this.loading = false
    }
  }

  set = (id: string, answer: DraftAnswer | undefined) => {
    if (!this.snapshot.loaded) return
    this.update({
      answers:
        answer === undefined
          ? Object.fromEntries(Object.entries(this.snapshot.answers).filter(([key]) => key !== id))
          : { ...this.snapshot.answers, [id]: answer },
      error: "",
    })
    this.pending = true
    void this.save()
  }

  retry = () => (this.snapshot.loaded ? this.save() : this.load())

  private async save() {
    if (this.running || !this.pending) return
    this.running = true
    this.update({ saving: true })
    try {
      while (this.pending) {
        this.pending = false
        const response = await this.request({ answers: this.snapshot.answers, revision: this.revision })
        this.revision = response.revision
      }
      this.migrated()
      this.update({ error: "" })
    } catch (error) {
      this.pending = true
      this.update({
        error: error instanceof Error ? error.message : "Could not save answers. Keep this page open and retry.",
      })
    } finally {
      this.running = false
      this.update({ saving: false })
    }
  }
}
