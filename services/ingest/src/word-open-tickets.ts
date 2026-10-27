import { createHash, randomBytes, randomUUID } from "node:crypto"

export const WORD_OPEN_TTL_MS = 60_000
export const wordOpenPath = (pathname: string) => /^\/word\/open\/[A-Za-z0-9_-]{43}$/.test(pathname)
export const wordOpenResultPath = (pathname: string) => /^\/word\/open\/[A-Za-z0-9_-]{43}\/result$/.test(pathname)
export type WordOpenState = "pending" | "redeemed" | "completed" | "failed" | "expired"
const digest = (nonce: string) => createHash("sha256").update(nonce).digest("hex")

/** Short-lived native read capabilities are separate from document write
 * grants. Only bounded metadata and nonce hashes live in this process. */
export class WordOpenTickets<T> {
  private readonly entries = new Map<string, { value: T; id: string; state: WordOpenState; readExpires: number; expires: number }>()
  constructor(private readonly now = Date.now, private readonly limit = 256) {}

  private prune() {
    for (const [key, item] of this.entries) {
      if (item.expires <= this.now()) this.entries.delete(key)
      else if (item.state === "pending" && item.readExpires <= this.now()) item.state = "expired"
    }
  }

  issue(value: T) {
    this.prune()
    while (this.entries.size >= this.limit) this.entries.delete(this.entries.keys().next().value!)
    const nonce = randomBytes(32).toString("base64url")
    const id = randomUUID()
    this.entries.set(digest(nonce), { value, id, state: "pending", readExpires: this.now() + WORD_OPEN_TTL_MS, expires: this.now() + 2 * WORD_OPEN_TTL_MS })
    return { nonce, id }
  }

  redeem(nonce: string): T | undefined {
    this.prune()
    if (!/^[A-Za-z0-9_-]{43}$/.test(nonce)) return
    const key = digest(nonce), item = this.entries.get(key)
    if (item?.state !== "pending") return
    item.state = "redeemed"
    return item.value
  }

  result(nonce: string, state: "opened" | "failed") {
    this.prune()
    const item = this.entries.get(digest(nonce)), next = state === "opened" ? "completed" : "failed"
    if (!item || !["redeemed", next].includes(item.state)) return false
    item.state = next
    return true
  }

  status(id: string, belongs: (value: T) => boolean): WordOpenState | undefined {
    this.prune()
    for (const item of this.entries.values()) if (item.id === id && belongs(item.value)) return item.state
  }
}
