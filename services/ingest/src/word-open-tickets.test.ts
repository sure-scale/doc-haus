import { expect, test } from "bun:test"
import { WordOpenTickets, WORD_OPEN_TTL_MS, wordOpenPath } from "./word-open-tickets"

test("native read tickets are one-use, expire, and evict oldest metadata at the bound", () => {
  let now = 1000
  const tickets = new WordOpenTickets<string>(() => now, 2)
  const { nonce: once, id } = tickets.issue("once")
  expect(tickets.status(id, value => value === "once")).toBe("pending")
  expect(tickets.result(once, "opened")).toBe(false)
  expect(wordOpenPath(`/word/open/${once}`)).toBe(true)
  expect(tickets.redeem(once)).toBe("once")
  expect(tickets.redeem(once)).toBeUndefined()
  expect(tickets.status(id, value => value === "once")).toBe("redeemed")
  expect(tickets.result(once, "opened")).toBe(true)
  expect(tickets.status(id, value => value === "once")).toBe("completed")
  expect(tickets.result(once, "opened")).toBe(true)
  expect(tickets.result(once, "failed")).toBe(false)
  expect(tickets.status(id, () => false)).toBeUndefined()
  const { nonce: expired, id: expiredId } = tickets.issue("expired")
  now += WORD_OPEN_TTL_MS
  expect(tickets.redeem(expired)).toBeUndefined()
  expect(tickets.status(expiredId, () => true)).toBe("expired")
  const { nonce: evicted } = tickets.issue("evicted"), { nonce: retained } = tickets.issue("retained"), { nonce: newest } = tickets.issue("newest")
  expect(tickets.redeem(evicted)).toBeUndefined()
  expect(tickets.redeem(retained)).toBe("retained")
  expect(tickets.redeem(newest)).toBe("newest")
  expect(wordOpenPath(`/word/open/${once}/extra`)).toBe(false)
  expect(tickets.redeem("invalid")).toBeUndefined()
  now += 2 * WORD_OPEN_TTL_MS
  expect(tickets.status(id, () => true)).toBeUndefined()
})
