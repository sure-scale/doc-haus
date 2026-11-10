import { expect, test } from "bun:test"
import { wordChatExcerpts, wordChatHash, wordChatPromptSources, wordChatQuoteVerified, WORD_CHAT_SOURCE_COVERAGE } from "./word-chat-context"
import type { WordChatContext, WordChatSource } from "../../../dochaus/lib/word-chat-contract"

const context: WordChatContext = { text: "Actual […excerpt gap…] words.", paragraphs: [{ id: "p1", text: "Actual […excerpt gap…] words." }], saved: false, capturedAt: "now", truncated: false }
const source = (extra: Partial<WordChatSource> = {}): WordChatSource => ({ id: "word", name: "Draft.docx", location: "Current Word buffer", provenance: "word-buffer", ...wordChatExcerpts(context.text, ""), ...extra })

test("canonical excerpt selection coalesces adjacent bytes and distinguishes sampled from complete captures", () => {
  const text = "a".repeat(1590) + "Literal boundary phrase" + "b".repeat(5000)
  const selected = wordChatExcerpts(text, "", 4000)
  expect(selected.complete).toBe(false)
  expect(selected.passages).toEqual([text.slice(0, 3200)])
  expect(selected.quote).toBe(text.slice(0, 3200))
  expect(wordChatQuoteVerified(source({ ...selected }), "Literal boundary phrase", { ...context, text })).toBe(true)
  expect(wordChatExcerpts(context.text, "", context.text.length)).toEqual({ quote: context.text, passages: [context.text], complete: true })
  const separated = wordChatExcerpts("Needle " + "a".repeat(1593) + "b".repeat(6000) + "\nNeedle last", "Needle", 4000)
  expect(separated.passages.length).toBeGreaterThan(1)
  expect(separated.quote.length).toBeLessThanOrEqual(4000)
  expect(separated.passages.every(passage => ("Needle " + "a".repeat(1593) + "b".repeat(6000) + "\nNeedle last").includes(passage))).toBe(true)
})

test("canonical quote admission preserves literal markers and refuses presentation joins and unavailable old snapshots", () => {
  expect(wordChatQuoteVerified(source(), "[…excerpt gap…]", context)).toBe(true)
  const joined = source({ quote: "First\n[…excerpt gap…]\nLast", passages: ["First", "Last"], complete: false })
  expect(wordChatQuoteVerified(joined, "First\n[…excerpt gap…]\nLast", context)).toBe(false)
  const old = source({ passages: undefined, complete: undefined })
  expect(wordChatQuoteVerified(old, context.text, context)).toBe(true)
  expect(wordChatQuoteVerified({ ...old, provenance: "saved-version", id: "source:1" }, context.text, context)).toBe(false)
  expect(wordChatQuoteVerified({ ...old, provenance: "instructions", id: "playbook:old" }, context.text, context)).toBe(false)
  const upload = { kind: "upload" as const, id: "one", name: "notes.txt", provenance: "request-upload" as const, hash: wordChatHash(context.text), text: context.text }
  const oldUpload = { ...old, id: "upload:one", contentHash: upload.hash, provenance: "request-upload" as const }
  expect(wordChatQuoteVerified(oldUpload, context.text, context, [upload])).toBe(true)
  expect(wordChatQuoteVerified(oldUpload, context.text, context, [{ ...upload, text: "Different capture" }])).toBe(false)
  expect(wordChatQuoteVerified({ ...oldUpload, id: "upload:other" }, context.text, context, [upload])).toBe(false)
})

test("model source projection supplies explicit coverage without duplicating canonical passages", () => {
  const selected = source({ quote: "Actual", passages: [context.text], complete: false })
  expect(wordChatPromptSources([selected])).toEqual([{ id: "word", name: "Draft.docx", location: "Current Word buffer", provenance: "word-buffer", quote: "Actual", complete: false }])
  expect(WORD_CHAT_SOURCE_COVERAGE).toContain("absent means completeness is unknown")
  expect(WORD_CHAT_SOURCE_COVERAGE).toContain("never complete document semantics")
})
