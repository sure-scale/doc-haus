import { expect, test } from "bun:test"
import { nativeWordChatTextEligible } from "./word-chat-apply"

const target = { id: "p", text: "beta", paragraphs: [{ id: "p", text: "Alpha beta" }], startOffset: 6, endOffset: 10, capturedAt: "", saved: false }
test("one-sided wording eligibility binds the selection and keeps a nonempty paragraph", () => {
  expect(nativeWordChatTextEligible(target, "")).toBe(true)
  expect(nativeWordChatTextEligible(target, "new beta")).toBe(true)
  expect(nativeWordChatTextEligible({ ...target, text: "Alpha beta", startOffset: 0 }, "")).toBe(false)
  expect(nativeWordChatTextEligible({ ...target, startOffset: 5 }, "")).toBe(false)
  expect(nativeWordChatTextEligible({ ...target, endOffset: 100 }, "")).toBe(false)
  expect(nativeWordChatTextEligible(target, "line\nbreak")).toBe(false)
  expect(nativeWordChatTextEligible({ ...target, paragraphs: [...target.paragraphs, { id: "q", text: "other" }] }, "")).toBe(false)
  expect(nativeWordChatTextEligible({ ...target, paragraphs: [{ id: "p", text: "A😀B" }], text: "\ude00", startOffset: 2, endOffset: 3 }, "")).toBe(false)
})
