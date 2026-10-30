import { expect, test } from "bun:test"
import { nativeWordChatTextEligible, wordChatApplyEligibility } from "./word-chat-apply"

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

test("inline formatting eligibility accepts exact partial and whole-paragraph spans, including turning a property off", () => {
  expect(wordChatApplyEligibility(target, target.text, undefined, true, { bold: false, italic: true })).toEqual({ supported: true })
  const whole = { ...target, text: "Alpha beta", startOffset: 0, endOffset: 10 }
  expect(wordChatApplyEligibility(whole, whole.text, undefined, true, { underline: false })).toEqual({ supported: true })
})

test("paragraph mark deletion needs an explicit action on the complete reviewed paragraph", () => {
  const whole = { ...target, text: "Alpha beta", startOffset: 0, endOffset: 10 }
  expect(wordChatApplyEligibility(whole, "", undefined, true, undefined, "delete").supported).toBe(true)
  expect(wordChatApplyEligibility(whole, "").supported).toBe(false)
  expect(wordChatApplyEligibility(target, "", undefined, true, undefined, "delete").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "changed", undefined, true, undefined, "delete").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "", [{paragraph:0,style:"Normal"}], true, undefined, "delete").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "", undefined, true, {bold:true}, "delete").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "", undefined, false, undefined, "delete").supported).toBe(false)
})

test("automatic list insertion binds its end and cannot fall through to ordinary insertion", () => {
  const whole = { ...target, text: "Alpha beta", startOffset: 0, endOffset: 10 }
  const list = { id: 17, format: "decimal" as const, members: [{ id: "earlier", text: "Earlier clause", level: 0, label: "1." },
    { id: "p", text: whole.text, level: 0, label: "2." }] }
  const after = whole.text + "\rNext clause."
  expect(wordChatApplyEligibility(whole, after, undefined, true, undefined, "append_list_item", list).supported).toBe(true)
  for (const changed of [undefined, { ...list, members: [...list.members].reverse() }, { ...list, members: [{ ...list.members[1]!, label: "1." }, { ...list.members[0]!, label: "2." }] }])
    expect(wordChatApplyEligibility(whole, after, undefined, true, undefined, "append_list_item", changed).supported).toBe(false)
  expect(wordChatApplyEligibility(whole, after, [{paragraph:1,style:"Normal"}], true, undefined, "append_list_item", list).supported).toBe(false)
  expect(wordChatApplyEligibility(whole, after, undefined, true, {bold:true}, "append_list_item", list).supported).toBe(false)
  expect(wordChatApplyEligibility(whole, after, undefined, false, undefined, "append_list_item", list).supported).toBe(false)
})

test("native paragraph insertion requires a complete unchanged anchor and one new paragraph", () => {
  const whole = { ...target, text: "Alpha beta", startOffset: 0, endOffset: 10 }
  expect(wordChatApplyEligibility(whole, "Alpha beta\rNew paragraph.").supported).toBe(true)
  expect(wordChatApplyEligibility(whole, "Alpha beta\rNew paragraph.", [{ paragraph: 1, style: "Normal" }]).supported).toBe(true)
  expect(wordChatApplyEligibility(target, "beta\rNew paragraph.").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "Changed anchor\rNew paragraph.").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "Alpha beta\rNew paragraph.\rAnother paragraph.").supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "Alpha beta\rNew paragraph.", undefined, false).supported).toBe(false)
  expect(wordChatApplyEligibility(whole, "Alpha beta\rNew paragraph.", [{ paragraph: 0, spacingAfter: 6 }]).supported).toBe(false)
})

test("inline formatting eligibility refuses changed wording, malformed properties, mixed formatting, and mismatched native offsets", () => {
  expect(wordChatApplyEligibility(target, "changed", undefined, true, { bold: true }).supported).toBe(false)
  expect(wordChatApplyEligibility(target, target.text, undefined, true, { bold: "true" } as never).supported).toBe(false)
  expect(wordChatApplyEligibility(target, target.text, [{ paragraph: 0, spacingAfter: 6 }], true, { italic: true }).supported).toBe(false)
  expect(wordChatApplyEligibility({ ...target, startOffset: 5 }, target.text, undefined, true, { underline: true }).supported).toBe(false)
})
