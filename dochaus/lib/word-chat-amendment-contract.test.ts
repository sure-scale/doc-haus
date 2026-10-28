import { expect, test } from "bun:test"
import { targetWordChatAmendment, wordChatParagraphBreaks, type WordChatSnapshot } from "./word-chat-amendment-contract"

const snapshot = (): WordChatSnapshot => ({text: "Pay within thirty days.\rPay within thirty days.",
  paragraphs: [{id:"first",text:"Pay within thirty days."},{id:"second",text:"Pay within thirty days."}],
  saved:false,capturedAt:"2026-01-01T00:00:00Z",truncated:false})

test("ambiguous document wording requires a selected target rather than choosing the first copy", () => {
  const value = snapshot()
  expect(targetWordChatAmendment(value,"thirty days")).toBeUndefined()
  value.selection = {id:"selected",text:"thirty days",paragraphs:[value.paragraphs[1]!],startOffset:11,endOffset:22,
    saved:false,capturedAt:value.capturedAt}
  const target = targetWordChatAmendment(value,"thirty days")!
  expect(target.paragraphs.map(p => p.id)).toEqual(["second"])
  expect([target.startOffset,target.endOffset]).toEqual([11,22])
  value.paragraphs[1]!.text = "A later local edit."
  expect(target.paragraphs[0]!.text).toBe("Pay within thirty days.")
})

test("a proposal cannot silently expand an attached selection to other words in its paragraph", () => {
  const value = snapshot()
  value.selection = {id:"selected",text:"thirty days",paragraphs:[value.paragraphs[1]!],startOffset:11,endOffset:22,
    saved:false,capturedAt:value.capturedAt}
  expect(targetWordChatAmendment(value,"Pay within")).toBeUndefined()
  expect(targetWordChatAmendment(value,"thirty days.\rPay")).toBeUndefined()
})

test("an exact multiline amendment retains adjacent paragraphs and partial boundary offsets", () => {
  const value: WordChatSnapshot = {...snapshot(),text:"Lead first words\rMiddle wording\rLast words tail",
    paragraphs:[{id:"first",text:"Lead first words"},{id:"middle",text:"Middle wording"},{id:"last",text:"Last words tail"}]}
  const target = targetWordChatAmendment(value,"first words\nMiddle wording\nLast words")!
  expect(target.text).toBe("first words\rMiddle wording\rLast words")
  expect(target.paragraphs.map(p => p.id)).toEqual(["first","middle","last"])
  expect([target.startOffset,target.endOffset]).toEqual([5,10])
  expect(wordChatParagraphBreaks("first\r\nsecond\nthird\vsoft break")).toBe("first\rsecond\rthird\vsoft break")
})

test("a selected multiline occurrence is exact and stays pinned when a later selection changes", () => {
  const paragraphs = [{id:"first",text:"Lead first words"},{id:"second",text:"Last words tail"},
    {id:"third",text:"Lead first words"},{id:"fourth",text:"Last words tail"}]
  const value: WordChatSnapshot = {...snapshot(),text:paragraphs.map(p => p.text).join("\r"),paragraphs}
  expect(targetWordChatAmendment(value,"first words\rLast words")).toBeUndefined()
  value.selection = {id:"chosen",text:"first words\rLast words",paragraphs:paragraphs.slice(2),
    startOffset:5,endOffset:10,saved:false,capturedAt:value.capturedAt}
  const target = targetWordChatAmendment(value,"first words\rLast words")!
  expect(target.paragraphs.map(p => p.id)).toEqual(["third","fourth"])
  expect(targetWordChatAmendment(value,"Lead first words\rLast words")).toBeUndefined()
  value.selection = {...value.selection,id:"moved",paragraphs:paragraphs.slice(0,2)}
  value.paragraphs[2]!.text = "Later local text"
  expect(target.paragraphs.map(p => [p.id,p.text])).toEqual([["third","Lead first words"],["fourth","Last words tail"]])
  expect([target.startOffset,target.endOffset]).toEqual([5,10])
})
