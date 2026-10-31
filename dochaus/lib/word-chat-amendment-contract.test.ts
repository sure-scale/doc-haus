import { expect, test } from "bun:test"
import { targetWordChatAmendment, validWordChatFormatting, validWordChatInlineFormatting, validWordChatParagraphInsertion, validWordChatParagraphDeletion, validWordChatParagraphMerge, validWordChatParagraphSplit, validWordChatListInsertion, wordChatListCaptureCurrent, validWordChatTableRowInsertion, validWordChatTableRowDeletion, wordChatApplyPlan, wordChatParagraphBreaks, wordChatTableCaptureCurrent, wordChatTargetCurrent, type WordChatSnapshot } from "./word-chat-amendment-contract"

test("inline formatting binds exact native UTF-16 targets without trusting model offsets", () => {
  const text = "Keep 😀 terms and terms."
  const paragraph = {id:"native-paragraph",text}
  const start = text.lastIndexOf("terms"), selection = {id:"native-selection",text:"terms",paragraphs:[paragraph],startOffset:start,endOffset:start+5,capturedAt:"now",saved:false}
  const context: WordChatSnapshot = {text,paragraphs:[paragraph],selection,capturedAt:"now",saved:false,truncated:false}
  const target = targetWordChatAmendment(context,"terms")!
  expect(validWordChatInlineFormatting(context,target,{bold:true,italic:false,underline:true})).toBe(true)
  expect(validWordChatInlineFormatting(context,target,{underline:false})).toBe(true)
  expect(validWordChatInlineFormatting(context,target,{})).toBe(false)
  expect(validWordChatInlineFormatting(context,target,{bold:undefined})).toBe(false)
  expect(validWordChatInlineFormatting(context,target,{bold:"true"})).toBe(false)
  expect(validWordChatInlineFormatting(context,target,{bold:true,offset:0})).toBe(false)
  expect(validWordChatInlineFormatting(context,target,{bold:true},"changed")).toBe(false)
  expect(validWordChatInlineFormatting(context,target,{bold:true},target.text,[{paragraph:0,style:"Heading1"}])).toBe(false)
  expect(validWordChatInlineFormatting({...context,selection:undefined},target,{bold:true})).toBe(false)
  expect(targetWordChatAmendment({...context,selection:undefined},"terms")).toBeUndefined()
  expect(validWordChatInlineFormatting(context,{...target,startOffset:text.indexOf("terms"),endOffset:text.indexOf("terms")+5},{bold:true})).toBe(false)
  expect(validWordChatInlineFormatting({...context,paragraphs:[{...paragraph,text:text+" changed"}]},target,{bold:true})).toBe(false)
  const unselected = {...context,selection:undefined}
  const full = targetWordChatAmendment(unselected,text)!, partial = targetWordChatAmendment(unselected,"Keep")!
  expect(validWordChatInlineFormatting(unselected,full,{bold:true})).toBe(true)
  expect(validWordChatInlineFormatting(unselected,partial,{bold:true})).toBe(true)
  const emoji = text.indexOf("😀"), split = targetWordChatAmendment(unselected,text.slice(emoji,emoji+1))!
  expect(validWordChatInlineFormatting(unselected,split,{bold:true})).toBe(false)
  const wholeEmoji = targetWordChatAmendment(unselected,"😀")!
  expect(validWordChatInlineFormatting(unselected,wholeEmoji,{italic:true})).toBe(true)
})

test("row insertion binds ordered native cell identities and literal cells to the captured table", () => {
  const paragraphs = [{id:"a",text:"One"},{id:"b",text:"Two"},{id:"c",text:"Three"},{id:"d",text:"Four"}]
  const table = {tableIndex:0,rows:[paragraphs.slice(0,2),paragraphs.slice(2)]}
  const context: WordChatSnapshot = {text:paragraphs.map(p => p.text).join("\r"),paragraphs,table,saved:false,capturedAt:"now",truncated:false}
  const proposal = {kind:"table_row_insert",id:"proposal",documentId:"document",explanation:"Append a row",openItems:[],createdAt:"now",position:"end",table,cells:["Five","Six"]}
  expect(wordChatTableCaptureCurrent(context, table)).toBe(true)
  expect(validWordChatTableRowInsertion(context, proposal)).toBe(true)
  expect(validWordChatTableRowInsertion(context, {...proposal, table:{...table,tableIndex:1}})).toBe(false)
  expect(validWordChatTableRowInsertion(context, {...proposal, cells:["Five"]})).toBe(false)
  expect(validWordChatTableRowInsertion(context, {...proposal, cells:[" Five","Six"]})).toBe(false)
  expect(validWordChatTableRowInsertion(context, {...proposal, cells:["Five\nSix",""]})).toBe(false)
  expect(validWordChatTableRowInsertion(context, {...proposal, position:"after"})).toBe(false)
  expect(wordChatTableCaptureCurrent(context, {...table, rows:[[paragraphs[0],paragraphs[0]],table.rows[1]]})).toBe(false)
  expect(wordChatTableCaptureCurrent(context, {...table, rows:[[paragraphs[1],paragraphs[0]],table.rows[1]]})).toBe(false)
  expect(wordChatTableCaptureCurrent(context, {...table, rows:[[...table.rows[0]!], [{...paragraphs[2]!,text:"Changed"},paragraphs[3]]]})).toBe(false)
  expect(wordChatTableCaptureCurrent(context, {...table, rows:[table.rows[0], [paragraphs[2]]]})).toBe(false)
})

const snapshot = (): WordChatSnapshot => ({text: "Pay within thirty days.\rPay within thirty days.",
  paragraphs: [{id:"first",text:"Pay within thirty days."},{id:"second",text:"Pay within thirty days."}],
  saved:false,capturedAt:"2026-01-01T00:00:00Z",truncated:false})

test("a formatting review keeps its wording and whole-paragraph authority", () => {
  const value = snapshot()
  value.paragraphs.push({ id: "unchanged-tail", text: "An unchanged following paragraph." })
  value.text = value.paragraphs.map(p => p.text).join("\r")
  const target = targetWordChatAmendment(value, value.text)!
  const formatting = [{ paragraph: 0, style: "Heading1" as const, spacingBefore: 12, spacingAfter: 6 }, { paragraph: 1, lineSpacing: 1.15 }]
  expect(validWordChatFormatting(target, formatting)).toBe(true)
  expect(wordChatApplyPlan(target, target.text, formatting)).toEqual({ supported: true, paragraphs: value.paragraphs.map(p => p.text) })
  expect(wordChatApplyPlan(target, target.text.replace("thirty", "sixty"), formatting).supported).toBe(true)
  const partial = { ...target, startOffset: 4, text: target.text.slice(4) }
  expect(wordChatApplyPlan(partial, partial.text, formatting).supported).toBe(false)
  expect(validWordChatFormatting({ ...target, endOffset: target.endOffset - 1, text: target.text.slice(0, -1) }, formatting)).toBe(false)
})

test("native formatting can include the final paragraph", () => {
  const value = snapshot(), target = targetWordChatAmendment(value, value.text)!
  const formatting = [{ paragraph: 0, spacingAfter: 6 }, { paragraph: 1, spacingAfter: 6 }]
  expect(validWordChatFormatting(target, formatting)).toBe(true)
  expect(wordChatApplyPlan(target, target.text, formatting)).toEqual({ supported: true, paragraphs: target.paragraphs.map(p => p.text) })
  expect(wordChatApplyPlan(target, target.text, formatting.slice(0, 1))).toEqual({ supported: true, paragraphs: target.paragraphs.map(p => p.text) })
  const one = { ...target, text: target.paragraphs[0]!.text, paragraphs: target.paragraphs.slice(0, 1) }
  expect(validWordChatFormatting(one, formatting.slice(0, 1))).toBe(true)
  expect(wordChatApplyPlan(one, one.text, formatting.slice(0, 1)).supported).toBe(true)
})

test("formatting operations cannot smuggle edits or conflicting paragraph properties", () => {
  const target = targetWordChatAmendment(snapshot(), snapshot().text)!
  for (const operations of [[], [{ paragraph: 2, spacingAfter: 6 }], [{ paragraph: 0, spacingAfter: 6 }, { paragraph: 0, style: "Title" }],
    [{ paragraph: 0 }], [{ paragraph: 0, style: "An arbitrary style" }], [{ paragraph: 0, spacingAfter: -1 }],
    [{ paragraph: 0, spacingBefore: 49 }], [{ paragraph: 0, lineSpacing: 0 }], [{ paragraph: 0, spacingAfter: Infinity }],
    [{ paragraph: 0, spacingAfter: 6, text: "Different wording" }], [{ paragraph: 0, ooxml: "<w:p/>" }]]) {
    expect(validWordChatFormatting(target, operations)).toBe(false)
  }
})

test("structural eligibility pins exact whole paragraphs and keeps the final import boundary", () => {
  const value = snapshot()
  value.paragraphs[1]!.text = "An unchanged qualification."
  const target = targetWordChatAmendment(value, value.paragraphs.map(p => p.text).join("\r"))!
  expect(wordChatApplyPlan(target, "A working note.\r" + target.text)).toEqual({ supported: true,
    paragraphs: ["A working note.", ...target.paragraphs.map(p => p.text)], structure: { kind: "insert", retained: [{ before: 0, after: 1 }, { before: 1, after: 2 }] } })
  expect(wordChatApplyPlan(target, target.paragraphs[1]!.text)).toMatchObject({ supported: true, structure: { kind: "delete", retained: [{ before: 1, after: 0 }] } })
  expect(wordChatApplyPlan(target, target.text + "\rA trailing note.").supported).toBe(false)
  expect(wordChatApplyPlan(target, "Pay within\rthirty days.\rAn unchanged qualification.").supported).toBe(false)
  const repeated = { ...target, paragraphs: [target.paragraphs[0]!, target.paragraphs[0]!, target.paragraphs[1]!], text: "Pay within thirty days.\rPay within thirty days.\rAn unchanged qualification." }
  expect(wordChatApplyPlan(repeated, target.text).supported).toBe(false)
  expect(wordChatApplyPlan({ ...target, startOffset: 4, text: target.text.slice(4) }, "A working note.\r" + target.text.slice(4)).supported).toBe(false)
})

test("mixed insertions index formatting against the result and reject ambiguous mappings", () => {
  const texts = ["The parties shall deliver the materials promptly.", "Unchanged qualification."]
  const target = targetWordChatAmendment({ ...snapshot(), text: texts.join("\r"), paragraphs: texts.map((text, i) => ({ id: String(i), text })) }, texts.join("\r"))!
  const replacement = ["Delivery", "The parties shall deliver the materials promptly and in writing.", texts[1]!].join("\r")
  expect(wordChatApplyPlan(target, replacement, [{ paragraph: 0, style: "Heading1" }])).toMatchObject({ supported: true, structure: { kind: "insert", retained: [{ before: 0, after: 1 }, { before: 1, after: 2 }] } })
  expect(validWordChatFormatting(target, [{ paragraph: 2, spacingAfter: 6 }], replacement)).toBe(true)
  expect(wordChatApplyPlan(target, replacement, [{ paragraph: 2, spacingAfter: 6 }]).supported).toBe(true)
  expect(wordChatApplyPlan(target, ["The parties shall deliver promptly.", "The parties shall deliver promptly.", texts[1]!].join("\r")).supported).toBe(false)
})

test("retained selection rejects stale native paragraphs even when its old words remain elsewhere", () => {
  const value = snapshot()
  const selected = { id: "selection", text: "thirty days", paragraphs: [{ ...value.paragraphs[1]! }], startOffset: 11, endOffset: 22, saved: false, capturedAt: value.capturedAt }
  expect(wordChatTargetCurrent(value, selected)).toBe(true)
  value.paragraphs[1]!.text = "Pay within forty days."
  value.text = value.paragraphs.map(paragraph => paragraph.text).join("\r")
  expect(value.text.includes(selected.text)).toBe(true)
  expect(wordChatTargetCurrent(value, selected)).toBe(false)
  value.paragraphs[1]!.text = selected.paragraphs[0]!.text
  expect(wordChatTargetCurrent(value, { ...selected, startOffset: 0 })).toBe(false)
  expect(wordChatTargetCurrent(value, { ...selected, endOffset: 500 })).toBe(false)
  expect(wordChatTargetCurrent(value, { ...selected, paragraphs: [{ ...selected.paragraphs[0]!, id: "missing" }] })).toBe(false)
  value.paragraphs[0]!.text = "An unrelated local edit."
  expect(wordChatTargetCurrent(value, { ...selected, id: "fresh-capture", capturedAt: "2026-01-02T00:00:00Z" })).toBe(true)
})

test("retained multiline selection requires its exact contiguous sequence and boundary text", () => {
  const value: WordChatSnapshot = { ...snapshot(), text: "Lead first words\rMiddle wording\rLast words tail", paragraphs: [
    { id: "first", text: "Lead first words" }, { id: "middle", text: "Middle wording" }, { id: "last", text: "Last words tail" },
  ] }
  const selected = targetWordChatAmendment(value, "first words\rMiddle wording\rLast words")!
  expect(wordChatTargetCurrent(value, selected)).toBe(true)
  value.paragraphs.splice(1, 0, { id: "new", text: "Inserted between selected paragraphs." })
  expect(wordChatTargetCurrent(value, selected)).toBe(false)
  value.paragraphs.splice(1, 1)
  expect(wordChatTargetCurrent(value, selected)).toBe(true)
  expect(wordChatTargetCurrent(value, { ...selected, text: "Changed selection" })).toBe(false)
  expect(wordChatTargetCurrent(value, { ...selected, paragraphs: [undefined as never] })).toBe(false)
})

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


test("paragraph insertion binds a unique complete native paragraph and literal new text", () => {
  const text="Existing paragraph.", context: WordChatSnapshot={text,paragraphs:[{id:"native",text}],capturedAt:"now",saved:false,truncated:false}
  const target=targetWordChatAmendment(context,text)!
  expect(validWordChatParagraphInsertion(context,target,text+"\rAdded 😀 paragraph.")).toBe(true)
  for (const after of [text+"\r",text+"\r  ",text+"\rOne\rTwo",text+"\rOne\tTwo",text+"\r\ud800", "Changed\rAdded"]) expect(validWordChatParagraphInsertion(context,target,after)).toBe(false)
  expect(validWordChatParagraphInsertion(context,target,text+"\rAdded",[{paragraph:1,style:"Normal"}])).toBe(true)
  for (const formatting of [[{paragraph:0,style:"Normal"}], [{paragraph:1,spacingAfter:8}], [{paragraph:1,style:"Heading1"}], [{paragraph:1,style:"Normal",spacingAfter:8}], [{paragraph:1,style:"Normal"},{paragraph:0,style:"Normal"}]])
    expect(validWordChatParagraphInsertion(context,target,text+"\rAdded",formatting)).toBe(false)
  expect(validWordChatParagraphInsertion(context,target,text+"\rAdded",undefined,{bold:true})).toBe(false)
  const partial=targetWordChatAmendment(context,"paragraph.")!
  expect(validWordChatParagraphInsertion(context,partial,partial.text+"\rAdded")).toBe(false)
  expect(validWordChatParagraphInsertion({...context,paragraphs:[{id:"native",text:"changed"}]},target,text+"\rAdded")).toBe(false)
  const repeated={...context,text:text+"\r"+text,paragraphs:[{id:"a",text},{id:"b",text}]}
  expect(targetWordChatAmendment(repeated,text)).toBeUndefined()
})


test("paragraph deletion requires explicit action and one unique complete native paragraph",()=>{
  const text="A complete paragraph.", paragraph={id:"native",text}, snapshot:WordChatSnapshot={text,paragraphs:[paragraph],capturedAt:"now",saved:false,truncated:false}
  const target=targetWordChatAmendment(snapshot,text)!
  expect(validWordChatParagraphDeletion(snapshot,target,"","delete")).toBe(true)
  expect(validWordChatParagraphDeletion(snapshot,target,"",undefined)).toBe(false)
  expect(validWordChatParagraphDeletion(snapshot,target,"","clear")).toBe(false)
  expect(validWordChatParagraphDeletion(snapshot,target,"changed","delete")).toBe(false)
  expect(validWordChatParagraphDeletion(snapshot,target,"","delete",[{paragraph:0,style:"Normal"}])).toBe(false)
  expect(validWordChatParagraphDeletion(snapshot,target,"","delete",undefined,{bold:true})).toBe(false)
  const partial=targetWordChatAmendment(snapshot,"complete")!
  expect(validWordChatParagraphDeletion(snapshot,partial,"","delete")).toBe(false)
  expect(validWordChatParagraphDeletion({...snapshot,paragraphs:[{id:"native",text:"changed"}]},target,"","delete")).toBe(false)
  expect(targetWordChatAmendment({...snapshot,text:text+"\r"+text,paragraphs:[paragraph,{id:"second",text}]},text)).toBeUndefined()
  const selected={...snapshot,selection:{...target,id:"selection"}}
  const resolved=targetWordChatAmendment(selected,text)!
  expect(validWordChatParagraphDeletion(selected,resolved,"","delete")).toBe(true)
})


test("decimal list capture and insertion bind the full contiguous sequence and last item",()=>{
  const paragraphs=[{id:"one",text:"First"},{id:"two",text:"Second"}]
  const list={id:7,format:"decimal" as const,members:paragraphs.map((p,i)=>({...p,level:0,label:`${i+1}.`}))}
  const snapshot:WordChatSnapshot={text:"First\rSecond",paragraphs,capturedAt:"now",saved:false,truncated:false,list}
  const target=targetWordChatAmendment(snapshot,"Second")!
  expect(wordChatListCaptureCurrent(snapshot,list)).toBe(true)
  expect(validWordChatListInsertion(snapshot,target,"Second\rThird","append_list_item")).toBe(true)
  expect(validWordChatListInsertion(snapshot,target,"Second\rThird","append_list_item",[{paragraph:1,style:"Normal"}])).toBe(false)
  expect(validWordChatListInsertion(snapshot,target,"Second\rThird","delete")).toBe(false)
  expect(validWordChatListInsertion(snapshot,target,"Second\rThird\rFourth","append_list_item")).toBe(false)
  expect(validWordChatListInsertion(snapshot,targetWordChatAmendment(snapshot,"First")!,"First\rThird","append_list_item")).toBe(false)
  for(const bad of [{...list,members:[list.members[0]!,{...list.members[1]!,id:"one"}]},{...list,members:[list.members[0]!,{...list.members[1]!,level:1}]},{...list,members:[list.members[0]!,{...list.members[1]!,label:"3."}]},{...list,extra:true}]) expect(wordChatListCaptureCurrent(snapshot,bad)).toBe(false)
  expect(wordChatListCaptureCurrent({...snapshot,paragraphs:[paragraphs[0]!,{id:"middle",text:"other"},paragraphs[1]!]},list)).toBe(false)
})


test("row deletion binds exact original cells and the captured row and refuses stale or sole rows", () => {
  const paragraphs = ["Header", "Amount", "Fee", " Ten "].map((text,i) => ({id:`cell-${i}`,text}))
  const table = {tableIndex:0, rows:[paragraphs.slice(0,2),paragraphs.slice(2)]}
  const context: WordChatSnapshot = {text:paragraphs.map(p=>p.text).join("\r"),paragraphs,table,capturedAt:"now",saved:false,truncated:false}
  const proposal = {kind:"table_row_delete",id:"p",documentId:"d",explanation:"Remove",openItems:[],createdAt:"now",table,rowIndex:1,cells:["Fee"," Ten "]}
  expect(validWordChatTableRowDeletion(context,proposal)).toBe(true)
  for (const bad of [{...proposal,rowIndex:0},{...proposal,rowIndex:2},{...proposal,cells:["Fee","Ten"]},{...proposal,unknown:true},{...proposal,table:{...table,rows:[table.rows[1]]}}]) expect(validWordChatTableRowDeletion(context,bad)).toBe(false)
  expect(validWordChatTableRowDeletion({...context,table:{...table,rows:[table.rows[1]!,table.rows[0]!] }},proposal)).toBe(false)
})


test("paragraph merge binds two full adjacent paragraphs and exact boundary", () => {
  const paragraphs=[{id:"a",text:"First sentence."},{id:"b",text:"Second sentence."},{id:"c",text:"Third."}]
  const snapshot:WordChatSnapshot={text:paragraphs.map(p=>p.text).join("\r"),paragraphs,capturedAt:"now",saved:false,truncated:false}
  const before="First sentence.\rSecond sentence.",after="First sentence. Second sentence.",target=targetWordChatAmendment(snapshot,before)!
  expect(validWordChatParagraphMerge(snapshot,target,after,"merge")).toBe(true)
  for (const [value, action, format, inline] of [["First sentence.Second sentence.","merge",undefined,undefined],[after,"delete",undefined,undefined],[after,"merge",[],undefined],[after,"merge",undefined,{bold:true}]] as const) expect(validWordChatParagraphMerge(snapshot,target,value,action,format,inline)).toBe(false)
  expect(validWordChatParagraphMerge(snapshot,{...target,startOffset:1},after,"merge")).toBe(false)
  expect(validWordChatParagraphMerge({...snapshot,paragraphs:[paragraphs[0]!,paragraphs[2]!,paragraphs[1]!] },target,after,"merge")).toBe(false)
  expect(validWordChatParagraphMerge({...snapshot,paragraphs:[paragraphs[0]!,{...paragraphs[1]!,text:"Changed"},paragraphs[2]!] },target,after,"merge")).toBe(false)
  expect(validWordChatParagraphMerge({...snapshot,list:{id:1,format:"decimal",members:[{...paragraphs[1]!,level:0,label:"1."}]}},target,after,"merge")).toBe(false)
  const partial={id:"selected",text:"sentence.\rSecond sentence.",paragraphs:paragraphs.slice(0,2),startOffset:6,endOffset:paragraphs[1]!.text.length,capturedAt:"now",saved:false}
  expect(validWordChatParagraphMerge({...snapshot,selection:partial},target,after,"merge")).toBe(false)
  const duplicates={...snapshot,paragraphs:[...paragraphs,...paragraphs.slice(0,2).map((p,i)=>({...p,id:`copy-${i}`}))]}
  expect(targetWordChatAmendment(duplicates,before)).toBeUndefined()
})


test("paragraph split binds a unique complete paragraph and exactly one existing space", () => {
  const paragraphs=[{id:"a",text:"First sentence. Second sentence."},{id:"b",text:"Other."}],snapshot:WordChatSnapshot={text:paragraphs.map(p=>p.text).join("\r"),paragraphs,capturedAt:"now",saved:false,truncated:false}
  const target=targetWordChatAmendment(snapshot,paragraphs[0]!.text)!,after="First sentence.\rSecond sentence."
  expect(validWordChatParagraphSplit(snapshot,target,after,"split")).toBe(true)
  for (const bad of ["First sentence.\rChanged sentence.","First sentence.\r\rSecond sentence.","\rFirst sentence. Second sentence.","First sentence.\tSecond sentence.","First sentence.\rSecond\n sentence."]) expect(validWordChatParagraphSplit(snapshot,target,bad,"split")).toBe(false)
  expect(validWordChatParagraphSplit(snapshot,target,after,"merge")).toBe(false)
  expect(validWordChatParagraphSplit(snapshot,target,after,"split",[])).toBe(false)
  expect(validWordChatParagraphSplit(snapshot,target,after,"split",undefined,{bold:true})).toBe(false)
  expect(validWordChatParagraphSplit(snapshot,{...target,startOffset:1},after,"split")).toBe(false)
  expect(validWordChatParagraphSplit({...snapshot,paragraphs:[{...paragraphs[0]!,text:"Changed"},paragraphs[1]!] },target,after,"split")).toBe(false)
  expect(validWordChatParagraphSplit({...snapshot,list:{id:1,format:"decimal",members:[{...paragraphs[0]!,level:0,label:"1."}]}},target,after,"split")).toBe(false)
  const selection={id:"partial",text:"sentence. Second sentence.",paragraphs:[paragraphs[0]!],startOffset:6,endOffset:paragraphs[0]!.text.length,capturedAt:"now",saved:false}
  expect(validWordChatParagraphSplit({...snapshot,selection},target,after,"split")).toBe(false)
  expect(targetWordChatAmendment({...snapshot,paragraphs:[...paragraphs,{id:"copy",text:paragraphs[0]!.text}]},paragraphs[0]!.text)).toBeUndefined()
})
