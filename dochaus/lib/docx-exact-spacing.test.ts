import { expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate"
import { renderMarkdownToDocx } from "./render"
import { createDocxRevision, applyDocxParagraphEdits } from "./docx-amendments"
import { exportDocxComments, readDocxComments } from "./docx-comments"
import { docxodus } from "./docxodus"
import { authorExactParagraphSpacing } from "./docx-exact-spacing"
const sha = (bytes:Uint8Array) => createHash("sha256").update(bytes).digest("hex")
test("original generated source refuses global style; narrowed synthetic source preserves paired bookmarks in either duplicate cell", async () => {
  const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |" )).bytes
  const original=await createDocxRevision(raw), rawRow=original.paragraphs.find(p=>p.text === "Same")!
  await expect(authorExactParagraphSpacing(raw,{sourceSha256:sha(raw),paragraphHandle:rawRow.paragraph_handle,paragraphText:rawRow.text,tableContext:rawRow.table,spacingAfterPoints:6,author:"Synthetic Counsel"})).rejects.toThrow("unresolved_spacing")
  const parts=unzipSync(raw)
  parts["word/styles.xml"]=strToU8(strFromU8(parts["word/styles.xml"]!).replace("<w:contextualSpacing />", ""))
  const source=zipSync(parts)
  const hash=sha(source), initial=await createDocxRevision(source)
  const rows=initial.paragraphs.filter(p=>p.text === "Same"); expect(rows).toHaveLength(2)
  for (const row of rows) {
    const result=await authorExactParagraphSpacing(source,{sourceSha256:hash,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"})
    expect(result.afterDirect).toBe("120"); expect(sha(source)).toBe(hash)
  }
})

test("bookmark and protected-range tampering refuses without changing the source", async () => {
  const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |" )).bytes
  const parts=unzipSync(raw)
  parts["word/styles.xml"]=strToU8(strFromU8(parts["word/styles.xml"]!).replace("<w:contextualSpacing />", ""))
  const doc=strFromU8(parts["word/document.xml"]!), paras=[...doc.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].filter(m=>m[0].includes(">Same</w:t>"))
  expect(paras).toHaveLength(2)
  const target=paras[1]![0], start=target.match(/<w:bookmarkStart\b[^>]*\/>/)![0], end=target.match(/<w:bookmarkEnd\b[^>]*\/>/)![0], firstStart=paras[0]![0].match(/<w:bookmarkStart\b[^>]*\/>/)![0]
  const id=start.match(/w:id="([^"]+)"/)![1]!, name=firstStart.match(/w:name="([^"]+)"/)![1]!
  const changed=(p:string)=>doc.replace(target,p)
  const cases: [string,string,string][] = [
    ["alias duplicate id","invalid_range",changed(target.replace(start,start.replace(`w:id="${id}"`,`w:id="0${paras[0]![0].match(/<w:bookmarkStart\b[^>]*w:id="([^"]+)"/)![1]}"`)))],
    ["missing name","invalid_bookmark_name",changed(target.replace(start,start.replace(/ w:name="[^"]+"/,"")))],
    ["duplicate name","invalid_bookmark_name",changed(target.replace(start,start.replace(/w:name="[^"]+"/,`w:name="${name}"`)))],
    ["reversed","invalid_range",changed(target.replace(start,"__START__").replace(end,start).replace("__START__",end))],
    ["nonempty","protected_range",changed(target.replace(start,start.replace("/>",">unexpected</w:bookmarkStart>")))],
    ["column","protected_range",changed(target.replace(start,start.replace("/>",' w:colFirst="0" w:colLast="1"/>')))],
    ["cross paragraph","protected_range",doc.replace(target, target.replace(end,"")+end)],
    ["enclosing","protected_range",doc.replace(target, start.replace(`w:id="${id}"`,'w:id="700"').replace(/w:name="[^"]+"/,'w:name="enclosing"')+target+end.replace(`w:id="${id}"`,'w:id="700"'))],
    ["comment","protected_range",changed(target.replace(start,'<w:commentRangeStart w:id="701"/>'+start).replace(end,end+'<w:commentRangeEnd w:id="701"/>'))],
    ["permission","protected_range",changed(target.replace(start,'<w:permStart w:id="702" w:edGrp="everyone"/>'+start).replace(end,end+'<w:permEnd w:id="702"/>'))],
    ["unsupported move","unsupported_range",changed(target.replace(start,'<w:moveFromRangeStart w:id="703" w:name="m"/>'+start).replace(end,end+'<w:moveFromRangeEnd w:id="703"/>'))],
    ["unbound QName","unbound_mc_prefix",changed(target.replace('<w:p>', '<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="missing">'))],
  ]
  const baseline=zipSync(parts), baselineProjection=await createDocxRevision(baseline), baselineRow=baselineProjection.paragraphs.filter(p=>p.text === "Same")[1]!
  const session=(await docxodus()).openDocxSession(baseline,{}), prototype=Object.getPrototypeOf(session); session.close()
  const setter=spyOn(prototype,"setParagraphFormat")
  try {
    for (const [name,code,xml] of cases) {
      expect(xml, name).not.toBe(doc)
      const source=zipSync({...parts,"word/document.xml":strToU8(xml)}), hash=sha(source)
      const row=code === "unbound_mc_prefix" ? baselineRow : (await createDocxRevision(source,{pendingRevisions:"layer"})).paragraphs.filter(p=>p.text === "Same")[1]!
      expect(row, name).toBeDefined()
      await expect(authorExactParagraphSpacing(source,{sourceSha256:hash,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"})).rejects.toThrow(`docx_exact_spacing:${code}`)
      expect(setter).toHaveBeenCalledTimes(0); expect(sha(source)).toBe(hash)
    }
    const requires=doc.replace(target,target.replace('<w:p>', '<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="missing"><w:r><w:t>X</w:t></w:r></mc:Choice></mc:AlternateContent>'))
    expect(requires).not.toBe(doc)
    const source=zipSync({...parts,"word/document.xml":strToU8(requires)}), hash=sha(source)
    await expect(authorExactParagraphSpacing(source,{sourceSha256:hash,paragraphHandle:baselineRow.paragraph_handle,paragraphText:baselineRow.text,tableContext:baselineRow.table,spacingAfterPoints:6,author:"Synthetic Counsel"})).rejects.toThrow("docx_exact_spacing:unbound_mc_prefix")
    expect(setter).toHaveBeenCalledTimes(0); expect(sha(source)).toBe(hash)
    const illegal = [
      '<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w:bogus">',
      '<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:ProcessContent="w:">',
      '<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:PreserveElements="w:x:y">',
      '<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w:bad"/></mc:AlternateContent>',
      '<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires=""/></mc:AlternateContent>',
    ]
    for (const injection of illegal) {
      const xml=doc.replace(target,target.replace('<w:p>',injection)); expect(xml).not.toBe(doc)
      const bytes=zipSync({...parts,"word/document.xml":strToU8(xml)}), digest=sha(bytes)
      await expect(authorExactParagraphSpacing(bytes,{sourceSha256:digest,paragraphHandle:baselineRow.paragraph_handle,paragraphText:baselineRow.text,tableContext:baselineRow.table,spacingAfterPoints:6,author:"Synthetic Counsel"})).rejects.toThrow("docx_exact_spacing:invalid_mc_token")
      expect(setter).toHaveBeenCalledTimes(0); expect(sha(bytes)).toBe(digest)
    }
  } finally { setter.mockRestore() }
})

test("source-bound selector refuses stale or altered authority before the SDK setter", async () => {
  const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes, parts=unzipSync(raw)
  parts["word/styles.xml"]=strToU8(strFromU8(parts["word/styles.xml"]!).replace("<w:contextualSpacing />",""))
  const source=zipSync(parts), hash=sha(source), projection=await createDocxRevision(source), row=projection.paragraphs.filter(p=>p.text === "Same")[1]!
  const request={sourceSha256:hash,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"}
  const other=projection.paragraphs.filter(p=>p.text === "Same")[0]!
  const session=(await docxodus()).openDocxSession(source,{}), proto=Object.getPrototypeOf(session); session.close()
  const spy=spyOn(proto,"setParagraphFormat")
  try {
    for (const [changed,code] of [
      [{...request,sourceSha256:"0".repeat(64)},"stale_source"],
      [{...request,paragraphHandle:"paragraph-not-offered"},"unknown_handle"],
      [{...request,paragraphText:"Different"},"target_changed"],
      [{...request,tableContext:other.table},"target_changed"],
      [{...request,spacingAfterPoints:4},"no_change"],
    ] as const) {
      await expect(authorExactParagraphSpacing(source,changed)).rejects.toThrow(`docx_exact_spacing:${code}`)
      expect(spy).toHaveBeenCalledTimes(0); expect(sha(source)).toBe(hash)
    }
  } finally {spy.mockRestore()}
})

test("SDK package or namespace tampering is refused before an artifact is returned", async () => {
  const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes, parts=unzipSync(raw)
  parts["word/styles.xml"]=strToU8(strFromU8(parts["word/styles.xml"]!).replace("<w:contextualSpacing />",""))
  const source=zipSync(parts), hash=sha(source), projection=await createDocxRevision(source), row=projection.paragraphs.filter(p=>p.text === "Same")[1]!
  const request={sourceSha256:hash,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"}
  const session=(await docxodus()).openDocxSession(source,{}), proto=Object.getPrototypeOf(session), original=proto.save; session.close()
  for (const mode of ["namespace","qname","part"] as const) {
    const spy=spyOn(proto,"save").mockImplementation(function(this:unknown,...args:unknown[]) {
      const saved=original.apply(this,args), modified=unzipSync(saved)
      if (mode === "namespace") modified["word/document.xml"]=strToU8(strFromU8(modified["word/document.xml"]!).replace('<w:document ', '<w:document xmlns:proof="urn:unexpected" '))
      else if (mode === "qname") modified["word/document.xml"]=strToU8(strFromU8(modified["word/document.xml"]!).replace('<w:document ', '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="missing" '))
      else modified["word/styles.xml"]=strToU8(strFromU8(modified["word/styles.xml"]!).replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="Tampered"/></w:styles>'))
      return zipSync(modified)
    })
    try { await expect(authorExactParagraphSpacing(source,request)).rejects.toThrow(mode === "namespace" ? "other_body_changed" : mode === "qname" ? "unbound_mc_prefix" : "other_part_changed"); expect(sha(source)).toBe(hash); expect(spy).toHaveBeenCalled() }
    finally {spy.mockRestore()}
  }
})

test("preserves prior same-author revision and separate comment through selective decisions", async () => {
  const raw=(await renderMarkdownToDocx("Prior paragraph.\n\nComment paragraph.\n\n| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes, parts=unzipSync(raw)
  parts["word/styles.xml"]=strToU8(strFromU8(parts["word/styles.xml"]!).replace("<w:contextualSpacing />",""))
  const first=zipSync(parts), projection=await createDocxRevision(first), old=projection.paragraphs.find(p=>p.text === "Prior paragraph.")!
  const revised=await applyDocxParagraphEdits(first,[{id:"prior",paragraph_handle:old.paragraph_handle,occurrence:0,before_quote:"Prior",replacement:"Earlier"}],undefined,{pendingRevisions:"layer",revisionAuthor:"Synthetic Counsel",ownAuthors:["Synthetic Counsel"]})
  const source=exportDocxComments(revised.bytes,[{id:"note",text:"Existing comment.",author:"Reviewer",anchor:{paragraph_text:"Comment paragraph."}}]).bytes
  const hash=sha(source), current=await createDocxRevision(source,{pendingRevisions:"layer"}), row=current.paragraphs.filter(p=>p.text === "Same")[1]!
  const result=await authorExactParagraphSpacing(source,{sourceSha256:hash,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"})
  const dx=await docxodus(), beforeSession=dx.openDocxSession(source,{persistAnchorIds:false}); const prior=beforeSession.listRevisions(); beforeSession.close()
  expect(prior.length).toBeGreaterThan(0); expect(readDocxComments(source).length).toBe(1)
  const carriers=(bytes:Uint8Array)=>strFromU8(unzipSync(bytes)["word/document.xml"]!).match(/<w:(?:ins|del)\b[\s\S]*?<\/w:(?:ins|del)>/g) ?? []
  const expectedCarriers=carriers(source); expect(expectedCarriers.length).toBeGreaterThan(0)
  expect(carriers(result.trackedBytes)).toEqual(expectedCarriers)
  for (const accept of [false,true]) {
    const s=dx.openDocxSession(result.trackedBytes,{persistAnchorIds:false})
    try { expect((accept?s.acceptRevision(result.revisionId):s.rejectRevision(result.revisionId)).success).toBe(true); const resolved=s.save(); expect(readDocxComments(resolved)).toEqual(readDocxComments(source)); expect(carriers(resolved)).toEqual(expectedCarriers); expect(s.listRevisions().map(r=>r.id).sort()).toEqual(prior.map(r=>r.id).sort()) }
    finally {s.close()}
  }
  expect(sha(source)).toBe(hash)
})
