import { expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { unzipSync, zipSync, strFromU8, strToU8, Zip, ZipPassThrough } from "fflate"
import { renderMarkdownToDocx } from "./render"
import { createDocxRevision, applyDocxParagraphEdits } from "./docx-amendments"
import { exportDocxComments, readDocxComments } from "./docx-comments"
import { docxodus } from "./docxodus"
import { authorExactParagraphSpacing, inspectExactParagraphSpacing, verifyExactParagraphSpacing, resolveExactParagraphSpacing } from "./docx-exact-spacing"
const sha = (bytes:Uint8Array) => createHash("sha256").update(bytes).digest("hex")
test("original generated source and irrelevant-style variant preserve paired bookmarks in either duplicate cell", async () => {
  const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |" )).bytes
  const original=await createDocxRevision(raw), rawRow=original.paragraphs.find(p=>p.text === "Same")!
  const originalResult=await authorExactParagraphSpacing(raw,{sourceSha256:sha(raw),paragraphHandle:rawRow.paragraph_handle,paragraphText:rawRow.text,tableContext:rawRow.table,spacingAfterPoints:6,author:"Synthetic Counsel"})
  expect(originalResult.afterDirect).toBe("120")
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
  const request={sourceSha256:hash,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"}
  const result=await authorExactParagraphSpacing(source,request)
  const verified=await verifyExactParagraphSpacing(source,result.trackedBytes,request)
  expect(verified.revisionId).toBe(result.revisionId)
  const resolvedByHelper=await resolveExactParagraphSpacing(source,result.trackedBytes,request,'accept')
  expect(readDocxComments(resolvedByHelper)).toEqual(readDocxComments(source))
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

test("only the selected paragraph's proven style closure admits spacing", async () => {
  const raw=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes, parts=unzipSync(raw)
  const body=strFromU8(parts["word/document.xml"]!), styles=strFromU8(parts["word/styles.xml"]!)
  const target=[...body.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].filter(m=>m[0].includes(">Same</w:t>"))[1]![0]
  const direct=(xml:string) => body.replace(target,target.replace('<w:pPr>',`<w:pPr>${xml}`))
  const normal='<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal" />'
  const cases:[string,string,string,string][]=[
    ["target contextual",direct('<w:contextualSpacing/>'),styles,"unresolved_spacing"],
    ["target grid",direct('<w:snapToGrid/>'),styles,"unresolved_spacing"],
    ["target numbering",direct('<w:numPr><w:numId w:val="1"/></w:numPr>'),styles,"unresolved_spacing"],
    ["target afterLines",body.replace(target,target.replace('w:after="80"','w:after="80" w:afterLines="100"')),styles,"unresolved_spacing"],
    ["defaults contextual",body,styles.replace('<w:pPrDefault><w:pPr>','<w:pPrDefault><w:pPr><w:contextualSpacing/>'),"unresolved_spacing"],
    ["default style auto",body,styles.replace(normal,normal+'<w:pPr><w:spacing w:afterAutospacing="1"/></w:pPr>'),"unresolved_spacing"],
    ["ancestor contextual",direct('<w:pStyle w:val="Heading2"/>'),styles.replace(normal,normal+'<w:pPr><w:contextualSpacing/></w:pPr>'),"unresolved_spacing"],
    ["style numbering",body,styles.replace(normal,normal+'<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>'),"unresolved_spacing"],
    ["missing style",direct('<w:pStyle w:val="Missing"/>'),styles,"unresolved_spacing"],
    ["cyclic style",body,styles.replace(normal,normal+'<w:basedOn w:val="Normal"/>'),"unresolved_spacing"],
    ["default table style",body,styles.replace('</w:styles>','<w:style w:type="table" w:default="1" w:styleId="Grid"/></w:styles>'),"unresolved_spacing"],
    ["explicit table style",body.replace('<w:tblPr>','<w:tblPr><w:tblStyle w:val="Grid"/>'),styles,"unsupported_table"],
    ["section grid",body.replace('<w:sectPr>','<w:sectPr><w:docGrid w:linePitch="360"/>'),styles,"unresolved_spacing"],
  ]
  const session=(await docxodus()).openDocxSession(raw,{}), proto=Object.getPrototypeOf(session); session.close()
  const setter=spyOn(proto,"setParagraphFormat")
  try {
    for (const [name,b,s,code] of cases) {
      expect(b !== body || s !== styles,name).toBe(true)
      const source=zipSync({...parts,"word/document.xml":strToU8(b),"word/styles.xml":strToU8(s)}), digest=sha(source)
      const projection=await createDocxRevision(source,{pendingRevisions:"layer"}), row=projection.paragraphs.filter(p=>p.text === "Same")[1]!
      expect(row,name).toBeDefined()
      await expect(authorExactParagraphSpacing(source,{sourceSha256:digest,paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"})).rejects.toThrow(`docx_exact_spacing:${code}`)
      expect(setter).toHaveBeenCalledTimes(0); expect(sha(source)).toBe(digest)
    }
  } finally {setter.mockRestore()}
})

test("OPC authoritative part bindings and style references fail closed", async () => {
  const source=(await renderMarkdownToDocx("| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes, parts=unzipSync(source), projection=await createDocxRevision(source)
  const row=projection.paragraphs.filter(p=>p.text === "Same")[1]!, request={sourceSha256:sha(source),paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:"Synthetic Counsel"}
  expect((await authorExactParagraphSpacing(source,request)).afterDirect).toBe("120")
  const root="_rels/.rels", rels="word/_rels/document.xml.rels", types="[Content_Types].xml"
  const original=(name:string)=>strFromU8(parts[name]!)
  const r=original(root), d=original(rels), c=original(types)
  const replace=(part:string,xml:string,changed:string,code:string) => ({part,xml,changed,code})
  const cases=[
    replace(root,r,r.replace('</Relationships>','<Relationship Id="strictRoot" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),"invalid_part_binding"),
    replace(root,r,r.replace('Target="word/document.xml"','Target="word/other.xml"'),"invalid_part_binding"),
    replace(root,r,r.replace('Target="word/document.xml"','Target="word/document.xml" TargetMode="External"'),"invalid_part_binding"),
    replace(root,r,r.replace('Target="word/document.xml"','Target="word/%64ocument.xml"'),"invalid_part_binding"),
    replace(rels,d,d.replace('</Relationships>','<Relationship Id="strictStyles" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/></Relationships>'),"invalid_part_binding"),
    replace(rels,d,d.replace('</Relationships>','<Relationship Id="strictSettings" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/settings" Target="settings.xml"/></Relationships>'),"invalid_part_binding"),
    replace(rels,d,d.replace('Target="styles.xml"','Target="other.xml"'),"invalid_part_binding"),
    replace(rels,d,d.replace('Target="settings.xml"','Target="../word/settings.xml"'),"invalid_part_binding"),
    replace(rels,d,d.replace('Target="styles.xml"','Target="styles.xml" TargetMode="External"'),"invalid_part_binding"),
    replace(rels,d,d.replace('</Relationships>','<Relationship Id="extra" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),"invalid_part_binding"),
    replace(rels,d,d.replace('Id="rId2"','Id="rId1"'),"invalid_relationships"),
    replace(types,c,c.replace('application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml','application/xml'),"invalid_content_types"),
    replace(types,c,c.replace('PartName="/word/settings.xml"','PartName="/word/Settings.xml"'),"invalid_content_types"),
    replace(types,c,c.replace('PartName="/word/settings.xml"','PartName="/word//settings.xml"'),"invalid_content_types"),
    replace("word/document.xml",original("word/document.xml"),original("word/document.xml").replace('<w:document ','<w:notDocument ').replace('</w:document>','</w:notDocument>'),"invalid_part_root"),
    replace("word/settings.xml",original("word/settings.xml"),original("word/settings.xml").replace('<w:settings','<w:notSettings').replace('</w:settings>','</w:notSettings>'),"invalid_part_root"),
    replace(types,c,c.replace('</Types>','<Override PartName="/word/styles.xml" ContentType="application/xml"/></Types>'),"invalid_content_types"),
  ]
  const session=(await docxodus()).openDocxSession(source,{}), proto=Object.getPrototypeOf(session);session.close(); const setter=spyOn(proto,"setParagraphFormat")
  try {
    for (const {part,xml,changed,code} of cases) {
      expect(changed).not.toBe(xml); const bytes=zipSync({...parts,[part]:strToU8(changed)}), digest=sha(bytes)
      await expect(authorExactParagraphSpacing(bytes,{...request,sourceSha256:digest})).rejects.toThrow(`docx_exact_spacing:${code}`)
      expect(setter).toHaveBeenCalledTimes(0); expect(sha(bytes)).toBe(digest)
    }
    const body=original("word/document.xml"), target=[...body.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].filter(m=>m[0].includes(">Same</w:t>"))[1]![0]
    for (const style of ["<w:pStyle/>",'<w:pStyle w:val=""/>']) {
      const changed=body.replace(target,target.replace("<w:pPr>",`<w:pPr>${style}`)); expect(changed).not.toBe(body)
      const bytes=zipSync({...parts,"word/document.xml":strToU8(changed)}), digest=sha(bytes), view=await createDocxRevision(bytes), selected=view.paragraphs.filter(p=>p.text === "Same")[1]!
      await expect(authorExactParagraphSpacing(bytes,{...request,sourceSha256:digest,paragraphHandle:selected.paragraph_handle,tableContext:selected.table})).rejects.toThrow("docx_exact_spacing:unresolved_spacing")
      expect(setter).toHaveBeenCalledTimes(0); expect(sha(bytes)).toBe(digest)
    }
    const alias=zipSync({...parts,"word//styles.xml":parts["word/styles.xml"]!}), aliasHash=sha(alias)
    await expect(authorExactParagraphSpacing(alias,{...request,sourceSha256:aliasHash})).rejects.toThrow("docx_exact_spacing:invalid_part_name")
    expect(setter).toHaveBeenCalledTimes(0); expect(sha(alias)).toBe(aliasHash)
    const duplicate=await new Promise<Uint8Array>((resolve,reject)=>{
      const chunks:Uint8Array[]=[]; const zip=new Zip((error,data,final)=>{if(error)reject(error);else {chunks.push(data);if(final){const result=new Uint8Array(chunks.reduce((n,c)=>n+c.length,0));let offset=0;for(const c of chunks){result.set(c,offset);offset+=c.length}resolve(result)}}})
      for (const [name,bytes] of [...Object.entries(parts),["word/styles.xml",parts["word/styles.xml"]!] as [string,Uint8Array]]) {const entry=new ZipPassThrough(name);zip.add(entry);entry.push(bytes,true)}
      zip.end()
    })
    const digest=sha(duplicate); await expect(authorExactParagraphSpacing(duplicate,{...request,sourceSha256:digest})).rejects.toThrow("docx_exact_spacing:duplicate_part")
    expect(setter).toHaveBeenCalledTimes(0); expect(sha(duplicate)).toBe(digest)
  } finally {setter.mockRestore()}
})

test("inspector preserves duplicate targets, proves every story, and never mutates", async () => {
  const raw=(await renderMarkdownToDocx("Same\n\n| A | B |\n| --- | --- |\n| X | Same |")).bytes, parts=unzipSync(raw)
  const w='http://schemas.openxmlformats.org/wordprocessingml/2006/main', o='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const para='<w:p><w:r><w:t>Same</w:t></w:r></w:p>'
  const stories:[string,string,string][]=[['header1','header',`<w:hdr xmlns:w="${w}">${para}</w:hdr>`],['footer1','footer',`<w:ftr xmlns:w="${w}">${para}</w:ftr>`],['footnotes','footnotes',`<w:footnotes xmlns:w="${w}"><w:footnote w:id="1">${para}</w:footnote></w:footnotes>`],['endnotes','endnotes',`<w:endnotes xmlns:w="${w}"><w:endnote w:id="1">${para}</w:endnote></w:endnotes>`],['comments','comments',`<w:comments xmlns:w="${w}"><w:comment w:id="1" w:author="Reviewer">${para}</w:comment></w:comments>`]]
  let rels=strFromU8(parts['word/_rels/document.xml.rels']!), types=strFromU8(parts['[Content_Types].xml']!), body=strFromU8(parts['word/document.xml']!)
  for(const [file,type,xml] of stories) {parts[`word/${file}.xml`]=strToU8(xml);rels=rels.replace('</Relationships>',`<Relationship Id="story-${file}" Type="${o}/${type}" Target="${file}.xml"/></Relationships>`);types=types.replace('</Types>',`<Override PartName="/word/${file}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/></Types>`)}
  body=body.replace('<w:sectPr>',`<w:sectPr><w:headerReference xmlns:r="${o}" w:type="default" r:id="story-header1"/><w:footerReference xmlns:r="${o}" w:type="default" r:id="story-footer1"/>`)
  body=body.replace('</w:body>',`<w:p><w:r><w:footnoteReference w:id="1"/><w:endnoteReference w:id="1"/></w:r><w:commentRangeStart w:id="1"/><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p></w:body>`)
  parts['word/document.xml']=strToU8(body);parts['word/_rels/document.xml.rels']=strToU8(rels);parts['[Content_Types].xml']=strToU8(types)
  const source=zipSync(parts), digest=sha(source), dx=await docxodus(), session=dx.openDocxSession(source,{}), proto=Object.getPrototypeOf(session);session.close()
  const setter=spyOn(proto,'setParagraphFormat'), save=spyOn(proto,'save')
  try {
    const result=await inspectExactParagraphSpacing(source,'Same')
    expect(result.sourceSha256).toBe(digest); expect(result.matches.length).toBe(7)
    expect(result.matches.filter(m=>m.eligibility.eligible)).toHaveLength(1)
    expect(result.matches.filter(m=>!m.eligibility.eligible && m.eligibility.reason === "unsupported_empty_properties")).toHaveLength(1)
    expect(new Set(result.matches.filter(m=>!m.eligibility.eligible && m.eligibility.reason === "unsupported_target").map(m=>m.scope)).size).toBe(5)
    expect(setter).toHaveBeenCalledTimes(0);expect(save).toHaveBeenCalledTimes(0);expect(sha(source)).toBe(digest)
    for (const content of ['<w:sym w:font="Symbol" w:char="F020"/>','<w:tab/>','<w:br/>','<w:cr/>','<w:contentPart/>','<w:fldChar w:fldCharType="begin"/>','<w:instrText>FIELD</w:instrText>']) {
      const changed=stories[0]![2].replace('</w:hdr>',`<w:p><w:r>${content}</w:r></w:p></w:hdr>`)
      expect(changed).not.toBe(stories[0]![2])
      const altered=zipSync({...parts,'word/header1.xml':strToU8(changed)}), h=sha(altered)
      await expect(inspectExactParagraphSpacing(altered,'Same')).rejects.toThrow('docx_exact_spacing:unsupported_story')
      expect(setter).toHaveBeenCalledTimes(0);expect(save).toHaveBeenCalledTimes(0);expect(sha(altered)).toBe(h)
    }
    let removed=0, bodyRetained=0, headerRetained=0
    const original=proto.grep, grep=spyOn(proto,'grep').mockImplementation(function(this:unknown,...args:unknown[]){
      const found=original.apply(this,args), kept=found.filter((m:{enclosingAnchor:{scope:string}})=>m.enclosingAnchor.scope !== 'hdr1')
      removed+=found.length-kept.length
      bodyRetained+=kept.filter((m:{enclosingAnchor:{scope:string}})=>m.enclosingAnchor.scope === 'body').length
      headerRetained+=kept.filter((m:{enclosingAnchor:{scope:string}})=>m.enclosingAnchor.scope === 'hdr1').length
      return kept
    })
    try {await expect(inspectExactParagraphSpacing(source,'Same')).rejects.toThrow('docx_exact_spacing:incomplete_story')}
    finally {grep.mockRestore()}
    expect(removed).toBe(1);expect(bodyRetained).toBeGreaterThan(0);expect(headerRetained).toBe(0)
    let altered=0
    const partial=spyOn(proto,'grep').mockImplementation(function(this:unknown,...args:unknown[]){
      const found=original.apply(this,args)
      return found.map((m:{text:string;enclosingAnchor:{scope:string}}) => {
        if (m.enclosingAnchor.scope === 'body' && m.text === 'Same' && altered === 0) {altered++;return {...m,text:'Sam'}}
        return m
      })
    })
    try {await expect(inspectExactParagraphSpacing(source,'Same')).rejects.toThrow('docx_exact_spacing:incomplete_story')}
    finally {partial.mockRestore()}
    expect(altered).toBe(1)
    expect(setter).toHaveBeenCalledTimes(0);expect(save).toHaveBeenCalledTimes(0);expect(sha(source)).toBe(digest)
  } finally {setter.mockRestore();save.mockRestore()}
})

test("inspector preserves raw direct values and author refuses semantic no-ops", async () => {
  const raw=(await renderMarkdownToDocx("| A | B |\n| --- | --- |\n| X | Same |")).bytes, parts=unzipSync(raw), doc=strFromU8(parts['word/document.xml']!)
  const target=[...doc.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].find(m=>m[0].includes('>Same</w:t>'))![0]
  const mutate=(value:string|null)=>zipSync({...parts,'word/document.xml':strToU8(doc.replace(target,target.replace('w:after="80"',value===null?'':`w:after="${value}"`)))})
  const source=mutate('0080'), digest=sha(source), session=(await docxodus()).openDocxSession(source,{}), proto=Object.getPrototypeOf(session);session.close(); const setter=spyOn(proto,'setParagraphFormat'),save=spyOn(proto,'save')
  try {
    const inspected=await inspectExactParagraphSpacing(source,'Same'), row=inspected.matches[0]!
    expect(row.directAfter).toEqual({kind:'twips',raw:'0080',normalized:'80'});expect(row.eligibility).toEqual({eligible:true})
    await expect(authorExactParagraphSpacing(source,{sourceSha256:digest,paragraphHandle:row.paragraphHandle,paragraphText:row.text,tableContext:row.tableContext,spacingAfterPoints:4,author:'Synthetic Counsel'})).rejects.toThrow('docx_exact_spacing:no_change')
    const absent=await inspectExactParagraphSpacing(mutate(null),'Same');expect(absent.matches[0]!.directAfter).toEqual({kind:'absent'})
    const bad=await inspectExactParagraphSpacing(mutate('80x'),'Same');expect(bad.matches[0]!.directAfter).toEqual({kind:'unavailable'});expect(bad.matches[0]!.eligibility).toEqual({eligible:false,reason:'unsupported_spacing'})
    expect((await inspectExactParagraphSpacing(source,'not present')).matches).toEqual([])
    await expect(inspectExactParagraphSpacing(source,'')).rejects.toThrow('docx_exact_spacing:invalid_query')
    expect(setter).toHaveBeenCalledTimes(0);expect(save).toHaveBeenCalledTimes(0);expect(sha(source)).toBe(digest)
  } finally {setter.mockRestore();save.mockRestore()}
})

test("inspector refuses excessive duplicates and excludes proven deleted-only text", async () => {
  const large=(await renderMarkdownToDocx(Array(65).fill('Same').join('\n\n'))).bytes, largeHash=sha(large)
  await expect(inspectExactParagraphSpacing(large,'Same')).rejects.toThrow('docx_exact_spacing:match_limit')
  expect(sha(large)).toBe(largeHash)
  const raw=(await renderMarkdownToDocx('Keep\n\nLost')).bytes, parts=unzipSync(raw), doc=strFromU8(parts['word/document.xml']!)
  const changed=doc.replace('<w:r><w:t xml:space="preserve">Lost</w:t></w:r>','<w:del w:id="901" w:author="Reviewer" w:date="2026-10-29T00:00:00Z"><w:r><w:delText xml:space="preserve">Lost</w:delText></w:r></w:del>')
  expect(changed).not.toBe(doc);const source=zipSync({...parts,'word/document.xml':strToU8(changed)}), digest=sha(source)
  expect((await inspectExactParagraphSpacing(source,'Keep')).matches).toHaveLength(1)
  expect((await inspectExactParagraphSpacing(source,'Lost')).matches).toHaveLength(0)
  expect(sha(source)).toBe(digest)
})

test("supplied formatting artifact is independently verified and selectively resolved", async () => {
  const source=(await renderMarkdownToDocx('| A | B |\n| --- | --- |\n| X | Same |')).bytes, view=await createDocxRevision(source), row=view.paragraphs.find(p=>p.text==='Same')!
  const request={sourceSha256:sha(source),paragraphHandle:row.paragraph_handle,paragraphText:row.text,tableContext:row.table,spacingAfterPoints:6,author:'Synthetic Counsel'}
  const authored=await authorExactParagraphSpacing(source,request), sourceHash=sha(source), artifactHash=sha(authored.trackedBytes)
  const proof=await verifyExactParagraphSpacing(source,authored.trackedBytes,request)
  expect(proof.afterDirect).toBe('120'); expect(proof.revisionId).toBe(authored.revisionId)
  const rejected=await resolveExactParagraphSpacing(source,authored.trackedBytes,request,'reject'), accepted=await resolveExactParagraphSpacing(source,authored.trackedBytes,request,'accept')
  expect(rejected).toEqual(source); expect(sha(source)).toBe(sourceHash); expect(sha(authored.trackedBytes)).toBe(artifactHash)
  const dx=await docxodus(), sess=dx.openDocxSession(accepted,{persistAnchorIds:false});try{expect(sess.listRevisions()).toHaveLength(0)}finally{sess.close()}
  const parts=unzipSync(authored.trackedBytes), xml=strFromU8(parts['word/document.xml']!)
  const mark=xml.match(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/)![0]
  const priorId=xml.match(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/)![1]!
  for (const [changed,code] of [
    [xml.replace('w:author="Synthetic Counsel"','w:author="Other"'),'invalid_revision'],
    [xml.replace(mark,mark.replace('w:after="80"','w:after="81"')),'revision_old_properties_changed'],
    [xml.replace(mark,mark+mark),'revision_membership'],
    [xml.replace(mark,mark.replace(/w:id="\d+"/,`w:id="${priorId}"`)),'revision_id_collision'],
    [xml.replace(mark,mark.replace(/w:date="[^"]+"/,'w:date="2026-15-01T00:00:00.000Z"')),'invalid_revision'],
    [xml.replace('>X</w:t>','>Y</w:t>'),'other_body_changed'],
    [xml.replace('w:after="120"','w:after="121"'),'wrong_after'],
    [xml.replace(/<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/,''),'revision_membership'],
    [xml.replace('w:author="Synthetic Counsel"','w:author="Synthetic Counsel" w:extra="x"'),'invalid_revision'],
  ]) {
    expect(changed).not.toBe(xml);const artifact=zipSync({...parts,'word/document.xml':strToU8(changed)})
    await expect(verifyExactParagraphSpacing(source,artifact,request)).rejects.toThrow(`docx_exact_spacing:${code}`)
    expect(sha(source)).toBe(sourceHash)
  }
  const changedStyles=zipSync({...parts,'word/styles.xml':strToU8(strFromU8(parts['word/styles.xml']!).replace('</w:styles>','<w:style w:type="paragraph" w:styleId="Injected"/></w:styles>'))})
  await expect(verifyExactParagraphSpacing(source,changedStyles,request)).rejects.toThrow('docx_exact_spacing:other_part_changed')
  const session=dx.openDocxSession(authored.trackedBytes,{}),proto=Object.getPrototypeOf(session);session.close()
  const realBatch=proto.executeBatch; let decisionApplied=false, rolledBack=false
  const spy=spyOn(proto,'executeBatch').mockImplementation(function(this:typeof session,...args:Parameters<typeof realBatch>) {
    const before=this.getPackageManifest().rawPackageBytesDigest, version=this.getVersion()
    const fail={tool:'paragraph',action:'injected-failure',mutation:()=>{decisionApplied=this.listRevisions().length===0;return {success:false,created:[],removed:[],modified:[],error:{code:'injected',message:'Synthetic failure'}}}}
    const result=realBatch.call(this,[...args[0],fail],args[1]), after=this.getPackageManifest().rawPackageBytesDigest
    rolledBack=!result.success && this.getVersion()===version && after.algorithm===before.algorithm && after.value===before.value
    return result
  })
  try {await expect(resolveExactParagraphSpacing(source,authored.trackedBytes,request,'accept')).rejects.toThrow('docx_exact_spacing:resolution_failed_rolled_back');expect(decisionApplied).toBe(true);expect(rolledBack).toBe(true)}
  finally {spy.mockRestore()}
  expect(sha(source)).toBe(sourceHash);expect(sha(authored.trackedBytes)).toBe(artifactHash)
})
