import { expect, spyOn, test } from "bun:test"
import { createHash } from "node:crypto"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync, Zip, ZipPassThrough } from "fflate"
import { createDocxRevision } from "./docx-amendments"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxCommentParagraphPositions, readDocxComments } from "./docx-comments"
import { authorExactParagraphSpacing, resolveExactParagraphSpacing, type ExactSpacingRequest } from "./docx-exact-spacing"
import { authorExactParagraphSpacingReview, resolveExactParagraphSpacingReview, verifyExactParagraphSpacingReview } from "./docx-exact-spacing-review"
import { renderMarkdownToDocx } from "./render"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", XML = "http://www.w3.org/XML/1998/namespace"
const R = "http://schemas.openxmlformats.org/package/2006/relationships", O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "http://schemas.openxmlformats.org/package/2006/content-types", BODY = "word/document.xml", COMMENTS = "word/comments.xml"
const RELS = "word/_rels/document.xml.rels", TYPES = "[Content_Types].xml", AUTHOR = "Synthetic Counsel"
const REASON = "Explains the requested paragraph spacing.\nLiteral **note** & <review> 😀."
const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
const xml = (node: Node) => new XMLSerializer().serializeToString(node)
const all = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS(W, name))
const hash = (bytes: Uint8Array) => createHash("sha256").update(bytes).digest("hex")
function changed(bytes: Uint8Array, part: string, change: (doc: Document) => void) {
  const parts = unzipSync(bytes), doc = parse(parts[part]!); change(doc)
  parts[part] = strToU8(xml(doc)); return zipSync(parts)
}
async function fixture(prior = false, text = "Same") {
  let bytes = (await renderMarkdownToDocx("Before paragraph.\n\nComment paragraph.\n\n| Label | Value |\n| --- | --- |\n| A | Same |\n| B | Same |")).bytes
  if (text !== "Same") bytes = changed(bytes, BODY, doc => {
    for (const node of all(doc, "t").filter(node => node.textContent === "Same")) { node.textContent = text; node.setAttributeNS(XML, "xml:space", "preserve") }
  })
  if (!prior) return bytes
  bytes = changed(bytes, BODY, doc => {
    const body = all(doc, "body")[0]!, stamp = `w:author="${AUTHOR}" w:date="2026-01-01T00:00:00Z"`
    const removed = parse(strToU8(`<w:p xmlns:w="${W}"><w:del w:id="700" ${stamp}><w:r><w:delText>Earlier deleted paragraph.</w:delText></w:r></w:del></w:p>`)).documentElement!
    body.insertBefore(doc.importNode(removed, true), body.firstChild)
    const before = all(doc, "p").find(p => all(p, "t").some(t => t.textContent === "Before paragraph."))!, run = all(before, "r")[0]!
    const insertion = doc.createElementNS(W, "w:ins"); insertion.setAttributeNS(W, "w:id", "701"); insertion.setAttributeNS(W, "w:author", AUTHOR); insertion.setAttributeNS(W, "w:date", "2026-01-01T00:00:00Z")
    before.replaceChild(insertion, run); insertion.appendChild(run)
  })
  bytes = exportDocxComments(bytes, [{ id: "prior", text: "Keep original note.", author: "Prior Reviewer", anchor: { paragraph_text: "Comment paragraph." } }]).bytes
  const parts = unzipSync(bytes), comments = parse(parts[COMMENTS]!), body = parse(parts[BODY]!), rels = parse(parts[RELS]!), types = parse(parts[TYPES]!)
  const w14 = "http://schemas.microsoft.com/office/word/2010/wordml", w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
  comments.documentElement!.setAttribute("xmlns:w14", w14); comments.documentElement!.setAttribute("xmlns:r", O)
  const first = all(comments, "p")[0]!; first.setAttributeNS(w14, "w14:paraId", "ABCDEF01")
  const link = comments.createElementNS(W, "w:hyperlink"), run = comments.createElementNS(W, "w:r"), word = comments.createElementNS(W, "w:t")
  link.setAttributeNS(O, "r:id", "priorLink"); word.textContent = " Link"; run.appendChild(word); link.appendChild(run); first.appendChild(link)
  comments.documentElement!.appendChild(comments.importNode(parse(strToU8(`<w:comment xmlns:w="${W}" xmlns:w14="${w14}" w:id="1" w:author="Prior Reviewer"><w:p w14:paraId="ABCDEF02"><w:r><w:t>Keep original reply.</w:t></w:r></w:p></w:comment>`)).documentElement!, true))
  const host = all(body, "p").find(p => all(p, "t").some(t => t.textContent === "Comment paragraph."))!, refRun = body.createElementNS(W, "w:r"), ref = body.createElementNS(W, "w:commentReference")
  ref.setAttributeNS(W, "w:id", "1"); refRun.appendChild(ref); host.appendChild(refRun)
  const rel = rels.createElementNS(R, "Relationship"); rel.setAttribute("Id", "priorThread"); rel.setAttribute("Type", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"); rel.setAttribute("Target", "commentsExtended.xml"); rels.documentElement!.appendChild(rel)
  const override = types.createElementNS(CT, "Override"); override.setAttribute("PartName", "/word/commentsExtended.xml"); override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"); types.documentElement!.appendChild(override)
  parts["word/commentsExtended.xml"] = strToU8(`<w15:commentsEx xmlns:w15="${w15}"><w15:commentEx w15:paraId="ABCDEF01" w15:done="0"/><w15:commentEx w15:paraId="ABCDEF02" w15:paraIdParent="ABCDEF01" w15:done="0"/></w15:commentsEx>`)
  parts["word/_rels/comments.xml.rels"] = strToU8(`<Relationships xmlns="${R}"><Relationship Id="priorLink" Type="${O}/hyperlink" Target="https://example.com/review" TargetMode="External"/></Relationships>`)
  for (const [name, doc] of [[COMMENTS, comments], [BODY, body], [RELS, rels], [TYPES, types]] as const) parts[name] = strToU8(xml(doc))
  return zipSync(parts)
}
async function request(bytes: Uint8Array, text = "Same"): Promise<ExactSpacingRequest> {
  const view = await createDocxRevision(bytes, { pendingRevisions: "layer" }), matches = view.paragraphs.filter(p => p.scope === "body" && p.text === text)
  expect(matches).toHaveLength(2)
  const selected = matches[1]!
  return { sourceSha256: hash(bytes), paragraphHandle: selected.paragraph_handle, paragraphText: selected.text, tableContext: selected.table, spacingAfterPoints: 6, author: AUTHOR }
}
async function registry(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return session.listRevisions().map(({ anchorId: _a, affectedAnchors: _b, ...record }) => record).sort((a, b) => a.id.localeCompare(b.id)) } finally { session.close() }
}
function physical(bytes: Uint8Array, names: string[], ids?: Set<string>) {
  const doc = parse(unzipSync(bytes)[BODY]!)
  return names.flatMap(name => all(doc, name).filter(node => !ids || ids.has(node.getAttributeNS(W, "id")!)).map(node => {
    const path: number[] = []; let current: Node = node
    while (current.parentNode && current.parentNode.nodeType !== 9) { path.unshift(Array.from(current.parentNode.childNodes).indexOf(current as ChildNode)); current = current.parentNode }
    return { name, path, subtree: xml(node) }
  }))
}
async function duplicate(bytes: Uint8Array) {
  const parts = unzipSync(bytes)
  return await new Promise<Uint8Array>((resolve, reject) => {
    const chunks: Uint8Array[] = [], zip = new Zip((error, data, final) => {
      if (error) reject(error)
      else { chunks.push(data); if (final) { const result = new Uint8Array(chunks.reduce((n, chunk) => n + chunk.length, 0)); let offset = 0; for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length } resolve(result) } }
    })
    for (const [name, bytes] of [...Object.entries(parts), [BODY, parts[BODY]!] as [string, Uint8Array]]) { const entry = new ZipPassThrough(name); zip.add(entry); entry.push(bytes, true) }
    zip.end()
  })
}

test("generated duplicate cells attach the reason to the verified physical formatting target and support both decisions", async () => {
  const source = await fixture(), original = new Uint8Array(source), input = await request(source), result = await authorExactParagraphSpacingReview(source, input, REASON)
  const proof = await verifyExactParagraphSpacingReview(source, result.bytes, input, REASON)
  expect(proof.commentId).toBe(result.commentId); expect(proof.afterDirect).toBe("120")
  const body = parse(unzipSync(result.bytes)[BODY]!), targets = all(body, "p").filter(p => all(p, "t").some(t => t.textContent === "Same"))
  expect(targets.map(p => all(p, "commentRangeStart").length)).toEqual([0, 1])
  expect(targets.map(p => all(p, "pPrChange").length)).toEqual([0, 1])
  expect(readDocxComments(result.bytes).map(comment => [comment.comment_id, comment.text, comment.markers.length])).toEqual([[result.commentId, REASON, 3]])
  const accepted = await resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "accept")
  expect(await registry(accepted)).toEqual([])
  expect(readDocxComments(accepted)).toEqual(readDocxComments(result.bytes))
  expect(await resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "reject")).toEqual(source)
  const core = await authorExactParagraphSpacing(source, input), clean = await resolveExactParagraphSpacing(source, core.trackedBytes, input, "accept")
  expect(readDocxComments(clean)).toEqual([])
  expect(source).toEqual(original)
})

test("whitespace and shifted accepted ordinals preserve old threads, dependencies and same-author carriers", async () => {
  const source = await fixture(true, "  "), input = await request(source, "  "), original = unzipSync(source), oldComments = readDocxComments(source), oldIds = new Set(oldComments.map(comment => comment.comment_id))
  const oldMarkers = physical(source, ["commentRangeStart", "commentRangeEnd", "commentReference"], oldIds), oldCarriers = physical(source, ["ins", "del"])
  expect(oldCarriers).toHaveLength(2)
  const result = await authorExactParagraphSpacingReview(source, input, REASON), own = readDocxComments(result.bytes).find(comment => comment.comment_id === result.commentId)!
  expect(own.markers.every(marker => marker.accepted_paragraph_text === "  ")).toBe(true)
  expect(own.markers.every(marker => marker.xml_paragraph_index! > marker.paragraph_index!)).toBe(true)
  const accepted = await resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "accept")
  for (const bytes of [result.bytes, accepted]) {
    expect(physical(bytes, ["ins", "del"])).toEqual(oldCarriers)
    expect(physical(bytes, ["commentRangeStart", "commentRangeEnd", "commentReference"], oldIds)).toEqual(oldMarkers)
    expect(readDocxComments(bytes).filter(comment => oldIds.has(comment.comment_id))).toEqual(oldComments)
    const parts = unzipSync(bytes)
    for (const name of [RELS, TYPES, "word/commentsExtended.xml", "word/_rels/comments.xml.rels", "word/styles.xml", "word/settings.xml"]) expect(parts[name]).toEqual(original[name])
  }
  expect(await registry(accepted)).toEqual(await registry(source))
  expect(await resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "reject")).toEqual(source)
})

test("foreign, moved, partial or duplicated reasons and modified source dependencies refuse supplied artifacts", async () => {
  const source = await fixture(true), input = await request(source), result = await authorExactParagraphSpacingReview(source, input, REASON), originalHash = hash(result.bytes)
  const own = (doc: Document, name: string) => all(doc, name).find(node => node.getAttributeNS(W, "id") === result.commentId)!
  const moved = changed(result.bytes, BODY, doc => {
    const target = all(doc, "p").filter(p => all(p, "t").some(t => t.textContent === "Same"))[0]!, props = Array.from(target.childNodes).find(node => node.nodeType === 1 && (node as Element).namespaceURI === W && (node as Element).localName === "pPr")
    const start = own(doc, "commentRangeStart"), end = own(doc, "commentRangeEnd"), reference = own(doc, "commentReference").parentNode!
    target.insertBefore(start, props?.nextSibling ?? target.firstChild); target.appendChild(end); target.appendChild(reference)
  })
  await expect(verifyExactParagraphSpacingReview(source, moved, input, REASON)).rejects.toThrow("reason_not_on_formatting_target")
  const partial = changed(result.bytes, BODY, doc => { const start = own(doc, "commentRangeStart"), parent = start.parentNode!; parent.removeChild(start); parent.insertBefore(start, own(doc, "commentRangeEnd")) })
  await expect(verifyExactParagraphSpacingReview(source, partial, input, REASON)).rejects.toThrow("owned_comment_range_changed")
  const tampered = [
    changed(result.bytes, COMMENTS, doc => own(doc, "comment").setAttributeNS(W, "w:author", "Other")),
    changed(result.bytes, COMMENTS, doc => { all(doc, "t").at(-1)!.textContent = "Different reason" }),
    moved, partial,
    changed(result.bytes, BODY, doc => { const marker = own(doc, "commentReference"); marker.parentNode!.appendChild(marker.cloneNode(true)) }),
    changed(result.bytes, RELS, doc => { Array.from(doc.getElementsByTagNameNS(R, "Relationship")).find(node => node.getAttribute("Type")?.endsWith("/comments"))!.setAttribute("Target", "other.xml") }),
    changed(result.bytes, COMMENTS, doc => { all(doc, "comment")[0]!.setAttributeNS(W, "w:author", "Changed prior author") }),
    changed(result.bytes, "word/commentsExtended.xml", doc => { doc.documentElement!.setAttribute("changed", "yes") }),
    changed(result.bytes, BODY, doc => { all(doc, "pPrChange")[0]!.getElementsByTagNameNS(W, "spacing")[0]!.setAttributeNS(W, "w:after", "81") }),
  ]
  for (const bytes of tampered) {
    expect(hash(bytes)).not.toBe(originalHash)
    await expect(verifyExactParagraphSpacingReview(source, bytes, input, REASON)).rejects.toThrow()
    await expect(resolveExactParagraphSpacingReview(source, bytes, input, REASON, "accept")).rejects.toThrow()
  }
  await expect(verifyExactParagraphSpacingReview(source, result.bytes, input, "Different reason")).rejects.toThrow("comment_content_or_prior_comments_changed")
  expect(hash(result.bytes)).toBe(originalHash)
  expect(await resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "reject")).toEqual(source)
})

test("raw duplicate ZIP entries refuse before annotation peeling or native decisions", async () => {
  const source = await fixture(), input = await request(source), result = await authorExactParagraphSpacingReview(source, input, REASON), duplicateArtifact = await duplicate(result.bytes), duplicateSource = await duplicate(source)
  const session = (await docxodus()).openDocxSession(result.bytes, {}), proto = Object.getPrototypeOf(session) as typeof session; session.close()
  const accept = spyOn(proto, "acceptRevision"), reject = spyOn(proto, "rejectRevision"), setter = spyOn(proto, "setParagraphFormat")
  try {
    await expect(verifyExactParagraphSpacingReview(source, duplicateArtifact, input, REASON)).rejects.toThrow("docx_exact_spacing:duplicate_part")
    await expect(resolveExactParagraphSpacingReview(source, duplicateArtifact, input, REASON, "reject")).rejects.toThrow("docx_exact_spacing:duplicate_part")
    await expect(authorExactParagraphSpacingReview(duplicateSource, { ...input, sourceSha256: hash(duplicateSource) }, REASON)).rejects.toThrow("docx_exact_spacing:duplicate_part")
    expect(accept).not.toHaveBeenCalled(); expect(reject).not.toHaveBeenCalled(); expect(setter).not.toHaveBeenCalled()
  } finally { accept.mockRestore(); reject.mockRestore(); setter.mockRestore() }
  const accepted = await resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "accept"), duplicateSaved = await duplicate(accepted), save = proto.save
  let injected = 0
  const saved = spyOn(proto, "save").mockImplementation(function(this: typeof session) {
    const bytes = save.call(this)
    if (this.listComments().some(comment => String(comment.id) === result.commentId) && this.listRevisions().length === 0) { injected++; return duplicateSaved }
    return bytes
  })
  try {
    await expect(resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, "accept")).rejects.toThrow("docx_exact_spacing:duplicate_part")
    expect(injected).toBe(1)
  } finally { saved.mockRestore() }
})

test("reason validation precedes authoring and one normalized author signs comment and format", async () => {
  const source = await fixture(), input = await request(source), session = (await docxodus()).openDocxSession(source, {}), proto = Object.getPrototypeOf(session) as typeof session; session.close()
  const setter = spyOn(proto, "setParagraphFormat")
  try {
    for (const reason of ["", " ", "x".repeat(4001), "Bad\u0000reason"]) await expect(authorExactParagraphSpacingReview(source, input, reason)).rejects.toThrow("invalid_reason")
    expect(setter).not.toHaveBeenCalled()
  } finally { setter.mockRestore() }
  const result = await authorExactParagraphSpacingReview(source, { ...input, author: ` ${AUTHOR} ` }, REASON)
  const comments = parse(unzipSync(result.bytes)[COMMENTS]!)
  expect(all(comments, "comment").map(comment => comment.getAttributeNS(W, "author"))).toEqual([AUTHOR])
  expect((await registry(result.bytes)).map(record => record.author)).toEqual([AUTHOR])
})

test("real native decision then failed step rolls back format and owned reason on the final artifact", async () => {
  const source = await fixture(), input = await request(source), result = await authorExactParagraphSpacingReview(source, input, REASON), sourceHash = hash(source), artifactHash = hash(result.bytes)
  const session = (await docxodus()).openDocxSession(result.bytes, {}), proto = Object.getPrototypeOf(session) as typeof session; session.close()
  const batch = proto.executeBatch
  for (const decision of ["accept", "reject"] as const) {
    let changed = false, restored = false
    const spy = spyOn(proto, "executeBatch").mockImplementation(function(this: typeof session, ...args: Parameters<typeof batch>) {
      // Core expectation checks also use a batch; inject only into a session
      // carrying the actual final artifact's owned explanation.
      if (!this.listComments().some(comment => String(comment.id) === result.commentId)) return batch.apply(this, args)
      const before = this.getPackageManifest().rawPackageBytesDigest, version = this.getVersion()
      const fail: Parameters<typeof batch>[0][number] = { tool: "paragraph", action: "injected_failure", mutation: () => {
        changed = this.listRevisions().length === 0 && this.listComments().some(comment => String(comment.id) === result.commentId) === (decision === "accept")
        return { success: false, created: [], removed: [], modified: [], error: { code: "internal_error", message: "Synthetic later failure" } }
      } }
      const outcome = batch.call(this, [...args[0], fail], args[1]), after = this.getPackageManifest().rawPackageBytesDigest
      restored = !outcome.success && this.getVersion() === version && after.algorithm === before.algorithm && after.value === before.value && this.listRevisions().length === 1 && this.listComments().some(comment => String(comment.id) === result.commentId)
      return outcome
    })
    try { await expect(resolveExactParagraphSpacingReview(source, result.bytes, input, REASON, decision)).rejects.toThrow("review_failed_rolled_back"); expect(changed).toBe(true); expect(restored).toBe(true) }
    finally { spy.mockRestore() }
  }
  expect(hash(source)).toBe(sourceHash); expect(hash(result.bytes)).toBe(artifactHash)
})
