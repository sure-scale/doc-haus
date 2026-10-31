import { expect, spyOn, test } from "bun:test"
import { DOMParser, XMLSerializer } from "@xmldom/xmldom"
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { docxodus } from "./docxodus"
import { exportDocxComments, readDocxComments } from "./docx-comments"
import { docxTableRowInsertionTargets, type DocxTableRowInsertion } from "./docx-table-row-insertion"
import { authorDocxTableRowReview, resolveDocxTableRowReview } from "./docx-table-row-review"

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const R = "http://schemas.openxmlformats.org/package/2006/relationships", CT = "http://schemas.openxmlformats.org/package/2006/content-types"
const BODY = "word/document.xml", COMMENTS = "word/comments.xml", RELS = "word/_rels/document.xml.rels", TYPES = "[Content_Types].xml"
const REASON = "Records the requested new row.\nLiteral **note** & <review> 😀."
const AUTHOR = "Review Editor"
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${content}</w:tc>`
const parse = (bytes: Uint8Array) => new DOMParser().parseFromString(strFromU8(bytes), "application/xml")
const serialize = (node: Node) => new XMLSerializer().serializeToString(node)
function source(withPrior = false) {
  const prior = withPrior ? `<w:p><w:del w:id="70" w:author="${AUTHOR}" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>Former</w:delText></w:r></w:del><w:ins w:id="71" w:author="${AUTHOR}" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Right</w:t></w:r></w:ins></w:p>` : p("Right")
  let bytes = zipSync({
    [TYPES]: strToU8(`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${BODY}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${R}"><Relationship Id="main" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`),
    [RELS]: strToU8(`<Relationships xmlns="${R}"><Relationship Id="styles" Type="${O}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${O}/settings" Target="settings.xml"/></Relationships>`),
    "word/styles.xml": strToU8(`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`),
    "word/settings.xml": strToU8(`<w:settings xmlns:w="${W}"/>`),
    [BODY]: strToU8(`<w:document xmlns:w="${W}" xmlns:r="${O}"><w:body>${p("Same")}<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr/>${cell(p("Same"))}${cell(prior)}</w:tr><w:tr><w:trPr/>${cell(p("Last"))}${cell(p("Stable"))}</w:tr></w:tbl>${p("After")}<w:sectPr/></w:body></w:document>`),
  }, { mtime: new Date(1980, 0, 1) })
  if (!withPrior) return bytes
  bytes = exportDocxComments(bytes, [{ id: "original", text: "Original reason.", author: "Prior Reviewer", anchor: { paragraph_text: "After" } }]).bytes
  const parts = unzipSync(bytes), comments = parse(parts[COMMENTS]!), body = parse(parts[BODY]!), rels = parse(parts[RELS]!), types = parse(parts[TYPES]!)
  const w14 = "http://schemas.microsoft.com/office/word/2010/wordml", w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
  comments.documentElement!.setAttribute("xmlns:w14", w14); comments.documentElement!.setAttribute("xmlns:r", O)
  const original = comments.getElementsByTagNameNS(W, "p")[0]!
  original.setAttributeNS(w14, "w14:paraId", "ABCDEF01")
  const hyperlink = comments.createElementNS(W, "w:hyperlink"); hyperlink.setAttributeNS(O, "r:id", "cLink")
  const run = comments.createElementNS(W, "w:r"), text = comments.createElementNS(W, "w:t")
  text.textContent = " Link"; run.appendChild(text); hyperlink.appendChild(run); original.appendChild(hyperlink)
  const reply = parse(strToU8(`<w:comment xmlns:w="${W}" xmlns:w14="${w14}" w:id="1" w:author="Prior Reviewer"><w:p w14:paraId="ABCDEF02"><w:r><w:t>Original reply.</w:t></w:r></w:p></w:comment>`)).documentElement!
  comments.documentElement!.appendChild(comments.importNode(reply, true))
  const host = Array.from(body.getElementsByTagNameNS(W, "p")).find(p => p.textContent === "After")!, referenceRun = body.createElementNS(W, "w:r"), reference = body.createElementNS(W, "w:commentReference")
  reference.setAttributeNS(W, "w:id", "1"); referenceRun.appendChild(reference); host.appendChild(referenceRun)
  const extended = rels.createElementNS(R, "Relationship")
  extended.setAttribute("Id", "thread"); extended.setAttribute("Type", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"); extended.setAttribute("Target", "commentsExtended.xml"); rels.documentElement!.appendChild(extended)
  const override = types.createElementNS(CT, "Override")
  override.setAttribute("PartName", "/word/commentsExtended.xml"); override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"); types.documentElement!.appendChild(override)
  parts["word/commentsExtended.xml"] = strToU8(`<w15:commentsEx xmlns:w15="${w15}"><w15:commentEx w15:paraId="ABCDEF01" w15:done="0"/><w15:commentEx w15:paraId="ABCDEF02" w15:paraIdParent="ABCDEF01" w15:done="0"/></w15:commentsEx>`)
  parts["word/_rels/comments.xml.rels"] = strToU8(`<Relationships xmlns="${R}"><Relationship Id="cLink" Type="${O}/hyperlink" Target="https://example.com/review" TargetMode="External"/></Relationships>`)
  for (const [name, doc] of [[COMMENTS, comments], [BODY, body], [RELS, rels], [TYPES, types]] as const) parts[name] = strToU8(serialize(doc))
  return zipSync(parts)
}
async function request(bytes: Uint8Array, cells = ["New left", "New right"]): Promise<DocxTableRowInsertion> {
  const packet = await docxTableRowInsertionTargets(bytes)
  return { source_sha256: packet.source_sha256, target_handle: packet.targets[0]!.target_handle, position: "after", cells, author: AUTHOR }
}
function changed(bytes: Uint8Array, part: string, action: (doc: Document) => void) {
  const parts = unzipSync(bytes), doc = parse(parts[part]!)
  action(doc); parts[part] = strToU8(serialize(doc)); return zipSync(parts)
}
async function registry(bytes: Uint8Array) {
  const session = (await docxodus()).openDocxSession(bytes, { persistAnchorIds: false })
  try { return session.listRevisions().map(({ anchorId: _anchor, affectedAnchors: _affected, ...record }) => record) } finally { session.close() }
}

test("final commented row artifact creates only its comment dependencies and passes grouped selective review", async () => {
  const bytes = source(), original = new Uint8Array(bytes), input = await request(bytes), result = await authorDocxTableRowReview(bytes, input, REASON)
  const comments = readDocxComments(result.bytes)
  expect(comments).toHaveLength(1)
  expect(comments[0]!.text).toBe(REASON)
  expect(comments[0]!.markers.map(marker => marker.kind)).toEqual(["commentRangeStart", "commentRangeEnd", "commentReference"])
  expect(new Set(comments[0]!.markers.map(marker => marker.accepted_paragraph_text))).toEqual(new Set(["New left"]))
  expect(result.revisions.map(r => r.family).sort()).toEqual(["content_insert", "content_insert", "row_insert"])
  const accepted = await resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "accept")
  expect(readDocxComments(accepted)).toEqual(comments)
  expect(await registry(accepted)).toEqual([])
  expect(await resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "reject")).toEqual(bytes)
  expect(bytes).toEqual(original)
})

test("reason targets first nonempty new cell despite duplicate words, preserving old threaded comments and same-author revisions", async () => {
  const bytes = source(true), original = unzipSync(bytes), input = await request(bytes, ["", "Same"]), result = await authorDocxTableRowReview(bytes, input, REASON)
  const comments = readDocxComments(result.bytes), own = comments.find(comment => comment.comment_id === result.commentId)!
  expect(comments.slice(0, 2).map(({ comment_id, text }) => ({ comment_id, text }))).toEqual(readDocxComments(bytes).map(({ comment_id, text }) => ({ comment_id, text })))
  expect(own.markers.every(marker => marker.accepted_paragraph_text === "Same")).toBe(true)
  const body = parse(unzipSync(result.bytes)[BODY]!), newRow = body.getElementsByTagNameNS(W, "tr")[1]!, cells = Array.from(newRow.getElementsByTagNameNS(W, "tc"))
  expect(cells[0]!.getElementsByTagNameNS(W, "commentRangeStart").length).toBe(0)
  expect(cells[1]!.getElementsByTagNameNS(W, "commentRangeStart")[0]!.getAttributeNS(W, "id")).toBe(result.commentId)
  for (const name of ["word/commentsExtended.xml", "word/_rels/comments.xml.rels", RELS, TYPES]) expect(unzipSync(result.bytes)[name]).toEqual(original[name])
  const accepted = await resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "accept")
  expect(await registry(accepted)).toEqual(await registry(bytes))
  expect(readDocxComments(accepted).map(comment => comment.text)).toEqual(comments.map(comment => comment.text))
  expect(await resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "reject")).toEqual(bytes)
})

test("foreign reason, moved or duplicated markers and dependency changes refuse final artifact review", async () => {
  const bytes = source(), input = await request(bytes), result = await authorDocxTableRowReview(bytes, input, REASON)
  const foreign = changed(result.bytes, COMMENTS, doc => doc.getElementsByTagNameNS(W, "comment")[0]!.setAttributeNS(W, "w:author", "Other"))
  const moved = changed(result.bytes, BODY, doc => { const marker = doc.getElementsByTagNameNS(W, "commentRangeStart")[0]!; doc.getElementsByTagNameNS(W, "p")[0]!.appendChild(marker) })
  const duplicated = changed(result.bytes, BODY, doc => { const marker = doc.getElementsByTagNameNS(W, "commentReference")[0]!; marker.parentNode!.appendChild(marker.cloneNode(true)) })
  const dependency = changed(result.bytes, RELS, doc => { Array.from(doc.getElementsByTagNameNS(R, "Relationship")).find(el => el.getAttribute("Type").endsWith("/comments"))!.setAttribute("Target", "other.xml") })
  const style = changed(result.bytes, "word/styles.xml", doc => doc.getElementsByTagNameNS(W, "name")[0]!.setAttributeNS(W, "w:val", "Changed"))
  for (const artifact of [foreign, moved, duplicated, dependency, style]) await expect(resolveDocxTableRowReview(bytes, artifact, input, REASON, "accept")).rejects.toThrow("docx_row_review:")
  await expect(resolveDocxTableRowReview(bytes, result.bytes, input, "Different reason.", "accept")).rejects.toThrow("comment_content_or_prior_comments_changed")
  expect(await resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "reject")).toEqual(bytes)
})

test("one canonical author attributes both row revisions and its reason", async () => {
  const bytes = source(), input = { ...await request(bytes), author: ` ${AUTHOR} ` }, result = await authorDocxTableRowReview(bytes, input, REASON)
  const commentXml = parse(unzipSync(result.bytes)[COMMENTS]!)
  expect(Array.from(commentXml.getElementsByTagNameNS(W, "comment")).map(comment => comment.getAttributeNS(W, "author"))).toEqual([AUTHOR])
  expect(new Set((await registry(result.bytes)).map(revision => revision.author))).toEqual(new Set([AUTHOR]))
  expect(await resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "reject")).toEqual(bytes)
})

test("invalid reason refuses before core row authoring", async () => {
  const bytes = source(), input = await request(bytes), temp = (await docxodus()).openDocxSession(bytes, {}), prototype = Object.getPrototypeOf(temp) as typeof temp
  temp.close(); const insert = spyOn(prototype, "insertTableRow")
  try {
    for (const reason of ["", " ", "x".repeat(4001), "Bad\u0000reason"]) await expect(authorDocxTableRowReview(bytes, input, reason)).rejects.toThrow("invalid_reason")
    expect(insert).not.toHaveBeenCalled()
  } finally { insert.mockRestore() }
})

test("failed later group member rolls back owned comment deletion and earlier content decision on final bytes", async () => {
  const bytes = source(), input = await request(bytes), result = await authorDocxTableRowReview(bytes, input, REASON), temp = (await docxodus()).openDocxSession(bytes, {})
  const prototype = Object.getPrototypeOf(temp) as typeof temp; temp.close()
  const remove = prototype.removeComment, reject = prototype.rejectRevision, batch = prototype.executeBatch, active = new WeakSet<object>()
  let removed = 0, calls = 0, restored = false
  const removal = spyOn(prototype, "removeComment").mockImplementation(function(this: typeof temp, ...args: Parameters<typeof remove>) {
    const outcome = remove.apply(this, args)
    if (outcome.success) { removed++; active.add(this) }
    return outcome
  })
  const failure = spyOn(prototype, "rejectRevision").mockImplementation(function(this: typeof temp, ...args: Parameters<typeof reject>) {
    if (active.has(this) && ++calls === 2) return { success: false, created: [], removed: [], modified: [], error: { code: "revision_not_found", message: "Synthetic later member failure" } }
    return reject.apply(this, args)
  })
  const observation = spyOn(prototype, "executeBatch").mockImplementation(function(this: typeof temp, ...args: Parameters<typeof batch>) {
    const before = this.getPackageManifest().rawPackageBytesDigest, version = this.getVersion(), outcome = batch.apply(this, args)
    if (active.has(this) && !outcome.success) {
      const after = this.getPackageManifest().rawPackageBytesDigest
      restored = before.algorithm === after.algorithm && before.value === after.value && version === this.getVersion()
    }
    return outcome
  })
  try {
    await expect(resolveDocxTableRowReview(bytes, result.bytes, input, REASON, "reject")).rejects.toThrow("review_resolution_failed_rolled_back")
    expect(removed).toBe(1); expect(calls).toBe(2); expect(restored).toBe(true)
  } finally { removal.mockRestore(); failure.mockRestore(); observation.mockRestore() }
})
